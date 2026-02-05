/**
 * ============================================
 * PIAF - Système de réservation
 * Backend Google Apps Script v5
 * ============================================
 *
 * NOUVEAUTÉS v5:
 * - Suppression automatique des réservations expirées (après 2 heures)
 * - Trigger automatique pour nettoyage quotidien à minuit
 * - API de nettoyage manuel et gestion des triggers
 *
 * HISTORIQUE v4:
 * - Contrainte 2h AVANT et APRÈS une réservation existante
 * - Ajout manuel de réservation (admin)
 * - Suppression de réservation (admin)
 */

// ============================================
// CONFIGURATION
// ============================================

const CONFIG = {
  SCHEDULE: {
    0: { open: true, services: [{ start: '11:00', end: '15:30', slots: ['11:00', '11:30', '12:30', '13:00', '13:30'] }] },  // Dimanche
    1: { open: false },  // Lundi
    2: { open: false },  // Mardi
    3: { open: true, services: [{ start: '12:00', end: '15:00', slots: ['12:00', '12:30', '13:00'] }] },  // Mercredi
    4: { open: true, services: [{ start: '18:30', end: '22:00', slots: ['18:30', '19:00', '19:30', '20:00'] }] },  // Jeudi
    5: { open: true, services: [{ start: '18:30', end: '22:30', slots: ['18:30', '19:00', '19:30', '20:00', '20:30'] }] },  // Vendredi
    6: { open: true, services: [{ start: '12:00', end: '15:30', slots: ['12:00', '12:30', '13:00', '13:30'] }] }  // Samedi
  },
  TABLES: {
    list: [1, 2, 10, 11, 12, 13, 14, 15, 16, 17],
    capacity: { 1: 4, 2: 2, 10: 4, 11: 4, 12: 2, 13: 2, 14: 4, 15: 2, 16: 2, 17: 2 },
    maxGroupSize: 8,
    minGuests: 2
  },
  OCCUPATION_DURATION: 120, // durée d'occupation en minutes (2 heures)
  MIN_ADVANCE_HOURS: 2,
  MAX_ADVANCE_DAYS: 60,  // 60 jours à l'avance (environ 2 mois)
  SLOT_INTERVAL: 15
};

// ============================================
// INITIALISATION
// ============================================

function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let reservationsSheet = ss.getSheetByName('Reservations');
  if (!reservationsSheet) {
    reservationsSheet = ss.insertSheet('Reservations');
    reservationsSheet.getRange('A1:L1').setValues([[
      'ID', 'Date', 'Heure', 'Personnes', 'Nom', 'Telephone', 
      'Email', 'Langue', 'Statut', 'Tables', 'Cree_le', 'Commentaire'
    ]]);
    reservationsSheet.getRange('A1:L1').setFontWeight('bold');
    reservationsSheet.setFrozenRows(1);
    reservationsSheet.getRange('B:B').setNumberFormat('@');
  }
  
  let closuresSheet = ss.getSheetByName('Fermetures');
  if (!closuresSheet) {
    closuresSheet = ss.insertSheet('Fermetures');
    closuresSheet.getRange('A1:D1').setValues([['ID', 'Date', 'Heure_debut', 'Heure_fin']]);
    closuresSheet.getRange('A1:D1').setFontWeight('bold');
    closuresSheet.setFrozenRows(1);
    closuresSheet.getRange('B:B').setNumberFormat('@');
  }
  
  let configSheet = ss.getSheetByName('Configuration');
  if (!configSheet) {
    configSheet = ss.insertSheet('Configuration');
    configSheet.getRange('A1:B5').setValues([
      ['Paramètre', 'Valeur'],
      ['EMAIL_ENVOI', ''],
      ['NOM_RESTAURANT', 'Piaf'],
      ['TELEPHONE', '+32 2 500 12 34'],
      ['ADMIN_URL_SECRET', generateSecret()]
    ]);
    configSheet.getRange('A1:B1').setFontWeight('bold');
  }
  
  Logger.log('Sheet initialisé!');
  Logger.log('URL secrète admin: ' + getConfig('ADMIN_URL_SECRET'));
}

function generateSecret() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < 32; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Configuration');
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

// ============================================
// UTILITAIRES
// ============================================

function normalizeDate(dateValue) {
  if (!dateValue) return null;

  // Fuseau horaire du restaurant (Bruxelles)
  const TIMEZONE = 'Europe/Brussels';

  if (typeof dateValue === 'string') {
    // Si c'est déjà au format YYYY-MM-DD, le retourner tel quel
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateValue)) return dateValue;

    // Essayer de parser et reformater avec le bon fuseau horaire
    const parsed = new Date(dateValue);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, TIMEZONE, 'yyyy-MM-dd');
    }
    return dateValue;
  }

  // Si c'est un objet Date, utiliser Utilities.formatDate avec le fuseau horaire du restaurant
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, TIMEZONE, 'yyyy-MM-dd');
  }

  // Dernière tentative
  try {
    const date = new Date(dateValue);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
    }
  } catch (e) {}

  return String(dateValue);
}

function formatDateToYYYYMMDD(date) {
  // Utiliser le fuseau horaire du restaurant pour la cohérence
  return Utilities.formatDate(date, 'Europe/Brussels', 'yyyy-MM-dd');
}

function timeToMinutes(time) {
  if (!time) return 0;
  const [h, m] = String(time).split(':').map(Number);
  return h * 60 + m;
}

function minutesToTime(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
}

function isTimeInRange(time, start, end) {
  const t = timeToMinutes(time);
  return t >= timeToMinutes(start) && t < timeToMinutes(end);
}

// ============================================
// POINT D'ENTRÉE WEB
// ============================================

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = e.parameter || {};
  const action = params.action;
  
  Logger.log('Action: ' + action);
  
  let result;
  
  try {
    switch (action) {
      case 'getAvailability':
        result = getAvailability(params.date, parseInt(params.guests));
        break;
        
      case 'createReservation':
        let postData = null;
        if (e.postData && e.postData.contents) {
          try { postData = JSON.parse(e.postData.contents); } 
          catch (err) { result = { success: false, error: 'JSON invalide' }; break; }
        } else if (params.date && params.time) {
          postData = { date: params.date, time: params.time, guests: params.guests, 
                       name: params.name, phone: params.phone, email: params.email, lang: params.lang };
        }
        if (!postData) { result = { success: false, error: 'Aucune donnée' }; break; }
        result = createReservation(postData);
        break;
        
      case 'test':
        result = { success: true, message: 'API OK', version: 'v5.1-debug', timestamp: new Date().toISOString() };
        break;
      
      case 'getReservations':
        if (!checkAdminAccess(params.secret)) { result = { success: false, error: 'Accès refusé' }; }
        else { result = getReservations(params.date, params.status); }
        break;
        
      case 'getAllReservations':
        if (!checkAdminAccess(params.secret)) { result = { success: false, error: 'Accès refusé' }; }
        else { result = getReservations(null, null); }
        break;
        
      case 'confirmReservation':
        if (!checkAdminAccess(params.secret)) { 
          result = { success: false, error: 'Accès refusé' }; 
        } else {
          const tables = params.tables || null;
          result = updateReservationStatus(params.id, 'confirmée', '', tables);
        }
        break;
        
      case 'refuseReservation':
        if (!checkAdminAccess(params.secret)) { result = { success: false, error: 'Accès refusé' }; }
        else {
          let refuseData = {};
          try { refuseData = JSON.parse(e.postData.contents); } catch (err) {}
          result = updateReservationStatus(params.id, 'refusée', refuseData.message || '');
        }
        break;
      
      // NOUVEAU: Suppression de réservation
      case 'deleteReservation':
        if (!checkAdminAccess(params.secret)) { 
          result = { success: false, error: 'Accès refusé' }; 
        } else {
          result = deleteReservation(params.id);
        }
        break;
      
      // NOUVEAU: Création manuelle de réservation (admin)
      case 'createManualReservation':
        if (!checkAdminAccess(params.secret)) { 
          result = { success: false, error: 'Accès refusé' }; 
        } else {
          let manualData = null;
          if (e.postData && e.postData.contents) {
            try { manualData = JSON.parse(e.postData.contents); } 
            catch (err) { result = { success: false, error: 'JSON invalide' }; break; }
          }
          if (!manualData) { result = { success: false, error: 'Aucune donnée' }; break; }
          result = createManualReservation(manualData);
        }
        break;
        
      case 'addClosure':
        if (!checkAdminAccess(params.secret)) { result = { success: false, error: 'Accès refusé' }; }
        else {
          let closureData;
          try { closureData = JSON.parse(e.postData.contents); } 
          catch (err) { result = { success: false, error: 'Données invalides' }; break; }
          result = addClosure(closureData);
        }
        break;
        
      case 'removeClosure':
        if (!checkAdminAccess(params.secret)) { result = { success: false, error: 'Accès refusé' }; }
        else { result = removeClosure(params.id); }
        break;
        
      case 'getClosures':
        if (!checkAdminAccess(params.secret)) { result = { success: false, error: 'Accès refusé' }; }
        else { result = getClosures(); }
        break;
        
      case 'getDashboardData':
        if (!checkAdminAccess(params.secret)) { result = { success: false, error: 'Accès refusé' }; }
        else { result = getDashboardData(params.date); }
        break;

      // NOUVEAU: Nettoyage manuel des réservations expirées
      case 'cleanupExpiredReservations':
        if (!checkAdminAccess(params.secret)) {
          result = { success: false, error: 'Accès refusé' };
        } else {
          result = deleteExpiredReservations();
        }
        break;

      // NOUVEAU: Installer le trigger de suppression automatique
      case 'setupAutoDeletion':
        if (!checkAdminAccess(params.secret)) {
          result = { success: false, error: 'Accès refusé' };
        } else {
          result = setupAutoDeletionTrigger();
        }
        break;

      // NOUVEAU: Désinstaller le trigger de suppression automatique
      case 'removeAutoDeletion':
        if (!checkAdminAccess(params.secret)) {
          result = { success: false, error: 'Accès refusé' };
        } else {
          result = removeAutoDeletionTrigger();
        }
        break;

      default:
        result = { success: false, error: 'Action inconnue: ' + action };
    }
  } catch (error) {
    Logger.log('ERREUR: ' + error.toString());
    result = { success: false, error: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function checkAdminAccess(secret) {
  return secret === getConfig('ADMIN_URL_SECRET');
}

// ============================================
// DISPONIBILITÉS (contrainte 2h avant et après)
// ============================================

function getAvailability(dateStr, guests) {
  const date = new Date(dateStr);
  const dayOfWeek = date.getDay();
  const schedule = CONFIG.SCHEDULE[dayOfWeek];
  
  if (!schedule.open) return { success: true, available: false, reason: 'closed' };
  
  const closures = getClosuresForDate(dateStr);
  const reservations = getReservationsForDate(dateStr);
  const slots = [];
  
  schedule.services.forEach(service => {
    generateTimeSlots(service.start, service.end, service.slots).forEach(slotTime => {
      const isClosed = closures.some(c => isTimeInRange(slotTime, c.start, c.end));
      
      if (!isClosed) {
        const availableTables = getAvailableTablesForSlot(dateStr, slotTime, reservations, guests);
        slots.push({
          time: slotTime,
          available: availableTables.canAccommodate,
          tables: availableTables
        });
      }
    });
  });
  
  return { success: true, date: dateStr, slots };
}

function generateTimeSlots(startTime, endTime, explicitSlots) {
  // Si des créneaux explicites sont définis, les utiliser directement
  if (explicitSlots && explicitSlots.length > 0) {
    return explicitSlots;
  }

  // Sinon, générer automatiquement les créneaux
  const slots = [];
  let current = timeToMinutes(startTime);
  const end = timeToMinutes(endTime) - CONFIG.OCCUPATION_DURATION;

  while (current <= end) {
    const h = Math.floor(current / 60);
    const m = current % 60;
    slots.push(`${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`);
    current += CONFIG.SLOT_INTERVAL;
  }

  return slots;
}

/**
 * Vérifie la disponibilité des tables pour un créneau
 * Prend en compte:
 * - Les chevauchements classiques (une réservation en cours)
 * - La contrainte "2h avant et après" : si une table est réservée à 14h00,
 *   on ne peut pas la réserver entre 12h00 et 16h00 (2h avant à 2h après)
 * - Inclut les réservations "en attente" et "confirmées" (exclut "refusée")
 */
function getAvailableTablesForSlot(dateStr, slotTime, reservations, guests) {
  const occupiedTables = new Set();
  const slotMinutes = timeToMinutes(slotTime);
  const slotEnd = slotMinutes + CONFIG.OCCUPATION_DURATION;
  
  reservations.forEach(res => {
    if (res.statut === 'refusée') return;
    
    const resStart = timeToMinutes(res.heure);
    const resEnd = resStart + CONFIG.OCCUPATION_DURATION;
    
    // Zone protégée : 2h avant et 2h après l'heure de réservation
    // Exemple: réservation à 14h00 → zone protégée de 12h00 à 16h00
    // protectedZoneStart = 14h00 - 2h = 12h00
    // protectedZoneEnd = 14h00 + 2h = 16h00
    const protectedZoneStart = resStart - CONFIG.OCCUPATION_DURATION;
    const protectedZoneEnd = resEnd;
    
    // Le créneau demandé occupe de slotMinutes à slotEnd
    // Il y a conflit si ces deux plages se chevauchent
    const hasConflict = (slotMinutes < protectedZoneEnd && slotEnd > protectedZoneStart);
    
    if (hasConflict && res.tables) {
      String(res.tables).split(',').forEach(t => occupiedTables.add(parseInt(t.trim())));
    }
  });
  
  const availableTables = [];
  CONFIG.TABLES.list.forEach(tableNum => {
    if (!occupiedTables.has(tableNum)) availableTables.push(tableNum);
  });

  // Calculer la capacité totale disponible
  const totalCapacity = availableTables.reduce((sum, t) => sum + CONFIG.TABLES.capacity[t], 0);

  // Trouver la meilleure combinaison de tables pour le groupe
  const canAccommodate = findTablesForGuests(availableTables, guests) !== null;

  return {
    tables: availableTables,
    canAccommodate: canAccommodate,
    totalCapacity: totalCapacity,
    canGroup: guests > 4 && totalCapacity >= guests
  };
}

/**
 * Trouve la meilleure combinaison de tables pour un nombre de convives
 * Retourne un tableau de numéros de tables ou null si impossible
 */
function findTablesForGuests(availableTables, guests) {
  // Trier les tables par capacité décroissante pour optimiser
  const sorted = [...availableTables].sort((a, b) =>
    CONFIG.TABLES.capacity[b] - CONFIG.TABLES.capacity[a]
  );

  // Essayer de trouver une seule table suffisante
  for (const table of sorted) {
    if (CONFIG.TABLES.capacity[table] >= guests) {
      return [table];
    }
  }

  // Sinon, chercher une combinaison de tables
  let bestCombo = null;
  let minWaste = Infinity;

  // Essayer toutes les combinaisons de 2 tables
  for (let i = 0; i < sorted.length; i++) {
    for (let j = i + 1; j < sorted.length; j++) {
      const capacity = CONFIG.TABLES.capacity[sorted[i]] + CONFIG.TABLES.capacity[sorted[j]];
      if (capacity >= guests) {
        const waste = capacity - guests;
        if (waste < minWaste) {
          minWaste = waste;
          bestCombo = [sorted[i], sorted[j]];
        }
      }
    }
  }

  return bestCombo;
}

// ============================================
// RÉSERVATIONS
// ============================================

function createReservation(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reservations');
  
  if (!sheet) return { success: false, error: 'Configuration incorrecte' };
  if (!data.date || !data.time || !data.guests || !data.name || !data.phone || !data.email) {
    return { success: false, error: 'Données manquantes' };
  }
  
  const guests = parseInt(data.guests);
  if (guests < CONFIG.TABLES.minGuests || guests > CONFIG.TABLES.maxGroupSize) {
    return { success: false, error: 'Nombre de personnes invalide' };
  }
  
  const normalizedDate = normalizeDate(data.date);
  const existingReservations = getReservationsForDate(normalizedDate);
  const availability = getAvailableTablesForSlot(normalizedDate, data.time, existingReservations, guests);
  
  if (!availability.canAccommodate) {
    return { success: false, error: 'Plus de tables disponibles' };
  }

  // Attribution automatique (sera modifiable lors de la confirmation)
  const suggestedTables = findTablesForGuests(availability.tables, guests);
  const assignedTables = suggestedTables ? suggestedTables.join(', ') : '';
  
  const id = 'RES-' + Date.now().toString(36).toUpperCase();
  
  sheet.appendRow([
    id, normalizedDate, data.time, guests, data.name, data.phone,
    data.email, data.lang || 'fr', 'en attente', assignedTables,
    new Date().toISOString(), ''
  ]);
  
  return { success: true, id, message: 'Réservation créée', tables: assignedTables };
}

/**
 * NOUVEAU: Création manuelle de réservation par l'admin
 * Permet de créer une réservation déjà confirmée avec les tables choisies
 */
function createManualReservation(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reservations');
  
  if (!sheet) return { success: false, error: 'Configuration incorrecte' };
  
  // Validation des champs requis
  if (!data.date || !data.time || !data.guests || !data.name) {
    return { success: false, error: 'Données manquantes (date, heure, personnes, nom requis)' };
  }
  
  const guests = parseInt(data.guests);
  if (guests < 1 || guests > 32) {
    return { success: false, error: 'Nombre de personnes invalide' };
  }
  
  const normalizedDate = normalizeDate(data.date);
  
  // Vérification optionnelle des conflits de tables si des tables sont spécifiées
  if (data.tables) {
    const existingReservations = getReservationsForDate(normalizedDate);
    const requestedTables = String(data.tables).split(',').map(t => parseInt(t.trim()));
    
    // Vérifier les conflits pour chaque table demandée
    const slotMinutes = timeToMinutes(data.time);
    const slotEnd = slotMinutes + CONFIG.OCCUPATION_DURATION;
    
    for (const res of existingReservations) {
      if (res.statut === 'refusée' || !res.tables) continue;
      
      const resStart = timeToMinutes(res.heure);
      const resEnd = resStart + CONFIG.OCCUPATION_DURATION;
      const protectedZoneStart = resStart - CONFIG.OCCUPATION_DURATION;
      
      const hasConflict = (slotMinutes < resEnd && slotEnd > protectedZoneStart);
      
      if (hasConflict) {
        const resTables = String(res.tables).split(',').map(t => parseInt(t.trim()));
        const conflict = requestedTables.filter(t => resTables.includes(t));
        
        if (conflict.length > 0) {
          return { 
            success: false, 
            error: `Conflit: Table(s) ${conflict.join(', ')} déjà réservée(s) à ${res.heure} par ${res.nom}`
          };
        }
      }
    }
  }
  
  const id = 'MAN-' + Date.now().toString(36).toUpperCase();
  
  // Le statut peut être spécifié (défaut: confirmée pour les réservations manuelles)
  const status = data.status || 'confirmée';
  
  sheet.appendRow([
    id, 
    normalizedDate, 
    data.time, 
    guests, 
    data.name, 
    data.phone || '', 
    data.email || '', 
    data.lang || 'fr', 
    status, 
    data.tables || '',
    new Date().toISOString(), 
    data.comment || 'Réservation manuelle'
  ]);
  
  return { 
    success: true, 
    id, 
    message: 'Réservation manuelle créée', 
    tables: data.tables || ''
  };
}

/**
 * NOUVEAU: Suppression d'une réservation
 */
function deleteReservation(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reservations');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      // Sauvegarder les infos pour le log
      const reservationInfo = {
        id: data[i][0],
        date: normalizeDate(data[i][1]),
        heure: data[i][2],
        nom: data[i][4],
        email: data[i][6]
      };
      
      // Supprimer la ligne
      sheet.deleteRow(i + 1);
      
      Logger.log('Réservation supprimée: ' + JSON.stringify(reservationInfo));
      
      return { 
        success: true, 
        message: 'Réservation supprimée',
        deleted: reservationInfo
      };
    }
  }
  
  return { success: false, error: 'Réservation non trouvée' };
}

/**
 * NOUVEAU: Suppression automatique des réservations expirées
 * Supprime toutes les réservations dont la date/heure + 2 heures est passée
 */
function deleteExpiredReservations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reservations');

  if (!sheet) {
    Logger.log('Erreur: Feuille Reservations non trouvée');
    return { success: false, error: 'Feuille Reservations non trouvée' };
  }

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const deletedReservations = [];

  // Parcourir de la fin vers le début pour éviter les décalages d'index lors de la suppression
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];

    // Vérifier que la ligne contient des données
    if (!row[0]) continue;

    const reservationDate = normalizeDate(row[1]);
    const reservationTime = row[2];

    if (!reservationDate || !reservationTime) continue;

    // Construire la date/heure de la réservation
    const [year, month, day] = reservationDate.split('-').map(Number);
    const [hours, minutes] = String(reservationTime).split(':').map(Number);
    const reservationDateTime = new Date(year, month - 1, day, hours, minutes);

    // Ajouter la durée d'occupation (2 heures)
    const expirationDateTime = new Date(reservationDateTime.getTime() + CONFIG.OCCUPATION_DURATION * 60 * 1000);

    // Si la réservation est expirée
    if (expirationDateTime < now) {
      const reservationInfo = {
        id: row[0],
        date: reservationDate,
        heure: reservationTime,
        nom: row[4],
        statut: row[8]
      };

      deletedReservations.push(reservationInfo);

      // Supprimer la ligne
      sheet.deleteRow(i + 1);

      Logger.log(`Réservation expirée supprimée: ${JSON.stringify(reservationInfo)}`);
    }
  }

  Logger.log(`Nettoyage terminé: ${deletedReservations.length} réservation(s) supprimée(s)`);

  return {
    success: true,
    message: `${deletedReservations.length} réservation(s) expirée(s) supprimée(s)`,
    deleted: deletedReservations
  };
}

/**
 * NOUVEAU: Configuration du déclencheur automatique pour la suppression des réservations expirées
 * À exécuter manuellement une fois pour installer le trigger
 */
function setupAutoDeletionTrigger() {
  // Supprimer les anciens triggers du même type pour éviter les doublons
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'deleteExpiredReservations') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Créer un nouveau trigger qui s'exécute tous les jours à minuit
  ScriptApp.newTrigger('deleteExpiredReservations')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();

  Logger.log('Trigger de nettoyage automatique installé (exécution quotidienne à minuit)');
  return { success: true, message: 'Trigger installé avec succès (exécution quotidienne à minuit)' };
}

/**
 * NOUVEAU: Suppression du déclencheur automatique
 * Pour désactiver la suppression automatique si nécessaire
 */
function removeAutoDeletionTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'deleteExpiredReservations') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  Logger.log(`${removed} trigger(s) de nettoyage supprimé(s)`);
  return { success: true, message: `${removed} trigger(s) supprimé(s)` };
}

function getReservations(date, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reservations');
  const data = sheet.getDataRange().getValues();

  const reservations = [];
  const normalizedFilterDate = date ? normalizeDate(date) : null;

  Logger.log('getReservations - Filtre date: ' + date + ' -> normalisé: ' + normalizedFilterDate);
  Logger.log('getReservations - Nombre de lignes dans le sheet: ' + (data.length - 1));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const rawDate = row[1];
    const rowDate = normalizeDate(row[1]);

    Logger.log('getReservations - Ligne ' + i + ': ID=' + row[0] + ', Date brute=' + rawDate + ' (type: ' + typeof rawDate + '), Date normalisée=' + rowDate);

    if (normalizedFilterDate && rowDate !== normalizedFilterDate) {
      Logger.log('getReservations - Ligne ' + i + ' ignorée: date ne correspond pas (' + rowDate + ' != ' + normalizedFilterDate + ')');
      continue;
    }
    if (status && row[8] !== status) continue;

    reservations.push({
      id: row[0], date: rowDate, heure: row[2], personnes: row[3],
      nom: row[4], telephone: row[5], email: row[6], langue: row[7],
      statut: row[8], tables: row[9], creeLe: row[10], commentaire: row[11]
    });
  }

  Logger.log('getReservations - Réservations retournées: ' + reservations.length);

  return { success: true, reservations };
}

function getReservationsForDate(dateStr) {
  return getReservations(dateStr, null).reservations || [];
}

/**
 * Met à jour le statut d'une réservation
 */
function updateReservationStatus(id, newStatus, comment = '', tables = null) {
  Logger.log('updateReservationStatus appelée - id: ' + id + ', newStatus: ' + newStatus + ', tables: ' + tables);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reservations');
  const data = sheet.getDataRange().getValues();

  // Convertir l'ID en string pour comparaison cohérente
  const idStr = String(id);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === idStr) {
      Logger.log('Réservation trouvée à la ligne ' + (i+1));
      // Mettre à jour le statut
      sheet.getRange(i + 1, 9).setValue(newStatus);
      
      // Si des tables sont spécifiées, les mettre à jour
      if (tables) {
        sheet.getRange(i + 1, 10).setValue(tables);
        Logger.log('Tables mises à jour pour ' + id + ': ' + tables);
      }
      
      // Commentaire si présent
      if (comment) {
        sheet.getRange(i + 1, 12).setValue(comment);
      }
      
      // Préparer les données pour l'email
      const reservation = {
        id: data[i][0],
        date: normalizeDate(data[i][1]),
        heure: data[i][2],
        personnes: data[i][3],
        nom: data[i][4],
        email: data[i][6],
        langue: data[i][7],
        tables: tables || data[i][9]
      };
      
      // Envoyer l'email approprié
      Logger.log('Données réservation pour email: ' + JSON.stringify(reservation));
      Logger.log('newStatus: "' + newStatus + '" - comparaison avec "confirmée": ' + (newStatus === 'confirmée'));

      if (newStatus === 'confirmée') {
        Logger.log('>>> Appel de sendConfirmationEmail');
        sendConfirmationEmail(reservation);
        Logger.log('>>> Retour de sendConfirmationEmail');
      } else if (newStatus === 'refusée') {
        sendRefusalEmail(reservation, comment);
      }
      
      return { 
        success: true, 
        message: 'Statut mis à jour',
        tables: tables || data[i][9]
      };
    }
  }
  
  return { success: false, error: 'Réservation non trouvée' };
}

// ============================================
// FERMETURES
// ============================================

function addClosure(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Fermetures');
  const id = 'CLO-' + Date.now().toString(36).toUpperCase();
  
  sheet.appendRow([id, normalizeDate(data.date), data.startTime, data.endTime]);
  return { success: true, id };
}

function removeClosure(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Fermetures');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Fermeture non trouvée' };
}

function getClosures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Fermetures');
  const data = sheet.getDataRange().getValues();
  
  const closures = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    closures.push({ id: data[i][0], date: normalizeDate(data[i][1]), start: data[i][2], end: data[i][3] });
  }
  return { success: true, closures };
}

function getClosuresForDate(dateStr) {
  return getClosures().closures.filter(c => c.date === normalizeDate(dateStr));
}

// ============================================
// DASHBOARD
// ============================================

function getDashboardData(dateStr) {
  const date = dateStr || formatDateToYYYYMMDD(new Date());

  Logger.log('getDashboardData - Date demandée: ' + dateStr + ' -> normalisée: ' + date);

  const reservations = getReservationsForDate(date);

  Logger.log('getDashboardData - Réservations trouvées: ' + reservations.length);

  const pending = reservations.filter(r => r.statut === 'en attente').length;
  const confirmed = reservations.filter(r => r.statut === 'confirmée').length;
  const totalGuests = reservations
    .filter(r => r.statut === 'confirmée')
    .reduce((sum, r) => sum + (parseInt(r.personnes) || 0), 0);

  return {
    success: true,
    date,
    stats: { pending, confirmed, totalGuests },
    reservations,
    closures: getClosuresForDate(date)
  };
}

// ============================================
// EMAILS
// ============================================

function sendConfirmationEmail(reservation) {
  Logger.log('sendConfirmationEmail appelée avec: ' + JSON.stringify(reservation));

  // Ne pas envoyer d'email si pas d'adresse client
  if (!reservation.email) {
    Logger.log('Pas d\'email client - abandon');
    return;
  }

  const restaurantName = getConfig('NOM_RESTAURANT');
  const restaurantPhone = getConfig('TELEPHONE');
  Logger.log('Restaurant: ' + restaurantName + ', envoi à: ' + reservation.email);

  const templates = {
    fr: {
      subject: `Confirmation de votre réservation - ${restaurantName}`,
      body: `Bonjour ${reservation.nom},

Nous avons le plaisir de vous confirmer votre réservation :

📅 Date : ${formatDateDisplay(reservation.date, 'fr')}
🕐 Heure : ${reservation.heure}
👥 Nombre de personnes : ${reservation.personnes}

Pour toute modification ou annulation, merci de nous contacter au ${restaurantPhone}.

À très bientôt !
L'équipe ${restaurantName}`
    },
    en: {
      subject: `Reservation confirmed - ${restaurantName}`,
      body: `Hello ${reservation.nom},

We are pleased to confirm your reservation:

📅 Date: ${formatDateDisplay(reservation.date, 'en')}
🕐 Time: ${reservation.heure}
👥 Guests: ${reservation.personnes}

For changes or cancellations, contact us at ${restaurantPhone}.

See you soon!
The ${restaurantName} team`
    },
    nl: {
      subject: `Reservering bevestigd - ${restaurantName}`,
      body: `Hallo ${reservation.nom},

Wij bevestigen uw reservering:

📅 Datum: ${formatDateDisplay(reservation.date, 'nl')}
🕐 Tijd: ${reservation.heure}
👥 Personen: ${reservation.personnes}

Voor wijzigingen, contacteer ons via ${restaurantPhone}.

Tot snel!
Het ${restaurantName} team`
    }
  };
  
  const template = templates[reservation.langue] || templates.fr;
  
  try {
    Logger.log('Tentative envoi email à ' + reservation.email);
    Logger.log('Sujet: ' + template.subject);
    MailApp.sendEmail(reservation.email, template.subject, template.body);
    Logger.log('Email envoyé avec succès à ' + reservation.email);
  } catch (e) {
    Logger.log('ERREUR email confirmation: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
  }
}

function sendRefusalEmail(reservation, reason) {
  const emailSender = getConfig('EMAIL_ENVOI');
  if (!emailSender) return;
  
  // Ne pas envoyer d'email si pas d'adresse
  if (!reservation.email) return;
  
  const restaurantName = getConfig('NOM_RESTAURANT');
  const restaurantPhone = getConfig('TELEPHONE');
  
  const templates = {
    fr: {
      subject: `Concernant votre réservation - ${restaurantName}`,
      body: `Bonjour ${reservation.nom},

Nous avons reçu votre demande pour le ${formatDateDisplay(reservation.date, 'fr')} à ${reservation.heure}.

Malheureusement, nous ne pouvons pas confirmer cette réservation.

${reason ? `Message :\n${reason}\n` : ''}
Contactez-nous au ${restaurantPhone} pour une alternative.

Cordialement,
L'équipe ${restaurantName}`
    },
    en: {
      subject: `Regarding your reservation - ${restaurantName}`,
      body: `Hello ${reservation.nom},

We received your request for ${formatDateDisplay(reservation.date, 'en')} at ${reservation.heure}.

Unfortunately, we cannot confirm this reservation.

${reason ? `Message:\n${reason}\n` : ''}
Contact us at ${restaurantPhone} for alternatives.

Best regards,
The ${restaurantName} team`
    },
    nl: {
      subject: `Betreft uw reservering - ${restaurantName}`,
      body: `Hallo ${reservation.nom},

Wij ontvingen uw aanvraag voor ${formatDateDisplay(reservation.date, 'nl')} om ${reservation.heure}.

Helaas kunnen wij deze reservering niet bevestigen.

${reason ? `Bericht:\n${reason}\n` : ''}
Contacteer ons via ${restaurantPhone} voor alternatieven.

Met vriendelijke groet,
Het ${restaurantName} team`
    }
  };
  
  const template = templates[reservation.langue] || templates.fr;
  
  try {
    MailApp.sendEmail({ to: reservation.email, subject: template.subject, body: template.body });
  } catch (e) {
    Logger.log('Erreur email: ' + e);
  }
}

function formatDateDisplay(dateStr, lang) {
  const date = new Date(dateStr);
  const locales = { fr: 'fr-FR', en: 'en-GB', nl: 'nl-NL' };
  return date.toLocaleDateString(locales[lang] || 'fr-FR', { weekday: 'long', day: 'numeric', month: 'long' });
}
