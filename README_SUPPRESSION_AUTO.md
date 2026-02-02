# Suppression Automatique des Réservations Expirées

## Vue d'ensemble

Cette fonctionnalité supprime automatiquement toutes les réservations expirées du Google Sheet après leur durée d'occupation de **2 heures**.

## Fonctionnalités

### 1. Suppression automatique des réservations expirées

Les réservations sont automatiquement supprimées du Google Sheet une fois que la date/heure de réservation + 2 heures est passée.

**Exemple :**
- Une réservation à 14h00 le 2 février 2026 sera automatiquement supprimée après 16h00 le même jour.

### 2. Trois méthodes d'activation

#### Méthode 1 : Trigger automatique (Recommandé)

Le trigger vérifie et supprime les réservations expirées **toutes les heures** automatiquement.

**Installation du trigger :**

1. Ouvrez votre Google Sheet contenant les réservations
2. Allez dans `Extensions > Apps Script`
3. Dans l'éditeur, sélectionnez la fonction `setupAutoDeletionTrigger` dans le menu déroulant
4. Cliquez sur "Exécuter" (▶️)
5. Autorisez les permissions si demandé

Le trigger est maintenant installé et s'exécutera automatiquement toutes les heures.

**Vérification du trigger :**
- Dans Apps Script, cliquez sur l'icône ⏰ (Déclencheurs) dans le menu de gauche
- Vous devriez voir un déclencheur pour la fonction `deleteExpiredReservations`

**Désinstallation du trigger :**
1. Dans Apps Script, sélectionnez la fonction `removeAutoDeletionTrigger`
2. Cliquez sur "Exécuter" (▶️)

#### Méthode 2 : Exécution manuelle via Apps Script

Pour nettoyer manuellement les réservations expirées :

1. Ouvrez `Extensions > Apps Script`
2. Sélectionnez la fonction `deleteExpiredReservations`
3. Cliquez sur "Exécuter" (▶️)
4. Vérifiez les logs pour voir combien de réservations ont été supprimées

#### Méthode 3 : Via l'API (pour les développeurs)

Trois nouvelles actions API sont disponibles pour les administrateurs :

**a) Nettoyer les réservations expirées :**
```
GET/POST: ?action=cleanupExpiredReservations&secret=VOTRE_CLE_SECRETE
```

**b) Installer le trigger automatique :**
```
GET/POST: ?action=setupAutoDeletion&secret=VOTRE_CLE_SECRETE
```

**c) Désinstaller le trigger automatique :**
```
GET/POST: ?action=removeAutoDeletion&secret=VOTRE_CLE_SECRETE
```

## Logs et Suivi

### Consultation des logs

Pour voir les réservations qui ont été supprimées :

1. Dans Apps Script, cliquez sur "Exécutions" dans le menu de gauche
2. Cliquez sur une exécution pour voir les détails
3. Les logs affichent :
   - Chaque réservation supprimée (ID, date, heure, nom, statut)
   - Le nombre total de réservations supprimées

**Exemple de log :**
```
Réservation expirée supprimée: {"id":"RES-123ABC","date":"2026-02-02","heure":"14:00","nom":"Dupont","statut":"confirmée"}
Nettoyage terminé: 3 réservation(s) supprimée(s)
```

## Sécurité

- ✅ Toutes les actions API sont protégées par la clé secrète administrateur
- ✅ Seules les réservations dont la durée complète (2h) est écoulée sont supprimées
- ✅ Les suppressions sont loggées pour traçabilité
- ✅ Le trigger ne peut être installé que manuellement (pas d'activation automatique)

## Durée d'occupation

La durée d'occupation est configurée dans `CONFIG.OCCUPATION_DURATION` :

```javascript
OCCUPATION_DURATION: 120, // durée en minutes (2 heures)
```

Si vous modifiez cette valeur, la suppression automatique s'adaptera automatiquement.

## Questions Fréquentes

**Q : Les réservations sont-elles supprimées immédiatement après les 2 heures ?**
R : Avec le trigger automatique, les réservations sont vérifiées toutes les heures. Une réservation expirée sera donc supprimée dans un délai maximal d'1 heure après son expiration.

**Q : Puis-je modifier la fréquence de nettoyage ?**
R : Oui, modifiez la ligne suivante dans `setupAutoDeletionTrigger()` :
```javascript
.everyHours(1)  // Changez 1 par 2, 3, etc. pour une fréquence différente
```

**Q : Les réservations refusées sont-elles aussi supprimées ?**
R : Oui, toutes les réservations expirées sont supprimées, quel que soit leur statut (en attente, confirmée, refusée).

**Q : Que se passe-t-il si j'installe plusieurs fois le trigger ?**
R : La fonction `setupAutoDeletionTrigger()` supprime automatiquement les anciens triggers avant d'en créer un nouveau pour éviter les doublons.

**Q : Comment savoir si le trigger est actif ?**
R : Allez dans Apps Script > Déclencheurs (icône ⏰). Si vous voyez `deleteExpiredReservations` dans la liste, le trigger est actif.

## Support

Pour toute question ou problème :
1. Vérifiez les logs d'exécution dans Apps Script
2. Assurez-vous que les permissions sont accordées
3. Vérifiez que la feuille "Reservations" existe dans le Google Sheet

## Version

Cette fonctionnalité a été ajoutée dans la **version 5** du système PIAF.
