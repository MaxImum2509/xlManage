# Liste des Fonctions Utilitaires pour modUtils.bas

## Introduction

Ce document liste l'ensemble des fonctions utilitaires à intégrer dans `modUtils.bas` pour le projet tbAffaires. Ces fonctions sont essentielles pour la gestion des erreurs, la manipulation des fichiers, et l'interaction avec les objets Excel. Elles sont utilisées par tous les autres modules VBA de l'application.

## Fonctions de Gestion des Erreurs

### 1. AfficherMessageErreur

- **Description** : Affiche un message d'erreur standardisé avec le code d'erreur, le message explicite, et l'action suggérée pour l'utilisateur.
- **Paramètres** :
    - `codeErreur` (String) : Code d'erreur (ex: "ERR-001")
    - `message` (String) : Message explicite de l'erreur
    - `actionUtilisateur` (String) : Action suggérée pour l'utilisateur
- **Retour** : Aucun
- **Exemple d'utilisation** :
    ```vba
    AfficherMessageErreur "ERR-001", "Utilisateur non configuré", "Contacter Patrick"
    ```

### 2. AfficherMessageInfo

- **Description** : Affiche un message informatif et enregistre l'action dans les logs.
- **Paramètres** :
    - `message` (String) : Message informatif à afficher
- **Retour** : Aucun
- **Exemple d'utilisation** :
    ```vba
    AfficherMessageInfo "Consolidation réussie - 50 affaires"
    ```

### 3. EnregistrerLog

- **Description** : Enregistre une action dans le fichier de log `tbAffaires.log`.
- **Paramètres** :
    - `niveau` (String) : Niveau de log (INFO, ERREUR, SUCCES)
    - `action` (String) : Action effectuée
    - `resultat` (String) : Résultat de l'action
- **Retour** : Aucun
- **Exemple d'utilisation** :
    ```vba
    EnregistrerLog "SUCCES", "Consolidation", "50 affaires consolidées (0.8 sec)"
    ```

## Fonctions de Vérification de Fichiers

### 4. FichierExiste

- **Description** : Vérifie si un fichier existe à l'emplacement spécifié.
- **Paramètres** :
    - `cheminFichier` (String) : Chemin complet du fichier
- **Retour** : Boolean (True si le fichier existe, False sinon)
- **Exemple d'utilisation** :
    ```vba
    If FichierExiste("C:\data\data.xlsx") Then
        ' Fichier existe
    End If
    ```

### 5. RepertoireExiste

- **Description** : Vérifie si un répertoire existe à l'emplacement spécifié.
- **Paramètres** :
    - `cheminRepertoire` (String) : Chemin complet du répertoire
- **Retour** : Boolean (True si le répertoire existe, False sinon)
- **Exemple d'utilisation** :
    ```vba
    If RepertoireExiste("C:\data") Then
        ' Répertoire existe
    End If
    ```

## Fonctions de Chargement d'Objets Excel

### 6. ChargerWorkbook

- **Description** : Charge un workbook Excel et retourne l'objet Workbook.
- **Paramètres** :
    - `cheminFichier` (String) : Chemin complet du fichier Excel
    - `lectureSeule` (Boolean, optionnel) : Si True, ouvre le fichier en lecture seule
- **Retour** : Workbook (objet Workbook chargé ou Nothing en cas d'erreur)
- **Exemple d'utilisation** :
    ```vba
    Dim wb As Workbook
    Set wb = ChargerWorkbook("C:\data\data.xlsx", True)
    ```

### 7. ChargerWorksheet

- **Description** : Charge une worksheet Excel et retourne l'objet Worksheet.
- **Paramètres** :
    - `wb` (Workbook) : Objet Workbook contenant la worksheet
    - `nomFeuille` (String) : Nom de la worksheet à charger
- **Retour** : Worksheet (objet Worksheet chargé ou Nothing en cas d'erreur)
- **Exemple d'utilisation** :
    ```vba
    Dim ws As Worksheet
    Set ws = ChargerWorksheet(wb, "ADV")
    ```

### 8. ChargerListObject

- **Description** : Charge un ListObject Excel et retourne l'objet ListObject.
- **Paramètres** :
    - `ws` (Worksheet) : Objet Worksheet contenant le ListObject
    - `nomListObject` (String) : Nom du ListObject à charger
- **Retour** : ListObject (objet ListObject chargé ou Nothing en cas d'erreur)
- **Exemple d'utilisation** :
    ```vba
    Dim lo As ListObject
    Set lo = ChargerListObject(ws, "tbADV")
    ```

## Fonctions de Gestion des Chemins

### 9. ObtenirCheminComplet

- **Description** : Retourne le chemin complet d'un fichier ou répertoire en combinant le chemin de base et le chemin relatif.
- **Paramètres** :
    - `cheminBase` (String) : Chemin de base
    - `cheminRelatif` (String) : Chemin relatif à ajouter
- **Retour** : String (chemin complet)
- **Exemple d'utilisation** :
    ```vba
    Dim chemin As String
    chemin = ObtenirCheminComplet("C:\data", "data.xlsx")
    ```

### 10. ObtenirNomFichier

- **Description** : Extrait le nom de fichier à partir d'un chemin complet.
- **Paramètres** :
    - `cheminFichier` (String) : Chemin complet du fichier
- **Retour** : String (nom du fichier)
- **Exemple d'utilisation** :
    ```vba
    Dim nomFichier As String
    nomFichier = ObtenirNomFichier("C:\data\data.xlsx")
    ```

## Fonctions de Gestion des Dates

### 11. FormaterDatePourLog

- **Description** : Formate une date pour l'enregistrement dans les logs.
- **Paramètres** :
    - `dateHeure` (Date) : Date et heure à formater
- **Retour** : String (date formatée au format "YYYY-MM-DD HH:MM:SS")
- **Exemple d'utilisation** :
    ```vba
    Dim dateFormatee As String
    dateFormatee = FormaterDatePourLog(Now())
    ```

### 12. ObtenirDateHeureActuelle

- **Description** : Retourne la date et l'heure actuelles.
- **Paramètres** : Aucun
- **Retour** : Date (date et heure actuelles)
- **Exemple d'utilisation** :
    ```vba
    Dim dateActuelle As Date
    dateActuelle = ObtenirDateHeureActuelle()
    ```

## Fonctions de Gestion des Chaînes de Caractères

### 13. EstVide

- **Description** : Vérifie si une chaîne de caractères est vide ou Null.
- **Paramètres** :
    - `chaine` (String) : Chaîne de caractères à vérifier
- **Retour** : Boolean (True si la chaîne est vide ou Null, False sinon)
- **Exemple d'utilisation** :
    ```vba
    If EstVide(texte) Then
        ' Chaîne est vide
    End If
    ```

### 14. RemplacerCaracteresSpeciaux

- **Description** : Remplace les caractères spéciaux dans une chaîne de caractères pour éviter les problèmes de formatage.
- **Paramètres** :
    - `chaine` (String) : Chaîne de caractères à traiter
- **Retour** : String (chaîne de caractères traitée)
- **Exemple d'utilisation** :
    ```vba
    Dim chaineTraitee As String
    chaineTraitee = RemplacerCaracteresSpeciaux("Texte avec caractères spéciaux")
    ```

## Fonctions de Gestion des Tableaux

### 15. TableauEstVide

- **Description** : Vérifie si un tableau est vide ou non initialisé.
- **Paramètres** :
    - `tableau` (Variant) : Tableau à vérifier
- **Retour** : Boolean (True si le tableau est vide ou non initialisé, False sinon)
- **Exemple d'utilisation** :
    ```vba
    If TableauEstVide(monTableau) Then
        ' Tableau est vide
    End If
    ```

### 16. ObtenirTailleTableau

- **Description** : Retourne la taille d'un tableau.
- **Paramètres** :
    - `tableau` (Variant) : Tableau à mesurer
- **Retour** : Long (taille du tableau)
- **Exemple d'utilisation** :
    ```vba
    Dim taille As Long
    taille = ObtenirTailleTableau(monTableau)
    ```

## Fonctions de Gestion des Erreurs VBA

### 17. GererErreur

- **Description** : Gère une erreur VBA et affiche un message d'erreur standardisé.
- **Paramètres** :
    - `codeErreur` (String) : Code d'erreur (ex: "ERR-001")
    - `message` (String) : Message explicite de l'erreur
    - `actionUtilisateur` (String) : Action suggérée pour l'utilisateur
- **Retour** : Aucun
- **Exemple d'utilisation** :

    ```vba
    On Error GoTo ErreurHandler
    ' Code susceptible de générer une erreur
    Exit Sub

    ErreurHandler:
    GererErreur "ERR-001", "Erreur lors du chargement", "Contacter Patrick"
    ```

### 18. EnregistrerErreurDansLog

- **Description** : Enregistre une erreur dans le fichier de log.
- **Paramètres** :
    - `codeErreur` (String) : Code d'erreur (ex: "ERR-001")
    - `message` (String) : Message explicite de l'erreur
    - `actionUtilisateur` (String) : Action suggérée pour l'utilisateur
- **Retour** : Aucun
- **Exemple d'utilisation** :
    ```vba
    EnregistrerErreurDansLog "ERR-001", "Erreur lors du chargement", "Contacter Patrick"
    ```

## Fonctions de Gestion des Paramètres

### 19. ObtenirParametre

- **Description** : Récupère un paramètre depuis la table tbParametres dans data.xlsx.
- **Paramètres** :
    - `nomParametre` (String) : Nom du paramètre à récupérer
- **Retour** : String (valeur du paramètre ou chaîne vide si non trouvé)
- **Exemple d'utilisation** :
    ```vba
    Dim cheminData As String
    cheminData = ObtenirParametre("CheminData")
    ```

### 20. MettreAJourParametre

- **Description** : Met à jour un paramètre dans la table tbParametres dans data.xlsx.
- **Paramètres** :
    - `nomParametre` (String) : Nom du paramètre à mettre à jour
    - `valeur` (String) : Nouvelle valeur du paramètre
- **Retour** : Boolean (True si la mise à jour a réussi, False sinon)
- **Exemple d'utilisation** :
    ```vba
    Dim succes As Boolean
    succes = MettreAJourParametre("CheminData", "C:\nouveau_chemin")
    ```

## Fonctions de Gestion des Utilisateurs

### 21. ObtenirTrigrammeUtilisateur

- **Description** : Récupère le trigramme de l'utilisateur actuel depuis la table tbADV dans data.xlsx.
- **Paramètres** : Aucun
- **Retour** : String (trigramme de l'utilisateur ou chaîne vide si non trouvé)
- **Exemple d'utilisation** :
    ```vba
    Dim trigramme As String
    trigramme = ObtenirTrigrammeUtilisateur()
    ```

### 22. EstAdmin

- **Description** : Vérifie si l'utilisateur actuel est un administrateur.
- **Paramètres** : Aucun
- **Retour** : Boolean (True si l'utilisateur est un administrateur, False sinon)
- **Exemple d'utilisation** :
    ```vba
    If EstAdmin() Then
        ' Utilisateur est un administrateur
    End If
    ```

## Fonctions de Gestion des Commentaires

### 23. ValiderLongueurCommentaire

- **Description** : Vérifie si la longueur d'un commentaire est valide.
- **Paramètres** :
    - `commentaire` (String) : Commentaire à valider
    - `longueurMax` (Long) : Longueur maximale autorisée
- **Retour** : Boolean (True si la longueur est valide, False sinon)
- **Exemple d'utilisation** :
    ```vba
    If ValiderLongueurCommentaire(commentaire, 255) Then
        ' Commentaire est valide
    End If
    ```

### 24. TronquerCommentaire

- **Description** : Tronque un commentaire à la longueur maximale autorisée.
- **Paramètres** :
    - `commentaire` (String) : Commentaire à tronquer
    - `longueurMax` (Long) : Longueur maximale autorisée
- **Retour** : String (commentaire tronqué)
- **Exemple d'utilisation** :
    ```vba
    Dim commentaireTronque As String
    commentaireTronque = TronquerCommentaire(commentaire, 255)
    ```

## Fonctions de Gestion des Backups

### 25. CreerBackup

- **Description** : Crée une sauvegarde d'un fichier dans le répertoire de backups.
- **Paramètres** :
    - `cheminFichierSource` (String) : Chemin complet du fichier source
    - `cheminBackup` (String) : Chemin complet du fichier de backup
- **Retour** : Boolean (True si la sauvegarde a réussi, False sinon)
- **Exemple d'utilisation** :
    ```vba
    Dim succes As Boolean
    succes = CreerBackup("C:\data\suivi.xlsx", "C:\data\backups\20260123_suivi.xlsx")
    ```

### 26. ObtenirCheminBackup

- **Description** : Génère le chemin complet pour une sauvegarde avec un horodatage.
- **Paramètres** :
    - `cheminFichierSource` (String) : Chemin complet du fichier source
- **Retour** : String (chemin complet du fichier de backup)
- **Exemple d'utilisation** :
    ```vba
    Dim cheminBackup As String
    cheminBackup = ObtenirCheminBackup("C:\data\suivi.xlsx")
    ```

## Fonctions de Gestion des Retries

### 27. AttendreDelaiAleatoire

- **Description** : Attend un délai aléatoire entre deux valeurs spécifiées.
- **Paramètres** :
    - `delaiMin` (Long) : Délai minimum en secondes
    - `delaiMax` (Long) : Délai maximum en secondes
- **Retour** : Aucun
- **Exemple d'utilisation** :
    ```vba
    AttendreDelaiAleatoire 0, 3
    ```

### 28. ExecuterAvecRetry

- **Description** : Exécute une fonction avec un mécanisme de retry en cas d'échec.
- **Paramètres** :
    - `fonction` (Function) : Fonction à exécuter
    - `maxTentatives` (Long) : Nombre maximum de tentatives
    - `delaiMin` (Long) : Délai minimum en secondes entre les tentatives
    - `delaiMax` (Long) : Délai maximum en secondes entre les tentatives
- **Retour** : Boolean (True si la fonction a réussi, False sinon)
- **Exemple d'utilisation** :
    ```vba
    Dim succes As Boolean
    succes = ExecuterAvecRetry AddressOf MaFonction, 5, 0, 3
    ```
