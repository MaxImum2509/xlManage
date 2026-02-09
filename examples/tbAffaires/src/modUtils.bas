Attribute VB_Name = "modUtils"
' modUtils.bas - Utilitaires fondamentaux tbAffaires
'
'-------------------------------------------------------------------------------
' GPL v3 - LICENSE
'-------------------------------------------------------------------------------
' tbAffaires - Application de suivi hebdomadaire des affaires pour ADV
' Copyright (C) 2026 Equipe tbAffaires
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'-------------------------------------------------------------------------------
' Auteur : Equipe tbAffaires
' Date : 2026-02-08
' Objet : Module utilitaire centralisé avec constantes, helpers et gestion d'erreurs
'         Implémente la Story 1.6
'
' Éléments publics :
'   - Constantes : Codes d'erreur ERR_* et leurs messages
'   - FichierExiste(chemin) As Boolean
'   - RepertoireExiste(chemin) As Boolean
'   - ChargerWorkbook(chemin) As Workbook
'   - ChargerWorksheet(wb, nomFeuille) As Worksheet
'   - ChargerListObject(ws, nomTable) As ListObject
'   - AfficherMessageErreur(codeErreur, detailsSupplementaires)
'   - AfficherMessageInfo(titre, message)
'-------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------
' CONSTANTES - CODES D'ERREUR ET MESSAGES
'-------------------------------------------------------------------------------

' Erreurs d'authentification et configuration utilisateur (ERR-001 à ERR-099)
Public Const ERR_001 As String = "ERR-001"
Public Const ERR_001_MSG As String = "Utilisateur non configuré dans tbADV."
Public Const ERR_001_ACTION As String = "Contacter Patrick pour ajouter votre compte."

Public Const ERR_002 As String = "ERR-002"
Public Const ERR_002_MSG As String = "Plusieurs administrateurs détectés dans tbADV."
Public Const ERR_002_ACTION As String = "Contacter Patrick pour corriger la configuration."

' Erreurs de validation des données (ERR-101 à ERR-199)
Public Const ERR_101 As String = "ERR-101"
Public Const ERR_101_MSG As String = "Colonne du mapping introuvable dans l'extraction ERP."
Public Const ERR_101_ACTION As String = "Vérifier le fichier d'extraction ou contacter Patrick pour mettre à jour le mapping."

Public Const ERR_102 As String = "ERR-102"
Public Const ERR_102_MSG As String = "Fichier d'extraction ERP introuvable ou non sélectionné."
Public Const ERR_102_ACTION As String = "Vérifier le chemin du fichier ou recommencer la sélection."

Public Const ERR_103 As String = "ERR-103"
Public Const ERR_103_MSG As String = "Format du fichier consolidé invalide (structure incorrecte ou colonnes manquantes)."
Public Const ERR_103_ACTION As String = "Choisir un autre fichier consolidé ou continuer sans commentaires historiques."

' Erreurs de consolidation et accès fichier (ERR-201 à ERR-299)
Public Const ERR_201 As String = "ERR-201"
Public Const ERR_201_MSG As String = "Fichier de consolidation occupé par un autre utilisateur."
Public Const ERR_201_ACTION As String = "Attendre quelques secondes, l'application réessaiera automatiquement."

Public Const ERR_202 As String = "ERR-202"
Public Const ERR_202_MSG As String = "Échec de la consolidation après 5 tentatives."
Public Const ERR_202_ACTION As String = "NE PAS FERMER l'application (vos données sont préservées). Contacter Patrick immédiatement."

' Erreurs de saisie utilisateur (ERR-301 à ERR-399)
Public Const ERR_301 As String = "ERR-301"
Public Const ERR_301_MSG As String = "Commentaire trop long (maximum 255 caractères)."
Public Const ERR_301_ACTION As String = "Raccourcir le commentaire et réessayer."

' Alertes mode administrateur (ERR-401 à ERR-499)
Public Const ERR_401 As String = "ERR-401"
Public Const ERR_401_MSG As String = "Mode Administrateur actif - Travail au nom d'un autre utilisateur."
Public Const ERR_401_ACTION As String = "Vérifier le trigramme affiché. Toutes les actions seront enregistrées sous votre nom."

'-------------------------------------------------------------------------------
' FichierExiste - Vérifie l'existence d'un fichier
' Parameters : chemin As String - Chemin complet du fichier à vérifier
' Return     : Boolean - True si le fichier existe, False sinon
'-------------------------------------------------------------------------------
Public Function FichierExiste(ByVal chemin As String) As Boolean
    On Error Resume Next
    FichierExiste = (Dir(chemin) <> "")
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' RepertoireExiste - Vérifie l'existence d'un répertoire
' Parameters : chemin As String - Chemin complet du répertoire à vérifier
' Return     : Boolean - True si le répertoire existe, False sinon
'-------------------------------------------------------------------------------
Public Function RepertoireExiste(ByVal chemin As String) As Boolean
    On Error Resume Next
    RepertoireExiste = (Dir(chemin, vbDirectory) <> "")
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' ChargerWorkbook - Ouvre un classeur Excel et retourne l'objet Workbook
' Parameters : chemin As String - Chemin complet du fichier Excel
'              lectureSeule As Boolean - Ouvrir en lecture seule (optionnel, défaut = False)
' Return     : Workbook - Objet workbook si succès, Nothing si échec
'-------------------------------------------------------------------------------
Public Function ChargerWorkbook(ByVal chemin As String, _
                                Optional ByVal lectureSeule As Boolean = False) As Workbook
    On Error GoTo ErrorHandler

    If Not FichierExiste(chemin) Then
        Call modLogging.EnregistrerErreur("ChargerWorkbook", "Fichier introuvable : " & chemin)
        Set ChargerWorkbook = Nothing
        Exit Function
    End If

    Set ChargerWorkbook = Workbooks.Open(Filename:=chemin, ReadOnly:=lectureSeule)
    Call modLogging.EnregistrerInfo("ChargerWorkbook", "Fichier ouvert : " & chemin)
    Exit Function

ErrorHandler:
    Call modLogging.EnregistrerErreur("ChargerWorkbook", "Erreur ouverture : " & Err.Description)
    Set ChargerWorkbook = Nothing
End Function

'-------------------------------------------------------------------------------
' ChargerWorksheet - Récupère une feuille par son nom dans un classeur
' Parameters : wb As Workbook - Classeur source
'              nomFeuille As String - Nom de la feuille à récupérer
' Return     : Worksheet - Objet worksheet si trouvé, Nothing si introuvable
'-------------------------------------------------------------------------------
Public Function ChargerWorksheet(ByVal wb As Workbook, _
                                 ByVal nomFeuille As String) As Worksheet
    On Error GoTo ErrorHandler

    If wb Is Nothing Then
        Call modLogging.EnregistrerErreur("ChargerWorksheet", "Workbook Nothing")
        Set ChargerWorksheet = Nothing
        Exit Function
    End If

    Set ChargerWorksheet = wb.Worksheets(nomFeuille)
    Exit Function

ErrorHandler:
    Call modLogging.EnregistrerErreur("ChargerWorksheet", "Feuille introuvable : " & nomFeuille)
    Set ChargerWorksheet = Nothing
End Function

'-------------------------------------------------------------------------------
' ChargerListObject - Récupère un ListObject (tableau Excel) par son nom
' Parameters : ws As Worksheet - Feuille source
'              nomTable As String - Nom du ListObject à récupérer
' Return     : ListObject - Objet ListObject si trouvé, Nothing si introuvable
'-------------------------------------------------------------------------------
Public Function ChargerListObject(ByVal ws As Worksheet, _
                                  ByVal nomTable As String) As ListObject
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Call modLogging.EnregistrerErreur("ChargerListObject", "Worksheet Nothing")
        Set ChargerListObject = Nothing
        Exit Function
    End If

    Set ChargerListObject = ws.ListObjects(nomTable)
    Exit Function

ErrorHandler:
    Call modLogging.EnregistrerErreur("ChargerListObject", "Table introuvable : " & nomTable)
    Set ChargerListObject = Nothing
End Function

'-------------------------------------------------------------------------------
' AfficherMessageErreur - Affiche un message d'erreur standardisé avec code, message et action
' Parameters : codeErreur As String - Code d'erreur (ex: ERR-101)
'              detailsSupplementaires As String - Détails additionnels optionnels
' Return     : None
'-------------------------------------------------------------------------------
Public Sub AfficherMessageErreur(ByVal codeErreur As String, _
                                 Optional ByVal detailsSupplementaires As String = "")
    Dim messageComplet As String
    Dim titre As String
    Dim messageErreur As String
    Dim actionCorrective As String

    titre = "Erreur - " & codeErreur

    ' Récupérer le message et l'action en fonction du code d'erreur
    Select Case codeErreur
        Case ERR_001
            messageErreur = ERR_001_MSG
            actionCorrective = ERR_001_ACTION
        Case ERR_002
            messageErreur = ERR_002_MSG
            actionCorrective = ERR_002_ACTION
        Case ERR_101
            messageErreur = ERR_101_MSG
            actionCorrective = ERR_101_ACTION
        Case ERR_102
            messageErreur = ERR_102_MSG
            actionCorrective = ERR_102_ACTION
        Case ERR_103
            messageErreur = ERR_103_MSG
            actionCorrective = ERR_103_ACTION
        Case ERR_201
            messageErreur = ERR_201_MSG
            actionCorrective = ERR_201_ACTION
        Case ERR_202
            messageErreur = ERR_202_MSG
            actionCorrective = ERR_202_ACTION
        Case ERR_301
            messageErreur = ERR_301_MSG
            actionCorrective = ERR_301_ACTION
        Case ERR_401
            messageErreur = ERR_401_MSG
            actionCorrective = ERR_401_ACTION
        Case Else
            messageErreur = "Erreur inconnue"
            actionCorrective = "Contacter Patrick"
    End Select

    ' Construire le message complet
    messageComplet = messageErreur & vbCrLf & vbCrLf

    If Len(Trim(detailsSupplementaires)) > 0 Then
        messageComplet = messageComplet & "Détails : " & detailsSupplementaires & vbCrLf & vbCrLf
    End If

    messageComplet = messageComplet & "Action : " & actionCorrective

    ' Afficher le message
    MsgBox messageComplet, vbCritical, titre

    ' Logger l'erreur
    Call modLogging.EnregistrerErreur("Message utilisateur", messageErreur, codeErreur)
End Sub

'-------------------------------------------------------------------------------
' AfficherMessageInfo - Affiche un message d'information et le logue
' Parameters : titre As String - Titre du message
'              message As String - Contenu du message
' Return     : None
'-------------------------------------------------------------------------------
Public Sub AfficherMessageInfo(ByVal titre As String, _
                               ByVal message As String)
    MsgBox message, vbInformation, titre
    Call modLogging.EnregistrerInfo("Message utilisateur", titre & " : " & message)
End Sub

'-------------------------------------------------------------------------------
' AfficherMessageSucces - Affiche un message de succès et le logue
' Parameters : titre As String - Titre du message
'              message As String - Contenu du message
'              tempsExecution As String - Temps d'exécution optionnel (ex: "0.8 sec")
' Return     : None
'-------------------------------------------------------------------------------
Public Sub AfficherMessageSucces(ByVal titre As String, _
                                 ByVal message As String, _
                                 Optional ByVal tempsExecution As String = "")
    Dim messageComplet As String

    messageComplet = message
    If Len(Trim(tempsExecution)) > 0 Then
        messageComplet = messageComplet & " (" & tempsExecution & ")"
    End If

    MsgBox messageComplet, vbInformation, titre
    Call modLogging.EnregistrerSucces(titre, messageComplet)
End Sub

'-------------------------------------------------------------------------------
' ConvertirCheminPOSIX - Convertit un chemin POSIX (/) en chemin Windows (\)
' Parameters : cheminPOSIX As String - Chemin avec séparateurs /
' Return     : String - Chemin avec séparateurs natifs de la plateforme
' Note       : Utilisé pour compatibilité des chemins stockés en POSIX dans data.xlsx
'-------------------------------------------------------------------------------
Public Function ConvertirCheminPOSIX(ByVal cheminPOSIX As String) As String
    ConvertirCheminPOSIX = Replace(cheminPOSIX, "/", Application.PathSeparator)
End Function

'-------------------------------------------------------------------------------
' ValiderChaineNonVide - Vérifie qu'une chaîne n'est pas vide après Trim
' Parameters : valeur As String - Chaîne à valider
'              nomChamp As String - Nom du champ (pour message d'erreur)
' Return     : Boolean - True si valide, False si vide
'-------------------------------------------------------------------------------
Public Function ValiderChaineNonVide(ByVal valeur As String, _
                                     ByVal nomChamp As String) As Boolean
    If Len(Trim(valeur)) = 0 Then
        Call modLogging.EnregistrerErreur("Validation", "Champ vide : " & nomChamp)
        ValiderChaineNonVide = False
    Else
        ValiderChaineNonVide = True
    End If
End Function
