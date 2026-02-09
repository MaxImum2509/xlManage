Attribute VB_Name = "modLogging"
' modLogging.bas - Logging tbAffaires
'
'-------------------------------------------------------------------------------
' GPL v3 - LICENSE
'-------------------------------------------------------------------------------
' Ce fichier fait partie de tbAffaires.
'
' tbAffaires est un logiciel libre : vous pouvez le redistribuer
' et/ou le modifier selon les termes de la Licence Publique Générale GNU
' telle que publiée par la Free Software Foundation, soit la version 3
' de la Licence, soit (à votre choix) toute version ultérieure.
'
' tbAffaires est distribué dans l'espoir qu'il sera utile,
' mais SANS AUCUNE GARANTIE ; sans même la garantie implicite de
' COMMERCIALISATION ou D'ADAPTATION À UN USAGE PARTICULIER.
' Voir la Licence Publique Générale GNU pour plus de détails.
'-------------------------------------------------------------------------------
'
' Auteur : Patrick
' Date : 2026-02-09
' Objet : Gestion du logging applicatif avec support des niveaux (INFO, ERREUR, SUCCES)
'         et du mode admin usurpé. Implémente la story 5.1.
'
' Éléments publics :
'   - Constantes : LOG_INFO, LOG_ERREUR, LOG_SUCCES
'   - EnregistrerLog(action, resultat, niveau) As Boolean
'   - EnregistrerErreur(action, messageErreur, codeErreur) As Boolean
'   - EnregistrerSucces(action, details) As Boolean
'   - EnregistrerInfo(action, details) As Boolean
'
'-------------------------------------------------------------------------------

Option Explicit

' Constantes pour les niveaux de log
Public Const LOG_INFO As String = "INFO"
Public Const LOG_ERREUR As String = "ERREUR"
Public Const LOG_SUCCES As String = "SUCCES"

' Chemin du fichier de log (utilisation de / pour compatibilité cross-platform)
Private Const NOM_FICHIER_LOG As String = "data/tbAffaires.log"

'-------------------------------------------------------------------------------
' EnregistrerLog
' Description : Enregistre une entrée dans le fichier de log
' Parameters  : action - String - Action effectuée
'               resultat - String - Résultat de l'action (ex: "50 affaires en 0.8 sec")
'               niveau - String - Niveau de log (LOG_INFO, LOG_ERREUR, LOG_SUCCES)
' Return      : Boolean - True si l'écriture s'est bien passée, False sinon
'-------------------------------------------------------------------------------
Public Function EnregistrerLog(action As String, resultat As String, Optional niveau As String = LOG_INFO) As Boolean
    On Error GoTo ErrorHandler

    Dim fichierLog As String
    Dim fileNum As Integer
    Dim ligneLog As String
    Dim utilisateurAffiche As String
    Dim cheminComplet As String

    ' Déterminer le chemin complet du fichier de log
    ' Utilisation de Replace pour convertir les / en \ pour Windows
    cheminComplet = ObtenirCheminLog()
    fichierLog = Replace(cheminComplet, "/", Application.PathSeparator)

    ' Obtenir le nom d'utilisateur formaté (avec support mode admin usurpé)
    utilisateurAffiche = ObtenirUtilisateurPourLog()

    ' Créer la ligne de log au format DATE | USER | ACTION | RESULTAT
    ' Format: "2026-01-23 14:32:15 | Patrick | Consolidation 50 affaires | SUCCES (0.8 sec)"
    ligneLog = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
               utilisateurAffiche & " | " & _
               action & " | " & _
               niveau

    ' Ajouter le résultat s'il n'est pas vide
    If Len(Trim(resultat)) > 0 Then
        ligneLog = ligneLog & " - " & resultat
    End If

    ' Obtenir un numéro de fichier libre
    fileNum = FreeFile

    ' Ouvrir le fichier en mode Append (ajout à la fin)
    ' Le fichier sera créé automatiquement s'il n'existe pas
    Open fichierLog For Append As #fileNum

    ' Écrire la ligne de log
    Print #fileNum, ligneLog

    ' Fermer le fichier
    Close #fileNum

    ' Retourner succès
    EnregistrerLog = True
    Exit Function

ErrorHandler:
    ' Si le fichier est inaccessible, ignorer silencieusement (selon exigence Story 5.1)
    ' Ne pas bloquer l'application à cause d'un problème de logging
    EnregistrerLog = False
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' ObtenirCheminLog
' Description : Construit le chemin complet du fichier de log
' Return      : String - Chemin complet du fichier de log
'-------------------------------------------------------------------------------
Private Function ObtenirCheminLog() As String
    On Error Resume Next

    Dim cheminFichier As String

    ' Essayer d'utiliser la configuration globale si disponible
    If modConfiguration.g_strFichierLog <> "" Then
        cheminFichier = ThisWorkbook.Path & "/" & modConfiguration.g_strFichierLog
    Else
        ' Utiliser le chemin par défaut si la configuration n'est pas chargée
        cheminFichier = ThisWorkbook.Path & "/" & NOM_FICHIER_LOG
    End If

    ' Réinitialiser la gestion d'erreur
    On Error GoTo 0

    ObtenirCheminLog = cheminFichier
End Function

'-------------------------------------------------------------------------------
' ObtenirUtilisateurPourLog
' Description : Obtient le nom d'utilisateur formaté pour le log (gestion mode admin usurpé)
' Return      : String - Nom d'utilisateur formaté pour les logs
'-------------------------------------------------------------------------------
Private Function ObtenirUtilisateurPourLog() As String
    On Error Resume Next

    Dim utilisateurReel As String
    Dim utilisateurAffiche As String

    ' Récupérer l'utilisateur système
    utilisateurReel = Environ("USERNAME")

    ' Par défaut, utiliser l'utilisateur réel
    utilisateurAffiche = utilisateurReel

    ' Vérifier si en mode admin usurpé (FR37)
    ' Format: "Patrick (pour HL)"
    If modConfiguration.g_boolEstAdminUsurpation Then
        If modConfiguration.g_strUtilisateurUsurpe <> "" Then
            utilisateurAffiche = utilisateurReel & " (pour " & modConfiguration.g_strUtilisateurUsurpe & ")"
        End If
    End If

    ' Réinitialiser la gestion d'erreur
    On Error GoTo 0

    ObtenirUtilisateurPourLog = utilisateurAffiche
End Function

'-------------------------------------------------------------------------------
' EnregistrerErreur
' Description : Enregistre une erreur dans le fichier de log avec le contexte complet (FR29)
' Parameters  : action - String - Action qui a causé l'erreur
'               messageErreur - String - Message d'erreur détaillé
'               codeErreur - String - Code d'erreur optionnel (ex: "ERR-101")
' Return      : Boolean - True si l'écriture réussie, False sinon
'-------------------------------------------------------------------------------
Public Function EnregistrerErreur(action As String, messageErreur As String, Optional codeErreur As String = "") As Boolean
    Dim resultat As String

    ' Formater le résultat avec code d'erreur si fourni
    If Len(Trim(codeErreur)) > 0 Then
        resultat = codeErreur & " - " & messageErreur
    Else
        resultat = messageErreur
    End If

    EnregistrerErreur = EnregistrerLog(action, resultat, LOG_ERREUR)
End Function

'-------------------------------------------------------------------------------
' EnregistrerSucces
' Description : Enregistre une opération réussie dans le fichier de log (FR30)
' Parameters  : action - String - Action effectuée (ex: "Consolidation")
'               details - String - Détails sur le résultat (ex: "50 affaires en 0.8 sec")
' Return      : Boolean - True si l'écriture réussie, False sinon
'-------------------------------------------------------------------------------
Public Function EnregistrerSucces(action As String, Optional details As String = "") As Boolean
    EnregistrerSucces = EnregistrerLog(action, details, LOG_SUCCES)
End Function

'-------------------------------------------------------------------------------
' EnregistrerInfo
' Description : Enregistre une information dans le fichier de log (FR30)
' Parameters  : action - String - Action ou événement à logger
'               details - String - Détails additionnels optionnels
' Return      : Boolean - True si l'écriture réussie, False sinon
'-------------------------------------------------------------------------------
Public Function EnregistrerInfo(action As String, Optional details As String = "") As Boolean
    EnregistrerInfo = EnregistrerLog(action, details, LOG_INFO)
End Function
