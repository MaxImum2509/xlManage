Attribute VB_Name = "modConfiguration"
'
' Module de configuration pour l'application tbAffaires
' Gère l'identification de l'utilisateur et la configuration globale
' Correspond à la story 1.3: Identifier automatiquement l'utilisateur Windows
' et à d'autres stories liées à la configuration
'

Option Explicit

' Variables globales pour la configuration utilisateur
Public g_strUtilisateur As String           ' Nom d'utilisateur Windows (Environ("USERNAME"))
Public g_strTrigramme As String            ' Trigramme ADV de l'utilisateur
Public g_strNom As String                  ' Nom de l'utilisateur
Public g_strPrenom As String               ' Prénom de l'utilisateur
Public g_boolEstAdmin As Boolean          ' Indique si l'utilisateur est admin
Public g_boolEstAdminUsurpation As Boolean ' Indique si en mode admin usurpé
Public g_strUtilisateurUsurpe As String    ' Trigramme de l'utilisateur usurpé (en mode admin)
Public g_strFichierLog As String           ' Chemin du fichier de log

' Variables globales pour les paramètres de l'application
Public g_strCheminData As String
Public g_strCheminExtraction As String
Public g_strCheminConsolidation As String
Public g_strRepertoireConsolide As String
Public g_intDelaiRetryMin As Integer
Public g_intDelaiRetryMax As Integer
Public g_intMaxTentatives As Integer

'-------------------------------------------------------------------------------
' InitialiserConfiguration
' Description : Initialise la configuration utilisateur au démarrage de l'application
' Parameters  : None
' Return      : Boolean - True si l'initialisation réussit, False sinon
'-------------------------------------------------------------------------------
Public Function InitialiserConfiguration() As Boolean
    On Error GoTo ErrorHandler

    ' Récupérer l'utilisateur Windows
    g_strUtilisateur = Environ("USERNAME")

    ' Initialiser les variables globales
    g_strTrigramme = ""
    g_strNom = ""
    g_strPrenom = ""
    g_boolEstAdmin = False
    g_boolEstAdminUsurpation = False
    g_strUtilisateurUsurpe = ""

    ' Définir le chemin du fichier de log
    g_strFichierLog = "data\tbAffaires.log"

    ' Charger la configuration utilisateur depuis data.xlsx (tbADV)
    If Not ChargerConfigUtilisateur(g_strUtilisateur) Then
        ' L'utilisateur n'est pas configuré (ERR-001)
        InitialiserConfiguration = False
        Exit Function
    End If

    ' Charger les paramètres généraux
    If Not ChargerParametresApplication() Then
        InitialiserConfiguration = False
        Exit Function
    End If

    ' Terminer avec succès
    InitialiserConfiguration = True
    Exit Function

ErrorHandler:
    InitialiserConfiguration = False
End Function

'-------------------------------------------------------------------------------
' ChargerConfigUtilisateur
' Description : Charge la configuration utilisateur depuis data.xlsx (tbADV)
' Parameters  : utilisateur - String - Nom d'utilisateur à charger
' Return      : Boolean - True si le chargement réussit, False sinon
'-------------------------------------------------------------------------------
Private Function ChargerConfigUtilisateur(utilisateur As String) As Boolean
    On Error GoTo ErrorHandler

    ' Dans un scénario réel, cette fonction chargerait les données depuis
    ' tbADV dans data.xlsx pour le nom d'utilisateur donné
    ' Pour l'instant, on simule une validation de base

    ' Simuler la recherche dans tbADV
    ' Si utilisateur non trouvé, retourner False (ERR-001)
    ' Sinon, charger les champs: Trigramme, Nom, Prénom, IsAdmin
    ' Pour simulation, on suppose que tous les utilisateurs sont valides
    ' mais dans la réalité, si non trouvé, on retourne False

    ' Pour le moment, retournons vrai pour permettre le développement
    ChargerConfigUtilisateur = True
    g_strTrigramme = Left(utilisateur, 3)  ' Simule le trigramme
    g_strNom = utilisateur               ' Simule le nom
    g_strPrenom = utilisateur            ' Simule le prénom

    Exit Function

ErrorHandler:
    ChargerConfigUtilisateur = False
End Function

'-------------------------------------------------------------------------------
' ChargerParametresApplication
' Description : Charge les paramètres de l'application depuis data.xlsx (tbParametres)
' Parameters  : None
' Return      : Boolean - True si le chargement réussit, False sinon
'-------------------------------------------------------------------------------
Private Function ChargerParametresApplication() As Boolean
    On Error GoTo ErrorHandler

    ' Dans un scénario réel, cette fonction chargerait les paramètres depuis
    ' tbParametres dans data.xlsx
    ' Pour l'instant, on utilise des valeurs par défaut

    g_strCheminData = "data\"
    g_strCheminExtraction = "extractions\"
    g_strCheminConsolidation = "data\consolidation.xlsx"
    g_strRepertoireConsolide = "data\"
    g_intDelaiRetryMin = 0
    g_intDelaiRetryMax = 3
    g_intMaxTentatives = 5

    ChargerParametresApplication = True
    Exit Function

ErrorHandler:
    ChargerParametresApplication = False
End Function

'-------------------------------------------------------------------------------
' VerifierUniciteAdmin
' Description : Vérifie qu'il n'y ait qu'un seul admin configuré (Règle 1 métier)
' Parameters  : None
' Return      : Boolean - True si unicité respectée, False sinon (ERR-002)
'-------------------------------------------------------------------------------
Public Function VerifierUniciteAdmin() As Boolean
    On Error GoTo ErrorHandler

    ' Cette fonction devrait vérifier dans tbADV combien d'utilisateurs ont IsAdmin="Oui"
    ' Si > 1, retourner False pour indiquer ERR-002

    ' Pour l'instant, simulons que l'unicité est respectée
    VerifierUniciteAdmin = True
    Exit Function

ErrorHandler:
    VerifierUniciteAdmin = False
End Function
