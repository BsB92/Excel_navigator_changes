Attribute VB_Name = "modNavigatorSettings"
Option Explicit

Private Const REG_APP As String = "ExcelNavigator46_alpha"
Private Const REG_SEC_SETTINGS As String = "Settings"
Private Const KEY_DEFAULT_FOLDER As String = "DefaultFolder"
Private Const HELP_FILE_NAME As String = "ExcelNavigator_Help.txt"

Public Function GetDefaultWorkingFolder() As String
    GetDefaultWorkingFolder = GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, ThisWorkbook.Path)
End Function

Public Sub SaveDefaultWorkingFolder(ByVal folderPath As String)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, folderPath
End Sub

Public Function ResolveInitialFolder(ByVal fallbackFolder As String) As String
    Dim configuredFolder As String
    configuredFolder = Trim$(GetDefaultWorkingFolder())

    If Len(configuredFolder) > 0 Then
        If Len(Dir$(configuredFolder, vbDirectory)) > 0 Then
            ResolveInitialFolder = configuredFolder
            Exit Function
        End If
    End If

    ResolveInitialFolder = fallbackFolder
End Function

Public Sub OpenHelpInstructions()
    Dim helpPath As String
    helpPath = EnsureHelpInstructionsFile()
    ThisWorkbook.FollowHyperlink helpPath
End Sub

Public Function EnsureHelpInstructionsFile() As String
    Dim helpPath As String
    helpPath = ThisWorkbook.Path & Application.PathSeparator & HELP_FILE_NAME

    If Len(Dir$(helpPath, vbNormal)) = 0 Then
        WriteHelpInstructions helpPath
    End If

    EnsureHelpInstructionsFile = helpPath
End Function

Private Sub WriteHelpInstructions(ByVal helpPath As String)
    Dim f As Integer
    f = FreeFile

    Open helpPath For Output As #f
    Print #f, "Excel Navigator - Instrukcja"
    Print #f, ""
    Print #f, "1) Search: filtruje listę otwartych plików po nazwie i pełnej ścieżce."
    Print #f, "2) Reload: odświeża listę aktywnych skoroszytów bez utraty ustawień formularza."
    Print #f, "3) Selection mode OFF/ON: przełącza między wyborem pojedynczego i wielu plików."
    Print #f, "4) Open: aktywuje zaznaczony skoroszyt."
    Print #f, "5) Open & Refresh: otwiera i odświeża dane w wybranych skoroszytach."
    Print #f, "6) Refresh / Save / Refresh+Save: operacje wsadowe dla zaznaczonych plików."
    Print #f, "7) Copy + Break Links: tworzy kopię pliku, a następnie zrywa połączenia zewnętrzne."
    Print #f, "8) Copy with suffix: tworzy kopię pliku z dopiskiem z pola Suffix."
    Print #f, "9) Open copied files / Open folder: szybkie otwieranie wyników kopiowania."
    Print #f, "10) Close selected: zamyka zaznaczone pliki z obsługą zapisanych/niezapisanych zmian."
    Print #f, "11) Lista arkuszy (panel po prawej): przejście do konkretnego arkusza."
    Print #f, "12) Settings: ustawia domyślną lokalizację folderu używaną przy operacjach Open/Copy."
    Print #f, ""
    Print #f, "Wskazówka: ustaw domyślny folder w Settings i kliknij Save settings."
    Close #f
End Sub
