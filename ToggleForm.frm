VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ToggleForm 
   Caption         =   "ToggleTemplateLocations v1.1"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "ToggleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ToggleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub userform_initialize()
    CurrentMappingLabel_Update
End Sub

Public Sub CurrentMappingLabel_Update()
    Dim sSiteDetermination  As String
    Dim i                   As Integer
    Dim sReadTemp           As String
    Dim sReadCust           As String
    Dim sReadSheet          As String
    
    ' set a default-case conclusion
    sSiteDetermination = "Undetermined"
    ' read user preferences
    sReadTemp = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates)
    sReadCust = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsCustomPropertyFile)
    sReadSheet = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsSheetFormat)
    ' attempt to match these preference values to a site in the list
    For i = 0 To UBound(dbSites)
        If UCase(sReadTemp) = UCase(dbSites(i).Templates) And _
           UCase(sReadCust) = UCase(dbSites(i).CustomProperties) And _
           UCase(sReadSheet) = UCase(dbSites(i).SheetFormats) Then
            sSiteDetermination = dbSites(i).Name
        End If
    Next i
    
    ' update label text to show site reflected by settings
    ToggleForm.CurrentMappingLabel.Caption = "Current site settings in use: " & sSiteDetermination
    ToggleForm.LblDocTemp.Caption = sReadTemp
    ToggleForm.LblCustProp.Caption = sReadCust
    ToggleForm.LblSheet.Caption = sReadSheet
End Sub

Private Sub BtnBangalore_Click()
    ChangeSite "Bangalore"
End Sub

Private Sub BtnHoffman_Click()
    ChangeSite "Hoffman"
End Sub

Private Sub BtnOldHoffman_Click()
    ChangeSite "Old Hoffman"
End Sub

Private Sub BtnStraubenhardt_Click()
    ChangeSite "Straubenhardt"
End Sub

Private Sub BtnTexa_Click()
    ChangeSite "Texa"
End Sub

Private Sub BtnWIP_Click()
    ChangeSite "MJG WIP"
End Sub

Public Sub ChangeSite(sSiteName As String)
    Dim i       As Integer
    Dim sTemp   As String
    Dim sCust   As String
    Dim sSheet  As String
    
    ' find matching settings
    For i = 0 To UBound(dbSites)
        If UCase(dbSites(i).Name) = UCase(sSiteName) Then
            sTemp = dbSites(i).Templates
            sCust = dbSites(i).CustomProperties
            sSheet = dbSites(i).SheetFormats
        End If
    Next i
    
    ' apply settings
    swApp.SetUserPreferenceStringValue swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates, sTemp
    swApp.SetUserPreferenceStringValue swUserPreferenceStringValue_e.swFileLocationsCustomPropertyFile, sCust
    swApp.SetUserPreferenceStringValue swUserPreferenceStringValue_e.swFileLocationsSheetFormat, sSheet
    ' update label to confirm change made
    CurrentMappingLabel_Update
End Sub

