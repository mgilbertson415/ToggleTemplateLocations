Attribute VB_Name = "ToggleTemplateLocations1"
Public Type SitePathMappings
    Name                As String
    Templates           As String
    CustomProperties    As String
    SheetFormats        As String
End Type

Public dbSites() As SitePathMappings
Public swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    ReDim dbSites(2)
    
    dbSites(0).Name = "Hoffman"
    dbSites(0).Templates = "\\nventco.com\global\SolidWorks_Templates\Hoffman\Templates"
    dbSites(0).CustomProperties = "\\nventco.com\Global\SolidWorks_Templates\Hoffman\Custom Properties"
    dbSites(0).SheetFormats = "\\nventco.com\global\SolidWorks_Templates\Hoffman\Templates\Sheet Formats"
    
    dbSites(1).Name = "Bangalore"
    dbSites(1).Templates = ""
    dbSites(1).CustomProperties = ""
    dbSites(1).SheetFormats = ""
    
    dbSites(2).Name = "Straubenhardt"
    dbSites(2).Templates = "\\destrap28\solidworks\Vorlagen\Dokumentvorlagen2024;\\destrap28\solidworks\Vorlagen\Dokumentvorlagen2024\Einzelteil;\\destrap28\solidworks\Vorlagen\Dokumentvorlagen2024\Zeichnungsvorlagen;\\destrap28\solidworks\Vorlagen\Dokumentvorlagen2024\Baugruppe"
    dbSites(2).CustomProperties = "\\destrap28\solidworks\Vorlagen\Dateien mit benutzerdefinierten Eigenschaften"
    dbSites(2).SheetFormats = "\\destrap28\solidworks\Vorlagen\Blattformate2024;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BF-A4;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BF-A3;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BF-A2;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BF-A1;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BF-A0;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\FLAT;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\PRINT;" & _
                            "\\destrap28\solidworks\Vorlagen\Platzhalter\Platzhalter_1;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BMK-A4;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BMK-A3;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BMK-A2;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BMK-A1;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\BMK-A0;"
                            dbSites(1).SheetFormats = dbSites(1).SheetFormats & _
                            "\\destrap28\solidworks\Vorlagen\Platzhalter\Platzhalter_2;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\FT-A4;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\FT-A3;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\FT-A2;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\FT-A1;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\FT-A0;" & _
                            "\\destrap28\solidworks\Vorlagen\Platzhalter\Platzhalter_3;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\PR-A4;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\PR-A3;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\PR-A2;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\PR-A1;" & _
                            "\\destrap28\solidworks\Vorlagen\Platzhalter\Platzhalter_4;" & _
                            "\\destrap28\solidworks\Vorlagen\Blattformate2024\Z-AZUBI"
    
    ToggleForm.Show False
    
End Sub
