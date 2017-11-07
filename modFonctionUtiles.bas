Attribute VB_Name = "modFonctionUtiles"
Option Compare Database
Option Explicit

Public Sub MiseAJourInfoMenage()
    Dim strsql As String
    DoCmd.SetWarnings False
    ' Récupérer le groupe d'âge du client principal
    ' Exclusion sur les clients sans date de naissance
    strsql = "SELECT ComptoirClient_Data.NoClient, IIf(Month(Date())>=Month([DateNaissance]) And Day(Date())>=Day([DateNaissance]),DateDiff('yyyy',[DateNaissance],Date()),DateDiff('yyyy',[DateNaissance],Date())-1) AS Age " & _
             "INTO CPT0000 IN """ & cnsCheminBDTemp & """ " & _
             "FROM ComptoirClient_Data " & _
             "WHERE (((ComptoirClient_Data.DateNaissance)<>#1/1/1900#));"
    
    DoCmd.RunSQL strsql
    
    ' Mise à jour de l'âge - Client
    strsql = "UPDATE CPT0000 INNER JOIN ComptoirClient_Data ON CPT0000.NoClient = ComptoirClient_Data.NoClient SET ComptoirClient_Data.AGE = [CPT0000].[Age];"
    DoCmd.RunSQL strsql
    
    ' Calculer l'âge des ménages
    strsql = "SELECT ComptoirMenage_Data.NoClient, ComptoirMenage_Data.NoMenage, IIf(Month(Date())>=Month([DateNaissance]) And Day(Date())>=Day([DateNaissance]),DateDiff('yyyy',[DateNaissance],Date()),DateDiff('yyyy',[DateNaissance],Date())-1) AS Age " & _
             "INTO CPT0000 IN """ & cnsCheminBDTemp & """ " & _
             "FROM ComptoirMenage_Data " & _
             "WHERE (((ComptoirMenage_Data.DateNaissance)<>#1/1/1900#));"

    DoCmd.RunSQL strsql
    
    ' Mise à jour de l'âge - Menage
    strsql = "UPDATE CPT0000 INNER JOIN ComptoirMenage_Data ON (CPT0000.NoMenage = ComptoirMenage_Data.NoMenage) AND (CPT0000.NoClient = ComptoirMenage_Data.NoClient) SET ComptoirMenage_Data.Age = [CPT0000].[Age];"
    DoCmd.RunSQL strsql
End Sub

Public Sub MiseEnProductionRapport()
    Dim fso As New FileSystemObject
    
    If fso.FileExists(cnsModeleHebdoMEP) Then
        fso.CopyFile cnsModeleHebdoMEP, cnsModeleRapportHebdo
        fso.DeleteFile cnsModeleHebdoMEP
    End If
    
    If fso.FileExists(cnsModeleRapportFacturationMEP) Then
        fso.CopyFile cnsModeleRapportFacturationMEP, cnsModeleRapportFacturation
        fso.DeleteFile cnsModeleRapportFacturationMEP
    End If
End Sub

Public Sub PriseBackup()
    Dim fso As New FileSystemObject
    
    fso.CopyFile cnsCheminBDData, cnsCheminBDData & "_" & Date
End Sub
' Calcul de l'âge du client
Function dhAge(pDtNaissance As Date, Optional dtmDate As Date = 0) As Integer
    ' This procedure is stored as dhAgeUnused in the sample
    ' module.
    Dim intAge As Integer
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    intAge = DateDiff("yyyy", pDtNaissance, dtmDate)
    If dtmDate < DateSerial(Year(dtmDate), Month(pDtNaissance), Day(pDtNaissance)) Then
        intAge = intAge - 1
    End If
    dhAge = intAge
End Function

Public Function retournerNombreClient() As Integer
    retournerNombreClient = DCount("NOCLIENT", "ComptoirClient_Data")
End Function

Public Function CheckNewField(pstrTable As String)
    Dim db As DAO.Database
    Dim dbCur As DAO.Database
    Dim fld As DAO.Field
    Dim tdf_Struct As DAO.TableDef
    Dim tdf As DAO.TableDef
    Dim strAllField As String
    Dim strAllField_Struct As String
    Dim strType As String
    Dim strDefaultValue As String
    
    On Error Resume Next
    Set dbCur = CurrentDb()
    Set db = OpenDatabase(cnsCheminBDData, False, False, ";PWD=legrenier_pwd")

    Set tdf_Struct = dbCur.TableDefs(pstrTable & "_STRUCT")
    Set tdf = db.TableDefs(pstrTable)

    For Each fld In tdf.Fields
        strAllField = strAllField & fld.Name & ";"
    Next
        
    For Each fld In tdf_Struct.Fields
        If InStr(1, strAllField, fld.Name) = 0 Then
            Select Case fld.Type
                Case dbBoolean
                    strType = "YESNO"
                    strDefaultValue = "-1"
                Case dbText
                    strType = "TEXT (255)"
                Case dbLong
                    strType = "INTEGER"
                Case dbDouble
                    strType = "DOUBLE"
                Case dbInteger
                    strType = "INTEGER"
                Case Else
            End Select
            db.Execute "ALTER TABLE " & pstrTable & " ADD " & fld.Name & " " & strType & ";"
            db.Execute "UPDATE " & pstrTable & " SET " & fld.Name & " = " & strDefaultValue & ";"
        End If
    Next
        
    For Each fld In tdf_Struct.Fields
        strAllField_Struct = strAllField_Struct & fld.Name & ";"
    Next
            
            
    If strAllField <> strAllField_Struct Then
        MsgBox "Erreur dans l'ajout d'un champ, voir avec Jonathan!", vbCritical, "Erreur"
    End If
    db.Close
    dbCur.Close
End Function

Public Sub CreerDBStats()
    Dim db As DAO.Database
    Dim wrkDefault  As Workspace
    Dim fso As New FileSystemObject
    Set wrkDefault = DBEngine.Workspaces(0)
    
    If fso.FileExists(cnsCheminBDStats) = False Then
        Set db = wrkDefault.CreateDatabase(cnsCheminBDStats, dbLangGeneral)
    End If
End Sub

Public Function remplacerNulleOuVide(pstrTable As String)
        
        
    Select Case pstrTable
        Case "CPTSTATS"
            
        Case Else
        
    End Select
    
    
End Function
