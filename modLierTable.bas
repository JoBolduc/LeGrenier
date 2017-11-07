Attribute VB_Name = "modLierTable"
Option Compare Database
Option Explicit

Public Sub RelierTable()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConnect As String
    Set db = CurrentDb
    For Each tdf In db.TableDefs
        ' ignore system and temporary tables
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
            If InStr(1, tdf.Connect, "legrenier_pwd") <> 0 Then
                If LinkTable(tdf.Name, tdf.Name, cnsConnectData) = False Then MsgBox "Erreur dans la liaison des tables.", vbCritical, "Liaison tables"
            Else
                If LinkTable(tdf.Name, tdf.Name, ";" & cnsConnectTemp) = False Then MsgBox "Erreur dans la liaison des tables.", vbCritical, "Liaison tables"
            End If
            
        End If
    Next


    Set tdf = Nothing
    Set db = Nothing
End Sub
Private Function LinkTable(LinkedTableName As String, TableToLink As String, connectString As String) As Boolean
    Dim tdf As New DAO.TableDef
    
    On Error GoTo LinkTable_Error
    
    DoCmd.RunSQL "DROP TABLE " & LinkedTableName & ";"
    
    With CurrentDb
    
        .TableDefs.Refresh
    
        Set tdf = .CreateTableDef(LinkedTableName)
        tdf.Connect = connectString
        tdf.SourceTableName = TableToLink
        .TableDefs.Append tdf
        .TableDefs.Refresh
    
    
    End With
    
    Set tdf = Nothing
    LinkTable = True
    
    Exit Function
LinkTable_Error:
    LinkTable = False
End Function
