Attribute VB_Name = "modRapports"
Option Compare Database
Option Explicit
Public mstrParamClient As String
Public mblnParamProvenance As Boolean
Public mblnParamBon As Boolean
Public mstrParamService As String
Private mstrWhere As String
Private mstrSelect As String
Private mstrSelectTemp As String
Private mstrGroupBy As String
Private mstrGroupByTemp As String
Private mstrTabColonneExcel() As String
Private nombreColonne As Integer
Public Sub test()
    Call ProductionRapportsHebdo(#1/1/2016#, #12/31/2016#, 0, "")
End Sub
Private Sub SetSelect()
    Dim strSelect As String
    Dim strND As String
    
    strSelect = ""
    mstrSelect = ""
    mstrSelectTemp = ""
    nombreColonne = 0
    strND = "Non déterminé"
    
'    If Form_frm_4_1_Rapport.chkSexe.Value = True Then
        strSelect = strSelect & "IIF(IsNull([Sexe]) Or [Sexe]='', """ & strND & """,[Sexe]) AS [Sexe (M ou F)], "
        mstrSelectTemp = mstrSelectTemp & "CPT0001.[Sexe (M ou F)], "
        nombreColonne = nombreColonne + 1
'    End If
    
'    If Form_frm_4_1_Rapport.chkNationalite.Value = -1 Then
        strSelect = strSelect & "IIF(IsNull([1ERNation_18]) AND IsNull([Imigration_18]), '', IIF([1ERNation_18]<>0, 'Première nation', IIF([Imigration_18]<>0, 'Immigrant', ''))) AS [Nationalité], "
        mstrSelectTemp = mstrSelectTemp & "CPT0001.[Nationalité], "
        nombreColonne = nombreColonne + 1
'    End If
    
'    If Form_frm_4_1_Rapport.chkAge.Value = -1 Then
        strSelect = strSelect & "IIf(IsNull([AGE]), """ & strND & """, IIF([AGE]>=18 And [AGE]<=30,'18-30 ans',IIf([AGE]>=31 And [AGE]<=44,'31-44 ans',IIf([AGE]>=45 And [AGE]<=64,'45-64 ans',IIf([AGE]>=65,'65 ans et plus',''))))) AS TrancheAge, "
        mstrSelectTemp = mstrSelectTemp & "CPT0001.[TrancheAge], "
        nombreColonne = nombreColonne + 1
'    End If
    
'    If Form_frm_4_1_Rapport.chkMenage.Value = -1 Then
        strSelect = strSelect & "IIF(IsNull(PILO_TYPE_MENAGE.CODE), """ & strND & """, PILO_TYPE_MENAGE.DESCRIPTION) AS TypMenage, "
        mstrSelectTemp = mstrSelectTemp & "CPT0001.[TypMenage], "
        nombreColonne = nombreColonne + 1
'    End If
    
'    If Form_frm_4_1_Rapport.chkRevenu.Value = -1 Then
        strSelect = strSelect & "IIF(IsNull(PILO_TYPE_REVENU.CODE), """ & strND & """,PILO_TYPE_REVENU.DESCRIPTION) AS SourceRevenu, "
        mstrSelectTemp = mstrSelectTemp & "CPT0001.[SourceRevenu], "
        nombreColonne = nombreColonne + 1
'    End If
    
'    If Form_frm_4_1_Rapport.chkLogement.Value = -1 Then
        strSelect = strSelect & "IIF(IsNull(PILO_TYPE_LOGEMENT.CODE),""" & strND & """, PILO_TYPE_LOGEMENT.DESCRIPTION) AS TypLogement, "
        mstrSelectTemp = mstrSelectTemp & "CPT0001.[TypLogement], "
        nombreColonne = nombreColonne + 1
'    End If
    
    
    mstrSelect = strSelect
End Sub
Private Sub SetWhere(pDateDebut As Date, pDateFin As Date, Optional pintProvenance As Integer)
    Dim strCrit As String
    
    mstrWhere = ""
    strCrit = ""
    If mblnParamProvenance Then
        strCrit = strCrit & "(ComptoirBon_Data.NoOrganisme=" & pintProvenance & ") AND "
    End If
    
    If mstrParamClient <> "" Then
        strCrit = strCrit & mstrParamClient & " AND "
    End If
    
    If mblnParamBon Then
        strCrit = strCrit & "(ComptoirBon_Data.DateBon Between #" & pDateDebut & "# And #" & pDateFin & "#) AND "
    End If
    
    If mstrParamService <> "" Then
        strCrit = strCrit & mstrParamService & " AND "
    End If

    If strCrit <> "" Then
        mstrWhere = "WHERE " & Left(strCrit, Len(strCrit) - 5)
    End If
End Sub
Private Sub SetGroupBy()
    Dim strND As String
    
    strND = "Non déterminé"
    mstrGroupBy = ""
    mstrGroupByTemp = ""
'    If Form_frm_4_1_Rapport.chkSexe.Value = -1 Then
        mstrGroupBy = mstrGroupBy & "IIF(IsNull([Sexe]) Or [Sexe]='', """ & strND & """,[Sexe]), "
        mstrGroupByTemp = mstrGroupByTemp & "CPT0001.[Sexe (M ou F)], "
'    End If
    
'    If Form_frm_4_1_Rapport.chkNationalite.Value = -1 Then
        mstrGroupBy = mstrGroupBy & "IIF(IsNull([1ERNation_18]) AND IsNull([Imigration_18]), '', IIF([1ERNation_18]<>0, 'Première nation', IIF([Imigration_18]<>0, 'Immigrant', ''))), "
        mstrGroupByTemp = mstrGroupByTemp & "CPT0001.[Nationalité], "
'    End If
    
'    If Form_frm_4_1_Rapport.chkAge.Value = -1 Then
        mstrGroupBy = mstrGroupBy & "IIf(IsNull([AGE]), """ & strND & """, IIF([AGE]>=18 And [AGE]<=30,'18-30 ans',IIf([AGE]>=31 And [AGE]<=44,'31-44 ans',IIf([AGE]>=45 And [AGE]<=64,'45-64 ans',IIf([AGE]>=65,'65 ans et plus',''))))), "
        mstrGroupByTemp = mstrGroupByTemp & "CPT0001.[TrancheAge], "
'    End If
    
'    If Form_frm_4_1_Rapport.chkMenage.Value = -1 Then
        mstrGroupBy = mstrGroupBy & "IIF(IsNull(PILO_TYPE_MENAGE.CODE), """ & strND & """, PILO_TYPE_MENAGE.DESCRIPTION), "
        mstrGroupByTemp = mstrGroupByTemp & "CPT0001.[TypMenage], "
'    End If
    
'    If Form_frm_4_1_Rapport.chkRevenu.Value = -1 Then
        mstrGroupBy = mstrGroupBy & "IIF(IsNull(PILO_TYPE_REVENU.CODE), """ & strND & """,PILO_TYPE_REVENU.DESCRIPTION), "
        mstrGroupByTemp = mstrGroupByTemp & "CPT0001.[SourceRevenu], "
'    End If
    
'    If Form_frm_4_1_Rapport.chkLogement.Value = -1 Then
        mstrGroupBy = mstrGroupBy & "IIF(IsNull(PILO_TYPE_LOGEMENT.CODE),""" & strND & """, PILO_TYPE_LOGEMENT.DESCRIPTION), "
        mstrGroupByTemp = mstrGroupByTemp & "CPT0001.[TypLogement], "
'    End If
    
    mstrGroupByTemp = Left(mstrGroupByTemp, Len(mstrGroupByTemp) - 2)
End Sub


Public Sub ProductionRapportsHebdo(pDateDebut As Date, pDateFin As Date, pintProvenance As Integer, pstrNomProvenance As String)
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim strsql As String
    Dim strNoClientAct As String
    Dim strCritere As String
    Dim i As Long
    Dim j As Long
    Dim lngNbVisite As Long
    Dim XL As Excel.Application
    Dim WB As Excel.Workbook
    Dim WKS As Excel.Worksheet
    Dim FichierExcel As Object
    
    On Error GoTo gestErr
    
    Call SetWhere(pDateDebut, pDateFin, pintProvenance)
    
    strsql = "SELECT ComptoirBon_Data.DateBon, PILO_TP_ORGANISME.[Description] AS ProvenanceBon, [ComptoirClient_Data].[NomClient] & ', ' & [ComptoirClient_Data].[PrenomClient] & IIf([ComptoirClient_Data].[Adresse] Is Null Or [ComptoirClient_Data].[Adresse]='','',' - ' & [ComptoirClient_Data].[Adresse]) AS Clients, " & _
             "ComptoirClient_Data.Ville, ComptoirClient_Data.TéléphonePrinc, Nz(ComptoirBon_Data.NB_ADULTE, 0) AS Adultes, Nz(ComptoirBon_Data.NB_ENFANT, 0) AS Enfants, Sum(ComptoirBon_Data.Panier) AS SommeDePanier, Count(ComptoirBon_Data.NoClient) AS Bon, Sum(ComptoirBon_Data.MontantDispo) AS Accordé, Sum(ComptoirBon_Data.MontantUtil) AS Utilisé, Sum(IIF(ComptoirBon_Data.MontantDispo - ComptoirBon_Data.MontantUtil < 0, Nz(ComptoirBon_Data.MontantDispo, 0), ComptoirBon_Data.MontantDispo - ComptoirBon_Data.MontantUtil)) AS Balance, Sum(ComptoirBon_Data.Surplus) AS [Surplus dépensé], Sum(ComptoirBon_Data.Pan_Noel) AS [PanierNoel] " & _
             "FROM (ComptoirClient_Data INNER JOIN ComptoirBon_Data ON ComptoirClient_Data.NoClient = ComptoirBon_Data.NoClient) INNER JOIN PILO_TP_ORGANISME ON ComptoirBon_Data.NoOrganisme = PILO_TP_ORGANISME.CODE " & _
             "" & mstrWhere & " " & _
             "GROUP BY ComptoirBon_Data.DateBon, PILO_TP_ORGANISME.Description, [ComptoirClient_Data].[NomClient] & ', ' & [ComptoirClient_Data].[PrenomClient] & IIf([ComptoirClient_Data].[Adresse] Is Null Or [ComptoirClient_Data].[Adresse]='','',' - ' & [ComptoirClient_Data].[Adresse]), ComptoirClient_Data.Ville, ComptoirClient_Data.TéléphonePrinc, Nz(ComptoirBon_Data.NB_ADULTE, 0), Nz(ComptoirBon_Data.NB_ENFANT, 0) " & _
             "ORDER BY [ComptoirClient_Data].[NomClient] & ', ' & [ComptoirClient_Data].[PrenomClient] & IIf([ComptoirClient_Data].[Adresse] Is Null Or [ComptoirClient_Data].[Adresse]='','',' - ' & [ComptoirClient_Data].[Adresse]);"
             
    Set rst = CurrentDb.OpenRecordset(strsql)
    
    If rst.RecordCount > 0 Then
        i = 5
        
        FileCopy cnsModeleRapportHebdo, cnsCheminDestRapport & pDateFin & " - " & pstrNomProvenance & " - Rapport Hebdo.xlsx"
        Set XL = New Excel.Application
'        XL.Visible = True
        Set WB = XL.Workbooks.Add(cnsCheminDestRapport & pDateFin & " - " & pstrNomProvenance & " - Rapport Hebdo.xlsx")
        WB.Activate
        Set WKS = WB.ActiveSheet
        
        ' Écrire titre du rapport selon la production
        WKS.Cells(1, 1) = "Rapport mensuel pour les bons de toutes provenances durant la période du " & Day(pDateDebut) & " " & MonthName(Month(pDateDebut)) & " au " & Day(pDateFin) & " " & MonthName(Month(pDateFin)) & " " & Year(pDateFin)
        While Not rst.EOF
            ' On alimente chacun des champs avec les bonnes valeurs; la disposition des champs du SQL doit être similaire au modèle Excel
            j = 0
            For Each fld In rst.Fields
                WKS.Cells(i, j + 1) = rst(j)
                If fld.Name = "Adultes" Or fld.Name = "Enfants" Then
                    WKS.Cells(i, j + 1) = CInt(rst(j))
                End If
                j = j + 1
            Next
            ' Incrémenter pour passer à la prochaine ligne
            i = i + 1
            ' Prendre pour acquis que le client a une deuxième visite
            lngNbVisite = lngNbVisite + 1
            
            rst.MoveNext
        Wend
        

        ' Effectuer les totaux du rapport
        WKS.Range("A" & i & ":O" & i).Borders(xlEdgeTop).LineStyle = xlContinuous
        WKS.Range("A" & i & ":O" & i).Borders(xlEdgeTop).Weight = xlMedium
        WKS.Range("E" & i) = "Total"
        WKS.Range("F" & i).FormulaR1C1 = "=SUM(R5C6:R" & i - 1 & "C6)" ' adultes
        WKS.Range("G" & i).FormulaR1C1 = "=SUM(R5C7:R" & i - 1 & "C7)" ' enfants
        WKS.Range("H" & i).FormulaR1C1 = "=SUM(R5C8:R" & i - 1 & "C8)" ' panier
        WKS.Range("I" & i).FormulaR1C1 = "=SUM(R5C9:R" & i - 1 & "C9)" ' bon
        WKS.Range("J" & i).FormulaR1C1 = "=SUM(R5C10:R" & i - 1 & "C10)" ' accordé
        WKS.Range("K" & i).FormulaR1C1 = "=SUM(R5C11:R" & i - 1 & "C11)" ' utilisé
        WKS.Range("L" & i).FormulaR1C1 = "=SUM(R5C12:R" & i - 1 & "C12)" ' Balance
        WKS.Range("M" & i).FormulaR1C1 = "=SUM(R5C13:R" & i - 1 & "C13)" ' Surplus
        WKS.Range("N" & i).FormulaR1C1 = "=SUM(R5C14:R" & i - 1 & "C14)" ' Panier Noël
        
        WKS.Range("G" & i + 1) = "Le total utilisé + le surplus est de"
        WKS.Range("J" & i + 1) = WKS.Range("K" & i) + WKS.Range("M" & i)

        ' Ajout de l'image
        If Dir(cnsImageGrenier) <> "" Then
            WKS.Pictures.Insert(cnsImageGrenier).Select
            WKS.Shapes.Range(Array("Picture 1")).Top = 0
            WKS.Shapes.Range(Array("Picture 1")).Left = 0
            WKS.Shapes.Range(Array("Picture 1")).Width = 80
            WKS.Shapes.Range(Array("Picture 1")).Height = 45
        End If

        DoCmd.Close acForm, "sfrmPatienter"
        MsgBox "Production du rapport terminée!", vbInformation, Form_frm_4_1_Rapport.Caption
        XL.Visible = True
    Else
        DoCmd.Close acForm, "sfrmPatienter"
        MsgBox "Aucun bon a été trouvé durant la période.", vbInformation, Form_frm_4_1_Rapport.Form.Caption
        Exit Sub
    End If

    Exit Sub
gestErr:
    MsgBox Err.Number & " : " & Err.Description
    DoCmd.Close acForm, "sfrmPatienter"
    Exit Sub
End Sub

Public Sub FacturationRapport(pDateDebut As Date, pDateFin As Date, pintProvenance As Integer, pstrNomProvenance As String, pblnFraisPoste As Boolean)
    Dim XL As Excel.Application
    Dim WB As Excel.Workbook
    Dim WKS As Excel.Worksheet
    Dim i As Long
    Dim j As Long
    Dim strsql As String
    Dim strCritere As String
    Dim rst As DAO.Recordset

    j = 1
    i = 6
            
    On Error GoTo gestErr
    

    Call SetWhere(pDateDebut, pDateFin, pintProvenance)

    strsql = "SELECT ComptoirBon_Data.DateBon, PILO_TP_ORGANISME.Description AS ProvenanceBon, PILO_TP_ORGANISME.RESPONSABLE, PILO_TP_ORGANISME.ADRESSE, [ComptoirClient_Data].[NomClient] & ', ' & [ComptoirClient_Data].[PrenomClient] & IIf([ComptoirClient_Data].[Adresse] Is Null Or [ComptoirClient_Data].[Adresse]='','',' - ' & [ComptoirClient_Data].[Adresse]) AS Clients, " & _
             "ComptoirClient_Data.Ville, ComptoirClient_Data.TéléphonePrinc, Nz(ComptoirBon_Data.NB_ADULTE, 0) AS Adultes, Nz(ComptoirBon_Data.NB_ENFANT, 0) AS Enfants, Sum(ComptoirBon_Data.Panier) AS SommeDePanier, Count(ComptoirBon_Data.NoClient) AS Bon, Sum(ComptoirBon_Data.Pan_Noel) AS PanierNoel, Sum(ComptoirBon_Data.MontantDispo) AS Accordé, Sum(IIf([MontantDispo]-[MontantUtil] <= 0, [MontantDispo], [MontantUtil])) AS Facturé " & _
             "FROM (ComptoirClient_Data INNER JOIN ComptoirBon_Data ON ComptoirClient_Data.NoClient = ComptoirBon_Data.NoClient) INNER JOIN PILO_TP_ORGANISME ON ComptoirBon_Data.NoOrganisme = PILO_TP_ORGANISME.CODE " & _
             "" & mstrWhere & " " & _
             "GROUP BY ComptoirBon_Data.DateBon, PILO_TP_ORGANISME.Description, PILO_TP_ORGANISME.RESPONSABLE, PILO_TP_ORGANISME.ADRESSE, [ComptoirClient_Data].[NomClient] & ', ' & [ComptoirClient_Data].[PrenomClient] & IIf([ComptoirClient_Data].[Adresse] Is Null Or [ComptoirClient_Data].[Adresse]='','',' - ' & [ComptoirClient_Data].[Adresse]), " & _
             "ComptoirClient_Data.Ville, ComptoirClient_Data.TéléphonePrinc, Nz(ComptoirBon_Data.NB_ADULTE, 0), Nz(ComptoirBon_Data.NB_ENFANT, 0) " & _
             "ORDER BY [ComptoirClient_Data].[NomClient] & ', ' & [ComptoirClient_Data].[PrenomClient] & IIf([ComptoirClient_Data].[Adresse] Is Null Or [ComptoirClient_Data].[Adresse]='','',' - ' & [ComptoirClient_Data].[Adresse]);"
    
    Set rst = CurrentDb.OpenRecordset(strsql)
    
    If rst.RecordCount > 0 Then
        FileCopy cnsModeleRapportFacturation, cnsCheminDestRapport & pDateFin & " - Facturation " & MonthName(Month(pDateFin)) & " " & Year(pDateFin) & " - " & pstrNomProvenance & ".xlsx"
        Set XL = New Excel.Application
        
        Set WB = XL.Workbooks.Add(cnsCheminDestRapport & pDateFin & " - Facturation " & MonthName(Month(pDateFin)) & " " & Year(pDateFin) & " - " & pstrNomProvenance & ".xlsx")
        WB.Activate
        Set WKS = WB.ActiveSheet
'        XL.Visible = True
        ' Écrire titre du rapport selon la production
        WKS.Range("A1") = rst("ProvenanceBon")
        WKS.Range("A2") = rst("RESPONSABLE")
        WKS.Range("A3") = rst("ADRESSE")
        WKS.Range("A4") = UCase("FACTURATION " & MonthName(Month(pDateFin)) & " " & Year(pDateFin))
        While Not rst.EOF

            WKS.Range("A" & i) = rst("DateBon")
            WKS.Range("B" & i) = rst("Clients")
            WKS.Range("C" & i) = rst("Ville")
            WKS.Range("D" & i) = rst("TéléphonePrinc")
            WKS.Range("E" & i) = CDbl(rst("Adultes"))
            WKS.Range("F" & i) = CDbl(rst("Enfants"))
            WKS.Range("G" & i) = rst("SommeDePanier")
            WKS.Range("H" & i) = rst("Bon")
            WKS.Range("I" & i) = rst("PanierNoel")
            WKS.Range("J" & i) = rst("Accordé")
            WKS.Range("K" & i) = rst("Facturé")
            ' Incrémenter pour passer à la prochaine ligne
            i = i + 1
            j = j + 1
            rst.MoveNext
        Wend

        If pblnFraisPoste Then
            WKS.Range("D" & i) = "FRAIS DE POSTE:"
            WKS.Range("K" & i) = 1.5
            WKS.Range("K" & i).Font.Bold = True
            WKS.Range("D" & i).Font.Bold = True
        End If
        
        i = i + 1
        WKS.Range("E" & i & ":K" & i).Font.Bold = True

        WKS.Range("E" & i).FormulaR1C1 = "=SUM(R6C5:R" & i - 2 & "C5)"
        WKS.Range("F" & i).FormulaR1C1 = "=SUM(R6C6:R" & i - 2 & "C6)"
        WKS.Range("G" & i).FormulaR1C1 = "=SUM(R6C7:R" & i - 2 & "C7)"
        WKS.Range("H" & i).FormulaR1C1 = "=SUM(R6C8:R" & i - 2 & "C8)"
        WKS.Range("I" & i).FormulaR1C1 = "=SUM(R6C9:R" & i - 2 & "C9)"
        WKS.Range("J" & i).FormulaR1C1 = "=SUM(R6C10:R" & i - 1 & "C10)"
        WKS.Range("K" & i).FormulaR1C1 = "=SUM(R6C11:R" & i - 1 & "C11)"
        
        With WKS.Range("A" & i & ":K" & i).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
            .Color = RGB(223, 183, 69)
        End With
        
        
        DoCmd.Close acForm, "sfrmPatienter"
        MsgBox "Production du rapport terminée!", vbInformation, Form_frm_4_1_Rapport.Caption
        XL.Visible = True
'        WB.Save
    Else
        DoCmd.Close acForm, "sfrmPatienter"
        MsgBox "Aucun bon a été trouvé durant la période.", vbInformation, Form_frm_4_1_Rapport.Form.Caption
        Exit Sub
    End If
    
    Exit Sub
gestErr:
    MsgBox "Erreur lors de la création du rapport." & vbCrLf & vbCrLf & Err.Number & " : " & Err.Description, vbCritical, "Facturation"
    DoCmd.Close acForm, "sfrmPatienter"
    Resume
    Exit Sub
End Sub

Public Function RapportStats(pDateDebut As Date, pDateFin As Date, pintProvenance As Integer) As Boolean
    Dim fso As New FileSystemObject
    Dim strsql As String
    Dim strFichier As String
    Dim strCritere As String
    Dim lngNbEnregistrement As Long
    Dim i As Long
    Dim rst As DAO.Recordset
    Dim XL As Excel.Application
    Dim WB As Excel.Workbook
    Dim WKS As Excel.Worksheet
                
    On Error GoTo gestErr
    
    mstrTabColonneExcel = Split("I,J,K,L,M,N", ",")
    
    Call SetSelect
    Call SetGroupBy
    Call SetWhere(pDateDebut, pDateFin, pintProvenance)
    
    strsql = "SELECT " & mstrSelect & "ComptoirClient_Data.NoClient, " & _
             "IIf([DateInscription] Between #" & pDateDebut & "# And #" & pDateFin & "#,1,0) AS Indicateur_NouvelleFamil, " & _
             "Sum(ComptoirBon_Data.Panier) AS SommeDePanier, Sum(IIf([DateBon] Between #" & pDateDebut & "# And #" & pDateFin & "#,[MontantDispo],0)) AS MontantAutorise, " & _
             "Sum(IIf([DateBon] Between #" & pDateDebut & "# And #" & pDateFin & "#,[MontantUtil],0)) AS MontantUtilise, " & _
             "Sum(IIf([DateBon] Between #" & pDateDebut & "# And #" & pDateFin & "#,[MontantBalance],0)) AS MontantBal, " & _
             "Sum(IIf([DateBon] Between #" & pDateDebut & "# And #" & pDateFin & "#,[NB_ADULTE],0)) AS NbAdulte, " & _
             "Sum(IIf([DateBon] Between #" & pDateDebut & "# And #" & pDateFin & "#,[NB_ENFANT],0)) AS NbEnfant " & _
             "INTO CPT0001 IN """ & cnsCheminBDTemp & """ " & _
             "FROM (((ComptoirClient_Data LEFT JOIN PILO_TYPE_LOGEMENT ON ComptoirClient_Data.TypeLogement = PILO_TYPE_LOGEMENT.CODE) LEFT JOIN PILO_TYPE_MENAGE ON ComptoirClient_Data.TypeMenage = PILO_TYPE_MENAGE.CODE) LEFT JOIN PILO_TYPE_REVENU ON ComptoirClient_Data.SourceRevenu = PILO_TYPE_REVENU.CODE) LEFT JOIN ComptoirBon_Data ON ComptoirClient_Data.NoClient = ComptoirBon_Data.NoClient " & _
             "" & mstrWhere & " " & _
             "GROUP BY " & mstrGroupBy & "ComptoirClient_Data.NoClient, IIf([DateInscription] Between #" & pDateDebut & "# And #" & pDateFin & "#,1,0);"
    
    DoCmd.RunSQL strsql
    
    strsql = "SELECT  " & mstrSelectTemp & " Count(CPT0001.NoClient) AS [Nombre de famille], Sum(CPT0001.Indicateur_NouvelleFamil) AS [Nombre de nouvelle famille], Sum(CPT0001.SommeDePanier) AS [Nombre de panier], Sum(CPT0001.MontantAutorise) AS [Montant accordé ($)], Sum(CPT0001.MontantUtilise) AS [Montant utilisé ($)], Sum(CPT0001.MontantBal) AS [Balance ($)], Sum(CPT0001.NbAdulte) AS [Nombre d'adulte désservie], Sum(CPT0001.NbEnfant) AS [Nombre d'enfant désservie] " & _
             "INTO CPTSTATS IN """ & cnsCheminBDStats & """ " & _
             "FROM CPT0001 " & _
             "GROUP BY " & mstrGroupByTemp & ";"
    DoCmd.RunSQL strsql
    
    Call remplacerNulleOuVide("CPTSTATS")
    
    strFichier = cnsCheminDestRapport & "Rapports de statistiques au " & pDateFin & ".xlsb"
    
    fso.CopyFile cnsModeleRapportStats, strFichier
    lngNbEnregistrement = DCount("[Nombre de nouvelle famille]", "CPTSTATS")
    ' Début de la création du rapport
    Set XL = New Excel.Application

    Set WB = XL.Workbooks.Open(strFichier)
    WB.Activate
    
    Set WKS = WB.Worksheets("Statistiques")
    WKS.Select
    WKS.Rows("1:4").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Création des en-têtes
    WKS.Range("A1").FormulaR1C1 = "Rapport de statistiques"
    WKS.Range("A2").FormulaR1C1 = "Comptoir Le Grenier"
    WKS.Range("A3").FormulaR1C1 = "Pour la période du " & pDateDebut & " au " & pDateFin

    WKS.Range("A1:A3").Font.Name = "Arial"
    WKS.Range("A1:A3").Font.Bold = True
    WKS.Range("A1:A3").Font.Size = 14

    WKS.Range("A3").Font.Size = 12

    ' Défini M1 selon nombre de colonne
    WKS.Range("A1:" & mstrTabColonneExcel(nombreColonne - 1) & "1").Merge
    WKS.Range("A1:" & mstrTabColonneExcel(nombreColonne - 1) & "1").HorizontalAlignment = xlCenter
    WKS.Range("A2:" & mstrTabColonneExcel(nombreColonne - 1) & "2").Merge
    WKS.Range("A2:" & mstrTabColonneExcel(nombreColonne - 1) & "2").HorizontalAlignment = xlCenter
    WKS.Range("A3:" & mstrTabColonneExcel(nombreColonne - 1) & "3").Merge
    WKS.Range("A3:" & mstrTabColonneExcel(nombreColonne - 1) & "3").HorizontalAlignment = xlCenter
    
    WKS.Columns("A:" & mstrTabColonneExcel(nombreColonne - 1) & "").ColumnWidth = 13
    WKS.Range("A5:" & mstrTabColonneExcel(nombreColonne - 1) & "" & lngNbEnregistrement + 4).WrapText = True

'    WKS.Range("F4").FormulaR1C1 = "=SUBTOTAL(9,R6C6:R" & lngNbEnregistrement + 5 & "C6)"
'    WKS.Range("G4").FormulaR1C1 = "=SUBTOTAL(9,R6C7:R" & lngNbEnregistrement + 5 & "C7)"
'    WKS.Range("H4").FormulaR1C1 = "=SUBTOTAL(9,R6C8:R" & lngNbEnregistrement + 5 & "C8)"
'    WKS.Range("I4").FormulaR1C1 = "=SUBTOTAL(9,R6C9:R" & lngNbEnregistrement + 5 & "C9)"
'    WKS.Range("J4").FormulaR1C1 = "=SUBTOTAL(9,R6C10:R" & lngNbEnregistrement + 5 & "C10)"
'    WKS.Range("K4").FormulaR1C1 = "=SUBTOTAL(9,R6C11:R" & lngNbEnregistrement + 5 & "C11)"
'    WKS.Range("L4").FormulaR1C1 = "=SUBTOTAL(9,R6C12:R" & lngNbEnregistrement + 5 & "C12)"
'    WKS.Range("M4").FormulaR1C1 = "=SUBTOTAL(9,R6C13:R" & lngNbEnregistrement + 5 & "C13)"
'    WKS.Columns("I:K").Style = "Currency"

    
    ' Ajout de l'image
    If Dir(cnsImageGrenier) <> "" Then
        WKS.Pictures.Insert(cnsImageGrenier).Select
        WKS.Shapes.Range(Array("Picture 1")).Top = 0
        WKS.Shapes.Range(Array("Picture 1")).Left = 0
        WKS.Shapes.Range(Array("Picture 1")).Width = 141.73
        WKS.Shapes.Range(Array("Picture 1")).Height = 68.03
    End If
    
    WB.RefreshAll
    WKS.Range("A1").Select
    
    ' Ajustement mise en page pour impression
    WKS.PageSetup.Orientation = xlLandscape
    XL.ActiveWindow.View = xlPageBreakPreview
    WKS.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    XL.ActiveWindow.View = xlNormalView
    XL.ActiveWindow.DisplayGridlines = False
    WKS.PageSetup.RightHeader = "&D"
    WKS.PageSetup.RightFooter = "Page &P"
    WKS.PageSetup.PrintTitleRows = "$1:$5"
    
    XL.Visible = True

    Exit Function
gestErr:
    Select Case Err.Number
        Case 3027
            MsgBox "Le rapport est déjà ouvert. Veuillez le fermer et relancer la production.", vbInformation, "Rapports stats"
        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbCritical, "Erreur rapports stats!"
            Resume
            MsgBox "Veuillez fermer l'application et tenter la création à nouveau!", vbInformation, "Rapports stats"
    End Select
End Function

Public Sub ExportListeClient()
    Dim XL As Excel.Application
    Dim WB As Excel.Workbook
    Dim WKS As Excel.Worksheet
    Dim strsql As String
    Dim strFichier As String
    Dim nombreClient As Integer
    On Error GoTo gestErr
    
    nombreClient = DCount("NoClient", "reqListeClientMenage")
    
    strFichier = cnsCheminDestRapport & Date & "_Liste des clients.xlsx"
    
    strsql = "SELECT [NomClient] & ', ' & [PrenomClient] AS Client, ComptoirClient_Data.Sexe, ComptoirClient_Data.TéléphonePrinc AS [Téléphone principale], ComptoirClient_Data.TéléphoneSec AS [Téléphone secondaire], ComptoirClient_Data.Adresse, ComptoirClient_Data.Ville, ComptoirClient_Data.DateNaissance AS [Date de naissance], [ComptoirMenage_Data].[Nom] & ', ' & [ComptoirMenage_Data].[Prénom] AS [Nom du ménage], ComptoirMenage_Data.Sexe AS [Sexe du ménage], ComptoirMenage_Data.DateNaissance AS [Date de naissance du ménage], ComptoirMenage_Data.Lien " & _
             "INTO CPT0002 IN """ & cnsCheminBDTemp & """ " & _
             "FROM ComptoirClient_Data LEFT JOIN ComptoirMenage_Data ON ComptoirClient_Data.NoClient = ComptoirMenage_Data.NoClient;"

    DoCmd.RunSQL strsql
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "CPT0002", strFichier, True
    
    Set XL = New Excel.Application

    Set WB = XL.Workbooks.Open(strFichier)
    Set WKS = XL.Worksheets(1)
    
    WB.Activate
    WKS.Columns("A:K").EntireColumn.AutoFit
    WKS.Range("A1:K1").AutoFilter
    WKS.Range("A1:K" & nombreClient + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    WKS.Range("A1:K" & nombreClient + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    WKS.Range("A1:K" & nombreClient + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    WKS.Range("A1:K" & nombreClient + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    WKS.Range("A1:K" & nombreClient + 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    XL.Visible = True
    
    Exit Sub
gestErr:
    MsgBox Err.Number & " - Erreur, veuillez ressayer à nouveau"
    If IsObject(XL) = False Then
        XL.Quit
        Set XL = Nothing
    End If
End Sub
