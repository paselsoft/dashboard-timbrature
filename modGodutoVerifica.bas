Attribute VB_Name = "modGodutoVerifica"
Option Explicit

'=============================================================================
' Modulo  : modGodutoVerifica  (macro "una tantum")
' Macro   : AggiungiPianificatoEVerifica
'
' 1) DASHBOARD: GODUTO E PIANIFICATO
'    - colonna J "Goduto"      = assenze fino a OGGI;
'    - nuova colonna L "Pianificato" = assenze nei giorni FUTURI;
'    - colonna K "Residuo"     = Plafond - Goduto - Pianificato (finale);
'    - lo stato delle ferie anno prec. ("da fruire entro / scadute") passa
'      da L7 a M7.
'    Nei mesi: colonna S (giorno futuro) e celle T11:T16 (totali futuri),
'    invisibili e fuori stampa.
'
' 2) VERIFICA CON IL CARTELLINO (ogni mese, riepilogo righe 23-27)
'    - Q24 EXT, Q25 FLE, Q26 Buoni: si scrivono i totali del cartellino;
'    - le celle diventano rosse se non coincidono con quelle del file;
'    - Q27 "OK" (verde) o "diverso" (rosso); "diverso" conta come anomalia.
'
' Prima delle modifiche salva un backup. Va eseguita una volta in OGNI file.
' Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const COL_ROSSO_FONDO As Long = 13551615     ' RGB(255,199,206)
Private Const COL_ROSSO_TESTO As Long = 393372       ' RGB(156,0,6)
Private Const COL_ARANCIO_FONDO As Long = 10284031   ' RGB(255,235,156)
Private Const COL_ARANCIO_TESTO As Long = 22428      ' RGB(156,87,0)
Private Const COL_VERDE_FONDO As Long = 13561798     ' RGB(198,239,206)
Private Const COL_VERDE_TESTO As Long = 24832        ' RGB(0,97,0)
Private Const COL_INPUT As Long = 13431551           ' RGB(255,242,204)

Public Sub AggiungiPianificatoEVerifica()

    Dim vMesi As Variant, v As Variant
    Dim wsDash As Worksheet, wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim F_S As String
    Dim F_M7 As String
    Dim F_I5 As String
    Dim F_Q27 As String
    Dim F_Q19 As String
    Dim R_Q24 As String
    Dim R_Q25 As String
    Dim R_Q26 As String
    Dim R_Q27_OK As String
    Dim R_Q27_KO As String
    Dim R_M7_ARANCIO As String
    Dim R_M7_ROSSO As String
    Dim F_T11 As String
    Dim F_T12 As String
    Dim F_T13 As String
    Dim F_T14 As String
    Dim F_T15 As String
    Dim F_T16 As String
    Dim sR24 As String, sR25 As String, sR26 As String, sOK As String, sKO As String
    Dim sAr As String, sRo As String
    Dim vRighe As Variant, vQ As Variant, vT As Variant
    Dim i As Long, r As Long
    Dim sBackup As String, sPrima As String, sDopo As String
    Dim eraProtetto As Boolean, bModificato As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim xlCalc As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set wsDash = OttieniFoglio("Dashboard")
    Set wsImp = OttieniFoglio("Impostazioni")
    If wsDash Is Nothing Or wsImp Is Nothing Then
        MsgBox "Servono i fogli Dashboard e Impostazioni.", vbExclamation, "Pianificato e verifica"
        Exit Sub
    End If
    If Not EsisteNome("FEP_Scadenza_Giorno") Then
        MsgBox "Esegui prima AggiornaFerieFEPeLimiteEXT (modulo modFerieFEPeLimite).", _
               vbExclamation, "Pianificato e verifica"
        Exit Sub
    End If
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Salva il file in una cartella del PC prima di eseguire la macro.", _
               vbExclamation, "Pianificato e verifica"
        Exit Sub
    End If
    If MsgBox("Verranno aggiunti in Dashboard la colonna 'Pianificato' e, nei mesi, " & _
              "la verifica con il cartellino." & vbCrLf & vbCrLf & _
              "Prima verra' salvata una copia di backup. Procedere?", _
              vbQuestion + vbYesNo, "Pianificato e verifica") <> vbYes Then Exit Sub

    sBackup = ThisWorkbook.Path & Application.PathSeparator & _
              NomeSenzaEstensione(ThisWorkbook.Name) & "_backup_" & _
              Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsm"
    On Error Resume Next
    ThisWorkbook.SaveCopyAs sBackup
    If Err.Number <> 0 Then
        MsgBox "Impossibile creare il backup: " & Err.Description, vbCritical, "Pianificato e verifica"
        Exit Sub
    End If
    On Error GoTo 0

    F_S = "=IF(ISNUMBER($B11),IF(DATE(Dashboard!$A$1,$A$1,$B11)>TODAY(),1,0),0)"
    F_M7 = "=IF(K7<=0,"""",IF(TODAY()<=DATE($A$1,FEP_Scadenza_Mese,FEP_Scadenza_Giorno)," & _
        """da fruire entro il ""&RIGHT(""0""&FEP_Scadenza_Giorno," & _
        "2)&""/""&RIGHT(""0""&FEP_Scadenza_Mese,2)&""/""&$A$1," & _
        """scadute il ""&RIGHT(""0""&FEP_Scadenza_Giorno,2)&""/""&RIGHT(""0""&FEP_Scadenza_Mese," & _
        "2)&""/""&$A$1))"
    F_I5 = "=IF(SUM(Gennaio:Dicembre!$Q$19)+(K7<0)=0,""nessuna"",SUM(Gennaio:Dicembre!$Q$19)+(K7<0))"
    F_Q27 = "=IF(COUNT($Q$24:$Q$26)=0,"""",IF(AND(OR(NOT(ISNUMBER($Q$24)),ROUND($Q$24*1440," & _
        "0)=ROUND($E$4*1440,0)),OR(NOT(ISNUMBER($Q$25)),ROUND($Q$25*1440,0)=ROUND($E$5*1440,0))," & _
        "OR(NOT(ISNUMBER($Q$26)),$Q$26=$Q$17)),""OK"",""diverso""))"
    F_Q19 = "=SUM($O$11:$O$41)+(COUNTIF($N$11:$N$41,""L. 104"")>Max_L104_Mese)+($Q$27=""diverso"")"
    R_Q24 = "=AND(ISNUMBER($Q$24),ROUND($Q$24*1440,0)<>ROUND($E$4*1440,0))"
    R_Q25 = "=AND(ISNUMBER($Q$25),ROUND($Q$25*1440,0)<>ROUND($E$5*1440,0))"
    R_Q26 = "=AND(ISNUMBER($Q$26),$Q$26<>$Q$17)"
    R_Q27_OK = "=$Q$27=""OK"""
    R_Q27_KO = "=$Q$27=""diverso"""
    R_M7_ARANCIO = "=AND($K$7>0,TODAY()<=DATE($A$1,FEP_Scadenza_Mese,FEP_Scadenza_Giorno))"
    R_M7_ROSSO = "=AND($K$7>0,TODAY()>DATE($A$1,FEP_Scadenza_Mese,FEP_Scadenza_Giorno))"
    F_T11 = "=COUNTIFS($N$11:$N$41,""Ferie"",$S$11:$S$41,1)"
    F_T12 = "=COUNTIFS($N$11:$N$41,""Ferie anno prec."",$S$11:$S$41,1)"
    F_T13 = "=COUNTIFS($N$11:$N$41,""L. 104"",$S$11:$S$41,1)"
    F_T14 = "=SUMIFS($E$11:$E$41,$N$11:$N$41,""Art. 23 (Visite)"",$S$11:$S$41,1)"
    F_T15 = "=SUMIFS($E$11:$E$41,$N$11:$N$41,""Art. 25 (Motivi pers.)"",$S$11:$S$41,1)"
    F_T16 = "=SUMIFS($E$11:$E$41,$N$11:$N$41,""Art. 34 (A recupero)"",$S$11:$S$41,1)"

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    xlCalc = Application.Calculation

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculate
    sPrima = Foto(wsDash, vMesi)
    Application.Calculation = xlCalculationManual
    bModificato = True

    '--- Regole di colore nella lingua di Excel (Impostazioni!F2) -------------
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect
    sR24 = FormulaLocale(wsImp, R_Q24)
    sR25 = FormulaLocale(wsImp, R_Q25)
    sR26 = FormulaLocale(wsImp, R_Q26)
    sOK = FormulaLocale(wsImp, R_Q27_OK)
    sKO = FormulaLocale(wsImp, R_Q27_KO)
    sAr = FormulaLocale(wsImp, R_M7_ARANCIO)
    sRo = FormulaLocale(wsImp, R_M7_ROSSO)
    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- Fogli mensili -----------------------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Pianificato e verifica: " & ws.Name & "..."
            DoEvents
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ws.Activate

            ' S: 1 se il giorno e' futuro; T11:T16 totali dei giorni futuri
            With ws.Range("S11:S41")
                .Formula = F_S
                .NumberFormat = ";;;"
            End With
            ws.Range("T11").Formula = F_T11
            ws.Range("T12").Formula = F_T12
            ws.Range("T13").Formula = F_T13
            ws.Range("T14").Formula = F_T14
            ws.Range("T15").Formula = F_T15
            ws.Range("T16").Formula = F_T16
            ws.Range("T11:T16").NumberFormat = ";;;"

            ' Verifica con il cartellino: righe 23-27 del riepilogo
            ws.Range("P10:Q10").Copy
            ws.Range("P23:Q23").PasteSpecial Paste:=xlPasteFormats
            ws.Range("P18:Q18").Copy
            ws.Range("P24:Q27").PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            ws.Range("P23").Value = "Verifica con il cartellino"
            ws.Range("Q23").Value = "Portale"
            ws.Range("P24").Value = "EXT del cartellino"
            ws.Range("P25").Value = "FLE del cartellino"
            ws.Range("P26").Value = "Buoni del cartellino"
            ws.Range("P27").Value = "Esito"
            With ws.Range("Q24:Q26")
                .Interior.Color = COL_INPUT
                .HorizontalAlignment = xlRight
                .Locked = False
            End With
            ws.Range("Q24:Q25").NumberFormat = "[h]:mm"
            ws.Range("Q26").NumberFormat = "0"
            With ws.Range("Q24:Q25").Validation
                .Delete
                .Add Type:=xlValidateTime, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:="0", Formula2:="=999/1000"
                .ErrorMessage = "Inserisci un totale in ore e minuti (es. 1:14)."
            End With
            With ws.Range("Q26").Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:="0", Formula2:="31"
                .ErrorMessage = "Inserisci il numero di buoni pasto (0-31)."
            End With
            With ws.Range("Q27")
                .Formula = F_Q27
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
            ws.Range("Q19").Formula = F_Q19

            ImpostaRegola ws.Range("Q24"), sR24, COL_ROSSO_FONDO, COL_ROSSO_TESTO
            ImpostaRegola ws.Range("Q25"), sR25, COL_ROSSO_FONDO, COL_ROSSO_TESTO
            ImpostaRegola ws.Range("Q26"), sR26, COL_ROSSO_FONDO, COL_ROSSO_TESTO
            ImpostaRegola ws.Range("Q27"), sOK, COL_VERDE_FONDO, COL_VERDE_TESTO
            ImpostaRegola ws.Range("Q27"), sKO, COL_ROSSO_FONDO, COL_ROSSO_TESTO

            ws.Range("A1").Select
            If eraProtetto Then ws.Protect
            eraProtetto = False
        End If
    Next v

    '--- Dashboard -----------------------------------------------------------------
    Set ws = wsDash
    sFoglioErr = wsDash.Name
    eraProtetto = wsDash.ProtectContents
    If eraProtetto Then wsDash.Unprotect
    wsDash.Activate

    ' stato ferie anno prec.: da L7 a M7
    wsDash.Range("L7").FormatConditions.Delete
    wsDash.Range("L7").ClearContents
    With wsDash.Range("M7")
        .Formula = F_M7
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    wsDash.Range("M7").FormatConditions.Delete
    ImpostaRegola wsDash.Range("M7"), sAr, COL_ARANCIO_FONDO, COL_ARANCIO_TESTO
    ImpostaRegola wsDash.Range("M7"), sRo, COL_ROSSO_FONDO, COL_ROSSO_TESTO

    ' colonna L "Pianificato", J = goduto fino a oggi, K = residuo finale
    wsDash.Range("K6").Copy
    wsDash.Range("L6").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    wsDash.Range("L6").Value = "Pianificato"
    vRighe = Array(7, 8, 9, 10, 11, 12)
    vQ = Array("Q20", "Q11", "Q12", "Q13", "Q14", "Q15")
    vT = Array("T12", "T11", "T13", "T14", "T15", "T16")
    For i = 0 To 5
        r = vRighe(i)
        wsDash.Range("J" & r).Formula = "=SUM(Gennaio:Dicembre!" & vQ(i) & ")-SUM(Gennaio:Dicembre!" & vT(i) & ")"
        wsDash.Range("L" & r).Formula = "=IF(SUM(Gennaio:Dicembre!" & vT(i) & ")=0,"""",SUM(Gennaio:Dicembre!" & vT(i) & "))"
        wsDash.Range("K" & r).Formula = "=I" & r & "-J" & r & "-N(L" & r & ")"
        wsDash.Range("L" & r).NumberFormat = wsDash.Range("J" & r).NumberFormat
    Next i
    With wsDash.Range("L7:L14")
        .HorizontalAlignment = xlCenter
        .Font.Italic = True
    End With
    With wsDash.Range("L6:L14").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    If wsDash.Columns("L").ColumnWidth < 11 Then wsDash.Columns("L").ColumnWidth = 11
    wsDash.Range("I5").Formula = F_I5

    wsDash.Range("A1").Select
    If eraProtetto Then wsDash.Protect
    eraProtetto = False
    sFoglioErr = ""

    Application.Calculation = xlCalc
    Application.Calculate
    sDopo = Foto(wsDash, vMesi)
    bModificato = False

Uscita:
    On Error Resume Next
    Application.CutCopyMode = False
    If eraProtetto Then ws.Protect
    Application.Calculation = xlCalc
    If Not shIniziale Is Nothing Then shIniziale.Activate
    Application.StatusBar = False
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, "") & vbCrLf & vbCrLf & _
               IIf(bModificato, "Chiudi il file SENZA salvare e riapri il backup:" & vbCrLf & sBackup, ""), _
               vbCritical, "Pianificato e verifica"
    Else
        MsgBox "Aggiornamento completato." & vbCrLf & vbCrLf & _
               "Prima: " & sPrima & vbCrLf & "Dopo:  " & sDopo & vbCrLf & vbCrLf & _
               "(il residuo deve restare uguale; il goduto ora esclude i giorni futuri)" & _
               vbCrLf & vbCrLf & "Backup: " & sBackup & vbCrLf & "Salva il file.", _
               vbInformation, "Pianificato e verifica"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


' Residui e saldi della Dashboard, per il confronto prima/dopo
Private Function Foto(ByVal wsDash As Worksheet, ByVal vMesi As Variant) As String
    Dim r As Long, s As String
    For r = 7 To 12
        s = s & wsDash.Range("K" & r).Text & " "
    Next r
    Foto = "residui " & s & "| saldo " & wsDash.Range("E22").Text
End Function


Private Sub ImpostaRegola(ByVal rng As Range, ByVal sFormula As String, _
                          ByVal lFondo As Long, ByVal lTesto As Long)
    Dim i As Long
    Dim fcs As FormatConditions
    Dim sF As String, sA As String

    rng.Cells(1, 1).Select
    Set fcs = rng.FormatConditions
    For i = fcs.Count To 1 Step -1
        sF = "": sA = ""
        On Error Resume Next
        sA = fcs(i).AppliesTo.Address
        sF = fcs(i).Formula1
        On Error GoTo 0
        If StrComp(sA, rng.Address, vbTextCompare) = 0 And _
           StrComp(Replace(sF, " ", ""), Replace(sFormula, " ", ""), vbTextCompare) = 0 Then
            fcs(i).Delete
        End If
    Next i

    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
        .Interior.Color = lFondo
        .Font.Color = lTesto
        .Font.Bold = True
        .StopIfTrue = False
    End With
End Sub


Private Function FormulaLocale(ByVal ws As Worksheet, ByVal sFormula As String) As String
    Dim c As Range
    Set c = ws.Range("F2")
    If Len(c.Formula) > 0 Then
        Err.Raise vbObjectError + 995, , "La cella Impostazioni!F2 deve essere vuota."
    End If
    c.Formula = sFormula
    FormulaLocale = c.FormulaLocal
    c.ClearContents
End Function


Private Function EsisteNome(ByVal sNome As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(sNome)
    EsisteNome = Not n Is Nothing
    On Error GoTo 0
End Function


Private Function OttieniFoglio(ByVal sNome As String) As Worksheet
    On Error Resume Next
    Set OttieniFoglio = ThisWorkbook.Worksheets(sNome)
    On Error GoTo 0
End Function


Private Function NomeSenzaEstensione(ByVal sNome As String) As String
    Dim p As Long
    p = InStrRev(sNome, ".")
    If p > 0 Then NomeSenzaEstensione = Left$(sNome, p - 1) Else NomeSenzaEstensione = sNome
End Function
