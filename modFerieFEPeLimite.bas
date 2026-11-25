Attribute VB_Name = "modFerieFEPeLimite"
Option Explicit

'=============================================================================
' Modulo  : modFerieFEPeLimite  (macro "una tantum")
' Macro   : AggiornaFerieFEPeLimiteEXT
'
' 1) FERIE ANNO PRECEDENTE COME SUL PORTALE (codice FEP)
'    - nuovo Tipo Assenza "Ferie anno prec." nel menu a tendina dei mesi
'      (copre l'orario previsto come le Ferie);
'    - riepilogo mese, riga 20: "Ferie anno prec. (gg)";
'    - Dashboard: Ferie Anno precedente = giorni "Ferie anno prec."
'                 Ferie Anno corrente   = giorni "Ferie";
'    - scadenza (giorno/mese) in Impostazioni, righe 28-29 (30/06);
'      "Ferie anno prec." dopo la scadenza e uso oltre il residuo sono
'      anomalie.
'
' 2) EXTRAPRESENZA MASSIMA NEI GIORNI SENZA RIENTRO
'    - Impostazioni riga 27 (0:30): nei giorni senza Entr 2/Usc 2 l'EXT
'      oltre il limite non viene conteggiato (come le "Ore non autorizzate"
'      del portale); nei giorni con rientro nessun limite;
'    - colonna R (invisibile, fuori stampa): EXT prima del limite;
'    - riepilogo mese, riga 21: "Ore non autorizzate";
'    - Piu'(EXT) in arancione nei giorni in cui il limite e' intervenuto.
'
' Prima delle modifiche salva un backup; alla fine confronta i totali.
' Va eseguita una volta in OGNI file. Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const FEP As String = "Ferie anno prec."
Private Const COL_ARANCIO_FONDO As Long = 10284031   ' RGB(255,235,156)
Private Const COL_ARANCIO_TESTO As Long = 22428      ' RGB(156,87,0)
Private Const COL_ROSSO_FONDO As Long = 13551615     ' RGB(255,199,206)
Private Const COL_ROSSO_TESTO As Long = 393372       ' RGB(156,0,6)

Public Sub AggiornaFerieFEPeLimiteEXT()

    Dim vMesi As Variant, v As Variant
    Dim wsDash As Worksheet, wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim F_R As String, F_K As String, F_L As String, F_O As String
    Dim F_Q20 As String, F_Q21 As String
    Dim F_J7 As String, F_J8 As String, F_L7 As String, F_I5 As String
    Dim R_K7_ARANCIO As String, R_K7_ROSSO As String, R_K_LIMITE As String
    Dim sArancio As String, sRosso As String, sLimite As String
    Dim sPrima As String, sDopo As String, sBackup As String
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim xlCalc As Long
    Dim bModificato As Boolean
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set wsDash = OttieniFoglio("Dashboard")
    Set wsImp = OttieniFoglio("Impostazioni")
    If wsDash Is Nothing Or wsImp Is Nothing Then
        MsgBox "Servono i fogli Dashboard e Impostazioni.", vbExclamation, "Ferie FEP e limite EXT"
        Exit Sub
    End If
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Salva il file in una cartella del PC prima di eseguire la macro.", _
               vbExclamation, "Ferie FEP e limite EXT"
        Exit Sub
    End If
    If MsgBox("Verranno aggiunti:" & vbCrLf & _
              "  - il tipo assenza '" & FEP & "' (ferie anno precedente, codice FEP)" & vbCrLf & _
              "  - il limite di extrapresenza nei giorni senza rientro" & vbCrLf & vbCrLf & _
              "Prima verra' salvata una copia di backup. Procedere?", _
              vbQuestion + vbYesNo, "Ferie FEP e limite EXT") <> vbYes Then Exit Sub

    sBackup = ThisWorkbook.Path & Application.PathSeparator & _
              NomeSenzaEstensione(ThisWorkbook.Name) & "_backup_" & _
              Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsm"
    On Error Resume Next
    ThisWorkbook.SaveCopyAs sBackup
    If Err.Number <> 0 Then
        MsgBox "Impossibile creare il backup: " & Err.Description, vbCritical, "Ferie FEP e limite EXT"
        Exit Sub
    End If
    On Error GoTo 0

    ' --- Formule (sintassi inglese, riferite alla riga 11) ---
    F_R = "=IF(AND(COUNT($F11:$I11)=0,IF(ISNUMBER($E11),$E11,IF(OR($N11=""Ferie"",$N11=""Malattia""," & _
        "$N11=""L. 104"",$N11=""Ferie anno prec.""),IF(ISNUMBER($D11),$D11,0),0))=0),0,MAX(0," & _
        "ROUND(($J11+IF(ISNUMBER($E11),$E11,IF(OR($N11=""Ferie"",$N11=""Malattia""," & _
        "$N11=""L. 104"",$N11=""Ferie anno prec.""),IF(ISNUMBER($D11),$D11,0)," & _
        "0))-IF(ISNUMBER($D11),$D11,0))*1440,0)/1440))"
    F_K = "=IF(AND(N(EXT_Max_Senza_Rientro)>0,COUNT($H11:$I11)=0),MIN($R11,EXT_Max_Senza_Rientro)," & _
        "$R11)"
    F_L = "=IF(AND(COUNT($F11:$I11)=0,IF(ISNUMBER($E11),$E11,IF(OR($N11=""Ferie"",$N11=""Malattia""," & _
        "$N11=""L. 104"",$N11=""Ferie anno prec.""),IF(ISNUMBER($D11),$D11,0),0))=0),0,MAX(0," & _
        "ROUND((IF(ISNUMBER($D11),$D11,0)-$J11-IF(ISNUMBER($E11),$E11,IF(OR($N11=""Ferie""," & _
        "$N11=""Malattia"",$N11=""L. 104"",$N11=""Ferie anno prec.""),IF(ISNUMBER($D11),$D11,0)," & _
        "0)))*1440,0)/1440))"
    F_O = "=IF(ISNUMBER($B11),AND(ISNUMBER($E11),$E11>0,$N11="""")+AND(OR($N11=""Art. 23 (Visite)""," & _
        "$N11=""Art. 25 (Motivi pers.)"",$N11=""Art. 34 (A recupero)""),OR(NOT(ISNUMBER($E11))," & _
        "$E11=0))+OR(AND(ISNUMBER($F11),ISNUMBER($G11),$G11<$F11),AND(ISNUMBER($G11)," & _
        "ISNUMBER($H11),$H11<$G11),AND(ISNUMBER($H11),ISNUMBER($I11)," & _
        "$I11<$H11))+AND(ISNUMBER($G11),ISNUMBER($H11),$H11>=$G11,ROUND(($H11-$G11)*1440," & _
        "0)<ROUND(Pausa_Minima*1440,0))+AND(ISODD(COUNT($F11:$I11)),DATE(Dashboard!$A$1,$A$1," & _
        "$B11)<TODAY())+AND(ISNUMBER($D11),$D11>0,DATE(Dashboard!$A$1,$A$1,$B11)<TODAY()," & _
        "COUNT($E11:$I11)=0,$N11="""")+AND($N11=""Ferie anno prec."",DATE(Dashboard!$A$1,$A$1," & _
        "$B11)>DATE(Dashboard!$A$1,FEP_Scadenza_Mese,FEP_Scadenza_Giorno)),0)"
    F_Q20 = "=COUNTIF($N$11:$N$41,""Ferie anno prec."")"
    F_Q21 = "=SUM($R$11:$R$41)-SUM($K$11:$K$41)"
    F_J7 = "=SUM(Gennaio:Dicembre!Q20)"
    F_J8 = "=SUM(Gennaio:Dicembre!Q11)"
    F_L7 = "=IF(I7-J7<=0,"""",IF(TODAY()<=DATE($A$1,FEP_Scadenza_Mese,FEP_Scadenza_Giorno)," & _
        """da fruire entro il ""&RIGHT(""0""&FEP_Scadenza_Giorno," & _
        "2)&""/""&RIGHT(""0""&FEP_Scadenza_Mese,2)&""/""&$A$1," & _
        """scadute il ""&RIGHT(""0""&FEP_Scadenza_Giorno,2)&""/""&RIGHT(""0""&FEP_Scadenza_Mese," & _
        "2)&""/""&$A$1))"
    F_I5 = "=IF(SUM(Gennaio:Dicembre!$Q$19)+(J7>I7)=0,""nessuna""," & _
        "SUM(Gennaio:Dicembre!$Q$19)+(J7>I7))"
    R_K7_ARANCIO = "=AND($K$7>0,TODAY()<=DATE($A$1,FEP_Scadenza_Mese,FEP_Scadenza_Giorno))"
    R_K7_ROSSO = "=AND($K$7>0,TODAY()>DATE($A$1,FEP_Scadenza_Mese,FEP_Scadenza_Giorno))"
    R_K_LIMITE = "=$R11>$K11"

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    xlCalc = Application.Calculation

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculate
    sPrima = Foto(vMesi)
    Application.Calculation = xlCalculationManual
    bModificato = True

    '--- 1) Impostazioni: righe 27-29 e nomi -----------------------------------
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect
    wsImp.Activate
    If Not EsisteNome("EXT_Max_Senza_Rientro") Then
        wsImp.Range("B26:D26").Copy
        wsImp.Range("B27:D29").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        wsImp.Range("B27").Value = "Extrapresenza massima senza rientro"
        wsImp.Range("C27").Value = TimeSerial(0, 30, 0)
        wsImp.Range("C27").NumberFormat = "[h]:mm"
        wsImp.Range("D27").Value = "Nei giorni senza Entr 2/Usc 2 l'EXT oltre il limite non conta (0:00 = nessun limite)."
        wsImp.Range("B28").Value = "Scadenza ferie anno prec. - giorno"
        wsImp.Range("C28").Value = 30
        wsImp.Range("C28").NumberFormat = "0"
        wsImp.Range("D28").Value = "Le 'Ferie anno prec.' vanno fruite entro questa data dell'anno."
        wsImp.Range("B29").Value = "Scadenza ferie anno prec. - mese"
        wsImp.Range("C29").Value = 6
        wsImp.Range("C29").NumberFormat = "0"
        wsImp.Range("D29").Value = ""
        With wsImp.Range("C27").Validation
            .Delete
            .Add Type:=xlValidateTime, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="0", Formula2:="=1/12"
            .ErrorMessage = "Inserisci una durata tra 0:00 e 2:00 (es. 0:30)."
        End With
        With wsImp.Range("C28").Validation
            .Delete
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="1", Formula2:="31"
            .ErrorMessage = "Giorno tra 1 e 31."
        End With
        With wsImp.Range("C29").Validation
            .Delete
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="1", Formula2:="12"
            .ErrorMessage = "Mese tra 1 e 12."
        End With
        wsImp.Range("C27:C29").Locked = False
        ThisWorkbook.Names.Add Name:="EXT_Max_Senza_Rientro", RefersTo:="=Impostazioni!$C$27"
        ThisWorkbook.Names.Add Name:="FEP_Scadenza_Giorno", RefersTo:="=Impostazioni!$C$28"
        ThisWorkbook.Names.Add Name:="FEP_Scadenza_Mese", RefersTo:="=Impostazioni!$C$29"
    End If
    sArancio = FormulaLocale(wsImp, R_K7_ARANCIO)
    sRosso = FormulaLocale(wsImp, R_K7_ROSSO)
    sLimite = FormulaLocale(wsImp, R_K_LIMITE)
    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- 2) Fogli mensili -------------------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Ferie FEP e limite EXT: " & ws.Name & "..."
            DoEvents
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ws.Activate

            AggiungiTipoAlMenu ws.Range("N11:N41")

            With ws.Range("R11:R41")                 ' EXT prima del limite
                .Formula = F_R
                .NumberFormat = ";;;"
                .Locked = True
            End With
            ws.Range("K11:K41").Formula = F_K
            ws.Range("L11:L41").Formula = F_L
            ws.Range("O11:O41").Formula = F_O

            ws.Range("P18:Q18").Copy                 ' formato riga senza regole di colore
            ws.Range("P20:Q21").PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            ws.Range("P20").Value = "Ferie anno prec. (gg)"
            ws.Range("Q20").Formula = F_Q20
            ws.Range("Q20").NumberFormat = "0"
            ws.Range("P21").Value = "Ore non autorizzate"
            ws.Range("Q21").Formula = F_Q21
            ws.Range("Q21").NumberFormat = "[h]:mm;;"
            ws.Range("Q20:Q21").HorizontalAlignment = xlRight

            ImpostaRegola ws.Range("K11:K41"), sLimite, COL_ARANCIO_FONDO, COL_ARANCIO_TESTO

            ws.Range("A1").Select
            If eraProtetto Then ws.Protect
            eraProtetto = False
        End If
    Next v

    '--- 3) Dashboard -------------------------------------------------------------
    Set ws = wsDash
    sFoglioErr = wsDash.Name
    eraProtetto = wsDash.ProtectContents
    If eraProtetto Then wsDash.Unprotect
    wsDash.Activate
    wsDash.Range("J7").Formula = F_J7
    wsDash.Range("J8").Formula = F_J8
    wsDash.Range("L7").Formula = F_L7
    wsDash.Range("I5").Formula = F_I5
    wsDash.Range("K7").FormatConditions.Delete
    wsDash.Range("L7").FormatConditions.Delete
    wsDash.Range("K7").Select
    ImpostaRegola wsDash.Range("K7"), sArancio, COL_ARANCIO_FONDO, COL_ARANCIO_TESTO
    ImpostaRegola wsDash.Range("K7"), sRosso, COL_ROSSO_FONDO, COL_ROSSO_TESTO
    wsDash.Range("L7").Select
    ImpostaRegola wsDash.Range("L7"), sArancio, COL_ARANCIO_FONDO, COL_ARANCIO_TESTO
    ImpostaRegola wsDash.Range("L7"), sRosso, COL_ROSSO_FONDO, COL_ROSSO_TESTO
    wsDash.Range("A1").Select
    If eraProtetto Then wsDash.Protect
    eraProtetto = False
    sFoglioErr = ""

    Application.Calculation = xlCalc
    Application.Calculate
    sDopo = Foto(vMesi)
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
               vbCritical, "Ferie FEP e limite EXT"
    Else
        MsgBox "Aggiornamento completato." & vbCrLf & vbCrLf & _
               "Prima: " & sPrima & vbCrLf & "Dopo:  " & sDopo & vbCrLf & vbCrLf & _
               "ORA segna come '" & FEP & "' i giorni che il portale indica FEP" & vbCrLf & _
               "(finche' non lo fai, le ferie dell'anno precedente risultano non godute)." & _
               vbCrLf & vbCrLf & "Backup: " & sBackup & vbCrLf & "Salva il file.", _
               vbInformation, "Ferie FEP e limite EXT"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


' Menu Tipo Assenza: elenco nelle celle Impostazioni!B32:B38 (nome Tipi_Assenza)
' e convalida che punta al nome, cosi' non dipende dal separatore di elenco
' della lingua di Excel. (Versione corretta: la prima versione scriveva
' l'elenco come testo e in Excel italiano diventava una voce unica.)
Private Sub AggiungiTipoAlMenu(ByVal rng As Range)
    Dim wsImp As Worksheet, vTipi As Variant, i As Long, bProt As Boolean
    vTipi = Array("Ferie", "Ferie anno prec.", "Malattia", "L. 104", _
                  "Art. 23 (Visite)", "Art. 25 (Motivi pers.)", "Art. 34 (A recupero)")
    If Not EsisteNome("Tipi_Assenza") Then
        Set wsImp = ThisWorkbook.Worksheets("Impostazioni")
        bProt = wsImp.ProtectContents
        If bProt Then wsImp.Unprotect
        wsImp.Range("B31").Value = "Tipi di assenza (menu a tendina)"
        wsImp.Range("B31").Font.Bold = True
        For i = 0 To UBound(vTipi)
            wsImp.Cells(32 + i, 2).Value = vTipi(i)
        Next i
        wsImp.Range("B31:B38").Borders.LineStyle = xlContinuous
        wsImp.Range("D31").Value = "Voci del menu Tipo Assenza: non modificarle, i nomi sono usati dalle formule."
        ThisWorkbook.Names.Add Name:="Tipi_Assenza", RefersTo:="=Impostazioni!$B$32:$B$38"
        If bProt Then wsImp.Protect
    End If
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=Tipi_Assenza"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub


Private Function Foto(ByVal vMesi As Variant) As String
    Dim v As Variant, s As String
    For Each v In vMesi
        s = s & ThisWorkbook.Worksheets(CStr(v)).Range("E4").Text & "/" & _
            ThisWorkbook.Worksheets(CStr(v)).Range("E5").Text & " "
    Next v
    Foto = "EXT/FLE mesi: " & s
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
        Err.Raise vbObjectError + 991, , "La cella Impostazioni!F2 deve essere vuota."
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
