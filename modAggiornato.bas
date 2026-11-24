Attribute VB_Name = "modAggiornato"
Option Explicit

'=============================================================================
' Modulo  : modAggiornato  (macro "una tantum", rieseguibile)
' Macro   : AggiungiDataAggiornamento
'
'   - in ogni mese, nel riepilogo, aggiunge la riga 18
'       "Ultimo giorno inserito"  -> ultimo giorno FINO A OGGI con timbrature,
'                                    ore di assenza o Tipo Assenza;
'     (le ferie gia' pianificate nei giorni futuri non contano)
'   - in Dashboard aggiunge  H4 "Dati aggiornati al:"  e in I4 la data piu'
'     recente fra tutti i mesi ("nessun dato" se l'anno e' vuoto);
'     la data diventa ARANCIONE se e' piu' vecchia di 7 giorni.
'
' Sono solo formule: funziona anche su Excel per iPad.
' Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const COL_ARANCIO_FONDO As Long = 10284031   ' RGB(255,235,156)
Private Const COL_ARANCIO_TESTO As Long = 22428      ' RGB(156,87,0)

Public Sub AggiungiDataAggiornamento()

    Dim vMesi As Variant, v As Variant
    Dim wsDash As Worksheet, wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim F_Q18 As String, F_I4 As String, R_VECCHIA As String, sRegola As String
    Dim COND As String
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim nFogli As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set wsDash = OttieniFoglio("Dashboard")
    Set wsImp = OttieniFoglio("Impostazioni")
    If wsDash Is Nothing Or wsImp Is Nothing Then
        MsgBox "Servono i fogli Dashboard e Impostazioni.", vbExclamation, "Data aggiornamento"
        Exit Sub
    End If
    If MsgBox("Salva il file prima di continuare (Ctrl+S)." & vbCrLf & vbCrLf & _
              "Aggiungere 'Dati aggiornati al' in Dashboard e 'Ultimo giorno inserito' nei mesi?", _
              vbQuestion + vbYesNo, "Data aggiornamento") <> vbYes Then Exit Sub

    ' riga con dati = almeno una cella compilata fra E:I e N
    COND = "((($E11:$E41<>"""")+($F11:$F41<>"""")+($G11:$G41<>"""")+($H11:$H41<>"""")" & _
           "+($I11:$I41<>"""")+($N11:$N41<>""""))>0)"
    ' ultimo giorno del mese, fino a oggi, con dati (LOOKUP(2;1/condizione;giorni))
    F_Q18 = "=IFERROR(DATE(Dashboard!$A$1,$A$1,LOOKUP(2,1/(" & COND & _
            "*(DATE(Dashboard!$A$1,$A$1,$B11:$B41)<=TODAY())),$B11:$B41)),"""")"
    F_I4 = "=IF(MAX(Gennaio:Dicembre!$Q$18)=0,""nessun dato"",MAX(Gennaio:Dicembre!$Q$18))"
    R_VECCHIA = "=AND(ISNUMBER($I$4),$I$4<TODAY()-7)"

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- Riepilogo dei mesi: riga 18 -------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Data aggiornamento: " & ws.Name & "..."
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ws.Activate
            ws.Range("P17:Q17").Copy
            ws.Range("P18:Q18").PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            ws.Range("P18").Value = "Ultimo giorno inserito"
            ScriviFormula ws.Range("Q18"), F_Q18
            ws.Range("Q18").NumberFormat = "dd/mm"
            ws.Range("Q18").HorizontalAlignment = xlRight
            ws.Range("A1").Select
            If eraProtetto Then ws.Protect
            eraProtetto = False
            nFogli = nFogli + 1
        End If
    Next v

    '--- Formula della regola nella lingua di Excel (Impostazioni!F2) ---------
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect
    sRegola = FormulaLocale(wsImp, R_VECCHIA)
    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- Dashboard: H4 / I4 ----------------------------------------------------
    Set ws = wsDash
    sFoglioErr = wsDash.Name
    eraProtetto = wsDash.ProtectContents
    If eraProtetto Then wsDash.Unprotect
    With wsDash.Range("H4")
        .Value = "Dati aggiornati al:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With wsDash.Range("I4")
        .Formula = F_I4
        .NumberFormat = "dd/mm/yyyy"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    wsDash.Activate
    ImpostaRegola wsDash.Range("I4"), sRegola
    wsDash.Range("A1").Select
    If eraProtetto Then wsDash.Protect
    eraProtetto = False
    sFoglioErr = ""

Uscita:
    On Error Resume Next
    Application.CutCopyMode = False
    If eraProtetto Then ws.Protect
    If Not shIniziale Is Nothing Then shIniziale.Activate
    Application.StatusBar = False
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, "") & vbCrLf & vbCrLf & _
               "Chiudi il file senza salvare e riaprilo.", vbCritical, "Data aggiornamento"
    Else
        MsgBox "Aggiunti:" & vbCrLf & _
               "  - Dashboard: 'Dati aggiornati al' = " & wsDash.Range("I4").Text & vbCrLf & _
               "  - " & nFogli & " mesi: 'Ultimo giorno inserito' nel riepilogo" & vbCrLf & vbCrLf & _
               "Salva il file.", vbInformation, "Data aggiornamento"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


' Scrive la formula con Formula2 (Excel 365: calcolo matriciale nativo);
' sulle versioni che non hanno Formula2 usa Formula.
Private Sub ScriviFormula(ByVal c As Range, ByVal sFormula As String)
    On Error Resume Next
    c.Formula2 = sFormula
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        c.Formula = sFormula
    End If
    On Error GoTo 0
End Sub


' Toglie la stessa regola se gia' presente e la aggiunge (cella attiva = prima cella)
Private Sub ImpostaRegola(ByVal rng As Range, ByVal sFormula As String)
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
        .Interior.Color = COL_ARANCIO_FONDO
        .Font.Color = COL_ARANCIO_TESTO
        .StopIfTrue = False
    End With
End Sub


' Converte la formula inglese nella lingua di Excel tramite la cella vuota F2
' di Impostazioni (il foglio deve essere gia' sprotetto).
Private Function FormulaLocale(ByVal ws As Worksheet, ByVal sFormula As String) As String
    Dim c As Range
    Set c = ws.Range("F2")
    If Len(c.Formula) > 0 Then
        Err.Raise vbObjectError + 960, , "La cella Impostazioni!F2 deve essere vuota."
    End If
    c.Formula = sFormula
    FormulaLocale = c.FormulaLocal
    c.ClearContents
End Function


Private Function OttieniFoglio(ByVal sNome As String) As Worksheet
    On Error Resume Next
    Set OttieniFoglio = ThisWorkbook.Worksheets(sNome)
    On Error GoTo 0
End Function
