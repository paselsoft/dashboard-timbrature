Attribute VB_Name = "modRiepilogoAnomalie"
Option Explicit

'=============================================================================
' Modulo  : modRiepilogoAnomalie  (macro "una tantum", rieseguibile)
' Macro   : AggiungiRiepilogoAnomalie
'
' Conta le anomalie con sole formule (funziona anche su Excel per iPad):
'
'   Fogli mensili
'     O11:O41  (colonna stretta, contenuto INVISIBILE) numero di anomalie
'              della riga:
'                - giorno lavorativo passato senza timbrature ne' assenze
'                - ore in Assenze senza Tipo Assenza
'                - permesso a ore (Art. 23/25/34) senza ore
'                - orari Entr/Usc non in sequenza
'                - pausa piu' breve della pausa minima
'                - numero dispari di timbrature (escluso il giorno di oggi)
'     P19:Q19  "Anomalie" = somma della colonna O + L. 104 oltre il massimo
'              mensile; rosso se maggiore di zero.
'
'   Dashboard
'     G6:G18   colonna "Anomalie" per mese (vuota se zero, rossa se > 0)
'     H5:I5    "Anomalie nell'anno:" totale ("nessuna" in verde, numero in rosso)
'
' L'elenco dettagliato resta disponibile con la macro VerificaAssenze.
' Va eseguita una volta in OGNI file. Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const COL_ROSSO_FONDO As Long = 13551615     ' RGB(255,199,206)
Private Const COL_ROSSO_TESTO As Long = 393372       ' RGB(156,0,6)
Private Const COL_VERDE_FONDO As Long = 13561798     ' RGB(198,239,206)
Private Const COL_VERDE_TESTO As Long = 24832        ' RGB(0,97,0)
Private Const COL_INTESTAZIONE As Long = 15917529    ' RGB(217,225,242)

Public Sub AggiungiRiepilogoAnomalie()

    Dim vMesi As Variant
    Dim wsDash As Worksheet, wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim F_O As String, F_Q19 As String, F_I5 As String
    Dim sRossoNum As String, sRossoQ19 As String, sRossoI5 As String, sVerdeI5 As String
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim i As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set wsDash = OttieniFoglio("Dashboard")
    Set wsImp = OttieniFoglio("Impostazioni")
    If wsDash Is Nothing Or wsImp Is Nothing Then
        MsgBox "Servono i fogli Dashboard e Impostazioni.", vbExclamation, "Riepilogo anomalie"
        Exit Sub
    End If
    If Not EsisteNome("Pausa_Minima") Or Not EsisteNome("Max_L104_Mese") Then
        MsgBox "Mancano i parametri Pausa_Minima o Max_L104_Mese in Impostazioni.", _
               vbExclamation, "Riepilogo anomalie"
        Exit Sub
    End If
    If MsgBox("Salva il file prima di continuare (Ctrl+S)." & vbCrLf & vbCrLf & _
              "Aggiungere il riepilogo delle anomalie in Dashboard e nei mesi?", _
              vbQuestion + vbYesNo, "Riepilogo anomalie") <> vbYes Then Exit Sub

    ' Numero di anomalie della riga (sintassi inglese, riferita alla riga 11)
    F_O = "=IF(ISNUMBER($B11),AND(ISNUMBER($E11),$E11>0,$N11="""")+AND(OR($N11=""Art. 23 (Visite)""," & _
        "$N11=""Art. 25 (Motivi pers.)"",$N11=""Art. 34 (A recupero)""),OR(NOT(ISNUMBER($E11))," & _
        "$E11=0))+OR(AND(ISNUMBER($F11),ISNUMBER($G11),$G11<$F11),AND(ISNUMBER($G11)," & _
        "ISNUMBER($H11),$H11<$G11),AND(ISNUMBER($H11),ISNUMBER($I11)," & _
        "$I11<$H11))+AND(ISNUMBER($G11),ISNUMBER($H11),$H11>=$G11,ROUND(($H11-$G11)*1440," & _
        "0)<ROUND(Pausa_Minima*1440,0))+AND(ISODD(COUNT($F11:$I11)),DATE(Dashboard!$A$1,$A$1," & _
        "$B11)<TODAY())+AND(ISNUMBER($D11),$D11>0,DATE(Dashboard!$A$1,$A$1,$B11)<TODAY()," & _
        "COUNT($E11:$I11)=0,$N11=""""),0)"
    F_Q19 = "=SUM($O$11:$O$41)+(COUNTIF($N$11:$N$41,""L. 104"")>Max_L104_Mese)"
    F_I5 = "=IF(SUM(Gennaio:Dicembre!$Q$19)=0,""nessuna"",SUM(Gennaio:Dicembre!$Q$19))"

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- Regole di colore nella lingua di Excel (Impostazioni!F2) -------------
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect
    sRossoNum = FormulaLocale(wsImp, "=AND(ISNUMBER($G7),$G7>0)")
    sRossoQ19 = FormulaLocale(wsImp, "=$Q$19>0")
    sRossoI5 = FormulaLocale(wsImp, "=ISNUMBER($I$5)")
    sVerdeI5 = FormulaLocale(wsImp, "=$I$5=""nessuna""")
    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- Fogli mensili -----------------------------------------------------------
    For i = 0 To 11
        Set ws = OttieniFoglio(CStr(vMesi(i)))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Riepilogo anomalie: " & ws.Name & "..."
            DoEvents
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ws.Activate

            ' colonna O: calcolo per riga, invisibile (formato ;;;)
            With ws.Range("O11:O41")
                .Formula = F_O
                .NumberFormat = ";;;"
            End With

            ' riepilogo: riga 19 "Anomalie"
            ws.Range("P18:Q18").Copy
            ws.Range("P19:Q19").PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            ws.Range("P19").Value = "Anomalie"
            With ws.Range("Q19")
                .Formula = F_Q19
                .NumberFormat = "0"
                .HorizontalAlignment = xlRight
            End With
            ImpostaRegola ws.Range("Q19"), sRossoQ19, COL_ROSSO_FONDO, COL_ROSSO_TESTO

            ws.Range("A1").Select
            If eraProtetto Then ws.Protect
            eraProtetto = False
        End If
    Next i

    '--- Dashboard -----------------------------------------------------------------
    Set ws = wsDash
    sFoglioErr = wsDash.Name
    eraProtetto = wsDash.ProtectContents
    If eraProtetto Then wsDash.Unprotect
    wsDash.Activate

    ' colonna G "Anomalie" nella tabella dei mesi
    With wsDash.Range("G6")
        .Value = "Anomalie"
        .Font.Bold = True
        .Interior.Color = COL_INTESTAZIONE
        .HorizontalAlignment = xlCenter
    End With
    For i = 0 To 11
        wsDash.Range("G" & (7 + i)).Formula = _
            "=IF(" & vMesi(i) & "!$Q$19=0,""""," & vMesi(i) & "!$Q$19)"
    Next i
    With wsDash.Range("G7:G18")
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0"
        .Font.Bold = True
    End With
    With wsDash.Range("G6:G18").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    If wsDash.Columns("G").ColumnWidth < 10 Then wsDash.Columns("G").ColumnWidth = 10
    ImpostaRegola wsDash.Range("G7:G18"), sRossoNum, COL_ROSSO_FONDO, COL_ROSSO_TESTO

    ' totale dell'anno sotto "Dati aggiornati al"
    With wsDash.Range("H5")
        .Value = "Anomalie nell'anno:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With wsDash.Range("I5")
        .Formula = F_I5
        .NumberFormat = "0"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ImpostaRegola wsDash.Range("I5"), sRossoI5, COL_ROSSO_FONDO, COL_ROSSO_TESTO
    ImpostaRegola wsDash.Range("I5"), sVerdeI5, COL_VERDE_FONDO, COL_VERDE_TESTO

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
               "Chiudi il file senza salvare e riaprilo.", vbCritical, "Riepilogo anomalie"
    Else
        MsgBox "Riepilogo anomalie aggiunto." & vbCrLf & vbCrLf & _
               "Dashboard - Anomalie nell'anno: " & wsDash.Range("I5").Text & vbCrLf & vbCrLf & _
               "Per l'elenco dettagliato usa la macro VerificaAssenze." & vbCrLf & _
               "Salva il file.", vbInformation, "Riepilogo anomalie"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


' Toglie la stessa regola se gia' presente e la aggiunge (cella attiva = prima cella)
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


' Converte la formula inglese nella lingua di Excel tramite la cella vuota F2
' di Impostazioni (il foglio deve essere gia' sprotetto).
Private Function FormulaLocale(ByVal ws As Worksheet, ByVal sFormula As String) As String
    Dim c As Range
    Set c = ws.Range("F2")
    If Len(c.Formula) > 0 Then
        Err.Raise vbObjectError + 980, , "La cella Impostazioni!F2 deve essere vuota."
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
