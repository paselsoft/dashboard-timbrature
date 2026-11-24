Attribute VB_Name = "modGiorniMancanti"
Option Explicit

'=============================================================================
' Modulo  : modGiorniMancanti  (macro "una tantum", rieseguibile senza doppioni)
' Macro   : AggiungiControlloGiorniMancanti
'
' Nei 12 mesi colora di GIALLO il giorno (colonna B) e le celle Entr/Usc
' (F:I) quando un giorno lavorativo GIA' PASSATO:
'   - ha un orario previsto (Target), quindi non e' festivo ne' weekend;
'   - non ha timbrature (F:I vuote) ne' ore di assenza (E);
'   - non ha un Tipo Assenza (N).
' I giorni di oggi e futuri non vengono mai segnalati.
'
' La stessa verifica e' inclusa in VerificaAssenze e nel controllo prima
' del PDF annuale (modulo modControlliAssenze aggiornato).
' Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const COL_GIALLO_FONDO As Long = 10092543    ' RGB(255,255,153)

Private Const R_GIORNO_MANCANTE As String = _
    "=AND(ISNUMBER($B11),ISNUMBER($D11),$D11>0," & _
    "DATE(Dashboard!$A$1,$A$1,$B11)<TODAY()," & _
    "COUNT($E11:$I11)=0,$N11="""")"

Public Sub AggiungiControlloGiorniMancanti()

    Dim vMesi As Variant, v As Variant
    Dim wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim sRegola As String
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim nFogli As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set wsImp = OttieniFoglio("Impostazioni")
    If wsImp Is Nothing Then
        MsgBox "Serve il foglio Impostazioni.", vbExclamation, "Giorni mancanti"
        Exit Sub
    End If
    If MsgBox("Salva il file prima di continuare (Ctrl+S)." & vbCrLf & vbCrLf & _
              "Aggiungere la segnalazione dei giorni lavorativi senza dati?", _
              vbQuestion + vbYesNo, "Giorni mancanti") <> vbYes Then Exit Sub

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- Formula nella lingua di Excel (cella di appoggio Impostazioni!F2) -----
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect
    sRegola = FormulaLocale(wsImp, R_GIORNO_MANCANTE)
    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- Regola nei mesi -------------------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Giorni mancanti: " & ws.Name & "..."
            DoEvents
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ws.Activate
            ImpostaRegola ws.Range("B11:B41"), sRegola
            ImpostaRegola ws.Range("F11:I41"), sRegola
            ws.Range("A1").Select
            If eraProtetto Then ws.Protect
            eraProtetto = False
            nFogli = nFogli + 1
        End If
    Next v
    sFoglioErr = ""

Uscita:
    On Error Resume Next
    If eraProtetto Then ws.Protect
    If Not shIniziale Is Nothing Then shIniziale.Activate
    Application.StatusBar = False
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, "") & vbCrLf & vbCrLf & _
               "Chiudi il file senza salvare e riaprilo.", vbCritical, "Giorni mancanti"
    Else
        MsgBox "Segnalazione aggiunta in " & nFogli & " mesi:" & vbCrLf & _
               "i giorni lavorativi gia' passati senza timbrature ne' assenze" & vbCrLf & _
               "ora sono evidenziati in giallo." & vbCrLf & vbCrLf & "Salva il file.", _
               vbInformation, "Giorni mancanti"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


' Toglie la stessa regola se gia' presente e la aggiunge (prima cella attiva).
' Solo riempimento giallo: il colore del testo resta quello delle altre regole.
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
        .Interior.Color = COL_GIALLO_FONDO
        .StopIfTrue = False
    End With
End Sub


' Converte la formula inglese nella lingua di Excel tramite la cella vuota F2
' di Impostazioni (il foglio deve essere gia' sprotetto).
Private Function FormulaLocale(ByVal ws As Worksheet, ByVal sFormula As String) As String
    Dim c As Range
    Set c = ws.Range("F2")
    If Len(c.Formula) > 0 Then
        Err.Raise vbObjectError + 950, , "La cella Impostazioni!F2 deve essere vuota."
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
