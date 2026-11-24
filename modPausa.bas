Attribute VB_Name = "modPausa"
Option Explicit

'=============================================================================
' Modulo  : modPausa  (macro "una tantum", si puo' rieseguire senza doppioni)
' Macro   : AggiungiControlloPausa
'
'   1. aggiunge in Impostazioni (riga 26) il parametro "Pausa minima"
'      (nome Pausa_Minima, valore iniziale 0:30);
'   2. nei 12 mesi colora di ROSSO Usc 1 ed Entr 2 (G:H) quando la pausa
'      tra le due timbrature e' piu' breve della pausa minima.
'      (Se Entr 2 e' prima di Usc 1 interviene gia' il controllo
'       "orari non in sequenza".)
'
' La stessa verifica e' inclusa in VerificaAssenze e nel controllo prima
' del PDF annuale (modulo modControlliAssenze aggiornato).
' Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const COL_ROSSO_FONDO As Long = 13551615     ' RGB(255,199,206)
Private Const COL_ROSSO_TESTO As Long = 393372       ' RGB(156,0,6)

Private Const R_PAUSA As String = _
    "=AND(ISNUMBER($G11),ISNUMBER($H11),$H11>=$G11," & _
    "ROUND(($H11-$G11)*1440,0)<ROUND(Pausa_Minima*1440,0))"

Public Sub AggiungiControlloPausa()

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
        MsgBox "Serve il foglio Impostazioni.", vbExclamation, "Pausa minima"
        Exit Sub
    End If
    If MsgBox("Salva il file prima di continuare (Ctrl+S)." & vbCrLf & vbCrLf & _
              "Aggiungere il controllo della pausa minima?", vbQuestion + vbYesNo, _
              "Pausa minima") <> vbYes Then Exit Sub

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- 1) Parametro in Impostazioni -----------------------------------------
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect
    If Not EsisteNome("Pausa_Minima") Then
        wsImp.Activate
        wsImp.Range("B25:C25").Copy
        wsImp.Range("B26:C26").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        wsImp.Range("B26").Value = "Pausa minima tra Usc 1 ed Entr 2"
        With wsImp.Range("C26")
            .Value = TimeSerial(0, 30, 0)
            .NumberFormat = "[h]:mm"
            .Locked = False
            .Validation.Delete
            .Validation.Add Type:=xlValidateTime, AlertStyle:=xlValidAlertStop, _
                            Operator:=xlBetween, Formula1:="0", Formula2:="=1/12"
            .Validation.ErrorMessage = "Inserisci una durata tra 0:00 e 2:00 (es. 0:30)."
        End With
        With wsImp.Range("D26")
            .Value = "Pause piu' brevi: Usc 1 ed Entr 2 diventano rosse."
            .Font.Italic = True
            .Font.Color = RGB(128, 128, 128)
        End With
        ThisWorkbook.Names.Add Name:="Pausa_Minima", RefersTo:="=Impostazioni!$C$26"
    End If
    sRegola = FormulaLocale(wsImp, R_PAUSA)
    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- 2) Regola di colore nei mesi -------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Pausa minima: " & ws.Name & "..."
            DoEvents
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ws.Activate
            ImpostaRegola ws.Range("G11:H41"), sRegola
            ws.Range("A1").Select
            If eraProtetto Then ws.Protect
            eraProtetto = False
            nFogli = nFogli + 1
        End If
    Next v
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
               "Chiudi il file senza salvare e riaprilo.", vbCritical, "Pausa minima"
    Else
        MsgBox "Controllo della pausa minima aggiunto in " & nFogli & " mesi." & vbCrLf & _
               "Valore modificabile in Impostazioni, riga 26 (ora " & _
               wsImp.Range("C26").Text & ")." & vbCrLf & vbCrLf & "Salva il file.", _
               vbInformation, "Pausa minima"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


' Toglie la stessa regola se gia' presente e la aggiunge (prima cella attiva)
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
        .Interior.Color = COL_ROSSO_FONDO
        .Font.Color = COL_ROSSO_TESTO
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
        Err.Raise vbObjectError + 900, , "La cella Impostazioni!F2 deve essere vuota."
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
