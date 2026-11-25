Attribute VB_Name = "modFerieScadenza"
Option Explicit

'=============================================================================
' Modulo  : modFerieScadenza  (macro "una tantum", rieseguibile)
' Macro   : AggiornaFerieAnnoPrecedente
'
' Gestisce separatamente le ferie dell'anno precedente, da fruire entro il
' 30 giugno:
'
'   Dashboard
'     J7  Goduto ferie anno prec. = solo ferie da GENNAIO a GIUGNO
'         (sempre consumate per prime), fino al massimo del residuo
'     J8  Goduto ferie anno       = tutte le ferie dell'anno meno J7
'         (quelle da luglio in poi vanno sempre sull'anno in corso)
'     L7  stato del residuo anno prec.:
'           "da fruire entro il 30/06/AAAA"  (arancione, fino al 30 giugno)
'           "scadute il 30/06/AAAA"          (rosso, dopo il 30 giugno)
'         K7 e L7 si colorano allo stesso modo.
'
'   Fogli mensili
'     B7  la nota in alto viene aggiornata alla regola attuale
'         (Ferie/Malattia/L. 104: basta il Tipo Assenza).
'
' Alla fine confronta i valori ferie prima/dopo.
' Va eseguita una volta in OGNI file (Paolo e Barbara).
' Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const COL_ARANCIO_FONDO As Long = 10284031   ' RGB(255,235,156)
Private Const COL_ARANCIO_TESTO As Long = 22428      ' RGB(156,87,0)
Private Const COL_ROSSO_FONDO As Long = 13551615     ' RGB(255,199,206)
Private Const COL_ROSSO_TESTO As Long = 393372       ' RGB(156,0,6)

Private Const F_J7 As String = "=MIN(I7,SUM(Gennaio:Giugno!Q11))"
Private Const F_J8 As String = "=SUM(Gennaio:Dicembre!Q11)-J7"
Private Const F_L7 As String = _
    "=IF(I7-J7<=0,"""",IF(TODAY()<=DATE($A$1,6,30)," & _
    """da fruire entro il 30/06/""&$A$1,""scadute il 30/06/""&$A$1))"
Private Const R_DA_FRUIRE As String = "=AND($K$7>0,TODAY()<=DATE($A$1,6,30))"
Private Const R_SCADUTE As String = "=AND($K$7>0,TODAY()>DATE($A$1,6,30))"
Private Const NOTA_MESI As String = _
    "Nota: per Ferie, Malattia e L. 104 basta il Tipo Assenza; per i permessi a ore " & _
    "(Art. 23, 25, 34) inserisci anche le ore in 'Assenze' (col. E)."

Public Sub AggiornaFerieAnnoPrecedente()

    Dim vMesi As Variant, v As Variant
    Dim wsDash As Worksheet, wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim sPrima As String, sDopo As String
    Dim sDaFruire As String, sScadute As String
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim nNote As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set wsDash = OttieniFoglio("Dashboard")
    Set wsImp = OttieniFoglio("Impostazioni")
    If wsDash Is Nothing Or wsImp Is Nothing Then
        MsgBox "Servono i fogli Dashboard e Impostazioni.", vbExclamation, "Ferie anno precedente"
        Exit Sub
    End If
    If MsgBox("Salva il file prima di continuare (Ctrl+S)." & vbCrLf & vbCrLf & _
              "Le ferie dell'anno precedente verranno scalate solo da gennaio a giugno " & _
              "(scadenza 30 giugno). Procedere?", vbQuestion + vbYesNo, _
              "Ferie anno precedente") <> vbYes Then Exit Sub

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Application.Calculate
    sPrima = FotoFerie(wsDash)

    '--- Regole di colore nella lingua di Excel (Impostazioni!F2) -------------
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect
    sDaFruire = FormulaLocale(wsImp, R_DA_FRUIRE)
    sScadute = FormulaLocale(wsImp, R_SCADUTE)
    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- Dashboard -------------------------------------------------------------
    Set ws = wsDash
    sFoglioErr = wsDash.Name
    eraProtetto = wsDash.ProtectContents
    If eraProtetto Then wsDash.Unprotect

    wsDash.Range("J7").Formula = F_J7
    wsDash.Range("J8").Formula = F_J8
    With wsDash.Range("L7")
        .Formula = F_L7
        .Font.Italic = True
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With

    ' K7: sostituisco le regole precedenti (es. "maggiore di 0") con le nuove
    wsDash.Activate
    wsDash.Range("K7").FormatConditions.Delete
    wsDash.Range("L7").FormatConditions.Delete
    wsDash.Range("K7").Select
    AggiungiRegola wsDash.Range("K7"), sDaFruire, COL_ARANCIO_FONDO, COL_ARANCIO_TESTO
    AggiungiRegola wsDash.Range("K7"), sScadute, COL_ROSSO_FONDO, COL_ROSSO_TESTO
    wsDash.Range("L7").Select
    AggiungiRegola wsDash.Range("L7"), sDaFruire, COL_ARANCIO_FONDO, COL_ARANCIO_TESTO
    AggiungiRegola wsDash.Range("L7"), sScadute, COL_ROSSO_FONDO, COL_ROSSO_TESTO
    wsDash.Range("A1").Select

    If eraProtetto Then wsDash.Protect
    eraProtetto = False

    '--- Nota in alto nei mesi -------------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            If InStr(1, CStr(ws.Range("B7").Value), "Nota", vbTextCompare) = 1 Then
                eraProtetto = ws.ProtectContents
                If eraProtetto Then ws.Unprotect
                ws.Range("B7").Value = NOTA_MESI
                If eraProtetto Then ws.Protect
                eraProtetto = False
                nNote = nNote + 1
            End If
        End If
    Next v
    sFoglioErr = ""

    Application.Calculate
    sDopo = FotoFerie(wsDash)

Uscita:
    On Error Resume Next
    If eraProtetto Then ws.Protect
    If Not shIniziale Is Nothing Then shIniziale.Activate
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, "") & vbCrLf & vbCrLf & _
               "Chiudi il file senza salvare e riaprilo.", vbCritical, "Ferie anno precedente"
    Else
        MsgBox "Ferie anno precedente aggiornate (scadenza 30 giugno)." & vbCrLf & _
               "Note aggiornate in " & nNote & " mesi." & vbCrLf & vbCrLf & _
               "Prima: " & sPrima & vbCrLf & _
               "Dopo:  " & sDopo & vbCrLf & vbCrLf & _
               IIf(sPrima = sDopo, "Nessun valore cambiato.", _
                   "I valori sono cambiati: le ferie prese dopo il 30 giugno ora vanno " & _
                   "sull'anno in corso.") & vbCrLf & vbCrLf & "Salva il file.", _
               vbInformation, "Ferie anno precedente"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


Private Function FotoFerie(ByVal ws As Worksheet) As String
    FotoFerie = "anno prec. goduto " & ws.Range("J7").Text & " residuo " & ws.Range("K7").Text & _
                " | anno goduto " & ws.Range("J8").Text & " residuo " & ws.Range("K8").Text
End Function


' Aggiunge una regola formula (gia' nella lingua di Excel) alla cella attiva
Private Sub AggiungiRegola(ByVal rng As Range, ByVal sFormula As String, _
                           ByVal lFondo As Long, ByVal lTesto As Long)
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
        Err.Raise vbObjectError + 970, , "La cella Impostazioni!F2 deve essere vuota."
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
