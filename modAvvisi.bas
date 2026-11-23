Attribute VB_Name = "modAvvisi"
Option Explicit

'=============================================================================
' Modulo  : modAvvisi  (macro "una tantum", si puo' rieseguire senza doppioni)
' Macro   : AggiungiAvvisi
'
' Aggiunge la formattazione condizionale per:
'
'   DASHBOARD, colonna Residuo (K8:K12: ferie anno, L. 104, Art. 23/25/34)
'     - ARANCIONE se il residuo e' sotto il 20% del plafond;
'     - ROSSO     se il residuo e' a zero o negativo.
'     I plafond sono quelli della colonna I, collegata al foglio Impostazioni.
'     (K7, ferie dell'anno precedente, ha gia' una sua regola e non si tocca.)
'
'   FOGLI MENSILI
'     - Entr/Usc (F:I) in ROSSO se gli orari non sono in sequenza:
'       Usc 1 prima di Entr 1, Entr 2 prima di Usc 1, Usc 2 prima di Entr 2;
'     - Tipo Assenza (N) in ROSSO sui giorni "L. 104" quando nel mese
'       superano il massimo mensile; anche il totale in Q12 diventa rosso.
'
' Il massimo mensile L. 104 diventa un parametro del foglio Impostazioni
' (riga 25, nome Max_L104_Mese, valore iniziale 3) e l'etichetta in
' Dashboard H9 lo mostra automaticamente.
'
' Requisito: il foglio Impostazioni (macro CreaFoglioImpostazioni).
'=============================================================================

Private Const COL_ROSSO_FONDO As Long = 13551615     ' RGB(255,199,206)
Private Const COL_ROSSO_TESTO As Long = 393372       ' RGB(156,0,6)
Private Const COL_ARANCIO_FONDO As Long = 10284031   ' RGB(255,235,156)
Private Const COL_ARANCIO_TESTO As Long = 22428      ' RGB(156,87,0)

' Regole in sintassi inglese (riferite alla prima riga della rispettiva area)
Private Const R_RESIDUO_ZERO As String = "=AND(ISNUMBER($I8),$I8>0,$K8<=0)"
Private Const R_RESIDUO_BASSO As String = "=AND(ISNUMBER($I8),$I8>0,$K8>0,$K8*5<$I8)"
Private Const R_ORARI As String = _
    "=OR(AND(ISNUMBER($F11),ISNUMBER($G11),$G11<$F11)," & _
    "AND(ISNUMBER($G11),ISNUMBER($H11),$H11<$G11)," & _
    "AND(ISNUMBER($H11),ISNUMBER($I11),$I11<$H11))"
Private Const R_L104_GIORNI As String = _
    "=AND($N11=""L. 104"",COUNTIF($N$11:$N$41,""L. 104"")>Max_L104_Mese)"
Private Const R_L104_TOTALE As String = "=$Q$12>Max_L104_Mese"

Public Sub AggiungiAvvisi()

    Dim vMesi As Variant, v As Variant
    Dim wsDash As Worksheet, wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim sZero As String, sBasso As String, sOrari As String
    Dim sL104G As String, sL104T As String
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim nFogli As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set wsDash = OttieniFoglio("Dashboard")
    Set wsImp = OttieniFoglio("Impostazioni")
    If wsDash Is Nothing Or wsImp Is Nothing Or Not EsisteNome("Plafond_Ferie") Then
        MsgBox "Serve prima il foglio Impostazioni (macro CreaFoglioImpostazioni).", _
               vbExclamation, "Avvisi"
        Exit Sub
    End If

    If MsgBox("Salva il file prima di continuare (Ctrl+S)." & vbCrLf & vbCrLf & _
              "Procedere con l'aggiunta degli avvisi?", vbQuestion + vbYesNo, "Avvisi") <> vbYes Then
        Exit Sub
    End If

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- 1) Parametro "massimo L. 104 al mese" in Impostazioni ------------------
    sFoglioErr = wsImp.Name
    If Not EsisteNome("Max_L104_Mese") Then
        eraProtetto = wsImp.ProtectContents
        If eraProtetto Then wsImp.Unprotect
        wsImp.Activate
        wsImp.Range("B24:C24").Copy
        wsImp.Range("B25:C25").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        wsImp.Range("B25").Value = "L. 104 - massimo giorni al mese"
        With wsImp.Range("C25")
            .Value = 3
            .NumberFormat = "0"
            .Validation.Delete
            .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                            Operator:=xlGreaterEqual, Formula1:="0"
            .Validation.ErrorMessage = "Inserisci un numero intero maggiore o uguale a zero."
        End With
        wsImp.Range("D25").Value = "Oltre questo numero i giorni L. 104 del mese diventano rossi."
        wsImp.Range("D25").Font.Italic = True
        wsImp.Range("D25").Font.Color = RGB(128, 128, 128)
        ThisWorkbook.Names.Add Name:="Max_L104_Mese", RefersTo:="=Impostazioni!$C$25"
        If eraProtetto Then wsImp.Protect
        eraProtetto = False
    End If

    '--- 2) Formule nella lingua di Excel (conversione una sola volta) ---------
    sZero = FormulaLocale(wsImp, R_RESIDUO_ZERO)
    sBasso = FormulaLocale(wsImp, R_RESIDUO_BASSO)
    sOrari = FormulaLocale(wsImp, R_ORARI)
    sL104G = FormulaLocale(wsImp, R_L104_GIORNI)
    sL104T = FormulaLocale(wsImp, R_L104_TOTALE)

    '--- 3) Dashboard ----------------------------------------------------------
    Set ws = wsDash
    sFoglioErr = ws.Name
    Application.StatusBar = "Avvisi: Dashboard..."
    eraProtetto = ws.ProtectContents
    If eraProtetto Then ws.Unprotect
    ws.Range("H9").Formula = "=""L. 104 (max ""&Max_L104_Mese&"" gg/mese)"""
    ws.Activate
    ImpostaRegola ws.Range("K8:K12"), sZero, COL_ROSSO_FONDO, COL_ROSSO_TESTO
    ImpostaRegola ws.Range("K8:K12"), sBasso, COL_ARANCIO_FONDO, COL_ARANCIO_TESTO
    ws.Range("A1").Select
    If eraProtetto Then ws.Protect
    eraProtetto = False

    '--- 4) Fogli mensili ------------------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Avvisi: " & ws.Name & "..."
            DoEvents
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ws.Activate
            ImpostaRegola ws.Range("F11:I41"), sOrari, COL_ROSSO_FONDO, COL_ROSSO_TESTO
            ImpostaRegola ws.Range("N11:N41"), sL104G, COL_ROSSO_FONDO, COL_ROSSO_TESTO
            ImpostaRegola ws.Range("Q12"), sL104T, COL_ROSSO_FONDO, COL_ROSSO_TESTO
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
               "Chiudi il file senza salvare e riaprilo.", vbCritical, "Avvisi"
    Else
        MsgBox "Avvisi aggiunti (Dashboard + " & nFogli & " mesi)." & vbCrLf & vbCrLf & _
               "Dashboard, colonna Residuo:" & vbCrLf & _
               "  - arancione sotto il 20% del plafond" & vbCrLf & _
               "  - rosso a zero o negativo" & vbCrLf & vbCrLf & _
               "Fogli mensili:" & vbCrLf & _
               "  - orari Entr/Usc non in sequenza in rosso" & vbCrLf & _
               "  - giorni L. 104 oltre il massimo mensile in rosso" & vbCrLf & _
               "    (massimo modificabile in Impostazioni, riga 25)" & vbCrLf & vbCrLf & _
               "Salva il file.", vbInformation, "Avvisi"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


'-----------------------------------------------------------------------------
' Toglie la stessa regola se gia' presente (riesecuzione senza doppioni) e la
' aggiunge. FormatConditions interpreta i riferimenti relativi rispetto alla
' cella attiva: seleziono la prima cella dell'area.
'-----------------------------------------------------------------------------
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


' Converte una formula in inglese nella lingua di Excel (E, O, VAL.NUMERO,
' punto e virgola...) usando la cella vuota F2 del foglio Impostazioni.
Private Function FormulaLocale(ByVal ws As Worksheet, ByVal sFormula As String) As String
    Dim c As Range
    Dim eraProtetto As Boolean
    Set c = ws.Range("F2")
    If Len(c.Formula) > 0 Then
        Err.Raise vbObjectError + 800, , "La cella Impostazioni!F2 deve essere vuota."
    End If
    eraProtetto = ws.ProtectContents
    If eraProtetto Then ws.Unprotect
    c.Formula = sFormula
    FormulaLocale = c.FormulaLocal
    c.ClearContents
    If eraProtetto Then ws.Protect
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
