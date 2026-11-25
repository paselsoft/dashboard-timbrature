Attribute VB_Name = "modControlliAssenze"
Option Explicit

'=============================================================================
' Modulo  : modControlliAssenze
' Scopo   : evidenzia in rosso, nei 12 fogli mensili, le celle "Assenze" (E)
'           e "Tipo Assenza" (N) quando:
'             1) ci sono ore in Assenze ma il Tipo Assenza e' vuoto;
'             2) il Tipo e' un permesso a ore (Art. 23, 25, 34) ma le ore
'                in Assenze mancano o sono zero.
'
' Macro:
'   AggiungiControlliAssenze  -> installa (o reinstalla) la formattazione
'                                condizionale e poi mostra le anomalie trovate.
'                                Basta eseguirla una volta; si puo' ripetere
'                                senza creare doppioni.
'   VerificaAssenze           -> elenca le anomalie presenti in tutti i mesi
'                                (assenze, orari non in sequenza, pause piu'
'                                brevi del minimo, L. 104 oltre il massimo
'                                mensile, giorni lavorativi passati senza dati,
'                                timbrature dispari),
'                                senza modificare nulla. Utile da collegare a
'                                un pulsante sulla Dashboard.
'=============================================================================

Private Const RIGA_PRIMA As Long = 11
Private Const RIGA_ULTIMA As Long = 41
Private Const AREA_ORE As String = "$E$11:$E$41"
Private Const AREA_TIPO As String = "$N$11:$N$41"
Private Const AREA_VECCHIA As String = "$E$11:$E$41,$N$11:$N$41"   ' versione 1
Private Const CELLA_APPOGGIO As String = "O11"   ' colonna vuota tra Tipo e Riepilogo

' Regole in sintassi inglese, riferite alla riga 11 (prima riga dell'area).
Private Const REGOLA_ORE_SENZA_TIPO As String = _
    "=AND(ISNUMBER($E11),$E11>0,$N11="""")"
Private Const REGOLA_TIPO_SENZA_ORE As String = _
    "=AND(OR($N11=""Art. 23 (Visite)"",$N11=""Art. 25 (Motivi pers.)""," & _
    "$N11=""Art. 34 (A recupero)""),OR(NOT(ISNUMBER($E11)),$E11=0))"


'-----------------------------------------------------------------------------
' Installa la formattazione condizionale nei 12 mesi
' (versione 2: aree singole, nessuna cella temporanea fuori tabella,
'  eventi e ricalcolo sospesi, avanzamento nella barra di stato)
'-----------------------------------------------------------------------------
Public Sub AggiungiControlliAssenze()

    Dim vMesi As Variant, v As Variant
    Dim ws As Worksheet, shIniziale As Object
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim xlCalc As Long
    Dim nFogli As Long
    Dim sRegola1 As String, sRegola2 As String
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    If MsgBox("Prima di continuare salva il file (Ctrl+S)." & vbCrLf & vbCrLf & _
              "Procedere con l'installazione dei controlli?", _
              vbQuestion + vbYesNo, "Controlli assenze") <> vbYes Then Exit Sub

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    xlCalc = Application.Calculation

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            Application.StatusBar = "Controlli assenze: " & ws.Name & "..."
            DoEvents

            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect

            ' Le formule vanno passate nella lingua di Excel: le converto una
            ' sola volta, sul primo foglio, usando la cella vuota O11.
            If Len(sRegola1) = 0 Then
                sRegola1 = FormulaLocale(ws, REGOLA_ORE_SENZA_TIPO)
                sRegola2 = FormulaLocale(ws, REGOLA_TIPO_SENZA_ORE)
            End If

            ' FormatConditions.Add interpreta i riferimenti relativi rispetto
            ' alla cella attiva: attivo il foglio e seleziono E11.
            ws.Activate
            ws.Range("E" & RIGA_PRIMA).Select

            RimuoviRegoleEsistenti ws
            AggiungiRegola ws.Range(AREA_ORE), sRegola1
            AggiungiRegola ws.Range(AREA_ORE), sRegola2
            AggiungiRegola ws.Range(AREA_TIPO), sRegola1
            AggiungiRegola ws.Range(AREA_TIPO), sRegola2

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
    Application.Calculation = xlCalc
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, ""), _
               vbCritical, "Controlli assenze"
    Else
        MsgBox "Controlli installati in " & nFogli & " fogli mensili." & vbCrLf & vbCrLf & _
               "Le celle Assenze e Tipo Assenza diventano rosse quando:" & vbCrLf & _
               "  - ci sono ore senza Tipo Assenza;" & vbCrLf & _
               "  - c'e' un permesso a ore (Art. 23, 25, 34) senza ore.", _
               vbInformation, "Controlli assenze"
        VerificaAssenze
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


'-----------------------------------------------------------------------------
' Mostra le anomalie di tutti i mesi (non modifica nulla)
'-----------------------------------------------------------------------------
Public Sub VerificaAssenze()
    Dim sElenco As String
    Dim n As Long

    sElenco = ElencoAnomalie()
    n = ContaVoci(sElenco)

    If n = 0 Then
        MsgBox "Nessuna anomalia nei 12 mesi (assenze, orari, pause, timbrature, L. 104, giorni mancanti).", _
               vbInformation, "Verifica assenze"
    Else
        MsgBox "Trovate " & n & " anomalie:" & Abbrevia(sElenco, 20), _
               vbExclamation, "Verifica assenze"
    End If
End Sub


'-----------------------------------------------------------------------------
' Restituisce l'elenco delle anomalie (una per riga, ciascuna preceduta da
' vbCrLf & "  - "), oppure "" se non ce ne sono.
' E' pubblica: la usa anche EsportaReportPDF prima di creare il PDF.
'-----------------------------------------------------------------------------
Public Function ElencoAnomalie() As String

    Dim vMesi As Variant, v As Variant
    Dim ws As Worksheet
    Dim r As Long
    Dim sTipo As String, vOre As Variant
    Dim bOre As Boolean
    Dim sElenco As String
    Dim nL104 As Long, vMax As Variant, lMaxL104 As Long
    Dim vPausa As Variant, lPausaMin As Long, lPausa As Long
    Dim vG As Variant, vH As Variant
    Dim lAnno As Long, lMese As Long, vGiorno As Variant, vTarget As Variant
    Dim dGiorno As Date

    ' massimo mensile L. 104 dal foglio Impostazioni (se c'e'), altrimenti 3
    lMaxL104 = 3
    On Error Resume Next
    vMax = Application.Evaluate("Max_L104_Mese")
    If Not IsError(vMax) Then If IsNumeric(vMax) Then lMaxL104 = CLng(vMax)
    On Error GoTo 0

    ' pausa minima in minuti (Impostazioni, nome Pausa_Minima); 0 = nessun controllo
    lPausaMin = 0
    On Error Resume Next
    vPausa = Application.Evaluate("Pausa_Minima")
    If Not IsError(vPausa) Then If IsNumeric(vPausa) Then lPausaMin = CLng(Round(CDbl(vPausa) * 1440, 0))
    On Error GoTo 0

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    ' anno (Dashboard!A1) per riconoscere i giorni gia' passati
    lAnno = 0
    On Error Resume Next
    lAnno = CLng(ThisWorkbook.Worksheets("Dashboard").Range("A1").Value)
    On Error GoTo 0

    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            nL104 = 0
            lMese = 0
            If IsNumeric(ws.Range("A1").Value) Then lMese = CLng(ws.Range("A1").Value)
            For r = RIGA_PRIMA To RIGA_ULTIMA
                vOre = ws.Cells(r, 5).Value
                sTipo = Trim$(CStr(ws.Cells(r, 14).Text))
                bOre = False
                If Not IsEmpty(vOre) And Not IsError(vOre) Then
                    If IsNumeric(vOre) Then bOre = (CDbl(vOre) > 0)
                End If

                If bOre And Len(sTipo) = 0 Then
                    sElenco = sElenco & vbCrLf & "  - " & ws.Name & " " & GiornoRiga(ws, r) & _
                              ": " & ws.Cells(r, 5).Text & " in Assenze senza Tipo Assenza"
                ElseIf EPermessoAOre(sTipo) And Not bOre Then
                    sElenco = sElenco & vbCrLf & "  - " & ws.Name & " " & GiornoRiga(ws, r) & _
                              ": " & sTipo & " senza ore in Assenze"
                End If

                ' orari non in sequenza (Entr1 < Usc1 < Entr2 < Usc2)
                If OrariFuoriSequenza(ws, r) Then
                    sElenco = sElenco & vbCrLf & "  - " & ws.Name & " " & GiornoRiga(ws, r) & _
                              ": orari di entrata/uscita non in sequenza"
                End If

                ' pausa tra Usc 1 ed Entr 2 piu' breve del minimo
                If lPausaMin > 0 Then
                    vG = ws.Cells(r, 7).Value
                    vH = ws.Cells(r, 8).Value
                    If Not IsEmpty(vG) And Not IsEmpty(vH) And Not IsError(vG) And Not IsError(vH) Then
                        If IsNumeric(vG) And IsNumeric(vH) Then
                            lPausa = CLng(Round((CDbl(vH) - CDbl(vG)) * 1440, 0))
                            If lPausa >= 0 And lPausa < lPausaMin Then
                                sElenco = sElenco & vbCrLf & "  - " & ws.Name & " " & GiornoRiga(ws, r) & _
                                          ": pausa di " & lPausa & " min (minimo " & lPausaMin & ")"
                            End If
                        End If
                    End If
                End If

                ' giorno lavorativo gia' passato senza timbrature ne' assenze
                If lAnno > 0 And lMese >= 1 And lMese <= 12 And Len(sTipo) = 0 Then
                    vGiorno = ws.Cells(r, 2).Value
                    vTarget = ws.Cells(r, 4).Value
                    If Not IsError(vGiorno) And Not IsError(vTarget) Then
                        If IsNumeric(vGiorno) And Not IsEmpty(vGiorno) And _
                           IsNumeric(vTarget) And Not IsEmpty(vTarget) Then
                            If CDbl(vTarget) > 0 Then
                                dGiorno = DateSerial(lAnno, lMese, CLng(vGiorno))
                                If dGiorno < Date And _
                                   Application.WorksheetFunction.Count(ws.Range(ws.Cells(r, 5), ws.Cells(r, 9))) = 0 Then
                                    sElenco = sElenco & vbCrLf & "  - " & ws.Name & " " & GiornoRiga(ws, r) & _
                                              ": giorno lavorativo senza timbrature n" & ChrW(233) & " assenze"
                                End If
                            End If
                        End If
                    End If
                End If

                ' numero dispari di timbrature (escluso il giorno di oggi)
                If lAnno > 0 And lMese >= 1 And lMese <= 12 Then
                    vGiorno = ws.Cells(r, 2).Value
                    If Not IsError(vGiorno) Then
                        If IsNumeric(vGiorno) And Not IsEmpty(vGiorno) Then
                            If DateSerial(lAnno, lMese, CLng(vGiorno)) < Date And _
                               (Application.WorksheetFunction.Count(ws.Range(ws.Cells(r, 6), ws.Cells(r, 9))) Mod 2) = 1 Then
                                sElenco = sElenco & vbCrLf & "  - " & ws.Name & " " & GiornoRiga(ws, r) & _
                                          ": numero dispari di timbrature (manca un'entrata o un'uscita)"
                            End If
                        End If
                    End If
                End If

                If sTipo = "L. 104" Then nL104 = nL104 + 1
            Next r

            If nL104 > lMaxL104 Then
                sElenco = sElenco & vbCrLf & "  - " & ws.Name & ": " & nL104 & _
                          " giorni di L. 104 (massimo " & lMaxL104 & ")"
            End If
        End If
    Next v

    ElencoAnomalie = sElenco
End Function


' Numero di voci in un elenco prodotto da ElencoAnomalie
Public Function ContaVoci(ByVal sElenco As String) As Long
    Const SEP As String = vbCrLf & "  - "
    If Len(sElenco) = 0 Then Exit Function
    ContaVoci = (Len(sElenco) - Len(Replace(sElenco, SEP, ""))) \ Len(SEP)
End Function


' Tiene solo le prime nMax voci (un MsgBox mostra al massimo ~1000 caratteri)
Public Function Abbrevia(ByVal sElenco As String, ByVal nMax As Long) As String
    Const SEP As String = vbCrLf & "  - "
    Dim v As Variant, i As Long, n As Long, s As String
    v = Split(sElenco, SEP)
    n = UBound(v)                                   ' v(0) e' vuoto
    For i = 1 To n
        If i > nMax Then
            s = s & vbCrLf & "  ... e altre " & (n - nMax) & " (usa VerificaAssenze per l'elenco)"
            Exit For
        End If
        s = s & SEP & v(i)
    Next i
    Abbrevia = s
End Function


'-----------------------------------------------------------------------------
' Supporto
'-----------------------------------------------------------------------------

' Elimina le regole create in precedenza da questa macro (riconosciute
' dall'area esatta a cui si applicano), lasciando intatte tutte le altre.
' Esamina solo le regole che toccano le colonne E e N.
Private Sub RimuoviRegoleEsistenti(ByVal ws As Worksheet)
    Dim i As Long
    Dim sArea As String
    Dim fcs As FormatConditions
    Dim vArea As Variant

    For Each vArea In Array(AREA_ORE, AREA_TIPO)
        Set fcs = ws.Range(vArea).FormatConditions
        For i = fcs.Count To 1 Step -1
            sArea = ""
            On Error Resume Next
            sArea = fcs(i).AppliesTo.Address
            On Error GoTo 0
            If StrComp(sArea, AREA_ORE, vbTextCompare) = 0 Or _
               StrComp(sArea, AREA_TIPO, vbTextCompare) = 0 Or _
               StrComp(sArea, AREA_VECCHIA, vbTextCompare) = 0 Then
                fcs(i).Delete
            End If
        Next i
    Next vArea
End Sub


' Aggiunge una regola "formula" (gia' nella lingua di Excel) all'area,
' con riempimento rosso chiaro e testo rosso scuro in grassetto.
Private Sub AggiungiRegola(ByVal rng As Range, ByVal sFormulaLocale As String)
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormulaLocale)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
        .StopIfTrue = False
    End With
End Sub


' FormatConditions.Add vuole la formula nella lingua di Excel (in italiano:
' E, O, VAL.NUMERO e il punto e virgola). La converto scrivendola nella
' cella vuota O11 (stessa riga delle regole) e rileggendola con FormulaLocal;
' poi svuoto la cella senza toccarne la formattazione.
Private Function FormulaLocale(ByVal ws As Worksheet, ByVal sFormula As String) As String
    Dim c As Range
    Set c = ws.Range(CELLA_APPOGGIO)
    If Len(c.Formula) > 0 Then
        Err.Raise vbObjectError + 600, , "La cella " & CELLA_APPOGGIO & " del foglio " & _
                  ws.Name & " non e' vuota: impossibile usarla come appoggio."
    End If
    c.Formula = sFormula
    FormulaLocale = c.FormulaLocal
    c.ClearContents
End Function


' True se due orari consecutivi compilati sono in ordine inverso
Private Function OrariFuoriSequenza(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim k As Long
    Dim v1 As Variant, v2 As Variant
    For k = 6 To 8                       ' coppie F-G, G-H, H-I
        v1 = ws.Cells(r, k).Value
        v2 = ws.Cells(r, k + 1).Value
        If Not IsEmpty(v1) And Not IsEmpty(v2) And Not IsError(v1) And Not IsError(v2) Then
            If IsNumeric(v1) And IsNumeric(v2) Then
                If CDbl(v2) < CDbl(v1) Then
                    OrariFuoriSequenza = True
                    Exit Function
                End If
            End If
        End If
    Next k
End Function


Private Function EPermessoAOre(ByVal sTipo As String) As Boolean
    Select Case sTipo
        Case "Art. 23 (Visite)", "Art. 25 (Motivi pers.)", "Art. 34 (A recupero)"
            EPermessoAOre = True
    End Select
End Function


Private Function GiornoRiga(ByVal ws As Worksheet, ByVal r As Long) As String
    If Len(ws.Cells(r, 2).Text) > 0 Then
        GiornoRiga = ws.Cells(r, 2).Text
    Else
        GiornoRiga = "(riga " & r & ")"
    End If
End Function


Private Function OttieniFoglio(ByVal sNome As String) As Worksheet
    On Error Resume Next
    Set OttieniFoglio = ThisWorkbook.Worksheets(sNome)
    On Error GoTo 0
End Function
