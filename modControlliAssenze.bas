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
'   VerificaAssenze           -> elenca le anomalie presenti in tutti i mesi,
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
' Elenca le anomalie di tutti i mesi (non modifica nulla)
'-----------------------------------------------------------------------------
Public Sub VerificaAssenze()

    Dim vMesi As Variant, v As Variant
    Dim ws As Worksheet
    Dim r As Long, nAnomalie As Long
    Dim sTipo As String, vOre As Variant
    Dim bOre As Boolean
    Dim sElenco As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
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
                    nAnomalie = nAnomalie + 1
                ElseIf EPermessoAOre(sTipo) And Not bOre Then
                    sElenco = sElenco & vbCrLf & "  - " & ws.Name & " " & GiornoRiga(ws, r) & _
                              ": " & sTipo & " senza ore in Assenze"
                    nAnomalie = nAnomalie + 1
                End If
            Next r
        End If
    Next v

    If nAnomalie = 0 Then
        MsgBox "Nessuna anomalia nelle assenze dei 12 mesi.", vbInformation, "Verifica assenze"
    Else
        MsgBox "Trovate " & nAnomalie & " anomalie nelle assenze:" & sElenco, _
               vbExclamation, "Verifica assenze"
    End If
End Sub


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
