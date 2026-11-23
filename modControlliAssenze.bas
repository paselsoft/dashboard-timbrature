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
Private Const AREA_CONTROLLO As String = "$E$11:$E$41,$N$11:$N$41"

' Regole in sintassi inglese, riferite alla riga 11 (prima riga dell'area).
Private Const REGOLA_ORE_SENZA_TIPO As String = _
    "=AND(ISNUMBER($E11),$E11>0,$N11="""")"
Private Const REGOLA_TIPO_SENZA_ORE As String = _
    "=AND(OR($N11=""Art. 23 (Visite)"",$N11=""Art. 25 (Motivi pers.)""," & _
    "$N11=""Art. 34 (A recupero)""),OR(NOT(ISNUMBER($E11)),$E11=0))"


'-----------------------------------------------------------------------------
' Installa la formattazione condizionale nei 12 mesi
'-----------------------------------------------------------------------------
Public Sub AggiungiControlliAssenze()

    Dim vMesi As Variant, v As Variant
    Dim ws As Worksheet, shIniziale As Object
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean
    Dim nFogli As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False

    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect

            ' FormatConditions.Add interpreta i riferimenti relativi rispetto
            ' alla cella attiva: attivo il foglio e seleziono E11.
            ws.Activate
            ws.Range("E" & RIGA_PRIMA).Select

            RimuoviRegoleEsistenti ws
            AggiungiRegola ws, REGOLA_ORE_SENZA_TIPO
            AggiungiRegola ws, REGOLA_TIPO_SENZA_ORE

            ws.Range("A1").Select
            If eraProtetto Then ws.Protect
            nFogli = nFogli + 1
        End If
    Next v
    sFoglioErr = ""

Uscita:
    On Error Resume Next
    If Not shIniziale Is Nothing Then shIniziale.Activate
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
    If eraProtetto And Not ws Is Nothing Then
        On Error Resume Next
        ws.Protect
    End If
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
Private Sub RimuoviRegoleEsistenti(ByVal ws As Worksheet)
    Dim i As Long
    Dim sArea As String
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        sArea = ""
        On Error Resume Next
        sArea = ws.Cells.FormatConditions(i).AppliesTo.Address
        On Error GoTo 0
        If StrComp(sArea, AREA_CONTROLLO, vbTextCompare) = 0 Then
            ws.Cells.FormatConditions(i).Delete
        End If
    Next i
End Sub


' Aggiunge una regola "formula" all'area di controllo, con riempimento rosso
' chiaro e testo rosso scuro, in cima alla lista delle priorita'.
Private Sub AggiungiRegola(ByVal ws As Worksheet, ByVal sFormulaInglese As String)
    Dim fc As FormatCondition
    Set fc = ws.Range(AREA_CONTROLLO).FormatConditions.Add( _
                Type:=xlExpression, Formula1:=FormulaLocale(ws, sFormulaInglese))
    With fc
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
        .StopIfTrue = False
        .SetFirstPriority
    End With
End Sub


' FormatConditions.Add vuole la formula nella lingua di Excel (in italiano:
' E, O, VAL.NUMERO e il punto e virgola). La converto scrivendola in una
' cella temporanea della riga 11 e rileggendola con FormulaLocal.
Private Function FormulaLocale(ByVal ws As Worksheet, ByVal sFormula As String) As String
    Dim c As Range
    Set c = ws.Cells(RIGA_PRIMA, ws.Columns.Count)   ' ultima colonna, riga 11
    c.Formula = sFormula
    FormulaLocale = c.FormulaLocal
    c.Clear
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
