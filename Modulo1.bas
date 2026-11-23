Attribute VB_Name = "Modulo1"
Option Explicit

'=============================================================================
' ImportaDaPortaleWeb
' Legge un file .txt con il copia-incolla del cartellino dal portale web e
' scrive le timbrature (Entr1, Usc1, Entr2, Usc2) nel foglio mensile attivo.
'
' Rispetto alla versione precedente:
'   - funziona solo sui 12 fogli mensili (non su Dashboard, Festivita', ...);
'   - rimette la protezione SOLO se il foglio era protetto all'avvio;
'   - gestione errori: il file .txt viene sempre chiuso e la protezione
'     ripristinata anche in caso di problemi;
'   - legge correttamente file con fine riga Windows, Mac o Unix e con BOM;
'   - salta (e conta) gli orari non validi invece di fermarsi;
'   - ignora i giorni che non esistono nel mese (es. 31 in un mese di 30).
'=============================================================================

Sub ImportaDaPortaleWeb()

    ' Costanti di configurazione
    Const RIGA_OFFSET As Long = 10          ' riga = RIGA_OFFSET + giorno
    Const COL_ENTR1 As Long = 6             ' F
    Const COL_USC2 As Long = 9              ' I
    Const MAX_TIMBRATURE As Long = 4

    Dim ws As Worksheet
    Dim fd As FileDialog
    Dim filePath As String
    Dim fileNum As Integer
    Dim bFileAperto As Boolean
    Dim sTesto As String
    Dim vRighe As Variant
    Dim lineText As String, cleanedLine As String, orarioStr As String
    Dim regExGiorno As Object
    Dim giornoCorrente As Long, rowDest As Long
    Dim countTrovati As Long, countScartati As Long
    Dim numTimbrature As Long
    Dim i As Long, k As Long
    Dim isTimeLine As Boolean
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean
    Dim sScartati As String
    Dim dOrario As Date
    Dim bCompletato As Boolean
    Dim lErr As Long, sErr As String

    Set ws = ActiveSheet

    '--- Il foglio attivo deve essere un mese ---------------------------------
    If Not EFoglioMensile(ws.Name) Then
        MsgBox "Spostati su un foglio mensile (es. Ottobre) prima di avviare l'importazione.", _
               vbExclamation, "Importa dal portale"
        Exit Sub
    End If

    '--- Scelta del file -------------------------------------------------------
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Seleziona il file TXT con il copia-incolla dal portale"
    fd.Filters.Clear
    fd.Filters.Add "File di Testo", "*.txt"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub
    filePath = fd.SelectedItems(1)

    ' Riga del giorno (es. "1 Gio", "26 Lun")
    Set regExGiorno = CreateObject("VBScript.RegExp")
    regExGiorno.Pattern = "^([1-9]|[12][0-9]|3[01])\s+(Lun|Mar|Mer|Gio|Ven|Sab|Dom)"
    regExGiorno.IgnoreCase = True

    bScreenUpd = Application.ScreenUpdating
    eraProtetto = ws.ProtectContents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False

    '--- Lettura dell'intero file ---------------------------------------------
    ' (lettura binaria: legge tutto il file in una volta, qualunque sia il
    '  tipo di fine riga)
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    bFileAperto = True
    sTesto = Space$(LOF(fileNum))
    If LOF(fileNum) > 0 Then Get #fileNum, , sTesto
    Close #fileNum
    bFileAperto = False

    ' toglie l'eventuale BOM (UTF-8 o "Unicode" di Blocco note) e uniforma
    ' i fine riga
    If Left$(sTesto, 3) = Chr(239) & Chr(187) & Chr(191) Then sTesto = Mid$(sTesto, 4)
    If Left$(sTesto, 2) = Chr(255) & Chr(254) Then sTesto = Replace(Mid$(sTesto, 3), Chr(0), "")
    sTesto = Replace(sTesto, vbCrLf, vbLf)
    sTesto = Replace(sTesto, vbCr, vbLf)
    vRighe = Split(sTesto, vbLf)

    If eraProtetto Then ws.Unprotect

    '--- Analisi riga per riga -------------------------------------------------
    For k = LBound(vRighe) To UBound(vRighe)
        lineText = Trim$(Replace(vRighe(k), Chr(160), " "))

        ' 1. Riga con l'intestazione del giorno?
        If regExGiorno.Test(lineText) Then
            giornoCorrente = CLng(regExGiorno.Execute(lineText)(0).SubMatches(0))
        End If

        ' 2. Pulizia da spazi, tabulazioni e spazi non-breaking
        cleanedLine = Replace(lineText, " ", "")
        cleanedLine = Replace(cleanedLine, vbTab, "")

        ' 3. Riga di soli orari "hh:mmhh:mm..." (lunghezza multipla di 5)?
        If giornoCorrente > 0 And Len(cleanedLine) >= 5 And Len(cleanedLine) Mod 5 = 0 Then

            isTimeLine = True
            For i = 1 To Len(cleanedLine)
                If Not (Mid$(cleanedLine, i, 1) Like "[0-9]" Or Mid$(cleanedLine, i, 1) = ":") Then
                    isTimeLine = False
                    Exit For
                End If
            Next i

            If isTimeLine Then
                rowDest = RIGA_OFFSET + giornoCorrente

                ' giorno inesistente nel mese (colonna B vuota): si salta
                If Len(Trim$(ws.Cells(rowDest, 2).Text)) = 0 Then
                    sScartati = sScartati & vbCrLf & "  - giorno " & giornoCorrente & " (non esiste nel mese)"
                    countScartati = countScartati + 1
                Else
                    ' pulisce i vecchi orari del giorno
                    ws.Range(ws.Cells(rowDest, COL_ENTR1), ws.Cells(rowDest, COL_USC2)).ClearContents

                    numTimbrature = Len(cleanedLine) \ 5
                    If numTimbrature > MAX_TIMBRATURE Then numTimbrature = MAX_TIMBRATURE

                    For i = 0 To numTimbrature - 1
                        orarioStr = Mid$(cleanedLine, (i * 5) + 1, 5)
                        If LeggiOrario(orarioStr, dOrario) Then
                            ws.Cells(rowDest, COL_ENTR1 + i).Value = dOrario
                        Else
                            sScartati = sScartati & vbCrLf & "  - giorno " & giornoCorrente & _
                                        ": orario non valido """ & orarioStr & """"
                            countScartati = countScartati + 1
                        End If
                    Next i

                    ws.Range(ws.Cells(rowDest, COL_ENTR1), ws.Cells(rowDest, COL_USC2)).NumberFormat = "hh:mm"
                    countTrovati = countTrovati + 1
                End If

                ' evita sovrascritture accidentali fino al prossimo giorno
                giornoCorrente = 0
            End If
        End If
    Next k
    bCompletato = True

'--- Ripristino (eseguito SEMPRE) ---------------------------------------------
Uscita:
    On Error Resume Next
    If bFileAperto Then Close #fileNum
    If eraProtetto And Not ws.ProtectContents Then ws.Protect
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If bCompletato Then
        If countScartati > 0 Then
            MsgBox "Importazione completata nel foglio '" & ws.Name & "'." & vbCrLf & _
                   "Giorni elaborati: " & countTrovati & vbCrLf & vbCrLf & _
                   "Elementi saltati (" & countScartati & "):" & sScartati, _
                   vbExclamation, "Importa dal portale"
        Else
            MsgBox "Importazione completata!" & vbCrLf & _
                   "Elaborati " & countTrovati & " giorni nel foglio '" & ws.Name & "'.", _
                   vbInformation, "Importa dal portale"
        End If
    ElseIf lErr <> 0 Then
        MsgBox "Importazione interrotta." & vbCrLf & vbCrLf & _
               "Errore " & lErr & ": " & sErr & vbCrLf & _
               "Giorni elaborati prima dell'errore: " & countTrovati, _
               vbCritical, "Importa dal portale"
    End If
    Exit Sub

GestioneErrori:
    lErr = Err.Number
    sErr = Err.Description
    Resume Uscita
End Sub


'-----------------------------------------------------------------------------
' True se il nome corrisponde a uno dei 12 fogli mensili
'-----------------------------------------------------------------------------
Private Function EFoglioMensile(ByVal sNome As String) As Boolean
    Dim v As Variant
    For Each v In Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                        "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")
        If StrComp(sNome, CStr(v), vbTextCompare) = 0 Then
            EFoglioMensile = True
            Exit Function
        End If
    Next v
End Function


'-----------------------------------------------------------------------------
' Converte "hh:mm" in un orario. Restituisce False se il testo non e' un
' orario valido (es. "25:70" o "7:3x"), senza generare errori.
'-----------------------------------------------------------------------------
Private Function LeggiOrario(ByVal s As String, ByRef dOrario As Date) As Boolean
    Dim hh As Long, mm As Long
    If Not s Like "##:##" Then Exit Function
    hh = CLng(Left$(s, 2))
    mm = CLng(Right$(s, 2))
    If hh > 23 Or mm > 59 Then Exit Function
    dOrario = TimeSerial(hh, mm, 0)
    LeggiOrario = True
End Function
