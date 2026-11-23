Attribute VB_Name = "modExportPDF"
Option Explicit

'=============================================================================
' Modulo  : modExportPDF
' Scopo   : esporta in un unico PDF il foglio "Dashboard" e i 12 fogli mensili
'           (Gennaio -> Dicembre), nell'ordine esatto indicato.
' Output  : <cartella del file>\Report_Annuale_AAAA-MM-GG.pdf
'
' NOTA IMPORTANTE SULL'ORDINE
'   Quando si esporta un gruppo di fogli, Excel scrive le pagine nel PDF
'   seguendo l'ordine delle SCHEDE nella cartella di lavoro, non l'ordine
'   dell'array passato a Sheets(...). Se le schede non sono gia' nell'ordine
'   richiesto, la macro le riordina temporaneamente e poi ripristina la
'   disposizione originale (le formule non vengono toccate).
'=============================================================================

' True  = imposta margini ridotti (vedi costanti MARGINE_* sotto)
' False = lascia invariati i margini attuali di ogni foglio
Private Const USA_MARGINI_STRETTI As Boolean = True

' True  = ogni foglio su UNA sola pagina (larghezza E altezza)
' False = una pagina in larghezza, altezza libera (piu' pagine se serve)
Private Const UNA_PAGINA_PER_FOGLIO As Boolean = True

' Margini in centimetri (usati solo se USA_MARGINI_STRETTI = True).
' Sopra/sotto sono ridotti al minimo perche' i fogli mensili sono "alti"
' (fino a 41 righe) e in orizzontale lo spazio verticale e' il piu' scarso.
Private Const MARGINE_SX_DX_CM As Double = 0.8
Private Const MARGINE_SOPRA_SOTTO_CM As Double = 0.8
Private Const MARGINE_INTEST_PIEDE_CM As Double = 0.3

' Prefisso del nome del file PDF
Private Const PREFISSO_FILE As String = "Report_Annuale_"


'-----------------------------------------------------------------------------
' Macro principale: da assegnare al pulsante sul foglio Dashboard
'-----------------------------------------------------------------------------
Public Sub EsportaReportPDF()

    Dim vElenco As Variant              ' elenco dei fogli desiderati, in ordine
    Dim aFogli() As String              ' fogli realmente esportabili
    Dim nFogli As Long                  ' quanti sono
    Dim sSaltati As String              ' fogli mancanti o nascosti (per il report)
    Dim aOrdineOriginale() As String    ' ordine originale di tutte le schede
    Dim bRiordinato As Boolean          ' True se abbiamo spostato delle schede
    Dim shIniziale As Object            ' foglio attivo all'avvio
    Dim bScreenUpdOrig As Boolean       ' stato originale di ScreenUpdating
    Dim bEventsOrig As Boolean          ' stato originale di EnableEvents
    Dim sPercorso As String             ' percorso completo del PDF
    Dim bEsportato As Boolean           ' True se l'export e' andato a buon fine
    Dim lErrNum As Long, sErrDesc As String
    Dim ws As Worksheet
    Dim sMsg As String
    Dim i As Long

    '--- 1) Il file deve essere gia' stato salvato su disco -------------------
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "La cartella di lavoro non e' ancora stata salvata." & vbCrLf & _
               "Salva il file e poi rilancia l'esportazione.", _
               vbExclamation, "Esporta PDF"
        Exit Sub
    End If

    ' Con OneDrive/SharePoint ThisWorkbook.Path puo' restituire un indirizzo
    ' web (https://...) su cui ExportAsFixedFormat non puo' scrivere.
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Il file e' aperto da un percorso web (OneDrive/SharePoint):" & vbCrLf & _
               ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "Aprilo dalla cartella sincronizzata sul PC e riprova.", _
               vbExclamation, "Esporta PDF"
        Exit Sub
    End If

    '--- 2) Elenco dei fogli nell'ordine voluto e verifica esistenza ----------
    vElenco = Array("Dashboard", "Gennaio", "Febbraio", "Marzo", "Aprile", _
                    "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", _
                    "Ottobre", "Novembre", "Dicembre")

    ReDim aFogli(0 To UBound(vElenco))
    nFogli = 0
    For i = LBound(vElenco) To UBound(vElenco)
        Set ws = OttieniFoglio(CStr(vElenco(i)))
        If ws Is Nothing Then
            sSaltati = sSaltati & vbCrLf & "  - " & vElenco(i) & " (non esiste)"
        ElseIf ws.Visible <> xlSheetVisible Then
            sSaltati = sSaltati & vbCrLf & "  - " & vElenco(i) & " (nascosto)"
        Else
            aFogli(nFogli) = ws.Name
            nFogli = nFogli + 1
        End If
    Next i

    If nFogli = 0 Then
        MsgBox "Nessuno dei fogli previsti e' presente e visibile." & _
               vbCrLf & "Esportazione annullata.", vbExclamation, "Esporta PDF"
        Exit Sub
    End If
    ReDim Preserve aFogli(0 To nFogli - 1)

    '--- 3) Nome e percorso del file PDF --------------------------------------
    ' Format con "yyyy-mm-dd" funziona anche con Excel in italiano
    sPercorso = ThisWorkbook.Path & Application.PathSeparator & _
                PREFISSO_FILE & Format$(Date, "yyyy-mm-dd") & ".pdf"

    '--- 4) Memorizzo lo stato iniziale e attivo la gestione errori ----------
    Set shIniziale = ThisWorkbook.ActiveSheet
    bScreenUpdOrig = Application.ScreenUpdating
    bEventsOrig = Application.EnableEvents

    On Error GoTo GestioneErrori

    Application.ScreenUpdating = False
    Application.EnableEvents = False     ' evita che Select/Move scatenino eventi
    Application.Cursor = xlWait

    '--- 5) Impaginazione di ciascun foglio -----------------------------------
    ' PrintCommunication = False raggruppa le modifiche a PageSetup e rende
    ' l'operazione molto piu' veloce (va SEMPRE rimesso a True).
    ImpostaPrintCommunication False
    For i = 0 To nFogli - 1
        With ThisWorkbook.Worksheets(aFogli(i)).PageSetup
            .Orientation = xlLandscape
            .Zoom = False                ' obbligatorio per usare FitToPages
            .FitToPagesWide = 1          ' tutte le colonne in una pagina
            If UNA_PAGINA_PER_FOGLIO Then
                .FitToPagesTall = 1      ' tutte le righe in una pagina
            Else
                .FitToPagesTall = False  ' altezza libera (piu' pagine se serve)
            End If
            .CenterHorizontally = True   ' tabella centrata sul foglio
            .CenterVertically = False
            If USA_MARGINI_STRETTI Then
                .LeftMargin = Application.CentimetersToPoints(MARGINE_SX_DX_CM)
                .RightMargin = Application.CentimetersToPoints(MARGINE_SX_DX_CM)
                .TopMargin = Application.CentimetersToPoints(MARGINE_SOPRA_SOTTO_CM)
                .BottomMargin = Application.CentimetersToPoints(MARGINE_SOPRA_SOTTO_CM)
                .HeaderMargin = Application.CentimetersToPoints(MARGINE_INTEST_PIEDE_CM)
                .FooterMargin = Application.CentimetersToPoints(MARGINE_INTEST_PIEDE_CM)
            End If
        End With
    Next i
    ImpostaPrintCommunication True

    '--- 6) Garantisco l'ordine delle schede (solo se necessario) -------------
    If Not SchedeInOrdine(aFogli) Then
        If ThisWorkbook.ProtectStructure Then
            Err.Raise vbObjectError + 513, , _
                "Le schede non sono nell'ordine richiesto e la struttura " & _
                "della cartella e' protetta: impossibile riordinarle."
        End If
        ' salvo l'ordine attuale di TUTTE le schede per il ripristino
        ReDim aOrdineOriginale(1 To ThisWorkbook.Sheets.Count)
        For i = 1 To ThisWorkbook.Sheets.Count
            aOrdineOriginale(i) = ThisWorkbook.Sheets(i).Name
        Next i
        bRiordinato = True
        ' sposto i fogli in fondo, uno dopo l'altro, nell'ordine desiderato
        For i = 0 To nFogli - 1
            ThisWorkbook.Sheets(aFogli(i)).Move _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Next i
    End If

    '--- 7) Selezione del gruppo ed esportazione -----------------------------
    ThisWorkbook.Activate
    ThisWorkbook.Sheets(aFogli).Select   ' raggruppa i fogli

    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=sPercorso, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    bEsportato = True

'--- 8) Ripristino dello stato iniziale (eseguito SEMPRE) --------------------
Uscita:
    On Error Resume Next                 ' il ripristino non deve mai bloccarsi
    ImpostaPrintCommunication True
    If bRiordinato Then RipristinaOrdineSchede aOrdineOriginale
    ThisWorkbook.Activate
    If Not shIniziale Is Nothing Then shIniziale.Select   ' separa il gruppo
    Application.Cursor = xlDefault
    Application.EnableEvents = bEventsOrig
    Application.ScreenUpdating = bScreenUpdOrig
    On Error GoTo 0

    '--- 9) Messaggio finale ---------------------------------------------------
    If bEsportato Then
        sMsg = "PDF creato correttamente (" & nFogli & " fogli):" & _
               vbCrLf & vbCrLf & sPercorso
        If Len(sSaltati) > 0 Then
            sMsg = sMsg & vbCrLf & vbCrLf & "Fogli saltati:" & sSaltati
        End If
        sMsg = sMsg & vbCrLf & vbCrLf & "Vuoi aprire il file adesso?"

        If MsgBox(sMsg, vbInformation + vbYesNo, "Esporta PDF") = vbYes Then
            On Error Resume Next
            ThisWorkbook.FollowHyperlink sPercorso
            If Err.Number <> 0 Then
                MsgBox "Impossibile aprire il PDF automaticamente." & vbCrLf & _
                       "Lo trovi qui: " & sPercorso, vbExclamation, "Esporta PDF"
            End If
            On Error GoTo 0
        End If
    ElseIf lErrNum <> 0 Then
        MsgBox "Esportazione non riuscita." & vbCrLf & vbCrLf & _
               "Errore " & lErrNum & ": " & sErrDesc & vbCrLf & vbCrLf & _
               "Verifica che il PDF non sia gia' aperto in un altro programma " & _
               "e di avere i permessi di scrittura nella cartella.", _
               vbCritical, "Esporta PDF"
    End If
    Exit Sub

GestioneErrori:
    ' salvo i dettagli PRIMA che On Error Resume Next azzeri l'oggetto Err
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita

End Sub


'-----------------------------------------------------------------------------
' Restituisce il foglio di lavoro con il nome indicato, oppure Nothing
' se non esiste (senza generare errori).
'-----------------------------------------------------------------------------
Private Function OttieniFoglio(ByVal sNome As String) As Worksheet
    On Error Resume Next
    Set OttieniFoglio = ThisWorkbook.Worksheets(sNome)
    On Error GoTo 0
End Function


'-----------------------------------------------------------------------------
' True se le schede dell'array compaiono nella cartella in ordine crescente
' di posizione (cioe' l'export rispettera' gia' l'ordine richiesto).
'-----------------------------------------------------------------------------
Private Function SchedeInOrdine(aNomi() As String) As Boolean
    Dim i As Long
    SchedeInOrdine = True
    For i = LBound(aNomi) + 1 To UBound(aNomi)
        If ThisWorkbook.Sheets(aNomi(i)).Index < _
           ThisWorkbook.Sheets(aNomi(i - 1)).Index Then
            SchedeInOrdine = False
            Exit Function
        End If
    Next i
End Function


'-----------------------------------------------------------------------------
' Riporta tutte le schede nell'ordine memorizzato prima del riordino.
' Sposta solo le schede fuori posto.
'-----------------------------------------------------------------------------
Private Sub RipristinaOrdineSchede(aOrdine() As String)
    Dim i As Long
    For i = LBound(aOrdine) To UBound(aOrdine)
        If ThisWorkbook.Sheets(i).Name <> aOrdine(i) Then
            ThisWorkbook.Sheets(aOrdine(i)).Move Before:=ThisWorkbook.Sheets(i)
        End If
    Next i
End Sub


'-----------------------------------------------------------------------------
' Imposta Application.PrintCommunication ignorando l'errore sulle versioni
' di Excel che non supportano la proprieta' (precedenti al 2010).
'-----------------------------------------------------------------------------
Private Sub ImpostaPrintCommunication(ByVal bValore As Boolean)
    On Error Resume Next
    Application.PrintCommunication = bValore
    On Error GoTo 0
End Sub
