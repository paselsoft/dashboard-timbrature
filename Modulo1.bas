Attribute VB_Name = "Modulo1"
Option Explicit

'=============================================================================
' ImportaDaPortaleWeb
' Importa il cartellino del portale web nel foglio mensile attivo, leggendo
' il testo COPIATO (appunti: Ctrl+A e Ctrl+C sul portale) oppure un file .txt.
' Prima di scrivere controlla che il cartellino sia del mese del foglio
' (giorni della settimana, numero di giorni e festivi). Dopo un'importazione
' dagli appunti salva il testo in "MM.Mese.txt" nella cartella del file.
'
' Rispetto alla versione precedente:
'   - funziona solo sui 12 fogli mensili (non su Dashboard, Festivita', ...);
'   - rimette la protezione SOLO se il foglio era protetto all'avvio;
'   - gestione errori: il file .txt viene sempre chiuso e la protezione
'     ripristinata anche in caso di problemi;
'   - legge correttamente file con fine riga Windows, Mac o Unix e con BOM;
'   - salta (e conta) gli orari non validi invece di fermarsi;
'   - ignora i giorni che non esistono nel mese (es. 31 in un mese di 30).
'
' Versione 3 - importa anche le ASSENZE con le ore (righe del portale come
' "9:00Ferie dalle 07:00 alle 16:00"):
'     Ferie                        -> Ferie
'     Ferie Esercizio Precedente   -> Ferie anno prec.
'     Malattia                     -> Malattia
'     ... VISITE ...               -> Art. 23 (Visite)      + ore in Assenze
'     PERMESSO PERSONALE / ART. 25 -> Art. 25 (Motivi pers.) + ore in Assenze
'     ART. 34 / RECUPERO           -> Art. 34 (A recupero)  + ore in Assenze
'     ... 104 ...                  -> L. 104
'   Extrapresenza, Flessibilita' e Ore non autorizzate vengono ignorate
'   (le calcola il file). Le richieste senza ore (es. ferie pianificate)
'   non vengono importate. Un giorno gia' segnato "Ferie anno prec." nel file
'   resta tale anche se il portale lo indica come Ferie dell'anno.
' - non cancella orari gia' presenti se il portale ne ha di meno (file .txt
'   salvato a giornata in corso): il giorno viene solo segnalato;
' - le timbrature di un giorno possono essere su piu' righe (il portale va a
'   capo quando sono molte): vengono raccolte tutte;
' - giorni con piu' di 4 timbrature: le coppie del pomeriggio vengono
'   accorpate in una sola con lo stesso tempo lavorato (orari reali in una
'   nota su Entr 2); segnala i giustificativi non riconosciuti.
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
    Dim regExGiust As Object, mGiust As Object
    Dim bOrariFatti As Boolean, nGiaPresenti As Long
    Dim sOrariGiorno As String
    Dim iScelta As VbMsgBoxResult, bDaAppunti As Boolean
    Dim sMotivo As String, bConferma As Boolean
    Dim sTestoAppunti As String, sFileSalvato As String
    Dim sDesc As String, sTipo As String, bConOre As Boolean
    Dim sPrecedente As String, sTipoImportato As String
    Dim countAssenze As Long, sAvvisi As String

    Set ws = ActiveSheet

    '--- Il foglio attivo deve essere un mese ---------------------------------
    If Not EFoglioMensile(ws.Name) Then
        MsgBox "Spostati su un foglio mensile (es. Ottobre) prima di avviare l'importazione.", _
               vbExclamation, "Importa dal portale"
        Exit Sub
    End If

    '--- Origine: appunti (testo copiato dal portale) oppure file .txt --------
    iScelta = MsgBox("Importare dagli APPUNTI il cartellino copiato dal portale?" & vbCrLf & vbCrLf & _
                     "  Si' = testo appena copiato (sul portale: Ctrl+A, poi Ctrl+C)" & vbCrLf & _
                     "  No  = scegli un file .txt", _
                     vbQuestion + vbYesNoCancel, "Importa dal portale")
    If iScelta = vbCancel Then Exit Sub

    If iScelta = vbYes Then
        sTesto = LeggiAppunti()
        If Len(Trim$(sTesto)) = 0 Then
            MsgBox "Negli appunti non c'e' testo." & vbCrLf & vbCrLf & _
                   "Apri il cartellino del mese sul portale, premi Ctrl+A e poi Ctrl+C, " & _
                   "quindi riprova.", vbExclamation, "Importa dal portale"
            Exit Sub
        End If
        bDaAppunti = True
        sTestoAppunti = sTesto          ' copia originale, da salvare in .txt
    Else
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        fd.Title = "Seleziona il file TXT con il copia-incolla dal portale"
        fd.Filters.Clear
        fd.Filters.Add "File di Testo", "*.txt"
        fd.AllowMultiSelect = False
        If fd.Show <> -1 Then Exit Sub
        filePath = fd.SelectedItems(1)
    End If

    ' Riga del giorno (es. "1 Gio", "26 Lun")
    Set regExGiorno = CreateObject("VBScript.RegExp")
    regExGiorno.Pattern = "^([1-9]|[12][0-9]|3[01])\s+(Lun|Mar|Mer|Gio|Ven|Sab|Dom)"
    regExGiorno.IgnoreCase = True

    ' Riga di un giustificativo con le ore (es. "6:00PERMESSO PERSONALE ...")
    Set regExGiust = CreateObject("VBScript.RegExp")
    regExGiust.Pattern = "^(\d{1,2}):(\d{2})\s*([A-Za-z].*)$"

    bScreenUpd = Application.ScreenUpdating
    eraProtetto = ws.ProtectContents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False

    '--- Lettura dell'intero file (se non si importa dagli appunti) -----------
    ' (lettura binaria: legge tutto il file in una volta, qualunque sia il
    '  tipo di fine riga)
    If Not bDaAppunti Then
        fileNum = FreeFile
        Open filePath For Binary Access Read As #fileNum
        bFileAperto = True
        sTesto = Space$(LOF(fileNum))
        If LOF(fileNum) > 0 Then Get #fileNum, , sTesto
        Close #fileNum
        bFileAperto = False
    End If

    ' toglie l'eventuale BOM (UTF-8 o "Unicode" di Blocco note) e uniforma
    ' i fine riga
    If Left$(sTesto, 3) = Chr(239) & Chr(187) & Chr(191) Then sTesto = Mid$(sTesto, 4)
    If Left$(sTesto, 2) = Chr(255) & Chr(254) Then sTesto = Replace(Mid$(sTesto, 3), Chr(0), "")
    sTesto = Replace(sTesto, vbCrLf, vbLf)
    sTesto = Replace(sTesto, vbCr, vbLf)
    vRighe = Split(sTesto, vbLf)

    '--- Il cartellino corrisponde al mese del foglio? -------------------------
    If Not ControllaMese(ws, vRighe, regExGiorno, sMotivo, bConferma) Then
        MsgBox sMotivo & vbCrLf & vbCrLf & "Nessuna modifica effettuata.", _
               vbExclamation, "Importa dal portale"
        GoTo Uscita
    End If
    If bConferma Then
        If MsgBox(sMotivo & vbCrLf & vbCrLf & "Importare nel foglio " & ws.Name & "?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Importa dal portale") <> vbYes Then GoTo Uscita
    End If

    If eraProtetto Then ws.Unprotect

    '--- Analisi riga per riga -------------------------------------------------
    For k = LBound(vRighe) To UBound(vRighe)
        lineText = Trim$(Replace(vRighe(k), Chr(160), " "))
        lineText = Trim$(Replace(lineText, vbTab, " "))

        ' 1. Intestazione del giorno: da qui in poi le righe sono di quel giorno
        If regExGiorno.Test(lineText) Then
            GoSub ElaboraOrari              ' completa gli orari del giorno precedente
            giornoCorrente = CLng(regExGiorno.Execute(lineText)(0).SubMatches(0))
            bOrariFatti = False
            sOrariGiorno = ""
            sTipoImportato = ""
            GoTo RigaSuccessiva
        End If
        If giornoCorrente = 0 Then GoTo RigaSuccessiva
        rowDest = RIGA_OFFSET + giornoCorrente

        ' 2. Righe di soli orari "hh:mmhh:mm...": si RACCOLGONO, perche' il portale
        '    puo' andare a capo (es. 6 timbrature su due righe)
        cleanedLine = Replace(lineText, " ", "")
        If Len(cleanedLine) = 0 Then GoTo RigaSuccessiva
        isTimeLine = False
        If Not bOrariFatti And Len(cleanedLine) >= 5 And Len(cleanedLine) Mod 5 = 0 Then
            isTimeLine = True
            For i = 1 To Len(cleanedLine)
                If Not (Mid$(cleanedLine, i, 1) Like "[0-9]" Or Mid$(cleanedLine, i, 1) = ":") Then
                    isTimeLine = False
                    Exit For
                End If
            Next i
        End If
        If isTimeLine Then
            sOrariGiorno = sOrariGiorno & cleanedLine
            GoTo RigaSuccessiva
        End If

        ' qualsiasi altra riga chiude l'elenco degli orari del giorno
        GoSub ElaboraOrari

        ' 3. Giustificativo con le ore (es. "9:00Ferie dalle 07:00 alle 16:00")
        If regExGiust.Test(lineText) Then
            Set mGiust = regExGiust.Execute(lineText)(0)
            sDesc = mGiust.SubMatches(2)
            If InStr(1, sDesc, " dalle ", vbTextCompare) > 0 Then
                sDesc = Left$(sDesc, InStr(1, sDesc, " dalle ", vbTextCompare) - 1)
            End If
            sDesc = Trim$(sDesc)
            sTipo = MappaGiustificativo(sDesc, bConOre)

            If sTipo = "?" Then
                sAvvisi = sAvvisi & vbCrLf & "  - giorno " & giornoCorrente & _
                          ": giustificativo non riconosciuto " & Chr(34) & sDesc & Chr(34)
            ElseIf Len(sTipo) > 0 Then
                If Len(Trim$(ws.Cells(rowDest, 2).Text)) > 0 Then
                    If Len(sTipoImportato) > 0 And sTipoImportato <> sTipo Then
                        sAvvisi = sAvvisi & vbCrLf & "  - giorno " & giornoCorrente & _
                                  ": due giustificativi (" & sTipoImportato & ", " & sTipo & _
                                  "), tenuto l'ultimo: verifica"
                    End If
                    sPrecedente = Trim$(CStr(ws.Cells(rowDest, 14).Value))
                    ' Ferie gia' imputate all'anno precedente nel file (es. criterio
                    ' "prima le ferie vecchie" indicato dall'ufficio del personale):
                    ' se il portale le segna come Ferie dell'anno, si mantiene la scelta del file.
                    If sTipo = "Ferie" And sPrecedente = "Ferie anno prec." Then sTipo = sPrecedente
                    If Len(sPrecedente) > 0 And sPrecedente <> sTipo Then
                        sAvvisi = sAvvisi & vbCrLf & "  - giorno " & giornoCorrente & ": " & _
                                  sPrecedente & " sostituito con " & sTipo
                    End If
                    ws.Cells(rowDest, 14).Value = sTipo
                    If bConOre Then
                        ws.Cells(rowDest, 5).Value = TimeSerial(CLng(mGiust.SubMatches(0)), _
                                                                CLng(mGiust.SubMatches(1)), 0)
                    Else
                        ws.Cells(rowDest, 5).ClearContents   ' giornata intera: niente ore
                    End If
                    sTipoImportato = sTipo
                    countAssenze = countAssenze + 1
                End If
            End If
        End If

RigaSuccessiva:
    Next k
    GoSub ElaboraOrari                  ' orari dell'ultimo giorno del file
    bCompletato = True

    ' Importazione dagli appunti riuscita: salvo anche il testo in un file
    ' .txt nella cartella del file Excel (es. 01.Gennaio.txt), come archivio.
    If bDaAppunti Then sFileSalvato = SalvaTestoMese(ws, sTestoAppunti)

'--- Ripristino (eseguito SEMPRE) ---------------------------------------------
Uscita:
    On Error Resume Next
    If bFileAperto Then Close #fileNum
    If eraProtetto And Not ws.ProtectContents Then ws.Protect
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If bCompletato Then
        MsgBox "Importazione completata nel foglio '" & ws.Name & "'." & vbCrLf & _
               "Giorni con timbrature: " & countTrovati & vbCrLf & _
               "Assenze importate: " & countAssenze & _
               IIf(bDaAppunti, vbCrLf & IIf(Len(sFileSalvato) > 0, "Testo salvato in: " & sFileSalvato, _
                   "Testo NON salvato in .txt (cartella non scrivibile)"), "") & _
               IIf(countScartati > 0, vbCrLf & vbCrLf & "Elementi saltati (" & countScartati & "):" & sScartati, "") & _
               IIf(Len(sAvvisi) > 0, vbCrLf & vbCrLf & "Da verificare:" & sAvvisi, ""), _
               IIf(countScartati > 0 Or Len(sAvvisi) > 0, vbExclamation, vbInformation), _
               "Importa dal portale"
    ElseIf lErr <> 0 Then
        MsgBox "Importazione interrotta." & vbCrLf & vbCrLf & _
               "Errore " & lErr & ": " & sErr & vbCrLf & _
               "Giorni elaborati prima dell'errore: " & countTrovati, _
               vbCritical, "Importa dal portale"
    End If
    Exit Sub

'--- Subroutine: scrive nel foglio gli orari raccolti per il giorno corrente ---
ElaboraOrari:
    If Len(sOrariGiorno) = 0 Or giornoCorrente = 0 Then Return
    cleanedLine = sOrariGiorno
    sOrariGiorno = ""
    bOrariFatti = True
    rowDest = RIGA_OFFSET + giornoCorrente
        If Len(Trim$(ws.Cells(rowDest, 2).Text)) = 0 Then
            sScartati = sScartati & vbCrLf & "  - giorno " & giornoCorrente & " (non esiste nel mese)"
            countScartati = countScartati + 1
            Return
        End If

        numTimbrature = Len(cleanedLine) \ 5
        nGiaPresenti = Application.WorksheetFunction.Count( _
                       ws.Range(ws.Cells(rowDest, COL_ENTR1), ws.Cells(rowDest, COL_USC2)))
        If numTimbrature < nGiaPresenti Then
            ' file salvato a giornata in corso: non cancello orari gia' inseriti
            sAvvisi = sAvvisi & vbCrLf & "  - giorno " & giornoCorrente & ": il portale ha " & _
                      numTimbrature & " timbrature, nel file ce ne sono " & nGiaPresenti & _
                      " (orari NON modificati)"
            Return
        End If
        If numTimbrature > MAX_TIMBRATURE Then
            ' Piu' di 4 timbrature (es. uscita e rientro per una visita nel
            ' pomeriggio): le coppie dopo la prima vengono accorpate in una sola,
            ' con lo stesso tempo lavorato; gli orari reali vanno in una nota.
            If ImportaPiuCoppie(ws, rowDest, cleanedLine, numTimbrature) Then
                sAvvisi = sAvvisi & vbCrLf & "  - giorno " & giornoCorrente & ": " & numTimbrature & _
                          " timbrature accorpate in 2 coppie (orari reali nella nota su Entr 2)"
                countTrovati = countTrovati + 1
                Return
            End If
            sAvvisi = sAvvisi & vbCrLf & "  - giorno " & giornoCorrente & ": " & numTimbrature & _
                      " timbrature non accorpabili, importate le prime " & MAX_TIMBRATURE & _
                      " (sistemale a mano)"
            numTimbrature = MAX_TIMBRATURE
        End If

        ws.Range(ws.Cells(rowDest, COL_ENTR1), ws.Cells(rowDest, COL_USC2)).ClearContents
        For i = 0 To numTimbrature - 1
            orarioStr = Mid$(cleanedLine, (i * 5) + 1, 5)
            If LeggiOrario(orarioStr, dOrario) Then
                ws.Cells(rowDest, COL_ENTR1 + i).Value = dOrario
            Else
                sScartati = sScartati & vbCrLf & "  - giorno " & giornoCorrente & _
                            ": orario non valido " & Chr(34) & orarioStr & Chr(34)
                countScartati = countScartati + 1
            End If
        Next i
        ws.Range(ws.Cells(rowDest, COL_ENTR1), ws.Cells(rowDest, COL_USC2)).NumberFormat = "hh:mm"
        countTrovati = countTrovati + 1
        Return

GestioneErrori:
    lErr = Err.Number
    sErr = Err.Description
    Resume Uscita
End Sub


'-----------------------------------------------------------------------------
' Giorno con piu' di 4 timbrature (numero pari, orari in ordine crescente):
'   Entr 1 / Usc 1 = prima coppia (reale)
'   Usc 2          = ultima uscita (reale)
'   Entr 2         = Usc 2 meno il tempo lavorato in TUTTE le coppie successive
' Cosi' il lavorato del giorno resta identico a quello del portale.
' Esempio: 07:38 13:38 14:08 14:15 15:55 17:08
'   pomeriggio = 0:07 + 1:13 = 1:20  ->  Entr 2 = 17:08 - 1:20 = 15:48
' Sulla cella Entr 2 viene aggiunta una nota con gli orari reali.
' Restituisce False (e non scrive nulla) se i dati non lo permettono.
'-----------------------------------------------------------------------------
Private Function ImportaPiuCoppie(ByVal ws As Worksheet, ByVal r As Long, _
                                  ByVal sOrari As String, ByVal n As Long) As Boolean
    Dim t() As Date, i As Long
    Dim dPom As Double, dEntr2 As Double
    Dim sNota As String

    If n Mod 2 <> 0 Then Exit Function
    ReDim t(1 To n)
    For i = 1 To n
        If Not LeggiOrario(Mid$(sOrari, (i - 1) * 5 + 1, 5), t(i)) Then Exit Function
        If i > 1 Then
            If t(i) < t(i - 1) Then Exit Function
        End If
    Next i

    For i = 3 To n - 1 Step 2
        dPom = dPom + (CDbl(t(i + 1)) - CDbl(t(i)))
        sNota = sNota & IIf(i > 3, " + ", "") & Format$(t(i), "hh:mm") & "-" & Format$(t(i + 1), "hh:mm")
    Next i
    dEntr2 = Round((CDbl(t(n)) - dPom) * 1440, 0) / 1440
    If dEntr2 < CDbl(t(2)) Then Exit Function

    With ws
        .Range(.Cells(r, 6), .Cells(r, 9)).ClearContents
        .Cells(r, 6).Value = t(1)
        .Cells(r, 7).Value = t(2)
        .Cells(r, 8).Value = CDate(dEntr2)
        .Cells(r, 9).Value = t(n)
        .Range(.Cells(r, 6), .Cells(r, 9)).NumberFormat = "hh:mm"
        If Not .Cells(r, 8).Comment Is Nothing Then .Cells(r, 8).Comment.Delete
        .Cells(r, 8).AddComment "Timbrature reali del pomeriggio: " & sNota & vbLf & _
                                "(Entr 2 calcolata dall'importazione)"
    End With
    ImportaPiuCoppie = True
End Function


'-----------------------------------------------------------------------------
' Salva il testo del cartellino nella cartella del file Excel con il nome
' "MM.Mese.txt" (es. 01.Gennaio.txt), sostituendo l'eventuale file esistente.
' Restituisce il nome del file salvato, oppure "" se non e' stato possibile.
'-----------------------------------------------------------------------------
Private Function SalvaTestoMese(ByVal ws As Worksheet, ByVal sTesto As String) As String
    Dim sNome As String, sPercorso As String, f As Integer
    Dim lMese As Long

    On Error GoTo NonSalvato
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Function

    lMese = 0
    If IsNumeric(ws.Range("A1").Value) Then lMese = CLng(ws.Range("A1").Value)
    If lMese < 1 Or lMese > 12 Then Exit Function

    sNome = Format$(lMese, "00") & "." & ws.Name & ".txt"
    sPercorso = ThisWorkbook.Path & Application.PathSeparator & sNome

    ' fine riga in stile Windows, cosi' il file si apre bene con Blocco note
    sTesto = Replace(sTesto, vbCrLf, vbLf)
    sTesto = Replace(sTesto, vbCr, vbLf)
    sTesto = Replace(sTesto, vbLf, vbCrLf)

    f = FreeFile
    Open sPercorso For Output As #f
    Print #f, sTesto;
    Close #f
    SalvaTestoMese = sNome
    Exit Function

NonSalvato:
    On Error Resume Next
    Close #f
    SalvaTestoMese = ""
End Function


'-----------------------------------------------------------------------------
' Testo presente negli appunti di Windows ("" se vuoti o non leggibili).
' Prova prima l'oggetto DataObject di Office, poi quello di Internet Explorer.
'-----------------------------------------------------------------------------
Private Function LeggiAppunti() As String
    Dim o As Object, s As String
    On Error Resume Next
    Set o = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")   ' MSForms.DataObject
    If Not o Is Nothing Then
        o.GetFromClipboard
        s = o.GetText
    End If
    If Len(s) = 0 Then
        Err.Clear
        s = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
    End If
    On Error GoTo 0
    LeggiAppunti = s
End Function


'-----------------------------------------------------------------------------
' Controlla che il cartellino copiato sia del mese del foglio attivo:
'   - ogni giorno del portale ("1 Dom", "2 Lun"...) deve esistere nel foglio
'     e cadere nello stesso giorno della settimana;
'   - il numero di giorni deve essere quello del mese (il portale li elenca
'     sempre tutti, anche quelli futuri);
'   - i codici orario del portale devono essere coerenti con il foglio:
'     "FS" (festivo) solo dove il foglio non ha orario previsto, "01"/"02"
'     (giorno lavorativo) solo dove il foglio ha un orario previsto.
' Restituisce False con il motivo se non corrisponde. Se corrisponde ma un
' altro mese dell'anno ha lo stesso calendario e il testo non contiene un
' festivo infrasettimanale che li distingua, chiede conferma (bConferma).
'-----------------------------------------------------------------------------
Private Function ControllaMese(ByVal ws As Worksheet, ByVal vRighe As Variant, _
                               ByVal regExGiorno As Object, ByRef sMotivo As String, _
                               ByRef bConferma As Boolean) As Boolean
    Dim vGG As Variant, vMesi As Variant
    Dim lAnno As Long, lMese As Long, nGiorniMese As Long
    Dim k As Long, g As Long, m As Long, nGiorni As Long
    Dim s As String, sAtteso As String, sCodice As String
    Dim bFestivoFeriale As Boolean, sSimili As String
    Dim vTarget As Variant, bLavorativo As Boolean
    Dim gMax As Long

    vGG = Array("Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom")
    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")
    bConferma = False
    sMotivo = ""

    On Error Resume Next
    lAnno = CLng(ThisWorkbook.Worksheets("Dashboard").Range("A1").Value)
    lMese = CLng(ws.Range("A1").Value)
    On Error GoTo 0
    If lAnno < 2000 Or lMese < 1 Or lMese > 12 Then
        ControllaMese = True                 ' impossibile verificare: si prosegue
        Exit Function
    End If
    nGiorniMese = Day(DateSerial(lAnno, lMese + 1, 0))

    For k = LBound(vRighe) To UBound(vRighe)
        s = Trim$(Replace(Replace(vRighe(k), Chr(160), " "), vbTab, " "))
        If regExGiorno.Test(s) Then
            g = CLng(regExGiorno.Execute(s)(0).SubMatches(0))
            If g > nGiorniMese Then
                sMotivo = "Il cartellino copiato contiene il giorno " & g & ", che in " & _
                          vMesi(lMese - 1) & " " & lAnno & " non esiste." & vbCrLf & _
                          "Spostati sul foglio del mese giusto oppure copia il mese corretto dal portale."
                Exit Function
            End If
            sAtteso = vGG(Weekday(DateSerial(lAnno, lMese, g), vbMonday) - 1)
            If StrComp(regExGiorno.Execute(s)(0).SubMatches(1), sAtteso, vbTextCompare) <> 0 Then
                sMotivo = "Il cartellino copiato non e' di " & vMesi(lMese - 1) & " " & lAnno & ":" & vbCrLf & _
                          "il giorno " & g & " sul portale e' " & regExGiorno.Execute(s)(0).SubMatches(1) & _
                          ", nel foglio e' " & sAtteso & "." & vbCrLf & _
                          "Spostati sul foglio del mese giusto oppure copia il mese corretto dal portale."
                Exit Function
            End If
            nGiorni = nGiorni + 1
            If g > gMax Then gMax = g
        ElseIf nGiorni > 0 And g > 0 Then
            ' codice orario del giorno: 01/02 = lavorativo, FS = festivo
            sCodice = UCase$(s)
            If sCodice = "FS" Or sCodice Like "##" Then
                vTarget = ws.Cells(10 + g, 4).Value
                bLavorativo = False
                If Not IsError(vTarget) Then
                    If IsNumeric(vTarget) And Not IsEmpty(vTarget) Then bLavorativo = (CDbl(vTarget) > 0)
                End If
                If sCodice = "FS" And bLavorativo Then
                    sMotivo = "Il cartellino copiato non sembra di " & vMesi(lMese - 1) & " " & lAnno & ":" & _
                              vbCrLf & "il giorno " & g & " sul portale e' festivo, nel foglio e' lavorativo." & _
                              vbCrLf & "Spostati sul foglio del mese giusto oppure copia il mese corretto dal portale."
                    Exit Function
                End If
                If sCodice Like "##" And Not bLavorativo And g <= nGiorniMese Then
                    If Weekday(DateSerial(lAnno, lMese, g), vbMonday) <= 5 Then
                        sMotivo = "Il cartellino copiato non sembra di " & vMesi(lMese - 1) & " " & lAnno & ":" & _
                                  vbCrLf & "il giorno " & g & " sul portale e' lavorativo, nel foglio e' festivo." & _
                                  vbCrLf & "Spostati sul foglio del mese giusto oppure copia il mese corretto dal portale."
                        Exit Function
                    End If
                End If
                If sCodice = "FS" And Weekday(DateSerial(lAnno, lMese, g), vbMonday) <= 5 Then bFestivoFeriale = True
            End If
        End If
    Next k

    If nGiorni = 0 Then
        sMotivo = "Nel testo non c'e' nessun giorno del cartellino (righe come " & Chr(34) & "1 Dom" & _
                  Chr(34) & ")." & vbCrLf & "Sul portale apri il cartellino del mese, premi Ctrl+A e Ctrl+C, poi riprova."
        Exit Function
    End If

    ' il portale elenca sempre TUTTI i giorni del mese (anche quelli futuri)
    If gMax <> nGiorniMese Then
        sMotivo = "Il cartellino copiato ha " & gMax & " giorni, " & vMesi(lMese - 1) & " " & lAnno & _
                  " ne ha " & nGiorniMese & ": non e' il mese di questo foglio." & vbCrLf & _
                  "Spostati sul foglio del mese giusto oppure copia il mese corretto dal portale."
        Exit Function
    End If

    ' altri mesi dello stesso anno con lo stesso calendario (stesso giorno di
    ' inizio e stessa lunghezza) che il testo non permette di distinguere?
    If Not bFestivoFeriale Then
        For m = 1 To 12
            If m <> lMese Then
                If Weekday(DateSerial(lAnno, m, 1)) = Weekday(DateSerial(lAnno, lMese, 1)) And _
                   Day(DateSerial(lAnno, m + 1, 0)) = nGiorniMese Then
                    sSimili = sSimili & IIf(Len(sSimili) > 0, ", ", "") & vMesi(m - 1)
                End If
            End If
        Next m
        If Len(sSimili) > 0 Then
            bConferma = True
            sMotivo = "Il cartellino copiato e' compatibile con " & vMesi(lMese - 1) & " " & lAnno & _
                      ", ma anche " & sSimili & " ha lo stesso calendario." & vbCrLf & _
                      "Controlla di aver copiato " & vMesi(lMese - 1) & " dal portale."
        End If
    End If
    ControllaMese = True
End Function


'-----------------------------------------------------------------------------
' Tipo Assenza del file corrispondente alla descrizione del portale.
' ""  = da ignorare (calcolato dal file), "?" = non riconosciuto.
' bConOre = True per i permessi a ore (le ore vanno in Assenze).
'-----------------------------------------------------------------------------
Private Function MappaGiustificativo(ByVal sDesc As String, ByRef bConOre As Boolean) As String
    Dim u As String
    u = UCase$(sDesc)
    bConOre = False
    If u Like "EXTRAPRESENZA*" Or u Like "FLESSIBILIT*" Or InStr(u, "NON AUTORIZZAT") > 0 Then
        MappaGiustificativo = ""
    ElseIf InStr(u, "ESERCIZIO PRECEDENTE") > 0 Then
        MappaGiustificativo = "Ferie anno prec."
    ElseIf u Like "FERIE*" Then
        MappaGiustificativo = "Ferie"
    ElseIf u Like "MALATTIA*" Then
        MappaGiustificativo = "Malattia"
    ElseIf InStr(u, "VISIT") > 0 Then
        MappaGiustificativo = "Art. 23 (Visite)": bConOre = True
    ElseIf InStr(u, "ART. 25") > 0 Or InStr(u, "PERSONAL") > 0 Then
        MappaGiustificativo = "Art. 25 (Motivi pers.)": bConOre = True
    ElseIf InStr(u, "ART. 34") > 0 Or InStr(u, "RECUPER") > 0 Then
        MappaGiustificativo = "Art. 34 (A recupero)": bConOre = True
    ElseIf InStr(u, "104") > 0 Then
        MappaGiustificativo = "L. 104"
    Else
        MappaGiustificativo = "?"
    End If
End Function


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
