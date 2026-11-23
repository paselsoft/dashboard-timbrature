Attribute VB_Name = "modAggiornaFormule"
Option Explicit

'=============================================================================
' Modulo  : modAggiornaFormule  (macro "una tantum")
' Scopo   : riscrive con formule IDENTICHE e indipendenti dal mese le colonne
'           B, C, D, J, K, L, M (righe 11-41), il titolo B2, i totali E4:E5 e
'           il riepilogo Q11:Q17 dei 12 fogli mensili, cosi' che:
'             - ogni foglio prenda mese e anno solo da A1 del foglio e da
'               Dashboard!A1 (niente piu' numeri del mese scritti a mano);
'             - i giorni oltre la fine del mese spariscano da soli
'               (e il 29 febbraio compaia negli anni bisestili);
'             - Ferie, Malattia e L. 104 senza ore in "Assenze" coprano
'               automaticamente l'orario previsto;
'             - un giorno senza timbrature e senza assenze conti 0
'               (giorni futuri o non ancora inseriti non falsano i saldi).
'
' Sicurezza: prima di modificare qualsiasi cosa salva una copia di backup
'            nella stessa cartella; al termine mostra il confronto dei
'            totali EXT/FLE di ogni mese prima e dopo.
'
' Dopo l'uso il modulo puo' essere rimosso (non serve piu').
'=============================================================================

Private Const RIGA_PRIMA As Long = 11       ' giorno 1
Private Const RIGA_ULTIMA As Long = 41      ' giorno 31

Public Sub AggiornaFormuleMesi()

    Dim vMesi As Variant
    Dim FOGLIO_FEST As String
    Dim F_B As String, F_C As String, F_D As String, F_J As String
    Dim F_K As String, F_L As String, F_M As String, F_TITOLO As String
    Dim ws As Worksheet, shIniziale As Object
    Dim i As Long, r As Long, ultimaFormattata As Long
    Dim eraProtetto As Boolean
    Dim aPrimaExt(1 To 12) As Variant, aPrimaFle(1 To 12) As Variant
    Dim bPresente(1 To 12) As Boolean
    Dim sBackup As String, sReport As String, sMancanti As String
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim xlCalc As Long
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    '--- Controlli preliminari ------------------------------------------------
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Salva il file in una cartella del PC prima di eseguire la macro " & _
               "(serve per creare la copia di backup).", vbExclamation, "Aggiorna formule"
        Exit Sub
    End If

    If MsgBox("La macro riscrivera' le formule delle colonne B, C, D, J, K, L, M " & _
              "(righe 11-41) e il titolo dei 12 fogli mensili." & vbCrLf & vbCrLf & _
              "Timbrature, assenze e tipi di assenza NON vengono toccati." & vbCrLf & _
              "Prima verra' salvata una copia di backup nella stessa cartella." & _
              vbCrLf & vbCrLf & "Procedere?", vbQuestion + vbYesNo, "Aggiorna formule") <> vbYes Then
        Exit Sub
    End If

    '--- Backup ---------------------------------------------------------------
    sBackup = ThisWorkbook.Path & Application.PathSeparator & _
              NomeSenzaEstensione(ThisWorkbook.Name) & "_backup_" & _
              Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsm"
    On Error Resume Next
    ThisWorkbook.SaveCopyAs sBackup
    If Err.Number <> 0 Then
        MsgBox "Impossibile creare la copia di backup:" & vbCrLf & Err.Description & _
               vbCrLf & vbCrLf & "Nessuna modifica effettuata.", vbCritical, "Aggiorna formule"
        Exit Sub
    End If
    On Error GoTo 0

    '--- Formule (sintassi inglese con virgole: valida con Excel in italiano) -
    ' Il nome "Festivita'" con l'accento e' costruito con ChrW per evitare
    ' problemi di codifica quando il modulo viene importato.
    FOGLIO_FEST = "Festivit" & ChrW(224)

    F_TITOLO = "=""Dati Cartellino - ""&UPPER(TEXT(DATE(Dashboard!$A$1,$A$1,1),""mmmm aaaa""))"

    ' B: numero del giorno, vuoto oltre la fine del mese
    F_B = "=IF(ROW()-ROW($B$10)>DAY(EOMONTH(DATE(Dashboard!$A$1,$A$1,1),0)),"""",ROW()-ROW($B$10))"

    ' C: giorno della settimana (lun, mar, ...)
    F_C = "=IF($B11="""","""",TEXT(DATE(Dashboard!$A$1,$A$1,$B11),""ggg""))"

    ' D: orario previsto (vuoto nei festivi e nel fine settimana).
    '    WEEKDAY(...;2): 1=lun ... 7=dom  -> lun/mer/ven 6:00, mar/gio 9:00
    F_D = "=IF($B11="""","""",IF(COUNTIF(" & FOGLIO_FEST & "!$A$2:$A$15,DATE(Dashboard!$A$1,$A$1," & _
        "$B11))>0,"""",CHOOSE(WEEKDAY(DATE(Dashboard!$A$1,$A$1,$B11),2),TIME(6,0,0),TIME(9,0,0),TIME(6," & _
        "0,0),TIME(9,0,0),TIME(6,0,0),"""","""")))"

    ' J: ore lavorate (coppie Entr1/Usc1 e Entr2/Usc2 complete)
    F_J = "=IF(AND(ISNUMBER($F11),ISNUMBER($G11)),$G11-$F11,0)+IF(AND(ISNUMBER($H11),ISNUMBER($I11))," & _
        "$I11-$H11,0)"

    ' K / L: Piu'(EXT) e Meno(FLE).
    '   Assenza effettiva = ore in "Assenze" se presenti, altrimenti l'intero
    '   orario previsto se il tipo e' Ferie, Malattia o L. 104, altrimenti 0.
    '   Se non ci sono ne' timbrature ne' assenze, il giorno conta 0.
    F_K = "=IF(AND(COUNT($F11:$I11)=0,IF(ISNUMBER($E11),$E11,IF(OR($N11=""Ferie"",$N11=""Malattia""," & _
        "$N11=""L. 104""),IF(ISNUMBER($D11),$D11,0),0))=0),0,MAX(0,ROUND(($J11+IF(ISNUMBER($E11),$E11," & _
        "IF(OR($N11=""Ferie"",$N11=""Malattia"",$N11=""L. 104""),IF(ISNUMBER($D11),$D11,0)," & _
        "0))-IF(ISNUMBER($D11),$D11,0))*1440,0)/1440))"

    F_L = "=IF(AND(COUNT($F11:$I11)=0,IF(ISNUMBER($E11),$E11,IF(OR($N11=""Ferie"",$N11=""Malattia""," & _
        "$N11=""L. 104""),IF(ISNUMBER($D11),$D11,0),0))=0),0,MAX(0,ROUND((IF(ISNUMBER($D11),$D11," & _
        "0)-$J11-IF(ISNUMBER($E11),$E11,IF(OR($N11=""Ferie"",$N11=""Malattia"",$N11=""L. 104"")," & _
        "IF(ISNUMBER($D11),$D11,0),0)))*1440,0)/1440))"

    ' M: Delta in testo (+h:mm / -h:mm)
    F_M = "=IF(AND($K11=0,$L11=0),"""",IF($K11>0,""+""&TEXT($K11,""[h]:mm""),""-""&TEXT($L11," & _
        """[h]:mm"")))"

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    '--- Stato iniziale -------------------------------------------------------
    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    xlCalc = Application.Calculation

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- Totali "prima" -------------------------------------------------------
    Application.Calculate
    For i = 1 To 12
        Set ws = OttieniFoglio(CStr(vMesi(i - 1)))
        If Not ws Is Nothing Then
            bPresente(i) = True
            aPrimaExt(i) = ws.Range("E4").Value
            aPrimaFle(i) = ws.Range("E5").Value
        Else
            sMancanti = sMancanti & " " & vMesi(i - 1)
        End If
    Next i

    '--- Riscrittura ----------------------------------------------------------
    For i = 1 To 12
        If bPresente(i) Then
            Set ws = ThisWorkbook.Worksheets(CStr(vMesi(i - 1)))
            sFoglioErr = ws.Name

            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect

            ' numero del mese del foglio (usato da tutte le formule)
            ws.Range("A1").Value = i

            ' Mesi corti (es. Febbraio): le righe finali non hanno bordi,
            ' convalide e formattazioni. Le copio dall'ultima riga formattata,
            ' cosi' il 29/30/31 compare correttamente quando serve.
            ultimaFormattata = RIGA_ULTIMA
            Do While ultimaFormattata > RIGA_PRIMA And _
                     Not ws.Cells(ultimaFormattata, 4).HasFormula
                ultimaFormattata = ultimaFormattata - 1
            Loop
            If ultimaFormattata < RIGA_ULTIMA Then
                ws.Activate
                ws.Range("B" & ultimaFormattata & ":N" & ultimaFormattata).Copy
                With ws.Range("B" & ultimaFormattata + 1 & ":N" & RIGA_ULTIMA)
                    .PasteSpecial Paste:=xlPasteFormats
                    .PasteSpecial Paste:=xlPasteValidation
                End With
                Application.CutCopyMode = False
                ws.Range("A1").Select
            End If

            ' Formule: assegnate all'intero intervallo, Excel adegua i
            ' riferimenti relativi riga per riga.
            ws.Range("B2").Formula = F_TITOLO
            ws.Range("B" & RIGA_PRIMA & ":B" & RIGA_ULTIMA).Formula = F_B
            ws.Range("C" & RIGA_PRIMA & ":C" & RIGA_ULTIMA).Formula = F_C
            ws.Range("D" & RIGA_PRIMA & ":D" & RIGA_ULTIMA).Formula = F_D
            ws.Range("J" & RIGA_PRIMA & ":J" & RIGA_ULTIMA).Formula = F_J
            ws.Range("K" & RIGA_PRIMA & ":K" & RIGA_ULTIMA).Formula = F_K
            ws.Range("L" & RIGA_PRIMA & ":L" & RIGA_ULTIMA).Formula = F_L
            ws.Range("M" & RIGA_PRIMA & ":M" & RIGA_ULTIMA).Formula = F_M

            ' Totali del mese e riepilogo: sempre sulle righe 11-41, cosi'
            ' anche il 29 febbraio (anni bisestili) viene conteggiato.
            ws.Range("E4").Formula = "=SUM(K11:K41)"
            ws.Range("E5").Formula = "=SUM(L11:L41)"
            ws.Range("Q11").Formula = "=COUNTIF($N$11:$N$41,""Ferie"")"
            ws.Range("Q12").Formula = "=COUNTIF($N$11:$N$41,""L. 104"")"
            ws.Range("Q13").Formula = "=SUMIF($N$11:$N$41,""Art. 23 (Visite)"",$E$11:$E$41)"
            ws.Range("Q14").Formula = "=SUMIF($N$11:$N$41,""Art. 25 (Motivi pers.)"",$E$11:$E$41)"
            ws.Range("Q15").Formula = "=SUMIF($N$11:$N$41,""Art. 34 (A recupero)"",$E$11:$E$41)"
            ws.Range("Q16").Formula = "=COUNTIF($N$11:$N$41,""Malattia"")"
            ws.Range("Q17").Formula = "=COUNTIFS($H$11:$H$41,"">0"",$I$11:$I$41,"">0"",$J$11:$J$41,"">""&TIME(6,0,0))"

            ' stesso formato ore per J, K, L in tutti i mesi (zero nascosto)
            ws.Range("J" & RIGA_PRIMA & ":L" & RIGA_ULTIMA).NumberFormat = "[h]:mm;;"

            If eraProtetto Then ws.Protect
        End If
    Next i
    sFoglioErr = ""

    '--- Totali "dopo" e report -----------------------------------------------
    Application.Calculation = xlCalc
    Application.Calculate

    sReport = "Mese" & vbTab & "EXT prima > dopo" & vbTab & "FLE prima > dopo" & vbCrLf
    For i = 1 To 12
        If bPresente(i) Then
            Set ws = ThisWorkbook.Worksheets(CStr(vMesi(i - 1)))
            sReport = sReport & vMesi(i - 1) & vbTab & _
                      Ore(aPrimaExt(i)) & " > " & Ore(ws.Range("E4").Value) & vbTab & _
                      Ore(aPrimaFle(i)) & " > " & Ore(ws.Range("E5").Value)
            If Ore(aPrimaExt(i)) <> Ore(ws.Range("E4").Value) Or _
               Ore(aPrimaFle(i)) <> Ore(ws.Range("E5").Value) Then
                sReport = sReport & "   <-- CAMBIATO"
            End If
            sReport = sReport & vbCrLf
        End If
    Next i

Uscita:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.Calculation = xlCalc
    If Not shIniziale Is Nothing Then shIniziale.Activate
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, "") & vbCrLf & vbCrLf & _
               "Il file potrebbe essere stato modificato solo in parte." & vbCrLf & _
               "Chiudilo SENZA salvare e riapri il backup:" & vbCrLf & sBackup, _
               vbCritical, "Aggiorna formule"
    Else
        If Len(sMancanti) > 0 Then sReport = sReport & vbCrLf & "Fogli non trovati:" & sMancanti & vbCrLf
        MsgBox "Formule aggiornate." & vbCrLf & vbCrLf & sReport & vbCrLf & _
               "Backup salvato in:" & vbCrLf & sBackup & vbCrLf & vbCrLf & _
               "Se i totali ti tornano, salva il file.", vbInformation, "Aggiorna formule"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


'-----------------------------------------------------------------------------
' Converte un valore orario in testo h:mm (per il report); "-" se non numerico
'-----------------------------------------------------------------------------
Private Function Ore(ByVal v As Variant) As String
    If IsNumeric(v) And Not IsEmpty(v) Then
        Dim m As Long
    m = CLng(Round(CDbl(v) * 1440, 0))
    Ore = CStr(m \ 60) & ":" & Format$(m Mod 60, "00")
    Else
        Ore = "-"
    End If
End Function


Private Function OttieniFoglio(ByVal sNome As String) As Worksheet
    On Error Resume Next
    Set OttieniFoglio = ThisWorkbook.Worksheets(sNome)
    On Error GoTo 0
End Function


Private Function NomeSenzaEstensione(ByVal sNome As String) As String
    Dim p As Long
    p = InStrRev(sNome, ".")
    If p > 0 Then NomeSenzaEstensione = Left$(sNome, p - 1) Else NomeSenzaEstensione = sNome
End Function
