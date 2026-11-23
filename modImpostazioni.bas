Attribute VB_Name = "modImpostazioni"
Option Explicit

'=============================================================================
' Modulo  : modImpostazioni  (macro "una tantum")
' Macro   : CreaFoglioImpostazioni
'
' Crea il foglio "Impostazioni" con i parametri oggi scritti dentro le
' formule e collega le formule a quel foglio tramite nomi:
'
'   Orario_Previsto   ore previste da lunedi' a domenica (7 celle)
'   Buono_Valore      valore del buono pasto in euro
'   Buono_Ore_Minime  ore lavorate da superare per maturare il buono
'   Plafond_Ferie     giorni di ferie annui
'   Plafond_L104      giorni di L. 104 annui
'   Plafond_Art23     ore annue Art. 23 (Visite)
'   Plafond_Art25     ore annue Art. 25 (Motivi personali)
'   Plafond_Art34     ore annue Art. 34 (A recupero)
'
' Formule aggiornate:
'   - fogli mensili: colonna D (Target) e Q17 (Buoni Pasto);
'   - Dashboard: I8:I12 (plafond) e J14 (valore buoni pasto).
'
' I valori iniziali sono quelli attuali: nessun totale deve cambiare. Prima
' delle modifiche viene salvato un backup; alla fine un confronto
' prima/dopo conferma che i numeri sono identici.
' Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Private Const NOME_FOGLIO As String = "Impostazioni"

Public Sub CreaFoglioImpostazioni()

    Dim vMesi As Variant, v As Variant
    Dim wsDash As Worksheet, wsImp As Worksheet, wsFest As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim FOGLIO_FEST As String, F_D As String, F_Q17 As String, DT As String, IDX As String
    Dim aPrima() As String, sDopo As String, sReport As String
    Dim i As Long, nDiff As Long
    Dim eraProtetto As Boolean
    Dim sBackup As String
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim xlCalc As Long
    Dim bModificato As Boolean, bFatto As Boolean
    Dim lErrNum As Long, sErrDesc As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")
    FOGLIO_FEST = "Festivit" & ChrW(224)

    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    xlCalc = Application.Calculation
    Set shIniziale = ActiveSheet

    '--- Controlli preliminari ------------------------------------------------
    If Not OttieniFoglio(NOME_FOGLIO) Is Nothing Then
        MsgBox "Il foglio '" & NOME_FOGLIO & "' esiste gia'. Nessuna modifica effettuata.", _
               vbInformation, "Impostazioni"
        Exit Sub
    End If
    Set wsDash = OttieniFoglio("Dashboard")
    Set wsFest = OttieniFoglio(FOGLIO_FEST)
    If wsDash Is Nothing Or wsFest Is Nothing Then
        MsgBox "Fogli Dashboard o " & FOGLIO_FEST & " non trovati.", vbCritical, "Impostazioni"
        Exit Sub
    End If
    For Each v In vMesi
        If OttieniFoglio(CStr(v)) Is Nothing Then
            MsgBox "Foglio '" & v & "' non trovato.", vbCritical, "Impostazioni"
            Exit Sub
        End If
    Next v
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Salva il file in una cartella del PC prima di eseguire la macro.", _
               vbExclamation, "Impostazioni"
        Exit Sub
    End If

    If MsgBox("Verra' creato il foglio 'Impostazioni' con orari, buono pasto e plafond " & _
              "attuali, e le formule verranno collegate a quel foglio." & vbCrLf & vbCrLf & _
              "Prima verra' salvata una copia di backup. Procedere?", _
              vbQuestion + vbYesNo, "Impostazioni") <> vbYes Then Exit Sub

    '--- Backup ---------------------------------------------------------------
    sBackup = ThisWorkbook.Path & Application.PathSeparator & _
              NomeSenzaEstensione(ThisWorkbook.Name) & "_backup_" & _
              Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsm"
    On Error Resume Next
    ThisWorkbook.SaveCopyAs sBackup
    If Err.Number <> 0 Then
        MsgBox "Impossibile creare il backup: " & Err.Description & vbCrLf & _
               "Nessuna modifica effettuata.", vbCritical, "Impostazioni"
        Exit Sub
    End If
    On Error GoTo GestioneErrori

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- Valori "prima" -------------------------------------------------------
    Application.Calculate
    ReDim aPrima(0 To 12)
    For i = 0 To 11
        aPrima(i) = Istantanea(ThisWorkbook.Worksheets(CStr(vMesi(i))))
    Next i
    aPrima(12) = IstantaneaDashboard(wsDash)

    Application.Calculation = xlCalculationManual
    bModificato = True

    '--- Foglio Impostazioni --------------------------------------------------
    Set wsImp = ThisWorkbook.Worksheets.Add(After:=wsFest)
    wsImp.Name = NOME_FOGLIO
    ActiveWindow.DisplayGridlines = False
    CompilaImpostazioni wsImp, wsDash

    ' Nomi (validi in tutta la cartella)
    AggiungiNome "Orario_Previsto", "=Impostazioni!$C$6:$C$12"
    AggiungiNome "Buono_Valore", "=Impostazioni!$C$15"
    AggiungiNome "Buono_Ore_Minime", "=Impostazioni!$C$16"
    AggiungiNome "Plafond_Ferie", "=Impostazioni!$C$20"
    AggiungiNome "Plafond_L104", "=Impostazioni!$C$21"
    AggiungiNome "Plafond_Art23", "=Impostazioni!$C$22"
    AggiungiNome "Plafond_Art25", "=Impostazioni!$C$23"
    AggiungiNome "Plafond_Art34", "=Impostazioni!$C$24"

    '--- Formule dei mesi -----------------------------------------------------
    ' D: festivo -> vuoto; altrimenti l'orario del giorno della settimana preso
    '    da Orario_Previsto (WEEKDAY(...;2): 1=lunedi' ... 7=domenica);
    '    orario vuoto o 0 -> vuoto (giorno non lavorativo).
    DT = "DATE(Dashboard!$A$1,$A$1,$B11)"
    IDX = "INDEX(Orario_Previsto,WEEKDAY(" & DT & ",2))"
    F_D = "=IF($B11="""","""",IF(COUNTIF(" & FOGLIO_FEST & "!$A$2:$A$15," & DT & ")>0,""""," & _
          "IF(N(" & IDX & ")=0,""""," & IDX & ")))"
    F_Q17 = "=COUNTIFS($H$11:$H$41,"">0"",$I$11:$I$41,"">0"",$J$11:$J$41,"">""&Buono_Ore_Minime)"

    For Each v In vMesi
        Set ws = ThisWorkbook.Worksheets(CStr(v))
        eraProtetto = ws.ProtectContents
        If eraProtetto Then ws.Unprotect
        ws.Range("D11:D41").Formula = F_D
        ws.Range("Q17").Formula = F_Q17
        If eraProtetto Then ws.Protect
        eraProtetto = False
    Next v

    '--- Formule della Dashboard ----------------------------------------------
    Set ws = wsDash
    eraProtetto = wsDash.ProtectContents
    If eraProtetto Then wsDash.Unprotect
    wsDash.Range("I8").Formula = "=Plafond_Ferie"
    wsDash.Range("I9").Formula = "=Plafond_L104"
    wsDash.Range("I10").Formula = "=Plafond_Art23"
    wsDash.Range("I11").Formula = "=Plafond_Art25"
    wsDash.Range("I12").Formula = "=Plafond_Art34"
    wsDash.Range("J14").Formula = "=I14*Buono_Valore"
    If eraProtetto Then wsDash.Protect
    eraProtetto = False

    '--- Confronto prima/dopo -------------------------------------------------
    Application.Calculation = xlCalc
    Application.Calculate
    For i = 0 To 12
        If i < 12 Then
            sDopo = Istantanea(ThisWorkbook.Worksheets(CStr(vMesi(i))))
        Else
            sDopo = IstantaneaDashboard(wsDash)
        End If
        If sDopo <> aPrima(i) Then
            nDiff = nDiff + 1
            sReport = sReport & vbCrLf & "  - " & IIf(i < 12, vMesi(i), "Dashboard") & _
                      ": prima " & aPrima(i) & vbCrLf & "      dopo  " & sDopo
        End If
    Next i
    bFatto = True

Uscita:
    On Error Resume Next
    If eraProtetto And Not ws Is Nothing Then ws.Protect
    Application.Calculation = xlCalc
    If Not wsImp Is Nothing And bFatto Then wsImp.Activate Else shIniziale.Activate
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & vbCrLf & vbCrLf & _
               IIf(bModificato, "Il file e' stato modificato solo in parte." & vbCrLf & _
                   "Chiudilo SENZA salvare e riapri il backup:" & vbCrLf & sBackup, ""), _
               vbCritical, "Impostazioni"
    ElseIf bFatto Then
        If nDiff = 0 Then
            MsgBox "Foglio 'Impostazioni' creato e formule collegate." & vbCrLf & vbCrLf & _
                   "Controllo: tutti i totali (mesi e Dashboard) sono identici a prima." & _
                   vbCrLf & vbCrLf & "Backup salvato in:" & vbCrLf & sBackup & vbCrLf & vbCrLf & _
                   "Salva il file per rendere definitive le modifiche.", _
                   vbInformation, "Impostazioni"
        Else
            MsgBox "Foglio creato, ma " & nDiff & " confronti sono diversi da prima:" & _
                   sReport & vbCrLf & vbCrLf & _
                   "Non salvare: chiudi il file e segnalami queste righe." & vbCrLf & _
                   "Backup: " & sBackup, vbExclamation, "Impostazioni"
        End If
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


'-----------------------------------------------------------------------------
' Contenuto e aspetto del foglio Impostazioni (valori iniziali = attuali)
'-----------------------------------------------------------------------------
Private Sub CompilaImpostazioni(ByVal ws As Worksheet, ByVal wsDash As Worksheet)
    Dim vGiorni As Variant, vOre As Variant
    Dim i As Long
    Dim colInput As Long, colHeader As Long, colTitolo As Long

    colInput = RGB(255, 242, 204)      ' giallo chiaro, come la colonna Assenze
    colHeader = RGB(221, 235, 247)     ' azzurro chiaro, come le intestazioni
    colTitolo = RGB(31, 78, 121)       ' blu dei titoli

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 60

    With ws.Range("B2")
        .Value = "IMPOSTAZIONI"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = colTitolo
    End With
    With ws.Range("B3")
        .Value = "Modifica solo le celle gialle: i valori valgono per tutti i mesi dell'anno " & _
                 "e per la Dashboard."
        .Font.Italic = True
        .Font.Color = RGB(128, 128, 128)
    End With

    ' --- Orario previsto ---
    Intestazione ws.Range("B5:C5"), "Orario previsto", "Ore", colHeader
    vGiorni = Array("Luned" & ChrW(236), "Marted" & ChrW(236), "Mercoled" & ChrW(236), _
                    "Gioved" & ChrW(236), "Venerd" & ChrW(236), "Sabato", "Domenica")
    vOre = Array(6, 9, 6, 9, 6, 0, 0)
    For i = 0 To 6
        ws.Cells(6 + i, 2).Value = vGiorni(i)
        If vOre(i) > 0 Then ws.Cells(6 + i, 3).Value = TimeSerial(vOre(i), 0, 0)
    Next i
    ws.Range("C6:C12").NumberFormat = "[h]:mm"
    ws.Range("D6").Value = "Lascia vuoto (o 0:00) per i giorni non lavorativi."
    ws.Range("D7").Value = "I festivi del foglio " & "Festivit" & ChrW(224) & _
                           " sono sempre senza orario."

    ' --- Buono pasto ---
    Intestazione ws.Range("B14:C14"), "Buono pasto", "Valore", colHeader
    ws.Range("B15").Value = "Valore del buono (" & ChrW(8364) & ")"
    ws.Range("C15").Value = 7
    ws.Range("C15").NumberFormat = ChrW(8364) & " #,##0.00"
    ws.Range("B16").Value = "Ore lavorate da superare"
    ws.Range("C16").Value = TimeSerial(6, 0, 0)
    ws.Range("C16").NumberFormat = "[h]:mm"
    ws.Range("D16").Value = "Il buono richiede anche il rientro (Entr 2 e Usc 2 compilati)."

    ' --- Plafond annui (valori attuali della Dashboard) ---
    Intestazione ws.Range("B19:C19"), "Plafond annui", "Valore", colHeader
    ws.Range("B20").Value = "Ferie (giorni)"
    ws.Range("B21").Value = "L. 104 (giorni)"
    ws.Range("B22").Value = "Art. 23 - Visite (ore)"
    ws.Range("B23").Value = "Art. 25 - Motivi personali (ore)"
    ws.Range("B24").Value = "Art. 34 - A recupero (ore)"
    ws.Range("C20").Value = wsDash.Range("I8").Value
    ws.Range("C21").Value = wsDash.Range("I9").Value
    ws.Range("C22").Value = wsDash.Range("I10").Value
    ws.Range("C23").Value = wsDash.Range("I11").Value
    ws.Range("C24").Value = wsDash.Range("I12").Value
    ws.Range("C20:C21").NumberFormat = "0"
    ws.Range("C22:C24").NumberFormat = "[h]:mm"
    ws.Range("D20").Value = "Le ferie non godute dell'anno precedente sono in Dashboard (riga " & _
                            "Ferie Anno precedente)."

    ' --- Aspetto comune ---
    With ws.Range("C6:C12,C15:C16,C20:C24")
        .Interior.Color = colInput
        .HorizontalAlignment = xlCenter
        .Locked = False
    End With
    With ws.Range("B5:C12,B14:C16,B19:C24")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ws.Range("D6:D24").Font.Color = RGB(128, 128, 128)
    ws.Range("D6:D24").Font.Italic = True

    ' --- Convalide: solo valori sensati ---
    With ws.Range("C6:C12").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="0", Formula2:="=95/96"
        .ErrorMessage = "Inserisci un orario tra 0:00 e 23:45 (es. 6:00)."
    End With
    With ws.Range("C15,C20:C21").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
             Operator:=xlGreaterEqual, Formula1:="0"
        .ErrorMessage = "Inserisci un numero maggiore o uguale a zero."
    End With
End Sub


Private Sub Intestazione(ByVal rng As Range, ByVal sTitolo As String, _
                         ByVal sColonna As String, ByVal lColore As Long)
    rng.Cells(1, 1).Value = sTitolo
    rng.Cells(1, 2).Value = sColonna
    rng.Interior.Color = lColore
    rng.Font.Bold = True
    rng.Cells(1, 2).HorizontalAlignment = xlCenter
End Sub


Private Sub AggiungiNome(ByVal sNome As String, ByVal sRif As String)
    On Error Resume Next
    ThisWorkbook.Names(sNome).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=sNome, RefersTo:=sRif
End Sub


'-----------------------------------------------------------------------------
' "Fotografie" dei valori per il confronto prima/dopo
'-----------------------------------------------------------------------------
Private Function Istantanea(ByVal ws As Worksheet) As String
    Dim r As Long, s As String
    s = "EXT " & ws.Range("E4").Text & " FLE " & ws.Range("E5").Text & _
        " Buoni " & ws.Range("Q17").Text & " Target:"
    For r = 11 To 41
        s = s & ws.Cells(r, 4).Text & "|"
    Next r
    Istantanea = s
End Function


Private Function IstantaneaDashboard(ByVal ws As Worksheet) As String
    Dim c As Variant, s As String
    For Each c In Array("I8", "I9", "I10", "I11", "I12", "K8", "K9", "K10", "K11", "K12", _
                        "I14", "J14", "E20", "E22")
        s = s & c & "=" & ws.Range(c).Text & " "
    Next c
    IstantaneaDashboard = s
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
