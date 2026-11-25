Attribute VB_Name = "modNuovoAnno"
Option Explicit

'=============================================================================
' Modulo  : modNuovoAnno
' Macro   : NuovoAnno
'
' Prepara il file per l'anno successivo SENZA toccare il file dell'anno in
' corso, che resta sul disco come archivio:
'
'   1. controlla che il file sia salvato (l'archivio deve essere completo);
'   2. calcola i valori da riportare:
'        - saldo orario finale  = Residuo anno prec. + EXT - FLE dell'anno
'          (Dashboard: C4 + C20 - D20), riportato per intero, anche negativo;
'        - ferie residue        = ferie non godute DELL'ANNO (Dashboard: K8),
'          riportate tutte; le ferie dell'anno precedente non godute entro
'          il 30 giugno (K7) sono scadute e NON vengono riportate;
'   3. mostra un riepilogo e chiede conferma;
'   4. svuota timbrature (F:I), ore di assenza (E), tipi di assenza (N) e
'      note di cella dei 12 fogli mensili (formule e formati restano);
'   5. scrive in Dashboard: anno + 1 (A1), saldo riportato (C4),
'      ferie residue come plafond "Ferie Anno <anno>" (I7);
'   6. salva con un NUOVO nome (es. ..._2026.xlsm -> ..._2027.xlsm).
'
' Calendario, festivita' (Pasqua compresa) e intestazioni si aggiornano da
' soli perche' tutte le formule dipendono da Dashboard!A1.
'=============================================================================

Private Const RIGA_PRIMA As Long = 11
Private Const RIGA_ULTIMA As Long = 41

Public Sub NuovoAnno()

    Dim wsDash As Worksheet, ws As Worksheet
    Dim vMesi As Variant, v As Variant
    Dim lAnno As Long, lNuovo As Long
    Dim dSaldo As Double, dFerieRes As Double, dFerieScadute As Double
    Dim sNomeNuovo As String, sPercorsoNuovo As String, sPercorsoVecchio As String
    Dim sMsg As String
    Dim c As Range
    Dim eraProtetto As Boolean, dashProtetta As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim xlCalc As Long
    Dim bModificato As Boolean, bFatto As Boolean
    Dim lErrNum As Long, sErrDesc As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    ' stato dell'applicazione, ripristinato sempre all'uscita
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    xlCalc = Application.Calculation

    '--- 1) Controlli preliminari ---------------------------------------------
    Set wsDash = OttieniFoglio("Dashboard")
    If wsDash Is Nothing Then
        MsgBox "Foglio Dashboard non trovato.", vbCritical, "Nuovo anno"
        Exit Sub
    End If
    For Each v In vMesi
        If OttieniFoglio(CStr(v)) Is Nothing Then
            MsgBox "Foglio '" & v & "' non trovato. Operazione annullata.", vbCritical, "Nuovo anno"
            Exit Sub
        End If
    Next v

    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Il file deve essere salvato in una cartella del PC.", vbExclamation, "Nuovo anno"
        Exit Sub
    End If

    If Not IsNumeric(wsDash.Range("A1").Value) Then
        MsgBox "Dashboard!A1 non contiene un anno valido.", vbCritical, "Nuovo anno"
        Exit Sub
    End If
    lAnno = CLng(wsDash.Range("A1").Value)
    lNuovo = lAnno + 1

    ' Il file attuale diventa l'archivio: deve essere salvato e completo.
    If Not ThisWorkbook.Saved Then
        If MsgBox("Il file ha modifiche non salvate." & vbCrLf & _
                  "Per creare un archivio completo del " & lAnno & " va salvato prima." & _
                  vbCrLf & vbCrLf & "Salvare adesso?", vbQuestion + vbYesNo, "Nuovo anno") <> vbYes Then
            Exit Sub
        End If
        ThisWorkbook.Save
    End If

    '--- 2) Valori da riportare ------------------------------------------------
    On Error GoTo GestioneErrori
    Application.Calculate
    dSaldo = ValoreNumerico(wsDash.Range("C4")) + ValoreNumerico(wsDash.Range("C20")) _
             - ValoreNumerico(wsDash.Range("D20"))
    dSaldo = Round(dSaldo * 1440, 0) / 1440          ' arrotondo al minuto
    ' Ferie: si riporta solo il residuo dell'anno (K8). Il residuo dell'anno
    ' precedente (K7) andava fruito entro il 30 giugno: e' scaduto.
    dFerieRes = ValoreNumerico(wsDash.Range("K8"))
    dFerieScadute = ValoreNumerico(wsDash.Range("K7"))
    If dFerieScadute < 0 Then dFerieScadute = 0

    ' Nome del nuovo file: sostituisco l'anno nel nome, se c'e'
    sPercorsoVecchio = ThisWorkbook.FullName
    If InStr(1, ThisWorkbook.Name, CStr(lAnno)) > 0 Then
        sNomeNuovo = Replace(ThisWorkbook.Name, CStr(lAnno), CStr(lNuovo))
    Else
        sNomeNuovo = NomeSenzaEstensione(ThisWorkbook.Name) & "_" & lNuovo & ".xlsm"
    End If
    If LCase$(Right$(sNomeNuovo, 5)) <> ".xlsm" Then sNomeNuovo = NomeSenzaEstensione(sNomeNuovo) & ".xlsm"
    sPercorsoNuovo = ThisWorkbook.Path & Application.PathSeparator & sNomeNuovo

    If Len(Dir$(sPercorsoNuovo)) > 0 Then
        MsgBox "Esiste gia' il file:" & vbCrLf & sPercorsoNuovo & vbCrLf & vbCrLf & _
               "Spostalo o rinominalo e rilancia la macro.", vbExclamation, "Nuovo anno"
        GoTo Uscita
    End If

    '--- 3) Riepilogo e conferma -----------------------------------------------
    sMsg = "Passaggio dal " & lAnno & " al " & lNuovo & vbCrLf & vbCrLf & _
           "Valori riportati nel " & lNuovo & ":" & vbCrLf & _
           "  - Saldo orario (Residuo " & lAnno & "):  " & OreConSegno(dSaldo) & vbCrLf & _
           "  - Ferie residue " & lAnno & ":  " & Format$(dFerieRes, "0") & " giorni" & _
           " (da fruire entro il 30/06/" & lNuovo & ")" & vbCrLf & _
           "  - Ferie " & lNuovo & " (plafond, invariato):  " & _
           Format$(ValoreNumerico(wsDash.Range("I8")), "0") & " giorni" & vbCrLf & vbCrLf & _
           "Verranno svuotati timbrature, assenze, tipi di assenza e note dei 12 mesi." & vbCrLf & vbCrLf & _
           "Il file del " & lAnno & " resta invariato come archivio:" & vbCrLf & sPercorsoVecchio & vbCrLf & _
           "Il nuovo anno verra' salvato come:" & vbCrLf & sPercorsoNuovo

    If Date < DateSerial(lNuovo, 1, 1) Then
        sMsg = sMsg & vbCrLf & vbCrLf & "ATTENZIONE: il " & lAnno & " non e' ancora finito." & _
               vbCrLf & "Le timbrature che inserirai dopo nel file " & lAnno & _
               " NON saranno riportate nel nuovo file."
    End If
    If dSaldo < 0 Then
        sMsg = sMsg & vbCrLf & vbCrLf & "Nota: il saldo e' negativo. In Dashboard C4 Excel " & _
               "lo mostrera' come ##### (gli orari negativi non sono visualizzabili), " & _
               "ma i calcoli resteranno corretti."
    End If
    If dFerieScadute > 0 Then
        sMsg = sMsg & vbCrLf & vbCrLf & "Nota: " & Format$(dFerieScadute, "0") & _
               " giorni di ferie " & (lAnno - 1) & " non fruiti entro il 30/06/" & lAnno & _
               " sono scaduti e NON vengono riportati."
    End If
    sMsg = sMsg & vbCrLf & vbCrLf & "Procedere?"

    If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Nuovo anno") <> vbYes Then GoTo Uscita

    '--- 4) Svuotamento dei mesi ----------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    bModificato = True

    For Each v In vMesi
        Set ws = ThisWorkbook.Worksheets(CStr(v))
        Application.StatusBar = "Nuovo anno: svuoto " & ws.Name & "..."
        eraProtetto = ws.ProtectContents
        If eraProtetto Then ws.Unprotect

        ' Solo i valori inseriti a mano: le eventuali formule restano
        For Each c In Union(ws.Range("E" & RIGA_PRIMA & ":I" & RIGA_ULTIMA), _
                            ws.Range("N" & RIGA_PRIMA & ":N" & RIGA_ULTIMA)).Cells
            If Not c.HasFormula Then
                If Not IsEmpty(c.Value) Then c.ClearContents
            End If
        Next c
        ' note/commenti sui giorni dell'anno concluso
        ws.Range("B" & RIGA_PRIMA & ":N" & RIGA_ULTIMA).ClearComments

        If eraProtetto Then ws.Protect
        eraProtetto = False
    Next v

    '--- 5) Dashboard ----------------------------------------------------------
    Application.StatusBar = "Nuovo anno: aggiorno la Dashboard..."
    dashProtetta = wsDash.ProtectContents
    If dashProtetta Then wsDash.Unprotect
    wsDash.Range("A1").Value = lNuovo
    wsDash.Range("C4").Value = dSaldo
    wsDash.Range("C4").NumberFormat = "[h]:mm"
    wsDash.Range("I7").Value = dFerieRes
    If dashProtetta Then wsDash.Protect
    dashProtetta = False

    Application.Calculation = xlCalc
    Application.Calculate

    '--- 6) Salvataggio con il nuovo nome --------------------------------------
    Application.StatusBar = "Nuovo anno: salvo " & sNomeNuovo & "..."
    wsDash.Activate
    ThisWorkbook.SaveAs Filename:=sPercorsoNuovo, FileFormat:=52   ' 52 = .xlsm
    bModificato = False
    bFatto = True

Uscita:
    On Error Resume Next
    If eraProtetto Then ws.Protect
    If dashProtetta Then wsDash.Protect
    Application.StatusBar = False
    Application.Calculation = xlCalc
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & vbCrLf & vbCrLf & _
               IIf(bModificato, _
                   "Il file aperto e' stato modificato solo in parte." & vbCrLf & _
                   "CHIUDILO SENZA SALVARE: il file del " & lAnno & " sul disco e' intatto.", _
                   ""), vbCritical, "Nuovo anno"
    ElseIf bFatto Then
        MsgBox "Fatto! Ora stai lavorando sul file del " & lNuovo & ":" & vbCrLf & _
               sPercorsoNuovo & vbCrLf & vbCrLf & _
               "Archivio del " & lAnno & " (invariato):" & vbCrLf & sPercorsoVecchio & vbCrLf & vbCrLf & _
               "Controlla in Dashboard il residuo " & lAnno & " e le ferie riportate." & vbCrLf & _
               "Se nel " & lNuovo & " cambiano plafond o festivita' locali, aggiornali ora.", _
               vbInformation, "Nuovo anno"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


'-----------------------------------------------------------------------------
' Supporto
'-----------------------------------------------------------------------------
Private Function ValoreNumerico(ByVal c As Range) As Double
    If IsError(c.Value) Then
        Err.Raise vbObjectError + 700, , "La cella " & c.Parent.Name & "!" & _
                  c.Address(False, False) & " contiene un errore."
    End If
    If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then ValoreNumerico = CDbl(c.Value)
End Function


Private Function OreConSegno(ByVal d As Double) As String
    Dim m As Long
    m = CLng(Round(Abs(d) * 1440, 0))
    OreConSegno = IIf(d < 0, "-", "+") & CStr(m \ 60) & ":" & Format$(m Mod 60, "00")
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
