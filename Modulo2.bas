Attribute VB_Name = "Modulo2"
Option Explicit

Sub EsportaMeseInPDF()
    Dim ws As Worksheet
    Dim ultimaRiga As Long
    Dim numMese As String
    Dim percorsoCartella As String
    Dim nomeFile As String
    Dim percorsoCompleto As String
    Dim eraProtetto As Boolean      ' il foglio era protetto all'avvio?

    Set ws = ActiveSheet

    ' Blocca l'esecuzione se ci si trova su Dashboard o Festività
    If ws.Name = "Dashboard" Or ws.Name = "Festività" Then
        MsgBox "Spostati su un foglio mensile (es. Novembre) prima di esportare in PDF.", vbExclamation, "Esportazione PDF"
        Exit Sub
    End If

    ' Mappa il numero del mese per avere i PDF in ordine cronologico nella cartella
    Select Case ws.Name
        Case "Gennaio":   numMese = "01"
        Case "Febbraio":  numMese = "02"
        Case "Marzo":     numMese = "03"
        Case "Aprile":    numMese = "04"
        Case "Maggio":    numMese = "05"
        Case "Giugno":    numMese = "06"
        Case "Luglio":    numMese = "07"
        Case "Agosto":    numMese = "08"
        Case "Settembre": numMese = "09"
        Case "Ottobre":   numMese = "10"
        Case "Novembre":  numMese = "11"
        Case "Dicembre":  numMese = "12"
        Case Else:        numMese = "00"
    End Select

    ' Trova l'ultimo giorno effettivo del mese (evita righe vuote a fine tabella)
    ultimaRiga = 41
    Do While ultimaRiga > 11 And Trim(ws.Cells(ultimaRiga, 2).Text) = ""
        ultimaRiga = ultimaRiga - 1
    Loop

    ' Percorso di salvataggio: stessa cartella del file Excel
    percorsoCartella = ThisWorkbook.Path
    If percorsoCartella = "" Then
        MsgBox "Salva prima il file Excel sul disco per poter esportare il PDF.", vbCritical
        Exit Sub
    End If

    nomeFile = "Timbrature_" & ThisWorkbook.Sheets("Dashboard").Range("A1").Value & "_" & numMese & "_" & ws.Name & ".pdf"
    percorsoCompleto = percorsoCartella & Application.PathSeparator & nomeFile

    ' Memorizza lo stato di protezione e sblocca SOLO se il foglio era protetto.
    ' Alla fine la protezione viene rimessa solo in quel caso: un foglio
    ' non protetto resta modificabile.
    eraProtetto = ws.ProtectContents
    If eraProtetto Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
    End If

    ' Configurazione pagina: Orizzontale, 1 pagina A4, da B2 a Q(ultima riga) per escludere i pulsanti in riga 1
    With ws.PageSetup
        .PrintArea = "$B$2:$Q$" & ultimaRiga
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
    End With

    ' Esportazione e apertura automatica del PDF
    On Error GoTo ErrorePDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=percorsoCompleto, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    If eraProtetto Then ws.Protect
    Exit Sub

ErrorePDF:
    On Error Resume Next
    If eraProtetto Then ws.Protect
    On Error GoTo 0
    MsgBox "Impossibile salvare il PDF. Verifica che il file '" & nomeFile & "' non sia già aperto in un altro programma.", vbCritical, "Errore Esportazione"
End Sub
