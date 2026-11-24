Attribute VB_Name = "modBackup"
Option Explicit

'=============================================================================
' Modulo  : modBackup
'
' Copia di sicurezza automatica: a ogni salvataggio riuscito il file viene
' copiato nella sottocartella "Backup" accanto al file, con la data nel nome:
'
'     Backup\Dashboard_Timbrature_2026_2026-11-24.xlsm
'
'   - una copia per giorno: i salvataggi dello stesso giorno aggiornano la
'     copia di quel giorno (resta l'ultima versione della giornata);
'   - si tengono solo le ultime NUM_COPIE giornate: le piu' vecchie vengono
'     cancellate da sole;
'   - non blocca MAI il salvataggio: in caso di problemi la copia viene
'     semplicemente saltata.
'
' Per attivarla serve l'evento Workbook_AfterSave nel modulo
' "Questa_cartella_di_lavoro" (vedi istruzioni).
' Per disattivarla: BACKUP_ATTIVO = False.
'
' Per ripristinare: apri la copia dalla cartella Backup e usa
' File > Salva con nome per rimetterla al posto del file originale.
'=============================================================================

Public Const BACKUP_ATTIVO As Boolean = True
Private Const NUM_COPIE As Long = 10
Private Const NOME_CARTELLA As String = "Backup"


'-----------------------------------------------------------------------------
' Chiamata da Workbook_AfterSave dopo ogni salvataggio riuscito
'-----------------------------------------------------------------------------
Public Sub CreaBackupGiornaliero()

    Dim sSep As String, sCartella As String, sBase As String, sEst As String
    Dim sCopia As String

    If Not BACKUP_ATTIVO Then Exit Sub

    On Error GoTo Fine                        ' mai bloccare il salvataggio

    ' file non salvato in locale (es. aperto da OneDrive via web): salto
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Sub

    sSep = Application.PathSeparator

    ' se il file aperto e' gia' una copia della cartella Backup, non ricopio
    If StrComp(Mid$(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, sSep) + 1), _
               NOME_CARTELLA, vbTextCompare) = 0 Then Exit Sub

    sCartella = ThisWorkbook.Path & sSep & NOME_CARTELLA
    If Len(Dir$(sCartella, vbDirectory)) = 0 Then MkDir sCartella

    sBase = NomeSenzaEstensione(ThisWorkbook.Name)
    sEst = Mid$(ThisWorkbook.Name, Len(sBase) + 1)          ' ".xlsm"
    sCopia = sCartella & sSep & sBase & "_" & Format$(Date, "yyyy-mm-dd") & sEst

    ThisWorkbook.SaveCopyAs sCopia
    Application.StatusBar = "Copia di sicurezza salvata: " & NOME_CARTELLA & sSep & _
                            Mid$(sCopia, Len(sCartella) + 2)

    EliminaCopieVecchie sCartella, sBase, sEst
    Application.OnTime Now + TimeSerial(0, 0, 5), "PulisciBarraStato"
    Exit Sub

Fine:
    Application.StatusBar = "Copia di sicurezza NON riuscita (" & Err.Description & ")"
End Sub


' Toglie il messaggio dalla barra di stato (chiamata dopo 5 secondi)
Public Sub PulisciBarraStato()
    Application.StatusBar = False
End Sub


'-----------------------------------------------------------------------------
' Tiene solo le NUM_COPIE copie piu' recenti di questo file.
' I nomi finiscono con la data AAAA-MM-GG, quindi l'ordine alfabetico e'
' anche l'ordine cronologico.
'-----------------------------------------------------------------------------
Private Sub EliminaCopieVecchie(ByVal sCartella As String, ByVal sBase As String, _
                                ByVal sEst As String)
    Dim sSep As String, sFile As String
    Dim aFile() As String, n As Long
    Dim i As Long, j As Long, sTmp As String

    sSep = Application.PathSeparator
    ReDim aFile(0 To 0)

    sFile = Dir$(sCartella & sSep & sBase & "_*" & sEst)
    Do While Len(sFile) > 0
        ' solo i file con il formato <base>_AAAA-MM-GG<est>
        If Mid$(sFile, Len(sBase) + 2, 10) Like "####-##-##" And _
           Len(sFile) = Len(sBase) + 11 + Len(sEst) Then
            ReDim Preserve aFile(0 To n)
            aFile(n) = sFile
            n = n + 1
        End If
        sFile = Dir$
    Loop
    If n <= NUM_COPIE Then Exit Sub

    ' ordinamento semplice (pochi file)
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If aFile(j) < aFile(i) Then
                sTmp = aFile(i): aFile(i) = aFile(j): aFile(j) = sTmp
            End If
        Next j
    Next i

    ' cancello le piu' vecchie
    For i = 0 To n - NUM_COPIE - 1
        Kill sCartella & sSep & aFile(i)
    Next i
End Sub


Private Function NomeSenzaEstensione(ByVal sNome As String) As String
    Dim p As Long
    p = InStrRev(sNome, ".")
    If p > 0 Then NomeSenzaEstensione = Left$(sNome, p - 1) Else NomeSenzaEstensione = sNome
End Function
