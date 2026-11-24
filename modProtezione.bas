Attribute VB_Name = "modProtezione"
Option Explicit

'=============================================================================
' Modulo  : modProtezione
'
' ProteggiFogli  -> protegge (SENZA password) Dashboard, Festivita',
'                   Impostazioni e i 12 mesi, lasciando modificabili solo le
'                   celle da compilare:
'                     - mesi:          Assenze, Entr/Usc (E11:I41) e
'                                      Tipo Assenza (N11:N41)
'                     - Impostazioni:  le celle gialle
'                     - Dashboard:     Residuo anno precedente (C4) e
'                                      Ferie anno precedente (I7)
'                     - Festivita':    la riga libera 15 (A15:B15)
'                   Tutto il resto (formule, intestazioni, anno) e' bloccato,
'                   cosi' un tocco sbagliato non cancella una formula.
'
' SbloccaFogli   -> toglie la protezione da tutti questi fogli in un colpo
'                   (per modifiche straordinarie). Poi rilancia ProteggiFogli.
'
' Le altre macro del file (importazione, PDF, nuovo anno, controlli)
' tolgono e rimettono da sole la protezione quando serve.
' Per sbloccare a mano un solo foglio: Revisione > Rimuovi protezione foglio.
'=============================================================================

Public Sub ProteggiFogli()

    Dim vMesi As Variant, v As Variant
    Dim ws As Worksheet
    Dim nFogli As Long
    Dim sMancanti As String
    Dim bScreenUpd As Boolean
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    If MsgBox("Verranno protetti (senza password) Dashboard, Festivit" & ChrW(224) & _
              ", Impostazioni e i 12 mesi." & vbCrLf & vbCrLf & _
              "Resteranno modificabili solo le celle da compilare." & vbCrLf & _
              "Procedere?", vbQuestion + vbYesNo, "Protezione") <> vbYes Then Exit Sub

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    bScreenUpd = Application.ScreenUpdating
    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False

    ' --- fogli mensili ---
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If ws Is Nothing Then
            sMancanti = sMancanti & " " & v
        Else
            sFoglioErr = ws.Name
            ProteggiConCelleLibere ws, "E11:I41,N11:N41"
            nFogli = nFogli + 1
        End If
    Next v

    ' --- Impostazioni: le celle gialle ---
    Set ws = OttieniFoglio("Impostazioni")
    If Not ws Is Nothing Then
        sFoglioErr = ws.Name
        ProteggiConCelleLibere ws, "C6:C12,C15:C16,C20:C26"
        nFogli = nFogli + 1
    Else
        sMancanti = sMancanti & " Impostazioni"
    End If

    ' --- Dashboard: valori riportati dall'anno precedente ---
    Set ws = OttieniFoglio("Dashboard")
    If Not ws Is Nothing Then
        sFoglioErr = ws.Name
        ProteggiConCelleLibere ws, "C4,I7"
        nFogli = nFogli + 1
    Else
        sMancanti = sMancanti & " Dashboard"
    End If

    ' --- Festivita': riga libera per una festivita' in piu' ---
    Set ws = OttieniFoglio("Festivit" & ChrW(224))
    If Not ws Is Nothing Then
        sFoglioErr = ws.Name
        ProteggiConCelleLibere ws, "A15:B15"
        nFogli = nFogli + 1
    Else
        sMancanti = sMancanti & " Festivit" & ChrW(224)
    End If
    sFoglioErr = ""

Uscita:
    On Error Resume Next
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, ""), _
               vbCritical, "Protezione"
    Else
        MsgBox "Protetti " & nFogli & " fogli (senza password)." & _
               IIf(Len(sMancanti) > 0, vbCrLf & "Non trovati:" & sMancanti, "") & vbCrLf & vbCrLf & _
               "Per una modifica straordinaria: Revisione > Rimuovi protezione foglio," & vbCrLf & _
               "oppure la macro SbloccaFogli." & vbCrLf & vbCrLf & "Salva il file.", _
               vbInformation, "Protezione"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


Public Sub SbloccaFogli()
    Dim v As Variant
    Dim ws As Worksheet
    Dim n As Long

    For Each v In Array("Dashboard", "Festivit" & ChrW(224), "Impostazioni", _
                        "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                        "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            If ws.ProtectContents Then
                ws.Unprotect
                n = n + 1
            End If
        End If
    Next v

    MsgBox "Protezione tolta da " & n & " fogli." & vbCrLf & _
           "Finite le modifiche, esegui ProteggiFogli.", vbInformation, "Protezione"
End Sub


'-----------------------------------------------------------------------------
' Blocca tutte le celle del foglio tranne quelle indicate, poi lo protegge
' senza password (stesse opzioni usate dalle altre macro con ws.Protect).
'-----------------------------------------------------------------------------
Private Sub ProteggiConCelleLibere(ByVal ws As Worksheet, ByVal sCelleLibere As String)
    If ws.ProtectContents Then ws.Unprotect
    ws.Cells.Locked = True
    ws.Range(sCelleLibere).Locked = False
    ws.Protect
End Sub


Private Function OttieniFoglio(ByVal sNome As String) As Worksheet
    On Error Resume Next
    Set OttieniFoglio = ThisWorkbook.Worksheets(sNome)
    On Error GoTo 0
End Function
