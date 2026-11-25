Attribute VB_Name = "modMenuTipiAssenza"
Option Explicit

'=============================================================================
' Modulo  : modMenuTipiAssenza  (macro "una tantum", rieseguibile)
' Macro   : CorreggiMenuTipoAssenza
'
' Ripristina il menu a tendina "Tipo Assenza" dei 12 mesi con le voci una
' sotto l'altra.
'
' L'elenco delle voci viene scritto nel foglio Impostazioni (B32:B38, nome
' Tipi_Assenza) e il menu di ogni mese punta a quelle celle: cosi' non
' dipende piu' dal separatore di elenco (virgola o punto e virgola) della
' lingua di Excel.
'
' I valori validi gia' inseriti nei mesi non vengono modificati. Le celle che
' contengono un valore non presente nell'elenco (per esempio la "voce unica"
' del menu rovinato) vengono svuotate ed elencate alla fine, cosi' si
' possono reinserire scegliendo dal menu corretto.
' Va eseguita una volta in OGNI file. Dopo l'uso il modulo puo' essere rimosso.
'=============================================================================

Public Sub CorreggiMenuTipoAssenza()

    Dim vMesi As Variant, v As Variant, vTipi As Variant
    Dim wsImp As Worksheet, ws As Worksheet
    Dim shIniziale As Object
    Dim eraProtetto As Boolean
    Dim bScreenUpd As Boolean, bEvents As Boolean
    Dim i As Long, nFogli As Long
    Dim c As Range, sTipo As String, bValido As Boolean, sPuliti As String
    Dim lErrNum As Long, sErrDesc As String, sFoglioErr As String

    vMesi = Array("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", _
                  "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")
    ' ATTENZIONE: i nomi devono restare identici, sono usati dalle formule
    vTipi = Array("Ferie", "Ferie anno prec.", "Malattia", "L. 104", _
                  "Art. 23 (Visite)", "Art. 25 (Motivi pers.)", "Art. 34 (A recupero)")

    Set wsImp = OttieniFoglio("Impostazioni")
    If wsImp Is Nothing Then
        MsgBox "Serve il foglio Impostazioni.", vbExclamation, "Menu Tipo Assenza"
        Exit Sub
    End If

    Set shIniziale = ActiveSheet
    bScreenUpd = Application.ScreenUpdating
    bEvents = Application.EnableEvents

    On Error GoTo GestioneErrori
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- Elenco in Impostazioni (B31:B38) ----------------------------------------
    Set ws = wsImp
    sFoglioErr = wsImp.Name
    eraProtetto = wsImp.ProtectContents
    If eraProtetto Then wsImp.Unprotect

    With wsImp.Range("B31")
        .Value = "Tipi di assenza (menu a tendina)"
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    For i = 0 To UBound(vTipi)
        wsImp.Cells(32 + i, 2).Value = vTipi(i)
    Next i
    With wsImp.Range("B31:B38").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    wsImp.Range("B32:B38").Locked = True
    With wsImp.Range("D31")
        .Value = "Voci del menu Tipo Assenza: non modificarle, i nomi sono usati dalle formule."
        .Font.Italic = True
        .Font.Color = RGB(128, 128, 128)
    End With
    ThisWorkbook.Names.Add Name:="Tipi_Assenza", RefersTo:="=Impostazioni!$B$32:$B$38"

    If eraProtetto Then wsImp.Protect
    eraProtetto = False

    '--- Menu dei 12 mesi --------------------------------------------------------
    For Each v In vMesi
        Set ws = OttieniFoglio(CStr(v))
        If Not ws Is Nothing Then
            sFoglioErr = ws.Name
            eraProtetto = ws.ProtectContents
            If eraProtetto Then ws.Unprotect
            ' valori non validi (es. la voce unica del menu rovinato)
            For Each c In ws.Range("N11:N41").Cells
                sTipo = Trim$(CStr(c.Value))
                If Len(sTipo) > 0 And Not c.HasFormula Then
                    bValido = False
                    For i = 0 To UBound(vTipi)
                        If StrComp(sTipo, CStr(vTipi(i)), vbBinaryCompare) = 0 Then bValido = True
                    Next i
                    If Not bValido Then
                        sPuliti = sPuliti & vbCrLf & "  - " & ws.Name & " " & ws.Cells(c.Row, 2).Text & _
                                  ": " & Chr(34) & Left$(sTipo, 30) & IIf(Len(sTipo) > 30, "...", "") & Chr(34)
                        c.ClearContents
                    End If
                End If
            Next c

            With ws.Range("N11:N41").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:="=Tipi_Assenza"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ErrorTitle = "Tipo Assenza"
                .ErrorMessage = "Scegli una voce dall'elenco."
            End With
            If eraProtetto Then ws.Protect
            eraProtetto = False
            nFogli = nFogli + 1
        End If
    Next v
    sFoglioErr = ""

Uscita:
    On Error Resume Next
    If eraProtetto Then ws.Protect
    If Not shIniziale Is Nothing Then shIniziale.Activate
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreenUpd
    On Error GoTo 0

    If lErrNum <> 0 Then
        MsgBox "Errore " & lErrNum & ": " & sErrDesc & _
               IIf(Len(sFoglioErr) > 0, vbCrLf & "Foglio: " & sFoglioErr, ""), _
               vbCritical, "Menu Tipo Assenza"
    Else
        MsgBox "Menu Tipo Assenza corretto in " & nFogli & " mesi (" & _
               UBound(vTipi) + 1 & " voci, una per riga)." & _
               IIf(Len(sPuliti) > 0, vbCrLf & vbCrLf & _
                   "Celle con un valore non valido, svuotate (scegli di nuovo il tipo dal menu):" & _
                   sPuliti, "") & vbCrLf & vbCrLf & _
               "Salva il file.", IIf(Len(sPuliti) > 0, vbExclamation, vbInformation), "Menu Tipo Assenza"
    End If
    Exit Sub

GestioneErrori:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Resume Uscita
End Sub


Private Function OttieniFoglio(ByVal sNome As String) As Worksheet
    On Error Resume Next
    Set OttieniFoglio = ThisWorkbook.Worksheets(sNome)
    On Error GoTo 0
End Function
