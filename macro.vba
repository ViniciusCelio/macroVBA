Sub AgendarCompromissoOutlook()
    Dim nomeColaborador, matriculaColaborador As String
    
    Dim horarioPadrao As Date
    Dim diaDoDesbloqueio As Date
    Dim dataCompromisso As Date
    
    Dim dias, col As Integer
    
    Dim ln As Long
    
    Dim ultCel As Range
    
    Dim W As Worksheet
    
    Set W = Sheets("Planilha1")
    
    Set ultCel = W.Cells(W.Rows.Count, 1).End(xlUp)
    
    'Criar instância do Outlook
    '------------------------------------------------------
    Set objOutlook = CreateObject("Outlook.Application")
    
    ln = 3
    col = 1
    
    Const olFolderCalendar = 9
    Const olAppointmentItem = 1
    
    Set objNamespace = objOutlook.GetNameSpace("MAPI")
    Set Items = objNamespace.GetDefaultFolder(olFolderCalendar).Items
    
    Set objCalendar = objNamespace.GetDefaultFolder(olFolderCalendar)
    
    Do While ln <= ultCel.Row
        If W.Cells(ln, col + 4) <> "X" Then
            Set objApt = objCalendar.Items.Add(olAppointmentItem)
            horarioPadrao = "08:45:00"
            diaDoDesbloqueio = W.Cells(ln, col + 3)
            nomeColaborador = W.Cells(ln, col + 2)
            matriculaColaborador = W.Cells(ln, col + 1)
            
            dataCompromisso = diaDoDesbloqueio + horarioPadrao
            
            If dataCompromisso <= Now Then
                dataCompromisso = Now
            End If
            
        dias = CInt(diaDoDesbloqueio - Date)
        
            If dias = 0 Then
                vdataenvio = DateAdd("n", 5, Now)
                horarioPadrao = Time + TimeValue("00:02:00")
            End If
            
            'Criar evento no Outlook
            '-----------------------------
            objApt.Subject = "Desbloquear usuário: " & matriculaColaborador & " " & nomeColaborador
            objApt.Start = dataCompromisso
            objApt.Duration = 1
            objApt.End = dataCompromisso + TimeValue("00:10:00")
            
            objApt.Save
            
            W.Cells(ln, col + 4) = "X"
            W.Cells(ln, col + 5) = Now
            
            Set objApt = Nothing
            
        End If
        
        ln = ln + 1
    Loop
    
    MsgBox "Compromissos agendados"
    
    Set W = Nothing
    Set ultCel = Nothing
    Set objApt = Nothing
    Set objCalendar = Nothing
    Set Items = Nothing
    Set objOutlook = Nothing
    
End Sub
