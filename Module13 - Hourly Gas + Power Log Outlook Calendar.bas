Public bWORKING As Boolean
Public bKEEPWORKING As Boolean
Public testCnt As Integer

Sub olCalSend(kw As Double, gas As Double, dtOut As Double, cntReads As Integer)

    Dim ol As outlook.Application
        Set ol = outlook.Application
    Dim Ns As outlook.Namespace
        Set Ns = outlook.Application.GetNamespace("MAPI")
    Dim folderCHP As outlook.MAPIFolder
        Set folderCHP = Ns.GetDefaultFolder(olFolderCalendar).Folders("CHP Log")
        
    Dim formatLog, calcJust As String
        formatLog = CStr(CDate(dtOut)) & " Log: " & CStr(Round(kw, 1)) & " kWh | " & CStr(Round(gas, 1)) & " SCFH"
        calcJust = "RGenPower and SCFHr read every 15 seconds and averaged for kWh and SCFH. Percent of Hour Read for Calcs: " & CStr(100 * (cntReads / 240)) & "%."
            
    Dim logAdd As outlook.AppointmentItem
    Set logAdd = folderCHP.Items.Add(olAppointmentItem)
        With logAdd
            .Subject = formatLog
            .Body = calcJust
            .Start = CDate(dtOut)
            .Duration = 60
            .Save
        End With
        
End Sub

Sub resetCalc_Click()

    If Sheet7.Cells(7, 2).Value = 0 Then
        
        Sheet7.Cells(7, 2).Value = 1
    
    Else
        
        Sheet7.Cells(7, 2).Value = 0
        Sheet7.Range("B9:B12").ClearContents
        Sheet7.Range("B14:C15").ClearContents
        Sheet7.Range("B71:C72").ClearContents
    
    End If
    
End Sub

Sub GasAndPower()
    
    Dim modStatus As String ' only calc if connection match mod open below

    Dim modOpen As String
        modOpen = "Connection success"



    Dim calcReady As Integer ' halt bit polled
    Dim dtOut  As Double
    
    Dim dtTrack, dtCheck, hrCheck As Date ' might not need timeDate or tdCheck

    Dim dayMatch, hmiMatch, goodDT, goodRead As Boolean ' determine if accurate datetime and/or current log tracking in hmi is consistent
    
    Dim cntReads As Integer ' counter for reads within the hour and output of count for accuracy %
        
    Dim kwRead, kwSum, kwAvg As Double
    Dim gasRead, gasSum, gasAvg As Double
    
    Dim dtStore As Date ' one for hmi tracking and one for olCalSend call NEEDS TOUCHUP
    
    Dim logNtitle, descNbody As String

    ' end declarations
    
    
    modStatus = Main.Cells(43, 2).Value
    calcReady = Sheet7.Cells(7, 2).Value

    ' bool flags to determine data legitimacy    
    dayMatch = False
    hmiMatch = False
    goodDT = False
    goodRead = False
    
    If bWORKING Then
        Exit Sub
    End If
    
    bWORKING = True
    
    
    If calcReady = 0 Then End
    
    cntReads = 0
    
    If modStatus = modOpen And calcReady = 1 Then
        
        bKEEPWORKING = True
        
        cntReads = 1
        cntReads = cntReads + Sheet7.Cells(12, 2).Value
        
        If Not IsEmpty(Sheet7.Cells(10, 2).Value) And Not IsEmpty(Sheet7.Cells(11, 2).Value) Then
            goodDT = True
            dtTrack = CDbl(CDate(Sheet7.Cells(10, 2).Value))
            dtTime = CDbl(CDate(Sheet7.Cells(11, 2).Value))
        End If
        
        
        dtCheck = CDbl(Date + (Hour(Now) / 24))
        
        If goodDT And dtTrack = dtCheck Then dtMatch = True
        If goodDT And dtTrack = dtTime Then hmiMatch = True

        If Not IsEmpty(Sheet7.Cells(14, 3).Value) And Not IsEmpty(Sheet7.Cells(15, 3).Value) And cntReads > 1 Then goodRead = True

        If dtMatch And hmiMatch And goodRead Then
            
            kwRead = Main.Cells(111, 22).Value + 0
            kwSum = Sheet7.Cells(14, 2).Value + kwRead
            kwAvg = kwSum / cntReads
            
            gasRead = Main.Cells(7, 22).Value + 0
            gasSum = Sheet7.Cells(15, 2).Value + gasRead
            gasAvg = gasSum / cntReads
            
            dtOut = CDbl(dtTrack)
            
            Sheet7.Cells(10, 2).Value = Date + (Hour(Time) / 24)
            
        ElseIf hmiMatch And goodRead And Not dtMatch Then
            
            kwAvg = Sheet7.Cells(14, 3).Value
            gasAvg = Sheet7.Cells(15, 3).Value
            logNtitle = CStr(kwAvg) & " " & CStr(gasAvg)
            
            descNbody = " please work"
            dtOut = CDbl(dtTrack)
            
            Call olCalSend(kwAvg, gasAvg, dtOut, cntReads)
            
            cntReads = 1
            
            kwRead = Main.Cells(111, 22).Value + 0
            kwSum = kwRead
            kwAvg = kwSum
            
            gasRead = Main.Cells(7, 22).Value + 0
            gasSum = gasRead
            gasAvg = gasSum
        
        Else
            cntReads = 1
        
            kwRead = Main.Cells(111, 22).Value + 0
            kwSum = kwRead
            kwAvg = kwSum
            
            gasRead = Main.Cells(7, 22).Value + 0
            gasSum = gasRead
            gasAvg = gasSum
            
        End If
        
        Sheet7.Cells(9, 2).Value = CDate(dtCheck)
        Sheet7.Cells(10, 2).Value = CDbl(Date + (Hour(Time) / 24))
        Sheet7.Cells(11, 2).Value = CDbl(dtCheck)
        Sheet7.Cells(12, 2).Value = cntReads
        'Sheet7.Cells(13, 1).Value = CDbl(dtCheck) + (1 / 24)
        Sheet7.Cells(14, 2).Value = kwSum
        Sheet7.Cells(14, 3).Value = kwAvg
        Sheet7.Cells(15, 2).Value = gasSum
        Sheet7.Cells(15, 3).Value = gasAvg
        
    End If
    
    bWORKING = False
        
    Application.OnTime Now + TimeSerial(0, 0, 10), "GasAndPower"
        
End Sub




