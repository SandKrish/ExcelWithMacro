Attribute VB_Name = "Module1"

Public Sub main()
    Dim Counter As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strDefectTitle As String
    Dim strStatus As String
        j = 1
        k = 1
        NumOfDefect Counter
        
            MainSheetAct
            Range("C2").Select
            Do While j <= Counter
            
                
                If ActiveCell.Value <> "" Then
                    CellOffset
                    
                        If j = Counter Then
                            Range("B2").Select
                            
                            Do While k <= Counter
                                
                                If ActiveCell.Value = "" Then
                                MsgBox "Defect Status is not filled, at" & k & "th Row Please fill and re-execute"
                                k = Counter + 1
                                                               
                                ElseIf ActiveCell.Value <> "" And k = Counter Then
                                Initialisation
                                Sorting
                                CountingSheetDefect
                                ToSendEmail
                                End If
                                
                                
                            CellOffset
                            k = k + 1
                            Loop
                            
                            
                        End If
                        
                ElseIf ActiveCell.Value = "" Then
                    
                        MsgBox "Sev is not filled, at " & j & "th Row Please fill and re-execute"
                        j = Counter + 1
                End If
                
            j = j + 1
            Loop
            

End Sub





Public Sub Sorting()
'MsgBox "RunningSorting"

  
    Dim i As Integer
    Dim Counter As Integer
    Dim strStatus As String
    Dim strDefectTitle As String
    
    
    i = 1
    NumOfDefect Counter
    
        MainSheetAct
        Range("C2").Select
        
        Do While i <= Counter
        CopystrStatus strStatus
        CopystrDefectTitle strDefectTitle
                
                If ActiveCell.Value = "Critical" Then
                    
                        SevCritical
                            
                            If ActiveCell.Value <> "" Then
                                'MsgBox " I am here"
                                PastestrStatus strStatus
                                PasteStrDefectTitle strDefectTitle
                                                              
                            ElseIf ActiveCell.Value = "" Then
                                 PasteStrStaForNull strStatus
                                  PasteStrDefForNull strDefectTitle
                                                        
                            End If
                            CellOffset
                                                                            
                                                      
                ElseIf ActiveCell.Value = "High" Then
                    
                            SevHigh
                                
                                If ActiveCell.Value <> "" Then
                                    PastestrStatus strStatus
                                    PasteStrDefectTitle strDefectTitle
                                
                                    ElseIf ActiveCell.Value = "" Then
                                        PasteStrDefForNull strDefectTitle
                                        PasteStrStaForNull strStatus
                                        
                                                            
                                End If
                                CellOffset
                
                ElseIf ActiveCell.Value = "Low" Then
                    
                    
                            SevLow
                                
                                If ActiveCell.Value <> "" Then
                                    PastestrStatus strStatus
                                    PasteStrDefectTitle strDefectTitle
                                    
                                    
                                    ElseIf ActiveCell.Value = "" Then
                                        PasteStrStaForNull strStatus
                                        PasteStrDefForNull strDefectTitle
                                                                
                                End If
                                CellOffset
               
            
            End If
            
            MainSheetAct
            CellOffset
           
            
        
        i = i + 1
        Loop
        


'MsgBox "End Of Sorting"
Range("A2").Select
End Sub

Public Function CopystrDefectTitle(CopystrDefTitle) As String
CopystrDefTitle = ActiveCell.Offset(0, -2).Value
End Function


Public Function CopystrStatus(CopystrStat) As String
CopystrStat = ActiveCell.Offset(0, -1).Value
'MsgBox CopystrStat & "Correctly Copied"
End Function


Public Function PastestrStatus(PasteStrStat) As String
ActiveCell.Offset(1, 1).Value = PasteStrStat
End Function

Public Function PasteStrDefectTitle(PasteStrDefTitle) As String
ActiveCell.Offset(1, 0).Value = PasteStrDefTitle
End Function

Public Function PasteStrDefForNull(PasteDefForNull) As String
ActiveCell.Offset(0, 0).Value = PasteDefForNull
'MsgBox PasteDefForNull & " Inside function"
End Function

Public Function PasteStrStaForNull(PasteStaForNull) As String
ActiveCell.Offset(0, 1).Value = PasteStaForNull
'MsgBox PasteStaForNull & " = Inside Function "
End Function

Public Sub CellOffset()
ActiveCell.Offset(1, 0).Select
End Sub

Public Function MainSheetAct()
Worksheets("Main").Activate
End Function
Public Function SevCritical()
Worksheets("SevCritical").Activate
End Function
Public Function SevHigh()
Worksheets("SevHigh").Activate
End Function
Public Function SevLow()
Worksheets("SevLow").Activate
End Function

Public Function DefectAnalysis()
Worksheets("DefectAnalysis").Activate
End Function



Sub Initialisation()
'
' Initialisation Macro
' Initialises the Sheet for Main to run
'

'
    'MsgBox "Starting to Intialise"
    SevCritical
    ClearContent
    
    
    SevHigh
    ClearContent
    
    SevLow
    ClearContent
    MainSheetAct
End Sub





Public Function OpenAndCloseDefect(l, n, a, b, c, d, e, Counter) As Integer
l = 0
n = 0
a = 0
b = 0
c = 0
d = 0
e = 0
Dim coun As Integer
Dim Val As String
AutFit
coun = 1

Range("B2").Select
        
        'MsgBox Counter & "InsideOpen"
        Do While coun <= Counter
                Val = ActiveCell.Value
                
                Select Case Val
                    
                    Case "Open"
                        l = l + 1
                        
                    Case "Close"
                        n = n + 1
                        
                    Case "Rejected"
                        a = a + 1
                        
                    Case "Fixed"
                        b = b + 1
                        
                    Case "Duplicate"
                        c = c + 1
                        
                    Case "Deferred"
                        d = d + 1
                        
                    Case "Reopened"
                        e = e + 1
                        
                    Case Else
                        coun = Counter + 1
                        
                End Select
                    
            
            
            CellOffset
        coun = coun + 1
        Loop
        


Range("A1").Select
Selection.CurrentRegion.Select
BorderForInsideSheet

DefectAnalysis

End Function

Public Function DefectValueInsertion(l, n, a, b, c, d, e) As Integer

ActiveCell.Offset(0, 1).Value = l 'Open
ActiveCell.Offset(0, 2).Value = n 'Close
ActiveCell.Offset(0, 3).Value = a 'Rejected
ActiveCell.Offset(0, 4).Value = b 'Fixed
ActiveCell.Offset(0, 5).Value = c 'Duplicate
ActiveCell.Offset(0, 6).Value = d 'Deferred
ActiveCell.Offset(0, 7).Value = e 'Reopened

End Function




Public Function NumOfDefect(Counter) As Integer
MainSheetAct

Counter = Range("C32").Value
End Function

Public Function CountingSheetDefect() As Integer
NumOfDefect Counter

SevCritical
Range("B2").Select
OpenAndCloseDefect l, n, a, b, c, d, e, Counter
Range("D4").Select
DefectValueInsertion l, n, a, b, c, d, e

SevHigh
Range("B2").Select
OpenAndCloseDefect l, n, a, b, c, d, e, Counter
Range("D5").Select
DefectValueInsertion l, n, a, b, c, d, e

SevLow
Range("B2").Select
OpenAndCloseDefect l, n, a, b, c, d, e, Counter
Range("D6").Select
DefectValueInsertion l, n, a, b, c, d, e

Range("A1").Select
End Function


Public Function AutFit()
'MsgBox "Going to AutoFit here"
Columns("A:B").EntireColumn.AutoFit
End Function


Public Sub ClearContent()
Range("A2:J100").Select
Selection.ClearContents
BoarderCleaning
Range("A2").Select
End Sub



Public Sub BorderForInsideSheet()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1:B1").Select
    Selection.AutoFilter
    Range("A1").Select
End Sub



Public Sub BoarderCleaning()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1:B1").Select
    Selection.AutoFilter
End Sub

Sub ToSendEmail()
    
    Dim i As Integer
    Dim Vari As String
    MainSheetAct
    Vari = Range("F9").Value
    i = 1
            
            
            If Vari = "" Then
                MsgBox "No Email Id Entered"
            End If
            
            
    Dim StartMsg As String
    Dim EndMsg As String
    Dim Signature As String
    
    Dim OutApp As Outlook.Application
    
        Set OutApp = New Outlook.Application
        
        
    If OutApp.DefaultProfileName <> "" Then
              
       If OutApp.Session.Accounts.Count > 0 Then
                         
               On Error Resume Next
                Dim OutMail As Outlook.MailItem
                    Set OutMail = OutApp.CreateItem(olMailItem)
                
                        OutMail.To = Vari
                        OutMail.CC = "TestwithSandhya@gmail.com"
                    
                        OutMail.Subject = "Auto Generated Email:Excel for Defect Analysis"
                        
                        StartMsg = "<font size='5' color='black'> Hi," & "<br> <br>" & "Please find the Defect Analysis Charts: " & "<br> <br> </font>"
                        EndMsg = "<font size='4' color='black'> Regards," & "<br> </font>"
                        Signature = "<font size='4' color='black'> Sandhya" & "<br> <br> </font>"
                        OutMail.HTMLBody = xStartMsg & xEndMsg & xSignature
                                               
                                                        
                            OutMail.Attachments.Add ActiveWorkbook.FullName
                            OutMail.Send
                                      
                    Set OutMail = Nothing
                    Set OutApp = Nothing
        
            ElseIf OutApp.Session.Accounts.Count <= 0 Then
            MsgBox "No Email configured in your Outlook"
            MsgBox "Excel couldnot be forwarded to your email"
            
         End If
            
     ElseIf OutApp.DefaultProfileName = "" Then
     MsgBox "No Outlook Found"
     MsgBox "Excel couldnot be forwarded to your email"
     End If
        

End Sub





