' The Passolo Macro is developed to track any "content Change" of Source String with original "Read-Only" state:
' 1. Display EN Source String together with its Source List to Passolo Messages output window.
' 2. Write EN Source String together with its Source List to a log file outside of Passolo application.
' 3. Remove Read-Only state from the Changed string and save its related Source List.
'
' Author: Pan, Jian-Hua(Jonathan)
' Email 1& MS Teams: hdpanjianhua@msn.com
' Email 2: fdpjh@126.com 
'
' If you have any question on the usage of Macro, please contact Author for support and help.
'

Sub Main
    Dim prj As PslProject
    Set prj = PSL.ActiveProject
    If prj Is Nothing Then Exit Sub
    
    prj.SuspendSaving

    'Clear the output messages window
	Dim outputWnd As PslOutputWnd
	Set outputWnd=PSL.OutputWnd(pslOutputWndMessages)
	outputWnd.Clear

    Dim resultFile As String

    resultFile = prj.Location & "\" & prj.Name &"_ReadOnly-Changed_SourceString_log.txt"
    prj.LogMessage ("The Source String(ReadOnly and Changed) has been written to " &resultFile &" now!")

    'Declare a FileSystemObject.
    'Dim fso As New FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Declare a TextStream.
    'Dim stream As TextStream
    Dim stream As Object
    'Create a TextStream.
    Set stream = fso.CreateTextFile(resultFile, True)

    Dim srcList As PslSourceList
    Dim srcStr As PslSourceString
    Dim i As Integer

    For Each srcList In prj.SourceLists
        For i = 1 To srcList.StringCount
            Set srcStr = srcList.String(i)

            If srcStr.State(pslStateReadOnly) = True And srcStr.State(pslStateChanged) = True Then

                PSL.Output srcList.SourceFile
                PSL.Output srcStr.Text

     	        stream.WriteLine srcList.SourceFile + "  :  " + srcStr.Text
                ' Automatically remove Read-Only state or not? Need team member's further input.
                srcStr.State(pslStateReadOnly) = False
                srcList.Save

            'ElseIf srcStr.State(pslStateReadOnly) = True And srcStr.State(pslStateChanged) = False Then

                'PSL.Output srcList.SourceFile + "  :  " + srcStr.Text + "  - no content Change this Time!"
                'stream.WriteLine srcList.SourceFile + "  :  " + srcStr.Text + "  - no content Change this Time!"

     	    End If
        Next i
    Next srcList

    'Close the file.
    stream.Close
    
    prj.ResumeSaving

End Sub
