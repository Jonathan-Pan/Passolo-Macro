' The Passolo Macro is developed to extract all of ReadOnly Source String of active Passolo lpu project file.
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

    resultFile = prj.Location & "\" & prj.Name &"_ReadOnlyAllString.txt"
    prj.LogMessage ("The Source String(ReadOnly) has been written to " &resultFile &" now!")

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

            If srcStr.State(pslStateReadOnly) = True Then

                PSL.Output srcStr.Text
                PSL.Output srcStr.ID
                'PSL.Output srcStr.IDName
                'PSL.Output srcStr.SourceList
                PSL.Output srcList.SourceFile


     	        stream.WriteLine srcStr.Text+"   |   "+Str(srcStr.ID)+"   |   "+srcList.SourceFile 
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
    
    MsgBox("Extracting all of ReadOnly Source Strings of active Passolo lpu is DONE now!") 

End Sub
