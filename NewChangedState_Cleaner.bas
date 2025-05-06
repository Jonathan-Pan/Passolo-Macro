'
' The Passolo Macro is developed to remove "New" or "Changed" or "New and Changed" state of Source String before update Source String List each time:
' 1. Display EN Source String with "New" or "Changed" or "New and Changed" state before update Source String List action to Passolo Messages output window.
' 2. Remove "New" or "Changed" or "New and Changed" state from EN Source String and save its related Source List.
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

    'Dim resultFile As String

    'resultFile = prj.Location & "\" & prj.Name &"_New_Changed_New-Changed_SourceString_log.txt"
    'prj.LogMessage ("The Source String(New, Changed or 'New And Changed') before update Source String List action has been written to " &resultFile &" now!")

    'Declare a FileSystemObject.
    'Dim fso As New FileSystemObject
    'Dim fso As Object
    'Set fso = CreateObject("Scripting.FileSystemObject")

    'Declare a TextStream.
    'Dim stream As TextStream
    'Dim stream As Object
    'Create a TextStream.
    'Set stream = fso.CreateTextFile(resultFile, True)

    Dim srcList As PslSourceList
    Dim srcStr As PslSourceString
    Dim i As Integer

    For Each srcList In prj.SourceLists
        For i = 1 To srcList.StringCount
            Set srcStr = srcList.String(i)

            If srcStr.State(pslStateNew) = True Or srcStr.State(pslStateChanged) = True Or (srcStr.State(pslStateNew) = True And srcStr.State(pslStateChanged) = True)Then

                PSL.Output srcList.SourceFile
                PSL.Output srcStr.Text

     	        'stream.WriteLine srcList.SourceFile + "  :  " + srcStr.Text
                srcStr.State(pslStateNew) = False
                srcStr.State(pslStateChanged) = False
                srcList.Save

     	    End If
        Next i
    Next srcList

    'Close the file.
    'stream.Close
    
    prj.ResumeSaving

    PSL.Output "Cleaning 'New', 'Changed' or 'New and Changed' state - Done now!"

End Sub
