'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime
'
' Author: Pan, Jian-Hua(Jonathan)
' Email 1& MS Teams: hdpanjianhua@msn.com
' Email 2: fdpjh@126.com 
'
' Exports the source list to the output window and external result file
	
Sub Main
  Dim prj As PslProject
  Set prj = PSL.ActiveProject
  If prj Is Nothing Then Exit Sub

	'Clear the output messages window
	Dim outputWnd As PslOutputWnd
	Set outputWnd=PSL.OutputWnd(pslOutputWndMessages)
	outputWnd.Clear

Dim resultFile As String

resultFile = prj.Location & "\" & prj.Name &"_filelist.txt"
prj.LogMessage ("Source list paths are written to" &resultFile)

' Declare a FileSystemObject.
Dim fso As New FileSystemObject
' Declare a TextStream.
Dim stream As TextStream
' Create a TextStream.
Set stream = fso.CreateTextFile(resultFile, True)

  Dim srcList As PslSourceList
  For Each srcList In prj.SourceLists
  	prj.LogMessage (srcList.SourceFile)
	stream.WriteLine srcList.SourceFile
  Next srcList

' Close the file.
stream.Close

End Sub
