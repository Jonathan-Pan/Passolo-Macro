' The Passolo Macro is developed to remove all of the Source String Lists which related EN source files do not exist any longer. 
'
' Author: Pan, Jian-Hua(Jonathan)
' Email 1& MS Teams: hdpanjianhua@msn.com
' Email 2: fdpjh@126.com 
'
' If you have any question on the usage of Macro, please contact Author for support and help.
'

Sub Main

    Dim prj As PslProject
	'Replace the below sample PSL lpu with your PSL lpu file path/file name	
	'Set prj = PSL.Projects.Open("C:\test.lpu")
	Set prj = PSL.ActiveProject
    If prj is Nothing Then Exit Sub
    
    prj.SuspendSaving
    Dim srcList As PslSourceList
    
    'Dim i As Integer
    'For i = 1 To prj.SourceLists.Count 
    For Each srcList In prj.SourceLists
        Dim srcFile As Object
        Set srcFile = CreateObject("Scripting.FileSystemObject")         
    
        'Dim srcList As PslSourceList
        'Set srcList = prj.SourceLists(i) 
        'If CStr(srcList.IsOpen) = "False" Then  
        If srcFile.FileExists(srcList.SourceFile) = False Then

            PSL.Output srcList.SourceFile
            prj.SourceLists.Remove(srcList)

        End If

    Next srcList
    'Next i 
   
    prj.ResumeSaving

    Msgbox "All of invalid Source String Lists(which EN source files do not exist) have been removed completely now!"	

End Sub
