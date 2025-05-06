'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime
'
' Author: Pan, Jian-Hua(Jonathan)
' Email 1& MS Teams: hdpanjianhua@msn.com
' Email 2: fdpjh@126.com 
' 
' Pretranslate non-100% TRN string list(s) per L10N drop

Sub Main

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '    Dim prj As PslProject
 '    Set prj = PSL.Projects.Open("C:\test.lpu")
 '    Set prj = PSL.ActiveProject
 '    If prj Is Nothing Then Exit Sub 
    
    'Clear the output messages window
	Dim outputWnd As PslOutputWnd
	Set outputWnd=PSL.OutputWnd(pslOutputWndMessages)
	outputWnd.Clear
    
    Dim prj As PslProject
    Dim src As PslSourceList
    Dim trn As PslTransList 
    Dim i As Integer 

    Dim StringListPretranslateLogFile As String
'   StringListPretranslateLogFile = prj.Location & "\" & prj.Name &"_src-trn_strlist_change_log.txt"     
    StringListPretranslateLogFile = "C:\test\psl_strlist_pretranslate_log.txt"

    ' Declare a FileSystemObject.
	Dim fso As New FileSystemObject
	' Declare a TextStream.
	Dim stream As TextStream
	' Create a TextStream.
	Set stream = fso.CreateTextFile(StringListPretranslateLogFile, True) 
    
    Dim today
    today = Date & "  "  &Time

	stream.WriteLine today
	stream.WriteLine " "  
    stream.WriteLine "******************************************************************"
    
    For i = 1 To PSL.Projects.Count
    Set prj = PSL.Projects(i) 
        
        prj.SuspendSaving

        stream.WriteLine " " 
        stream.WriteLine prj.Name 
        stream.WriteLine " "

		' Dim trn As PslTransList
		For Each trn In prj.TransLists
		    If (Dir(trn.SourceList.SourceFile) <> "") And (trn.Size <> 0) And (trn.TransRate <> 100) And (trn.Language.LangCode <> "ita") And (trn.Language.LangCode <> "nld") And (trn.Language.LangCode <> "ptb") Then

                PSL.Output trn.Language.LangCode & " : " & trn.Title
                stream.WriteLine trn.Language.LangCode & " : " & trn.Title
                ' stream.WriteLine "trn src FileDate: " & trn.SourceList.FileDate
                ' stream.WriteLine "trn LastChange: " & trn.LastChange
                ' stream.WriteLine " "
                ' stream.WriteLine "trn src LastUpdate: " & trn.SourceList.LastUpdate
                ' stream.WriteLine "trn LastUpdate: " & trn.LastUpdate
				' stream.WriteLine "trn src LastChange: " & trn.SourceList.LastChange

				' trn.Update
				trn.AutoTranslate
				trn.Save

		    End If
		Next trn     

	    PSL.Output(" ")
	    PSL.Output "The non-100% Translation String List(s) has been 'pretranslated' now!"
	    PSL.Output("Click [[shell:C:\test\psl_strlist_pretranslate_log.txt|here]] for pretranslated non-100% TRN string lists in details.")
	    stream.WriteLine " "
	    stream.WriteLine "The non-100% Translation String List(s) has been 'pretranslated' now!"
	    stream.WriteLine " "
	    stream.WriteLine "******************************************************************" 
        
        prj.ResumeSaving 

	Next i
    
    ' Close text log file
    stream.Close 

End Sub

