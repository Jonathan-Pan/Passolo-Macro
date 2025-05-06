'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime
'
' Author: Pan, Jian-Hua(Jonathan)
' Email 1& MS Teams: hdpanjianhua@msn.com
' Email 2: fdpjh@126.com 
'
' Update only changed EN source/TRN target string list(s) per L10N drop

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

    Dim StringListChangeLogFile As String
'   StringListChangeLogFile = prj.Location & "\" & prj.Name &"_src-trn_strlist_change_log.txt"     
    StringListChangeLogFile = "C:\test\psl_strlist_change_log.txt"

    ' Declare a FileSystemObject.
	Dim fso As New FileSystemObject
	' Declare a TextStream.
	Dim stream As TextStream
	' Create a TextStream.
	Set stream = fso.CreateTextFile(StringListChangeLogFile, True) 
    
    Dim today
    today = Date & "  "  &Time

	stream.WriteLine today
	stream.WriteLine " "  
    stream.WriteLine "******************************************************************"
    
    For i = 1 To PSL.Projects.Count
    Set prj = PSL.Projects(i) 
        
        stream.WriteLine " " 
        stream.WriteLine prj.Name 
        stream.WriteLine " "

		' Update all string lists that need to be updated
		' Dim src As PslSourceList
		For Each src In prj.SourceLists
		    If src.FileDate > src.LastUpdate Then

                stream.WriteLine src.Title
                stream.WriteLine "src FileDate: " & src.FileDate
                stream.WriteLine "src LastUpdate: " & src.LastUpdate
                stream.WriteLine " "

  		        src.Update
		        ' prj.LogMessage (src.Title)
		        ' PSL.Output "The "& src.Title & " string list has been updated now."

		    End If
		Next src

	    PSL.Output(" ")
	    PSL.Output "The changed Source String List(s) is 'all updated' now!"
	    ' PSL.Output("Click [[shell:C:\test\psl_strlist_change_log.txt|here]] for changed EN source file in details.") 
	    stream.WriteLine " "
	    stream.WriteLine  "The changed Source String List(s) is 'all updated' now!"
	    stream.WriteLine " " 
	    stream.WriteLine "******************************************************************"

	    ' Update all translation lists that need to be updated
		' Dim trn As PslTransList
		For Each trn In prj.TransLists
		    'If (trn.SourceList.FileDate > trn.LastChange) And (trn.Language.LangCode <> "ita") And (trn.Language.LangCode <> "nld") And (trn.Language.LangCode <> "ptb") Then

                stream.WriteLine trn.Language.LangCode & " : " & trn.Title
                stream.WriteLine "trn src FileDate: " & trn.SourceList.FileDate
                stream.WriteLine "trn LastChange: " & trn.LastChange
                stream.WriteLine " "
                'stream.WriteLine "trn src LastUpdate: " & trn.SourceList.LastUpdate
                'stream.WriteLine "trn LastUpdate: " & trn.LastUpdate
				'stream.WriteLine "trn src LastChange: " & trn.SourceList.LastChange

				trn.Update
				' trn.AutoTranslate

		    'End If
		Next trn     

	    PSL.Output(" ")
	    PSL.Output "The changed Translation String List(s) is also 'all updated' now!"
	    PSL.Output("Click [[shell:C:\test\psl_strlist_change_log.txt|here]] for changed EN and TRN string list in details.")
	    stream.WriteLine " "
	    stream.WriteLine "The changed Translation String List(s) is also 'all updated' now!"
	    stream.WriteLine " "
	    stream.WriteLine "******************************************************************"
	
	Next i
    
    ' Close text log file
    stream.Close 

End Sub
