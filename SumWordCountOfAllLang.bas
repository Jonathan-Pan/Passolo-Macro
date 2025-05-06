'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime
'
' Author: Pan, Jian-Hua(Jonathan)
' Email 1& MS Teams: hdpanjianhua@msn.com
' Email 2: fdpjh@126.com 
' 
' Get word count info of all *.lpu project files openned in Passolo application. 

Sub Main

    Dim Today
    Today = Date & "  "  &Time

    Dim prj As PslProject
    Dim i As Integer

    For i = 1 To PSL.Projects.Count
    Set prj = PSL.Projects(i)

        PSL.Output Today
        PSL.Output " "
        PSL.Output prj.Name
        PSL.Output " "
        PSL.Output "---------------------------------------------------"
        PSL.Output "The project contains " & Str(prj.Languages.Count) & " target languages."
        PSL.Output "---------------------------------------------------"
        PSL.Output " "

        Dim lang As PslLanguage

            For Each lang In prj.Languages

            Dim langWC As PslStatistics
            Set langWC = lang.GetStatistics

            PSL.Output lang.LangCode & " Untranslated:  " & langWC.ToTranslate.WordCount
            PSL.Output lang.LangCode & " Untranslated Repetitions:  " & langWC.Repeats.WordCount
            PSL.Output "---------------------------------------------------"
            PSL.Output " "

        Next lang
    Next i

    PSL.Output(" ")
    PSL.Output(" ")
    PSL.Output("All of opened PSL lpus word count statistics info is displayed now!")
    ' Need replace the below log file path with yours
    PSL.Output("Please click [[shell:C:\test\AllPSLlpu_WordCount.log|here]] to open word count log file to paste data for further estimation later.")
    PSL.Output(" ")
    PSL.Output(" ")

End Sub
