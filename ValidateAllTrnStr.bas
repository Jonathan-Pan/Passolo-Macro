'
' The Passolo Macro is developed to change all of translated strings to "Validated" state of all of translation string lists of all of "Opened" Passolo lpus.
'
' Author: Pan, Jian-Hua(Jonathan)
' Email 1& MS Teams: hdpanjianhua@msn.com
' Email 2: fdpjh@126.com
'
' If you have any question on the usage of Macro, please contact Author for support and help.
'

Sub Main

    Dim prj As PslProject
    Dim prji As Integer

    For prji = 1 To PSL.Projects.Count
        Set prj = PSL.Projects(prji)

        'Set prj = PSL.ActiveProject
        'If prj Is Nothing Then Exit Sub

        prj.SuspendSaving
    
        Dim trnList As PslTransList
        Dim trnStr As PslTransString
        Dim i As Integer

        For Each trnList In prj.TransLists

            For i = 1 To trnList.StringCount
            
                Set trnStr = trnList.String(i)

                If trnStr.State(pslStateReview) = True Then

                    trnStr.State(pslStateReview) = False
                    trnList.Save

     	        End If
            Next i
        Next trnList

        prj.ResumeSaving

    Next prji

    PSL.Output "Change all of translation strings to 'Validated' state - Done now!"

End Sub
