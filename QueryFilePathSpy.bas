' The Passolo Macro is developed to display EN Source file path of a Translation list to Passolo Messages output window.
' The Translation String List is Selected or Opened by translator for Query purpose. 
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
    
	  'Clear the output messages window
	  Dim outputWnd As PslOutputWnd
	  Set outputWnd=PSL.OutputWnd(pslOutputWndMessages)
	  outputWnd.Clear

    Dim trn As PslTransList
    ' Only "one" Translation String List is "Selected" or "Opened".
    Set trn = PSL.ActiveTransList

    If trn Is Nothing Then
        PSL.Output "No active translation list found."
    Else
        PSL.Output trn.SourceList.SourceFile
    End If

End Sub
