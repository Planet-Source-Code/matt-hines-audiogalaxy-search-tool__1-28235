Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub VisitAG(frm As Form)
ShellExecute frm.hwnd, "open", "http://www.audiogalaxy.com", "", "", 1
End Sub

Public Sub VisitQ(frm As Form)
ShellExecute frm.hwnd, "open", "http://audiogalaxy.com/satellite/queue.php?", "", "", 1
End Sub

Public Sub SearchFTP(frm As Form)
ShellExecute frm.hwnd, "open", "http://audiogalaxy.com/list/searches.php?SID=2142345c46c14d3bfcda77294caa79cf&searchType=1&searchStr=" & Form1.Text1.Text, "", "", 1
End Sub

Public Sub SearchMusic(frm As Form)
ShellExecute frm.hwnd, "open", "http://audiogalaxy.com/list/searches.php?SID=2142345c46c14d3bfcda77294caa79cf&searchType=0&searchStr=" & Form1.Text1.Text, "", "", 1
End Sub
