Attribute VB_Name = "Module1"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal Flags As Long, ByVal ExtraInfo As Long)

Public Type POINTAPI
    x As Long
    y As Long
    End Type
Dim Excel As Excel.Application  ' This is the excel program
Dim ExcelWBk As Excel.Workbook  ' This is the work book
Dim ExcelWS As Excel.Worksheet  ' This is one sheet
Dim ExcelWS2 As Excel.Worksheet
Public PptApp As PowerPoint.Application
Public Present As PowerPoint.Presentation
Public Slds As PowerPoint.Slides
Public Sld As PowerPoint.Slide
Public Shp As PowerPoint.Shape


Public continue As Boolean
Public aaa

Public Sub Main()
'GoTo dasa
'SendKeys "^(a)"
DetectIE
For i = 1 To 3
SendKeys "^{TAB}", True
SendKeys "^(c)", True

aaa = Clipboard.GetText
posit = InStr(1, aaa, "http", vbTextCompare)

If posit = 1 Then Exit For
Next
posit2 = InStr(8, aaa, "/", vbTextCompare)
If posit2 = 0 Then
posit2 = Len(aaa)
End If
aaa = Replace(aaa, ".", "", 1, -1, vbTextCompare)
aaa = Replace(aaa, "\", "", 1, -1, vbTextCompare)
aaa = Replace(aaa, "/", "", 1, -1, vbTextCompare)
aaa = Replace(aaa, ":", "", 1, -1, vbTextCompare)

aaa = Mid(aaa, 7, posit2 - 7)
ScreenToClipboard
Do
DoEvents
Loop While continue = False
'dasa:
Form1.Left = GetX * 15
Form1.Top = GetY * 15
Form1.Show
'form1.Left = screen.
End Sub
Sub ConvertWordDoc()
   'On Error Resume Next
    Static WordObj As Word.Application
    Set WordObj = CreateObject("Word.Application")

  WordObj.Documents.Add , , wdFormatDocument
  WordObj.ActiveDocument.content.Paste
  If Form1.Check1.Value = 1 Then
  WordObj.Visible = True
Else
   WordObj.Visible = False
    WordObj.ActiveDocument.SaveAs App.Path & "\" & aaa & ".doc", wdFormatDocument
    WordObj.Quit savechanges:=False
End If
    Set WordObj = Nothing
End Sub
Sub ScreenToClipboard()
continue = False
    Const VK_SNAPSHOT = &H2C
    'Call keybd_event(VK_SNAPSHOT, 1, 0&, 0&) 'for whole screen
    Call keybd_event(VK_SNAPSHOT, 0&, 0&, 0&) 'for active window
continue = True
End Sub
Public Function GetX() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetX = n.x
End Function


Public Function GetY() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetY = n.y
End Function

Sub ConvertExcel()
   'On Error Resume Next
    Static WordObj As Excel.Application
   Set WordObj = CreateObject("Excel.Application")
Set ExcelWBk = WordObj.Workbooks.Add
Set ExcelWS = WordObj.Worksheets(1)

ExcelWS.Paste ExcelWS.Range("A1")
If Form1.Check1.Value = 1 Then
  WordObj.Visible = True

Else
   WordObj.Visible = False
    On Error Resume Next
    Kill App.Path & "\" & aaa & ".xls"
    WordObj.ActiveSheet.SaveAs App.Path & "\" & aaa & ".xls"

ExcelWBk.Close False
  End If

End Sub
Sub ConvertPower()

  Set PptApp = New PowerPoint.Application
 PptApp.Activate
PptApp.WindowState = ppWindowMinimized
Set Present = PptApp.Presentations.Add
 Set Slds = Present.Slides
  Slds.Add Slds.Count + 1, ppLayoutBlank
PptApp.Windows(1).View.Paste
   If Form1.Check1.Value = 1 Then
   PptApp.WindowState = ppWindowNormal
   Else
   PptApp.Presentations(1).SaveAs App.Path & "\" & aaa & ".ppt"
   PptApp.Presentations(1).Close
PptApp.Quit
End If
End Sub

