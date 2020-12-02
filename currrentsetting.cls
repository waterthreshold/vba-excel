VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Hello()
'Range("A1:A4").Value = 5
'MsgBox ("Hello, world!")
 Set myXML = CreateObject("MSXML2.ServerXMLHTTP")

    myXML.Open "GET", "http://192.168.1.1/currentsetting.htm", False

   '通常不需要RequestHeader，若要使用需自行修改'
'    myXML.setRequestHeader "Content-Type", "text/html;charset=big5"
    myXML.send ("")
    Dim resp As String
    Dim arr As Variant
    Dim arr2 As Variant
     Dim delimeter As String
    resp = myXML.responseText
    MsgBox (resp)
    delimeter = "" & Chr(13) & Chr(10)
    arr = Split(resp, delimeter)
    For i = LBound(arr, 1) To UBound(arr, 1)
    'If InStr(1, arrSplitStrings2(i), "Can", vbTextCompare) > 0 Then
     '   strSingleString3 = arrSplitStrings2(i)
      '  Exit For
   ' End If
    If Not InStr(arr(i), "=") > 0 Then
       GoTo ContinueForLoop
    End If
    arr2 = Split(arr(i), "=")
    'MsgBox (i)
    
    'Cells(i + 1, 1).Font.Color = RGB(255, 0, 0)
    Cells(i + 1, 1).Value = arr2(0)
    'Cells(i + 1, 2).HorizontalAlignment = xlRight
    Cells(i + 1, 2).Value = arr2(1)
ContinueForLoop:
    Next i
End Sub
