Option Explicit

On Error Resume Next

ExcelMacroExample

Sub ExcelMacroExample() 

  Dim xlApp 
  Dim xlBook 

  Set xlApp = CreateObject("Excel.Application") 
  Set xlBook = xlApp.Workbooks.Open(WScript.Arguments.Item(0), 0, True) 
  xlApp.Run "ImportXMLtoList",WScript.Arguments.Item(2),WScript.Arguments.Item(1)
  xlApp.Quit 

  Set xlBook = Nothing 
  Set xlApp = Nothing 

End Sub 