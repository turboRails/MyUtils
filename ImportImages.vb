Sub Insert()
  Dim strFolder As String
  Dim strFileName As String
  Dim shpPicture As Shape
  Dim strFilePath As String
    
  Dim rngCell As Range
  Dim H As Integer
  Dim Top As Integer
  strFolder = "C:\Users\" & Environ$("Username") & "\Pictures\Screenshots"
  
  If Right(strFolder, 1) <> "\" Then
      strFolder = strFolder & "\"
  End If

  strFileName = Dir(strFolder & "*.png", vbNormal)
  ExtraSpace = 100
  Top = 50
  Do While Len(strFileName) > 0
  
    strFilePath = strFolder & strFileName
    Set shpPicture = ActiveSheet.Shapes.AddPicture( _
    Filename:=strFolder & strFileName, _
    LinkToFile:=False, _
    SaveWithDocument:=True, _
    Left:=Selection.Left, _
    Top:=Top, _
    Width:=-1, _
    Height:=-1)
    
    Top = Top + ExtraSpace + shpPicture.Height
    strFileName = Dir
  Loop
End Sub
