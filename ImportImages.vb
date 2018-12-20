Sub Insert()
  Dim strFolder As String
  Dim strFileName As String
  Dim objPic As Picture
  Dim rngCell As Range
  Dim H As Integer
  Dim Top As Integer
  strFolder = "C:\Users\msc\Pictures\Screenshots\遺伝子試験項目\" 'change the path accordingly
  If Right(strFolder, 1) <> "\" Then
      strFolder = strFolder & "\"
  End If
  Set rngCell = Range("B1") 'starting cell
  strFileName = Dir(strFolder & "*.png", vbNormal) 'filter for .png files
  H = 500
  Top = 50
  Do While Len(strFileName) > 0
      Set objPic = ActiveSheet.Pictures.Insert(strFolder & strFileName)
      With objPic
          .Left = rngCell.Left
          .Top = Top
          .Height = H
          .Placement = xlMoveAndSize
          Top = Top + H + 50
      End With
      Set rngCell = rngCell.Offset(1, 0)
      strFileName = Dir
  Loop
End Sub
