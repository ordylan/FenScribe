Attribute VB_Name = "ORDYLAN"
Sub 双栏图片溢出自动调()
    Dim colWidth As Single
    Dim img As InlineShape
    Dim ratio As Single
    colWidth = Selection.PageSetup.TextColumns(1).Width
    For Each img In ActiveDocument.InlineShapes
        ratio = img.Height / img.Width
        If img.Width > colWidth Then
            img.Width = colWidth
            img.Height = colWidth * ratio
        End If
    Next img
End Sub
