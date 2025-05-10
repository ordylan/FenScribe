Sub FenScribeMacro()
   Selection.PageSetup.Orientation = wdOrientLandscape

    Selection.WholeStory
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = True
        .Width = InchesToPoints(5.35)
        .Spacing = InchesToPoints(0)
    End With
    Dim colWidth As Single
    Dim img As InlineShape
    Dim ratio As Single
    
    ' 获取双栏文档中的每一列的宽度
    colWidth = Selection.PageSetup.TextColumns(1).Width
    
    ' 遍历文档中的每一个内联形状（即图片）
    For Each img In ActiveDocument.InlineShapes
        ' 计算图片的宽高比
        ratio = img.Height / img.Width
        ' 检查图片宽度是否大于列宽
        If img.Width > colWidth Then
            ' 如果是，则将图片宽度设置为列宽
            img.Width = colWidth
            img.Height = colWidth * ratio
        End If
    Next img

End Sub