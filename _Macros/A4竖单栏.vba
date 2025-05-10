Sub FenScribeMacro()
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