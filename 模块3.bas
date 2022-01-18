Attribute VB_Name = "模块3"
Sub 设置全文图片大小() '设置图片尺寸

Title = "设置全文图片大小，最好新开一个文档再处理"

Message = "设置图片宽度，单位厘米"

MyValue = InputBox(Message, Title, Default)

MyValue = MyValue * 28.35

Dim n '图片个数

On Error Resume Next '忽略错误

For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes 类型 图片

ActiveDocument.InlineShapes(n).Width = MyValue '设置图片宽度 10cm，其中，Word中1cm=28.35px

Next n
    MsgBox "尺寸设置完成"
End Sub
