Attribute VB_Name = "模块11"

Sub 设置样式()

    '删除所有自定义样式
    Dim i As Style
    For Each i In ThisDocument.Styles
        On Error Resume Next
        i.Delete
        Err.Clear
    Next i

      '论文正文样式无缩进
    Application.ScreenUpdating = False '关闭屏幕更新
    ActiveDocument.Styles.Add Name:="论文正文无缩进", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("论文正文无缩进").AutomaticallyUpdate = False
    With ActiveDocument.Styles("论文正文无缩进").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles("论文正文无缩进").ParagraphFormat
        .Alignment = wdAlignParagraphJustify '两端对齐
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '段前为0行
        .LineUnitAfter = 0 '段后为0行
        .LineSpacingRule = wdLineSpace1pt5 '1.5倍行距
    End With

    '论文正文样式有缩进
    Application.ScreenUpdating = False '关闭屏幕更新
    ActiveDocument.Styles.Add Name:="论文正文有缩进", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("论文正文有缩进").AutomaticallyUpdate = False
    With ActiveDocument.Styles("论文正文有缩进").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles("论文正文有缩进").ParagraphFormat
        .Alignment = wdAlignParagraphJustify '两端对齐
        .CharacterUnitFirstLineIndent = 2
        .LineUnitBefore = 0  '段前为0行
        .LineUnitAfter = 0 '段后为0行
        .LineSpacingRule = wdLineSpace1pt5 '1.5倍行距
    End With


    '标题1样式
    With ActiveDocument.Styles(wdStyleHeading1).Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 12  '字号小四，请输入对应数字
        .Bold = 1 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles(wdStyleHeading1).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '两端对齐
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0.5  '段前为0.5行
        .LineUnitAfter = 0.5 '段后为0.5行
        .LineSpacingRule = wdLineSpace1pt5 '1.5倍行距
    End With

    '标题2样式
    With ActiveDocument.Styles(wdStyleHeading2).Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '字号五号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles(wdStyleHeading2).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '两端对齐
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '段前为0行
        .LineUnitAfter = 0 '段后为0行
        .LineSpacingRule = wdLineSpaceSingle
    End With

    '标题3样式
    With ActiveDocument.Styles(wdStyleHeading3).Font
        .NameFarEast = "黑体"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles(wdStyleHeading3).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '两端对齐
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '段前为0.8行
        .LineUnitAfter = 0 '段后为0.5行
        .LineSpacingRule = wdLineSpaceSingle
    End With

    '标题4样式
    With ActiveDocument.Styles(wdStyleHeading4).Font
        .NameFarEast = "黑体"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles(wdStyleHeading4).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '两端对齐
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '段前为0.8行
        .LineUnitAfter = 0 '段后为0.5行
        .LineSpacingRule = wdLineSpaceSingle
    End With



    '论文公式样式
    ActiveDocument.Styles.Add Name:="论文公式", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("论文公式").AutomaticallyUpdate = False
    With ActiveDocument.Styles("论文公式").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 12  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles("论文公式").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .Alignment = wdAlignParagraphCenter
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '行间距设定为固定值23
    End With
    ActiveDocument.Styles("论文公式").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(7.41), Alignment:=wdAlignTabCenter, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("论文公式").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(14.81), Alignment:=wdAlignTabRight, Leader:= _
        wdTabLeaderSpaces

    '表格标题样式
    ActiveDocument.Styles.Add Name:="论文表格标题", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("论文表格标题").AutomaticallyUpdate = False
    With ActiveDocument.Styles("论文表格标题").Font
        .NameFarEast = "黑体"
        .NameAscii = "Times New Roman"
        .Size = 12  '字号，请输入对应数字
        .Bold = 1 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles("论文表格标题").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .Alignment = wdAlignParagraphCenter
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '行间距设定为固定值23
    End With

    '表格内容样式
    ActiveDocument.Styles.Add Name:="论文表格内容", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("论文表格内容").AutomaticallyUpdate = False
    With ActiveDocument.Styles("论文表格内容").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles("论文表格内容").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
        .Alignment = wdAlignParagraphCenter '两端对齐
        .LineUnitBefore = 0  '段前为0.8行
        .LineUnitAfter = 0 '段后为0.5行
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1
    End With

    '图表注释样式
    ActiveDocument.Styles.Add Name:="论文图表注释", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("论文图表注释").AutomaticallyUpdate = False
    With ActiveDocument.Styles("论文图表注释").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles("论文图表注释").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .Alignment = wdAlignParagraphJustify '两端对齐
        .CharacterUnitFirstLineIndent = 2
        .LineUnitBefore = 0  '段前为0.8行
        .LineUnitAfter = 0 '段后为0.5行
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '行间距设定为固定值23
    End With

    '图片标题样式
    ActiveDocument.Styles.Add Name:="论文图片标题", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("论文图片标题").AutomaticallyUpdate = False
    With ActiveDocument.Styles("论文图片标题").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 12  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0
    End With
    With ActiveDocument.Styles("论文图片标题").ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '行间距设定为固定值23
    End With
    Application.ScreenUpdating = True
    'MsgBox "设置样式成功"
End Sub

Sub 页边距单位厘米()
    Dim i, j, m, n
    '修改部分
    i = 2.5     '上边距，单位厘米
    j = 2       '下边距，单位厘米
    m = 3       '下边距，单位厘米
    n = 3       '左边距，单位厘米

    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait

        .TopMargin = CentimetersToPoints(i)
        .BottomMargin = CentimetersToPoints(j)
        .LeftMargin = CentimetersToPoints(m)
        .RightMargin = CentimetersToPoints(n)
    End With
    MsgBox "设置页边距成功"
End Sub

Sub 删除空行()
    Dim myRange As Range
    '选择区域为插入点
    If Selection.Type = wdSelectionIP Then
        MsgBox "未选定区域！"
    Else
        Set myRange = Selection.Range
        myRange.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll
        myRange.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll
        myRange.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll
        MsgBox "空行删除完毕！"
    End If
End Sub

Sub 删除空格()
    Dim myRange As Range
    '选择区域为插入点
    If Selection.Type = wdSelectionIP Then
        MsgBox "未选定区域！"
    Else
        Set myRange = Selection.Range
        myRange.Find.Execute FindText:=" ", ReplaceWith:="", Replace:=wdReplaceAll
        MsgBox "空格删除完毕！"
    End If
End Sub

Sub 清除所有格式()

    Selection.ClearFormatting

    With Selection.Font
        '字体设置 (字体  字号  加粗)
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 12  '字号，请输入对应数字
        .Bold = 0 '加粗为1，不加粗为0

    End With
    '取消缩进
    With Selection.ParagraphFormat
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
        .Alignment = wdAlignParagraphJustify
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '行间距设定为固定值23

    End With
    MsgBox "清除格式完成"
End Sub

Sub 各级标题正文自动化()
    '表外文字设置为正文
    Application.ScreenUpdating = False '关闭屏幕更新
    Dim j&, k&
    With ActiveDocument
        If .Tables.Count = 0 Then
            .Select
        Else
            If Not .Paragraphs(1).Range.Information(12) Then .Range(Start:=0, End:=.Tables(1).Range.Start).Editors.Add -1
            k = .Tables.Count
            For j = 1 To k
                If j = k Then Exit For
                .Range(Start:=.Tables(j).Range.End, End:=.Tables(j + 1).Range.Start).Editors.Add -1
            Next j
            .Range(Start:=.Tables(k).Range.End, End:=.Content.End).Editors.Add -1
            .SelectAllEditableRanges -1
            .DeleteAllEditableRanges -1
        End If
    End With
'    Selction.Style = ActiveDocument.Styles(wdStyleNormal)
    Selection.Style = "论文正文"

    '图片设置为单倍行距
    Dim image As InlineShape
    For Each image In ActiveDocument.InlineShapes
        'image.Height = 100 '图片高度属性可以自己调整
'        image.Width = 400 '图片宽度属性可以自己调整
        image.Range.Select
        Selection.ClearFormatting
        Selection.Range.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    Next

    '一二三四级标题自动识别并设置
    Dim para As Paragraph
    Application.ScreenUpdating = False
    For Each para In ActiveDocument.Paragraphs
        If para.Range Like "#.#.#.#*" = True Then
            para.Style = wdStyleHeading4
        ElseIf para.Range Like "#.#.#*" = True Then
            para.Style = wdStyleHeading3
        ElseIf para.Range Like "#.#*" = True Then
            para.Style = wdStyleHeading2
        ElseIf para.Range Like "# *" = True Then
            para.Style = wdStyleHeading1
'        Else
'            para.Style = wdStyleNormal
        End If
    Next
    Application.ScreenUpdating = True
    MsgBox "各级标题和正文自动设置完成"
End Sub

Sub 插入公式编号()
    With CaptionLabels("公式")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorPeriod
    End With
    Selection.InsertCaption Label:="公式", TitleAutoText:="InsertCaption1", _
        Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=1
End Sub

Sub 新图编号()  '修改系统插入“题注”命令

   '功能：自动删除标签与编号间的空格（英文除外），并在题注数字后添加一个空格；适用于：Word 2003 - 2013，不兼容WPS文字！
  '真正从原理上协同系统插入题注，无任何前提条件；用户照常插入题注即可，甚至感觉不到程序的存在！
   'Endlesswx于2015年8月4日
   
  '另,如果插入的始终未域代码而不是数字，非程序问题，Alt+F9一次即可
   
   Dim Lab As String, startPt As Long, endPt As Long, myrang As Range
   'On Error Resume Next  '发生错误时让程序继续执行下一句代码
'    Application.ScreenUpdating = False     '关闭屏幕更新，2013在此处关闭更新会导致输入框灰色不可选，故修正在调出对话框之后
   
   startPt = Selection.Start  'startPt标注起始点
      
   '***将if条件隐藏隐藏即可实现----手动替换题注空格***
   If Application.Dialogs(357).Show = -1 Then '插入“题注”对话框秀出来,如果按确定结束时执行以下程序，避免按取消后的空格,357也可换成wdDialogInsertCaption
      
      Application.ScreenUpdating = False     '关闭屏幕更新
      
       Lab = Dialogs(357).Label
       endPt = Selection.Start  'endPt标记插入的题注部分终点
      Selection.Start = startPt  '选定插入的整个题注
      
      '删除标签与编号间的空格（英文后的保留）
       With Selection.Find
          .Text = Lab & " "
          .Forward = True   'False=向上查找,(True=向下查找)
          .MatchWildcards = False '不使用通配符
          If Lab Like "*[0-9a-zA-Z.]" Then  '此处判断标签的最后一个字符是否为英文或数字，是则不删除空格
          Else
             .Replacement.Text = Lab
             .Execute Replace:=wdReplaceOne  '替换找到的第一个，此处用作删除空格
             endPt = endPt - 1 '删除空格后，末位减1
             Selection.End = endPt
          End If
       End With
      
      '在题注数字后添加一个空格
      Selection.Fields.ToggleShowCodes  '切换域代码，这样才能用^d查找域
       With Selection.Find
          .Text = "^d"
          .Replacement.Text = "^& "
          .Forward = False   'False=向上查找,(True=向下查找)
          .MatchWildcards = False '不使用通配符
          .Execute Replace:=wdReplaceOne  '替换找到的第一个，此处用作添加空格
       End With
      
      '选定整个插入的题注内容，将域代码切换回来
       endPt = endPt + 1 '增加空格后，末位加1
       With Selection
          .Start = startPt
          .End = endPt
          .Fields.ToggleShowCodes   '切换域代码（切换回来）
       End With
      
      '将光标定位至题注所在段尾处
'       Selection.MoveRight Unit:=wdCharacter, Count:=1  '此句光标返回插入题注前的原始位置，对于已经输好标题的情况并不合适
      '选择段尾回车符
       With Selection.Find
          .Text = "^13"
          .Forward = True   'False=向上查找,(True=向下查找)
          .MatchWildcards = False  '不使用通配符
          .Wrap = wdFindContinue '继续查找
          .Execute
       End With
      Selection.MoveLeft Unit:=wdCharacter, Count:=1  '定位到段尾回车前

   End If
   Application.ScreenUpdating = True          '恢复屏幕更新
   
End Sub

Sub 插入表编号()
    CaptionLabels.Add Name:="表"
    With CaptionLabels("表")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorHyphen
    End With
    Selection.InsertCaption Label:="表", TitleAutoText:="InsertCaption2", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
End Sub

Sub 表格自动三线表()
    Application.ScreenUpdating = False '关闭屏幕更新
    Dim t As Table
    For Each t In ActiveDocument.Tables
        With t
            .Range.Style = "论文表格内容"

            '去除所有边框
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone

            '设置上下边框
            Options.DefaultBorderLineWidth = wdLineWidth150pt
            .Borders(wdBorderTop).LineStyle = Options.DefaultBorderLineStyle
            .Borders(wdBorderTop).LineWidth = Options.DefaultBorderLineWidth
            .Borders(wdBorderTop).Color = Options.DefaultBorderColor

            Options.DefaultBorderLineWidth = wdLineWidth150pt
            .Borders(wdBorderBottom).LineStyle = Options.DefaultBorderLineStyle
            .Borders(wdBorderBottom).LineWidth = Options.DefaultBorderLineWidth
            .Borders(wdBorderBottom).Color = Options.DefaultBorderColor

            '设置中间边框
            Options.DefaultBorderLineWidth = wdLineWidth050pt
            .Cell(1, 1).Select
            With Selection
                .SelectRow
                .Borders(wdBorderBottom).LineStyle = Options.DefaultBorderLineStyle
                .Borders(wdBorderBottom).LineWidth = Options.DefaultBorderLineWidth
                .Borders(wdBorderBottom).Color = Options.DefaultBorderColor
            End With
        End With
    Next

    '自动匹配表格标题并设置格式
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "表 ^#.^#"
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40

        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "论文表格标题"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i

    '自动匹配图表注释并设置格式
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "注："
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40

        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "论文图表注释"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i
    Application.ScreenUpdating = True '关闭屏幕更新
    MsgBox "表格自动设置完成"
End Sub

Sub 自动设置所有图片()
    Application.ScreenUpdating = False '关闭屏幕更新
    Dim image As InlineShape
    For Each image In ActiveDocument.InlineShapes
        image.Height = 100 '图片高度属性可以自己调整
        image.Width = 400 '图片宽度属性可以自己调整
        image.Range.Select
        Selection.ClearFormatting
        Selection.Range.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    Next

    '自动匹配图片标题并设置格式
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "图 ^#.^#"
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40
        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "论文图片标题"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i

    '自动匹配图表注释并设置格式
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "注："
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40

        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "论文图表注释"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i
    Application.ScreenUpdating = True '关闭屏幕更新
    MsgBox "所有图片自动设置完成"
End Sub
Sub 插入分节符()
    Selection.InsertBreak Type:=wdSectionBreakNextPage
End Sub

Sub 自动生成目录()

    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
            LowerHeadingLevel:=3, IncludePageNumbers:=True, AddedStyles:="", _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        .TablesOfContents.Format = wdIndexIndent
    End With
End Sub

Sub 多级列表()

    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1 "
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 1"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 2"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 3"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%4."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
        .NumberFormat = "%5)"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6)
        .NumberFormat = "%6."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseRoman
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignRight
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 5
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(7)
        .NumberFormat = "%7."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 6
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(8)
        .NumberFormat = "%8)"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 7
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(9)
        .NumberFormat = "%9."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseRoman
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignRight
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 8
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
    MsgBox "多级列表自动设置完成"
End Sub

Sub 全文图片宽度() '设置图片尺寸.
Message = "设置图片宽度，单位厘米"
Title = "全文调整，慎用!!!!!!Xuxu"
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
Dim n '图片个数
On Error Resume Next '忽略错误
For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes 类型 图片
ActiveDocument.InlineShapes(n).Width = mmm '设置图片宽度 10cm，其中，Word中1cm=28.35px
Next n

End Sub

Sub 全文图片高度() '设置图片尺寸.
Message = "设置图片高度，单位厘米"
Title = "全文调整，慎用!!!!!!Xuxu"
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
Dim n '图片个数
On Error Resume Next '忽略错误
For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes 类型 图片
ActiveDocument.InlineShapes(n).Height = mmm '设置图片宽度 10cm，其中，Word中1cm=28.35px
Next n

End Sub

Sub SetPicWidth() '设置图片大小

Title = "图片大小,选中哪个调哪个,安全，XUXU"
Message = "设置图片宽度，单位厘米"
a = Selection.ShapeRange.Count '获取选中的图片数
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
For n = 1 To Selection.InlineShapes.Count 'InlineShapes 类型 图片
ActiveDocument.InlineShapes(n).Width = mmm '设置图片宽度 10cm，其中，Word中1cm=28.35px
Next n

End Sub

Sub SetPicHeight() '设置图片大小

Title = "图片大小,选中哪个调哪个,安全，XUXU"
Message = "设置图片宽度，单位厘米"
a = Selection.ShapeRange.Count '获取选中的图片数
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
For n = 1 To Selection.InlineShapes.Count 'InlineShapes 类型 图片
ActiveDocument.InlineShapes(n).Height = mmm '设置图片宽度 10cm，其中，Word中1cm=28.35px
Next n

End Sub

Sub TablesThree() ' 三线表格式设置
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
 
    
    Options.DefaultBorderLineWidth = wdLineWidth100pt
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Options.DefaultBorderLineWidth = wdLineWidth100pt
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
    Options.DefaultBorderLineWidth = wdLineWidth025pt
    With Selection.Rows(1).Borders(wdBorderBottom) '第一行的底边框
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
End Sub



Sub 选中所有表格()
    Dim t As Table
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    For Each t In ActiveDocument.Tables
        t.Range.Editors.Add wdEditorEveryone
    Next
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
End Sub


Sub 表格处理()
    On Error Resume Next
    Dim i As Long
    i = ActiveDocument.Tables.Count
    If i = 0 Then MsgBox "当前文档无表格！", vbOKOnly + vbCritical, "表格处理": Exit Sub
    Dim a As Long, b As Long, c As String, h As String, s As String, t As Table, n As Long
    c = MsgBox("是：自动    否：自定义    取消：放弃", vbYesNoCancel + vbExclamation, "表格处理")
    If c = vbYes Then
        h = 0.9
        s = 12
        a = 1
        b = 1
    ElseIf c = vbNo Then
        h = InputBox("请输入表格行高值：(0.7-1.2 厘米比较美观)", "表格处理", "0.9")
        If h = "" Then Exit Sub
        s = InputBox("请输入表格内文字字号：(比正文小半号比较美观)" & vbCr & "三号/16磅，小三/15磅，四号/14磅，小四/12磅，五号/10.5磅", "表格处理", "12")
        If s = "" Then Exit Sub
        If MsgBox("根据内容调整表格吗？", vbYesNo + vbExclamation, "自动调整") = vbYes Then a = 1
        If MsgBox("所有表格表头加粗吗？", vbYesNo + vbExclamation, "表头加粗") = vbYes Then b = 1
    Else
        Exit Sub
    End If
    If Selection.Information(wdWithInTable) = True Then Selection.Tables(1).Select: n = 1
    For Each t In ActiveDocument.Tables
        If n = 1 Then Set t = Selection.Tables(1) Else t.Select
' 表格标准化
        With t
            With .Rows
                .WrapAroundText = False
                .Alignment = wdAlignRowLeft
                .HeightRule = wdRowHeightAtLeast
                .Height = CentimetersToPoints(h)
            End With
            .AutoFitBehavior (wdAutoFitWindow)
            .AutoFitBehavior (wdAutoFitWindow)
            With .Range
                With .Cells
                    .DistributeWidth
                    .VerticalAlignment = wdCellAlignVerticalCenter
                End With
                .Font.Size = s
                With .ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                    .CharacterUnitFirstLineIndent = 0
                    .FirstLineIndent = CentimetersToPoints(0)
                    .Space1
                End With
            End With
            .Shading.BackgroundPatternColor = wdColorAutomatic
' 根据内容调整表格
            If a = 1 Then
                .AutoFitBehavior (wdAutoFitContent)
                .AutoFitBehavior (wdAutoFitContent)
            End If
            .Select
            .AutoFitBehavior (wdAutoFitWindow)
            .AutoFitBehavior (wdAutoFitWindow)
' 表头加粗
            If b = 1 Then
                With .Rows(1).Range.Font
                    .Name = "黑体"
                    .Name = "Times New Roman"
                    .Bold = True
                End With
            End If
        End With
    Next
    If n <> 1 Then Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
