Attribute VB_Name = "ģ��11"

Sub ������ʽ()

    'ɾ�������Զ�����ʽ
    Dim i As Style
    For Each i In ThisDocument.Styles
        On Error Resume Next
        i.Delete
        Err.Clear
    Next i

      '����������ʽ������
    Application.ScreenUpdating = False '�ر���Ļ����
    ActiveDocument.Styles.Add Name:="��������������", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("��������������").AutomaticallyUpdate = False
    With ActiveDocument.Styles("��������������").Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles("��������������").ParagraphFormat
        .Alignment = wdAlignParagraphJustify '���˶���
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '��ǰΪ0��
        .LineUnitAfter = 0 '�κ�Ϊ0��
        .LineSpacingRule = wdLineSpace1pt5 '1.5���о�
    End With

    '����������ʽ������
    Application.ScreenUpdating = False '�ر���Ļ����
    ActiveDocument.Styles.Add Name:="��������������", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("��������������").AutomaticallyUpdate = False
    With ActiveDocument.Styles("��������������").Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles("��������������").ParagraphFormat
        .Alignment = wdAlignParagraphJustify '���˶���
        .CharacterUnitFirstLineIndent = 2
        .LineUnitBefore = 0  '��ǰΪ0��
        .LineUnitAfter = 0 '�κ�Ϊ0��
        .LineSpacingRule = wdLineSpace1pt5 '1.5���о�
    End With


    '����1��ʽ
    With ActiveDocument.Styles(wdStyleHeading1).Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 12  '�ֺ�С�ģ��������Ӧ����
        .Bold = 1 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles(wdStyleHeading1).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '���˶���
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0.5  '��ǰΪ0.5��
        .LineUnitAfter = 0.5 '�κ�Ϊ0.5��
        .LineSpacingRule = wdLineSpace1pt5 '1.5���о�
    End With

    '����2��ʽ
    With ActiveDocument.Styles(wdStyleHeading2).Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '�ֺ���ţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles(wdStyleHeading2).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '���˶���
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '��ǰΪ0��
        .LineUnitAfter = 0 '�κ�Ϊ0��
        .LineSpacingRule = wdLineSpaceSingle
    End With

    '����3��ʽ
    With ActiveDocument.Styles(wdStyleHeading3).Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles(wdStyleHeading3).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '���˶���
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '��ǰΪ0.8��
        .LineUnitAfter = 0 '�κ�Ϊ0.5��
        .LineSpacingRule = wdLineSpaceSingle
    End With

    '����4��ʽ
    With ActiveDocument.Styles(wdStyleHeading4).Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles(wdStyleHeading4).ParagraphFormat
        .Alignment = wdAlignParagraphJustify '���˶���
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0  '��ǰΪ0.8��
        .LineUnitAfter = 0 '�κ�Ϊ0.5��
        .LineSpacingRule = wdLineSpaceSingle
    End With



    '���Ĺ�ʽ��ʽ
    ActiveDocument.Styles.Add Name:="���Ĺ�ʽ", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("���Ĺ�ʽ").AutomaticallyUpdate = False
    With ActiveDocument.Styles("���Ĺ�ʽ").Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 12  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles("���Ĺ�ʽ").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .Alignment = wdAlignParagraphCenter
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '�м���趨Ϊ�̶�ֵ23
    End With
    ActiveDocument.Styles("���Ĺ�ʽ").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(7.41), Alignment:=wdAlignTabCenter, Leader:= _
        wdTabLeaderSpaces
    ActiveDocument.Styles("���Ĺ�ʽ").ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(14.81), Alignment:=wdAlignTabRight, Leader:= _
        wdTabLeaderSpaces

    '��������ʽ
    ActiveDocument.Styles.Add Name:="���ı�����", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("���ı�����").AutomaticallyUpdate = False
    With ActiveDocument.Styles("���ı�����").Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 12  '�ֺţ��������Ӧ����
        .Bold = 1 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles("���ı�����").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .Alignment = wdAlignParagraphCenter
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '�м���趨Ϊ�̶�ֵ23
    End With

    '���������ʽ
    ActiveDocument.Styles.Add Name:="���ı������", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("���ı������").AutomaticallyUpdate = False
    With ActiveDocument.Styles("���ı������").Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles("���ı������").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
        .Alignment = wdAlignParagraphCenter '���˶���
        .LineUnitBefore = 0  '��ǰΪ0.8��
        .LineUnitAfter = 0 '�κ�Ϊ0.5��
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1
    End With

    'ͼ��ע����ʽ
    ActiveDocument.Styles.Add Name:="����ͼ��ע��", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("����ͼ��ע��").AutomaticallyUpdate = False
    With ActiveDocument.Styles("����ͼ��ע��").Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles("����ͼ��ע��").ParagraphFormat
        .OutlineLevel = wdOutlineLevelBodyText
        .Alignment = wdAlignParagraphJustify '���˶���
        .CharacterUnitFirstLineIndent = 2
        .LineUnitBefore = 0  '��ǰΪ0.8��
        .LineUnitAfter = 0 '�κ�Ϊ0.5��
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '�м���趨Ϊ�̶�ֵ23
    End With

    'ͼƬ������ʽ
    ActiveDocument.Styles.Add Name:="����ͼƬ����", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("����ͼƬ����").AutomaticallyUpdate = False
    With ActiveDocument.Styles("����ͼƬ����").Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 12  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0
    End With
    With ActiveDocument.Styles("����ͼƬ����").ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '�м���趨Ϊ�̶�ֵ23
    End With
    Application.ScreenUpdating = True
    'MsgBox "������ʽ�ɹ�"
End Sub

Sub ҳ�߾൥λ����()
    Dim i, j, m, n
    '�޸Ĳ���
    i = 2.5     '�ϱ߾࣬��λ����
    j = 2       '�±߾࣬��λ����
    m = 3       '�±߾࣬��λ����
    n = 3       '��߾࣬��λ����

    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait

        .TopMargin = CentimetersToPoints(i)
        .BottomMargin = CentimetersToPoints(j)
        .LeftMargin = CentimetersToPoints(m)
        .RightMargin = CentimetersToPoints(n)
    End With
    MsgBox "����ҳ�߾�ɹ�"
End Sub

Sub ɾ������()
    Dim myRange As Range
    'ѡ������Ϊ�����
    If Selection.Type = wdSelectionIP Then
        MsgBox "δѡ������"
    Else
        Set myRange = Selection.Range
        myRange.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll
        myRange.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll
        myRange.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll
        MsgBox "����ɾ����ϣ�"
    End If
End Sub

Sub ɾ���ո�()
    Dim myRange As Range
    'ѡ������Ϊ�����
    If Selection.Type = wdSelectionIP Then
        MsgBox "δѡ������"
    Else
        Set myRange = Selection.Range
        myRange.Find.Execute FindText:=" ", ReplaceWith:="", Replace:=wdReplaceAll
        MsgBox "�ո�ɾ����ϣ�"
    End If
End Sub

Sub ������и�ʽ()

    Selection.ClearFormatting

    With Selection.Font
        '�������� (����  �ֺ�  �Ӵ�)
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 12  '�ֺţ��������Ӧ����
        .Bold = 0 '�Ӵ�Ϊ1�����Ӵ�Ϊ0

    End With
    'ȡ������
    With Selection.ParagraphFormat
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
        .Alignment = wdAlignParagraphJustify
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 23 '�м���趨Ϊ�̶�ֵ23

    End With
    MsgBox "�����ʽ���"
End Sub

Sub �������������Զ���()
    '������������Ϊ����
    Application.ScreenUpdating = False '�ر���Ļ����
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
    Selection.Style = "��������"

    'ͼƬ����Ϊ�����о�
    Dim image As InlineShape
    For Each image In ActiveDocument.InlineShapes
        'image.Height = 100 'ͼƬ�߶����Կ����Լ�����
'        image.Width = 400 'ͼƬ������Կ����Լ�����
        image.Range.Select
        Selection.ClearFormatting
        Selection.Range.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    Next

    'һ�����ļ������Զ�ʶ������
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
    MsgBox "��������������Զ��������"
End Sub

Sub ���빫ʽ���()
    With CaptionLabels("��ʽ")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorPeriod
    End With
    Selection.InsertCaption Label:="��ʽ", TitleAutoText:="InsertCaption1", _
        Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=1
End Sub

Sub ��ͼ���()  '�޸�ϵͳ���롰��ע������

   '���ܣ��Զ�ɾ����ǩ���ż�Ŀո�Ӣ�ĳ��⣩��������ע���ֺ����һ���ո������ڣ�Word 2003 - 2013��������WPS���֣�
  '������ԭ����Эͬϵͳ������ע�����κ�ǰ���������û��ճ�������ע���ɣ������о���������Ĵ��ڣ�
   'Endlesswx��2015��8��4��
   
  '��,��������ʼ��δ�������������֣��ǳ������⣬Alt+F9һ�μ���
   
   Dim Lab As String, startPt As Long, endPt As Long, myrang As Range
   'On Error Resume Next  '��������ʱ�ó������ִ����һ�����
'    Application.ScreenUpdating = False     '�ر���Ļ���£�2013�ڴ˴��رո��»ᵼ��������ɫ����ѡ���������ڵ����Ի���֮��
   
   startPt = Selection.Start  'startPt��ע��ʼ��
      
   '***��if�����������ؼ���ʵ��----�ֶ��滻��ע�ո�***
   If Application.Dialogs(357).Show = -1 Then '���롰��ע���Ի��������,�����ȷ������ʱִ�����³��򣬱��ⰴȡ����Ŀո�,357Ҳ�ɻ���wdDialogInsertCaption
      
      Application.ScreenUpdating = False     '�ر���Ļ����
      
       Lab = Dialogs(357).Label
       endPt = Selection.Start  'endPt��ǲ������ע�����յ�
      Selection.Start = startPt  'ѡ�������������ע
      
      'ɾ����ǩ���ż�Ŀո�Ӣ�ĺ�ı�����
       With Selection.Find
          .Text = Lab & " "
          .Forward = True   'False=���ϲ���,(True=���²���)
          .MatchWildcards = False '��ʹ��ͨ���
          If Lab Like "*[0-9a-zA-Z.]" Then  '�˴��жϱ�ǩ�����һ���ַ��Ƿ�ΪӢ�Ļ����֣�����ɾ���ո�
          Else
             .Replacement.Text = Lab
             .Execute Replace:=wdReplaceOne  '�滻�ҵ��ĵ�һ�����˴�����ɾ���ո�
             endPt = endPt - 1 'ɾ���ո��ĩλ��1
             Selection.End = endPt
          End If
       End With
      
      '����ע���ֺ����һ���ո�
      Selection.Fields.ToggleShowCodes  '�л�����룬����������^d������
       With Selection.Find
          .Text = "^d"
          .Replacement.Text = "^& "
          .Forward = False   'False=���ϲ���,(True=���²���)
          .MatchWildcards = False '��ʹ��ͨ���
          .Execute Replace:=wdReplaceOne  '�滻�ҵ��ĵ�һ�����˴�������ӿո�
       End With
      
      'ѡ�������������ע���ݣ���������л�����
       endPt = endPt + 1 '���ӿո��ĩλ��1
       With Selection
          .Start = startPt
          .End = endPt
          .Fields.ToggleShowCodes   '�л�����루�л�������
       End With
      
      '����궨λ����ע���ڶ�β��
'       Selection.MoveRight Unit:=wdCharacter, Count:=1  '�˾��귵�ز�����עǰ��ԭʼλ�ã������Ѿ���ñ���������������
      'ѡ���β�س���
       With Selection.Find
          .Text = "^13"
          .Forward = True   'False=���ϲ���,(True=���²���)
          .MatchWildcards = False  '��ʹ��ͨ���
          .Wrap = wdFindContinue '��������
          .Execute
       End With
      Selection.MoveLeft Unit:=wdCharacter, Count:=1  '��λ����β�س�ǰ

   End If
   Application.ScreenUpdating = True          '�ָ���Ļ����
   
End Sub

Sub �������()
    CaptionLabels.Add Name:="��"
    With CaptionLabels("��")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorHyphen
    End With
    Selection.InsertCaption Label:="��", TitleAutoText:="InsertCaption2", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
End Sub

Sub ����Զ����߱�()
    Application.ScreenUpdating = False '�ر���Ļ����
    Dim t As Table
    For Each t In ActiveDocument.Tables
        With t
            .Range.Style = "���ı������"

            'ȥ�����б߿�
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone

            '�������±߿�
            Options.DefaultBorderLineWidth = wdLineWidth150pt
            .Borders(wdBorderTop).LineStyle = Options.DefaultBorderLineStyle
            .Borders(wdBorderTop).LineWidth = Options.DefaultBorderLineWidth
            .Borders(wdBorderTop).Color = Options.DefaultBorderColor

            Options.DefaultBorderLineWidth = wdLineWidth150pt
            .Borders(wdBorderBottom).LineStyle = Options.DefaultBorderLineStyle
            .Borders(wdBorderBottom).LineWidth = Options.DefaultBorderLineWidth
            .Borders(wdBorderBottom).Color = Options.DefaultBorderColor

            '�����м�߿�
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

    '�Զ�ƥ������Ⲣ���ø�ʽ
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "�� ^#.^#"
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40

        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "���ı�����"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i

    '�Զ�ƥ��ͼ��ע�Ͳ����ø�ʽ
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ע��"
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40

        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "����ͼ��ע��"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i
    Application.ScreenUpdating = True '�ر���Ļ����
    MsgBox "����Զ��������"
End Sub

Sub �Զ���������ͼƬ()
    Application.ScreenUpdating = False '�ر���Ļ����
    Dim image As InlineShape
    For Each image In ActiveDocument.InlineShapes
        image.Height = 100 'ͼƬ�߶����Կ����Լ�����
        image.Width = 400 'ͼƬ������Կ����Լ�����
        image.Range.Select
        Selection.ClearFormatting
        Selection.Range.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    Next

    '�Զ�ƥ��ͼƬ���Ⲣ���ø�ʽ
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ͼ ^#.^#"
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40
        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "����ͼƬ����"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i

    '�Զ�ƥ��ͼ��ע�Ͳ����ø�ʽ
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ע��"
        .Replacement.Text = ""
    End With
    Selection.Find.Execute

    For i = 1 To 40

        If Selection.Find.Found = True Then
            Selection.MoveUp Unit:=wdParagraph
            Selection.MoveDown Unit:=wdParagraph, Extend:=wdExtend
            Selection.Style = "����ͼ��ע��"
        End If
        Selection.Find.Execute
        Selection.Find.Execute
    Next i
    Application.ScreenUpdating = True '�ر���Ļ����
    MsgBox "����ͼƬ�Զ��������"
End Sub
Sub ����ֽڷ�()
    Selection.InsertBreak Type:=wdSectionBreakNextPage
End Sub

Sub �Զ�����Ŀ¼()

    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
            LowerHeadingLevel:=3, IncludePageNumbers:=True, AddedStyles:="", _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        .TablesOfContents.Format = wdIndexIndent
    End With
End Sub

Sub �༶�б�()

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
        .LinkedStyle = "���� 1"
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
        .LinkedStyle = "���� 2"
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
        .LinkedStyle = "���� 3"
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
    MsgBox "�༶�б��Զ��������"
End Sub

Sub ȫ��ͼƬ���() '����ͼƬ�ߴ�.
Message = "����ͼƬ��ȣ���λ����"
Title = "ȫ�ĵ���������!!!!!!Xuxu"
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
Dim n 'ͼƬ����
On Error Resume Next '���Դ���
For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes ���� ͼƬ
ActiveDocument.InlineShapes(n).Width = mmm '����ͼƬ��� 10cm�����У�Word��1cm=28.35px
Next n

End Sub

Sub ȫ��ͼƬ�߶�() '����ͼƬ�ߴ�.
Message = "����ͼƬ�߶ȣ���λ����"
Title = "ȫ�ĵ���������!!!!!!Xuxu"
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
Dim n 'ͼƬ����
On Error Resume Next '���Դ���
For n = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes ���� ͼƬ
ActiveDocument.InlineShapes(n).Height = mmm '����ͼƬ��� 10cm�����У�Word��1cm=28.35px
Next n

End Sub

Sub SetPicWidth() '����ͼƬ��С

Title = "ͼƬ��С,ѡ���ĸ����ĸ�,��ȫ��XUXU"
Message = "����ͼƬ��ȣ���λ����"
a = Selection.ShapeRange.Count '��ȡѡ�е�ͼƬ��
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
For n = 1 To Selection.InlineShapes.Count 'InlineShapes ���� ͼƬ
ActiveDocument.InlineShapes(n).Width = mmm '����ͼƬ��� 10cm�����У�Word��1cm=28.35px
Next n

End Sub

Sub SetPicHeight() '����ͼƬ��С

Title = "ͼƬ��С,ѡ���ĸ����ĸ�,��ȫ��XUXU"
Message = "����ͼƬ��ȣ���λ����"
a = Selection.ShapeRange.Count '��ȡѡ�е�ͼƬ��
mmm = InputBox(Message, Title, "9")
mmm = mmm * 28.35
For n = 1 To Selection.InlineShapes.Count 'InlineShapes ���� ͼƬ
ActiveDocument.InlineShapes(n).Height = mmm '����ͼƬ��� 10cm�����У�Word��1cm=28.35px
Next n

End Sub

Sub TablesThree() ' ���߱��ʽ����
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
    With Selection.Rows(1).Borders(wdBorderBottom) '��һ�еĵױ߿�
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
End Sub



Sub ѡ�����б��()
    Dim t As Table
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    For Each t In ActiveDocument.Tables
        t.Range.Editors.Add wdEditorEveryone
    Next
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
End Sub


Sub �����()
    On Error Resume Next
    Dim i As Long
    i = ActiveDocument.Tables.Count
    If i = 0 Then MsgBox "��ǰ�ĵ��ޱ��", vbOKOnly + vbCritical, "�����": Exit Sub
    Dim a As Long, b As Long, c As String, h As String, s As String, t As Table, n As Long
    c = MsgBox("�ǣ��Զ�    ���Զ���    ȡ��������", vbYesNoCancel + vbExclamation, "�����")
    If c = vbYes Then
        h = 0.9
        s = 12
        a = 1
        b = 1
    ElseIf c = vbNo Then
        h = InputBox("���������и�ֵ��(0.7-1.2 ���ױȽ�����)", "�����", "0.9")
        If h = "" Then Exit Sub
        s = InputBox("���������������ֺţ�(������С��űȽ�����)" & vbCr & "����/16����С��/15�����ĺ�/14����С��/12�������/10.5��", "�����", "12")
        If s = "" Then Exit Sub
        If MsgBox("�������ݵ��������", vbYesNo + vbExclamation, "�Զ�����") = vbYes Then a = 1
        If MsgBox("���б���ͷ�Ӵ���", vbYesNo + vbExclamation, "��ͷ�Ӵ�") = vbYes Then b = 1
    Else
        Exit Sub
    End If
    If Selection.Information(wdWithInTable) = True Then Selection.Tables(1).Select: n = 1
    For Each t In ActiveDocument.Tables
        If n = 1 Then Set t = Selection.Tables(1) Else t.Select
' ����׼��
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
' �������ݵ������
            If a = 1 Then
                .AutoFitBehavior (wdAutoFitContent)
                .AutoFitBehavior (wdAutoFitContent)
            End If
            .Select
            .AutoFitBehavior (wdAutoFitWindow)
            .AutoFitBehavior (wdAutoFitWindow)
' ��ͷ�Ӵ�
            If b = 1 Then
                With .Rows(1).Range.Font
                    .Name = "����"
                    .Name = "Times New Roman"
                    .Bold = True
                End With
            End If
        End With
    Next
    If n <> 1 Then Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
