Attribute VB_Name = "Module11"
Public Const FontSize As Integer = 14, FirstLine As Integer = 1.25, Tabs As Integer = 3


Sub SingleLineSpacing(control As IRibbonControl)
'
' ОднострочныйИнтервал Макрос
'
'
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
End Sub
Sub OneHalfLineSpacing(control As IRibbonControl)
'
' ПолуторныйИнтервал Макрос
'
'
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
End Sub
Sub TimesNewRoman(control As IRibbonControl)
'
' TimesNewRomanFont Макрос
'
'
    Selection.Font.Name = "Times New Roman"
End Sub
Sub FontSize8(control As IRibbonControl)
'
' Макрос1 Макрос
'
'
    Selection.Font.Size = 8
End Sub
Sub FontSize10(control As IRibbonControl)
'
' Макрос1 Макрос
'
'
    Selection.Font.Size = 10
End Sub
Sub FontSize12(control As IRibbonControl)
'
' Макрос1 Макрос
'
'
    Selection.Font.Size = 12
End Sub
Sub FontSize14(control As IRibbonControl)
    Selection.Font.Size = 14
End Sub
Sub UsualStyle(control As IRibbonControl)
    Call Usual
End Sub
Sub Title1Style(control As IRibbonControl)
    Call Title1
End Sub
Sub Title2Style(control As IRibbonControl)
    Call Title2
End Sub
Sub Title3Style(control As IRibbonControl)
    Call Title3
End Sub
Sub GOST_A_Style(control As IRibbonControl)
    Call GOST_A
End Sub
Sub PagesALL(control As IRibbonControl)
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "NUMPAGES  \* Arabic ", PreserveFormatting:=True
End Sub
Sub PageNum(control As IRibbonControl)
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "PAGE  \* Arabic ", PreserveFormatting:=True
End Sub


Function numListFormat(f)
     With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(f)
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(FirstLine)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = CentimetersToPoints(Tabs)
        .StartAt = 1
        With .Font
            .Bold = False
            .Italic = False
            .StrikeThrough = False
            .Subscript = False
            .Superscript = False
            .Shadow = False
            .Outline = False
            .Emboss = False
            .Engrave = False
            .AllCaps = False
            .Hidden = False
            .Underline = False
            .Color = wdColorBlack
            .Size = FontSize
            .Animation = False
            .DoubleStrikeThrough = False
            .Name = "Times New Roman"
        End With
    End With
End Function
Function TitleStyle()
    With Selection.Font
        .Name = "Times New Roman"
        .Size = FontSize
        If Selection.Style.NameLocal Like "Заголовок 1*" Then
            .Bold = True
        Else
            .Bold = False
        End If
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.5)
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(Tabs)
End Function
Sub NumList()
    MsgBox "numlist"
    a = 1
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1"
        .ResetOnHigher = 0
        .LinkedStyle = "Заголовок 1"
        numListFormat (a)
    End With
    a = 2
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2"
        .ResetOnHigher = 1
        .LinkedStyle = "Заголовок 2"
        numListFormat (a)
    End With
    a = 3
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .ResetOnHigher = 2
        numListFormat (a)
        .LinkedStyle = "Заголовок 3"
    End With
    a = 4
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .ResetOnHigher = 3
        numListFormat (a)
        .LinkedStyle = "Заголовок 4"
    End With
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = "My style"
    Selection.WholeStory
    'For i = 1 To 4
        Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
            ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
            ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, _
            DefaultListBehavior:=wdWord10ListBehavior ', ApplyLevel:=i
        'wdListApplyToSelection wdListApplyToThisPointForward
    'Next i
        
End Sub

Private Sub Title1()
    Selection.Style = ActiveDocument.Styles("Заголовок 1")
    TitleStyle
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceAfter = 6
        .FirstLineIndent = CentimetersToPoints(FirstLine)
    End With
End Sub
Private Sub Title2()
    Selection.Style = ActiveDocument.Styles("Заголовок 2")
    TitleStyle
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .SpaceBefore = 3
        .SpaceAfter = 3
        .FirstLineIndent = CentimetersToPoints(FirstLine)
    End With
End Sub
Private Sub Title3()
    Selection.Style = ActiveDocument.Styles("Заголовок 3")
    TitleStyle
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .SpaceBefore = 3
        .SpaceAfter = 3
        .FirstLineIndent = CentimetersToPoints(FirstLine)
    End With
End Sub
Private Sub Usual()
    Selection.Style = ActiveDocument.Styles("Обычный")
    TitleStyle
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = CentimetersToPoints(FirstLine)
    End With
End Sub
Private Sub GOST_A()
    Selection.Style = ActiveDocument.Styles("Обычный")
    With Selection.Font
        .Name = "GOST type A"
        .Size = FontSize
        .Bold = False
        .Italic = True
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1#)
    Selection.ParagraphFormat.TabStops.ClearAll
    'Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(0)
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
    End With
End Sub
Private Sub Test2()
 
    With ActiveDocument.Content.Find
         .ClearFormatting: .Style = "Заголовок 3"
         With .Replacement
              .ClearFormatting: .Style = "О"
         End With
         .Execute FindText:="", ReplaceWith:="", Format:=True, Replace:=wdReplaceAll
    End With
   
End Sub
Sub test()
 
For par = 1 To ActiveDocument.Paragraphs.Count
    Select Case True
    Case ActiveDocument.Paragraphs(par).Style.NameLocal Like "Заголовок 3*"
        ActiveDocument.Paragraphs(par).Range.Select
        Title3
    Case ActiveDocument.Paragraphs(par).Style.NameLocal Like "Заголовок 2*"
        ActiveDocument.Paragraphs(par).Range.Select
        Title2
    Case ActiveDocument.Paragraphs(par).Style.NameLocal Like "Заголовок 1*"
        ActiveDocument.Paragraphs(par).Range.Select
        Title1
    Case Else
        ActiveDocument.Paragraphs(par).Range.Select
        Usual
    End Select
Next
 
End Sub


