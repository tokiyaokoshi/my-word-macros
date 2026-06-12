' ==========================================
' 1. 初期設定（ショートカットキー一括登録）
' ==========================================
Sub SetupWordShortcuts()
    ' Normalテンプレートにショートカットを保存
    CustomizationContext = NormalTemplate
    
    ' 1. Alt + : で空の数式を挿入（自作マクロ）
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeySemiColon, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="InsertEmptyEquation"
    
    ' 2. Alt + D で表の行を削除（Word標準コマンド）
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyD, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="TableDeleteRow"
    
    ' 3. Alt + 1 で表の上に行を挿入（Word標準コマンド）
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey1, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="TableInsertRowAbove"

    ' --- Ctrl + Shift 系 (自作マクロ) ---
    KeyBindings.Add KeyCode:=BuildKeyCode(vbKeyUp, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="TableAutoFitToContents"
    KeyBindings.Add KeyCode:=BuildKeyCode(vbKeyRight, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="TableAutoFitToWindow"
    KeyBindings.Add KeyCode:=BuildKeyCode(vbKeyDown, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="TableAutoFitFixed"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyA, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FillGray"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyD, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="DeleteBorders_Word"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FillAqua"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyG, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="SetNormalBorderAroundSelection_Word"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyI, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="AddBorders_Word"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FillLightGreen"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyR, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FillLightBlue"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyT, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FillGray_And_BottomDoubleLine_Word"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyW, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FillYellow"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyX, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FillLightRed"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyZ, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="ClearFill"
    
    ' --- Ctrl + Shift 系 (Word標準コマンド) ---
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyC, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryCommand, Command:="CopyFormat"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyV, wdKeyControl, wdKeyShift), KeyCategory:=wdKeyCategoryCommand, Command:="PasteFormat"
    
    ' --- Ctrl 系 (Word標準コマンド) ---
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeySemiColon, wdKeyControl), KeyCategory:=wdKeyCategoryCommand, Command:="InsertEquation"
    
    ' --- Alt 単体 系 (自作マクロ) ---
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyReturn, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="InsertSectionBreakNextPage"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyA, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="InsertEquationSuperscript"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="ColorBlue"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyG, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="ColorGreen"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyN, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="RemoveLineBreaks"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyP, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="ColorMagenta"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyU, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="InsertEquationSubscript"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyX, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="ColorRed"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyZ, wdKeyAlt), KeyCategory:=wdKeyCategoryMacro, Command:="ColorBlack"
    
    ' --- Alt 単体 系 (Word標準コマンド) ---
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey2, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="TableSplitCells"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey3, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="TableMergeCells"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey4, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="FormatBordersAndShading"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey5, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="FormatParagraph"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyC, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="TableDeleteColumn"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyT, wdKeyAlt), KeyCategory:=wdKeyCategoryCommand, Command:="PasteTextOnly"
    
    ' --- Alt + Ctrl 系 (Word標準コマンド) ---
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyV, wdKeyAlt, wdKeyControl), KeyCategory:=wdKeyCategoryCommand, Command:="EditPasteSpecial"
    
    ' --- Alt + Shift 系 (自作マクロ) ---
    KeyBindings.Add KeyCode:=BuildKeyCode(vbKeyLeft, wdKeyAlt, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FitImageToTableCell"
    KeyBindings.Add KeyCode:=BuildKeyCode(vbKeyRight, wdKeyAlt, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="FitImageToMargins"
    KeyBindings.Add KeyCode:=BuildKeyCode(vbKeyDown, wdKeyAlt, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="MaximizeImageSizeWithinPage"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyI, wdKeyAlt, wdKeyShift), KeyCategory:=wdKeyCategoryMacro, Command:="AddBorders_OuterBold_InnerThin_Word"
    
    MsgBox "Wordのショートカットキー一括設定が完了しました！", vbInformation
End Sub

' ==========================================
' 2. 文字色変更マクロ
' ==========================================
Sub ColorBlack()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = wdColorBlack
    Else
        MsgBox "文字が選択されていません。", vbExclamation
    End If
End Sub

Sub ColorBlue()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = wdColorBlue
    Else
        MsgBox "文字が選択されていません。", vbExclamation
    End If
End Sub

Sub ColorGreen()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = wdColorGreen
    Else
        MsgBox "文字が選択されていません。", vbExclamation
    End If
End Sub

Sub ColorRed()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = wdColorRed
    Else
        MsgBox "文字が選択されていません。", vbExclamation
    End If
End Sub

Sub ColorMagenta()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = RGB(255, 0, 255)
    Else
        MsgBox "文字が選択されていません。", vbExclamation
    End If
End Sub


' ==========================================
' 3. 背景色（塗りつぶし）マクロ
' ==========================================
Private Sub ApplyShadingRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long)
    With Selection.Range.Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = RGB(r, g, b)
    End With
End Sub

Public Sub FillYellow()
    ApplyShadingRGB 255, 255, 0
End Sub

Public Sub FillGray()
    ApplyShadingRGB 191, 191, 191
End Sub

Public Sub FillLightGreen()
    ApplyShadingRGB 204, 255, 204
End Sub

Public Sub FillLightBlue()
    ApplyShadingRGB 221, 235, 247
End Sub

Public Sub FillLightRed()
    ApplyShadingRGB 255, 199, 206
End Sub

Public Sub FillAqua()
    ApplyShadingRGB 204, 255, 255
End Sub

Public Sub ClearFill()
    With Selection.Range.Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = wdColorAutomatic
    End With
End Sub


' ==========================================
' 4. 罫線操作マクロ
' ==========================================
Sub DeleteBorders_Word()
    Dim rng As Range
    Set rng = Selection.Range
    If Selection.Information(wdWithInTable) Then
        With Selection.Borders
            .OutsideLineStyle = wdLineStyleNone
            .InsideLineStyle = wdLineStyleNone
        End With
    Else
        With rng.ParagraphFormat.Borders
            .OutsideLineStyle = wdLineStyleNone
            .InsideLineStyle = wdLineStyleNone
        End With
    End If
End Sub

Sub AddBorders_Word()
    Dim rng As Range
    Set rng = Selection.Range
    If Selection.Information(wdWithInTable) Then
        Call SetAllBorders_Table(wdLineStyleSingle, wdLineWidth050pt, wdColorAutomatic, wdLineStyleSingle, wdLineWidth050pt, wdColorAutomatic)
    Else
        Dim p As Paragraph
        For Each p In rng.Paragraphs
            With p.Borders(wdBorderLeft):   .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
            With p.Borders(wdBorderTop):    .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
            With p.Borders(wdBorderRight):  .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
            With p.Borders(wdBorderBottom): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
        Next p
    End If
End Sub

Sub AddBorders_OuterBold_InnerThin_Word()
    Dim rng As Range
    Set rng = Selection.Range
    If Selection.Information(wdWithInTable) Then
        Call SetAllBorders_Table(wdLineStyleSingle, wdLineWidth150pt, wdColorAutomatic, wdLineStyleSingle, wdLineWidth050pt, wdColorAutomatic)
    Else
        With rng.ParagraphFormat.Borders
            .Enable = True
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth150pt
            .outsideColor = wdColorAutomatic
            .InsideLineStyle = wdLineStyleNone
        End With
    End If
End Sub

Sub SetNormalBorderAroundSelection_Word()
    Dim rng As Range
    Set rng = Selection.Range
    If Selection.Information(wdWithInTable) Then
        With Selection.Borders
            With .Item(wdBorderLeft):   .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
            With .Item(wdBorderTop):    .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
            With .Item(wdBorderRight):  .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
            With .Item(wdBorderBottom): .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt: .Color = wdColorAutomatic: End With
            .Item(wdBorderHorizontal).LineStyle = wdLineStyleNone
            .Item(wdBorderVertical).LineStyle = wdLineStyleNone
        End With
    Else
        With rng.ParagraphFormat.Borders
            .Enable = True
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth050pt
            .outsideColor = wdColorAutomatic
            .InsideLineStyle = wdLineStyleNone
        End With
    End If
End Sub

Sub FillGray_And_BottomDoubleLine_Word()
    Dim rng As Range
    Set rng = Selection.Range
    With rng.Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = wdColorGray05
    End With
    If Selection.Information(wdWithInTable) Then
        With Selection.Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
    Else
        With rng.ParagraphFormat.Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth150pt
            .Color = wdColorAutomatic
        End With
    End If
End Sub

Private Sub SetAllBorders_Table(ByVal outsideStyle As WdLineStyle, ByVal outsideWidth As WdLineWidth, ByVal outsideColor As WdColor, ByVal insideStyle As WdLineStyle, ByVal insideWidth As WdLineWidth, ByVal insideColor As WdColor)
    Dim nCells As Long, nRows As Long, nCols As Long
    nCells = Selection.Cells.Count
    If nCells > 0 Then
        On Error Resume Next
        nRows = Selection.Rows.Count
        nCols = Selection.Columns.Count
        On Error GoTo 0
        With Selection.Cells.Borders
            SafeSetBorder .Item(wdBorderLeft), outsideStyle, outsideWidth, outsideColor
            SafeSetBorder .Item(wdBorderTop), outsideStyle, outsideWidth, outsideColor
            SafeSetBorder .Item(wdBorderRight), outsideStyle, outsideWidth, outsideColor
            SafeSetBorder .Item(wdBorderBottom), outsideStyle, outsideWidth, outsideColor
            If nRows > 1 Then SafeSetBorder .Item(wdBorderHorizontal), insideStyle, insideWidth, insideColor
            If nCols > 1 Then SafeSetBorder .Item(wdBorderVertical), insideStyle, insideWidth, insideColor
        End With
        Exit Sub
    End If
    If Selection.Range.Tables.Count > 0 Then
        With Selection.Range.Tables(1).Borders
            SafeSetBorder .Item(wdBorderLeft), outsideStyle, outsideWidth, outsideColor
            SafeSetBorder .Item(wdBorderTop), outsideStyle, outsideWidth, outsideColor
            SafeSetBorder .Item(wdBorderRight), outsideStyle, outsideWidth, outsideColor
            SafeSetBorder .Item(wdBorderBottom), outsideStyle, outsideWidth, outsideColor
            SafeSetBorder .Item(wdBorderHorizontal), insideStyle, insideWidth, insideColor
            SafeSetBorder .Item(wdBorderVertical), insideStyle, insideWidth, insideColor
        End With
    End If
End Sub

Private Sub SafeSetBorder(ByVal brd As Border, ByVal st As WdLineStyle, ByVal wd As WdLineWidth, ByVal clr As WdColor)
    On Error Resume Next
    With brd
        .LineStyle = st
        .LineWidth = wd
        .Color = clr
    End With
    On Error GoTo 0
End Sub


' ==========================================
' 5. 画像サイズ調整マクロ
' ==========================================
Sub FitImageToMargins()
    Dim docSec As Section
    Dim targetWidth As Single
    If Selection.InlineShapes.Count = 0 And Selection.ShapeRange.Count = 0 Then
        MsgBox "画像を選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    Set docSec = Selection.Sections(1)
    With docSec.PageSetup
        targetWidth = .pageWidth - .leftMargin - .rightMargin
    End With
    If Selection.InlineShapes.Count > 0 Then
        With Selection.InlineShapes(1)
            .LockAspectRatio = msoTrue
            .Width = targetWidth
        End With
    ElseIf Selection.ShapeRange.Count > 0 Then
        With Selection.ShapeRange(1)
            .LockAspectRatio = msoTrue
            .Width = targetWidth
            .Left = wdShapeLeftMargin
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
        End With
    End If
End Sub

Sub FitImageToTableCell()
    Dim targetCell As Cell
    Dim targetWidth As Single
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "表の中の画像を選択してください。", vbExclamation
        Exit Sub
    End If
    If Selection.InlineShapes.Count = 0 Then
        MsgBox "行内配置の画像を選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    Set targetCell = Selection.Cells(1)
    targetWidth = targetCell.Width - targetCell.LeftPadding - targetCell.RightPadding
    If targetWidth <= 0 Then targetWidth = targetCell.Width
    With Selection.InlineShapes(1)
        .LockAspectRatio = msoTrue
        .Width = targetWidth
    End With
End Sub

Sub MaximizeImageSizeWithinPage()
    Dim targetShape As InlineShape
    Dim docSec As Section
    Dim pageHeight As Single, pageWidth As Single
    Dim bottomMargin As Single, leftMargin As Single, rightMargin As Single
    Dim currentTopPosition As Single
    Dim remainingHeight As Single, availableWidth As Single
    Dim ratioHeight As Single, ratioWidth As Single, minRatio As Single
    
    If Selection.InlineShapes.Count = 0 Then
        MsgBox "「行内配置」の画像を選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    Set targetShape = Selection.InlineShapes(1)
    Set docSec = Selection.Sections(1)
    
    On Error Resume Next
    pageHeight = docSec.PageSetup.pageHeight
    bottomMargin = docSec.PageSetup.bottomMargin
    leftMargin = docSec.PageSetup.leftMargin
    rightMargin = docSec.PageSetup.rightMargin
    If pageHeight = 0 Then pageHeight = CentimetersToPoints(29.7)
    If bottomMargin = 0 Then bottomMargin = CentimetersToPoints(3)
    If leftMargin = 0 Then leftMargin = CentimetersToPoints(3)
    If rightMargin = 0 Then rightMargin = CentimetersToPoints(3)
    pageWidth = CentimetersToPoints(21)
    On Error GoTo 0
    
    currentTopPosition = Selection.Information(wdVerticalPositionRelativeToPage)
    remainingHeight = pageHeight - currentTopPosition - bottomMargin - 15
    availableWidth = pageWidth - leftMargin - rightMargin
    
    If remainingHeight < 50 Then
        remainingHeight = pageHeight - docSec.PageSetup.topMargin - bottomMargin - 15
    End If
    
    targetShape.LockAspectRatio = msoTrue
    ratioWidth = availableWidth / targetShape.Width
    ratioHeight = remainingHeight / targetShape.Height
    
    If ratioWidth < ratioHeight Then
        minRatio = ratioWidth
    Else
        minRatio = ratioHeight
    End If
    
    targetShape.Width = targetShape.Width * minRatio
End Sub


' ==========================================
' 6. 表の幅調整マクロ
' ==========================================
Public Sub TableAutoFitToContents()
    Dim t As Table
    If Selection.Range.Tables.Count = 0 Then
        MsgBox "表内にカーソルを置くか、表を選択してください。", vbInformation
        Exit Sub
    End If
    For Each t In Selection.Range.Tables: t.AutoFitBehavior wdAutoFitContent: Next t
End Sub

Public Sub TableAutoFitToWindow()
    Dim t As Table
    If Selection.Range.Tables.Count = 0 Then
        MsgBox "表内にカーソルを置くか、表を選択してください。", vbInformation
        Exit Sub
    End If
    For Each t In Selection.Range.Tables: t.AutoFitBehavior wdAutoFitWindow: Next t
End Sub

Public Sub TableAutoFitFixed()
    Dim t As Table
    If Selection.Range.Tables.Count = 0 Then
        MsgBox "表内にカーソルを置くか、表を選択してください。", vbInformation
        Exit Sub
    End If
    For Each t In Selection.Range.Tables: t.AutoFitBehavior wdAutoFitFixed: Next t
End Sub


' ==========================================
' 7. 便利機能・ユーティリティ
' ==========================================
Sub RemoveLineBreaks()
    Dim rng As Range: Set rng = Selection.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^p": .Replacement.Text = "": .Execute Replace:=wdReplaceAll
    End With
End Sub

' ==========================================
' 数式挿入マクロ（完全ネイティブ動作版）
' ==========================================
Sub InsertEmptyEquation()
    On Error GoTo SafeExit
    
    ' 1. 現在のカーソル位置に数式の枠を挿入します
    ' （文字を選択している場合は、その文字を数式化します）
    Selection.OMaths.Add Range:=Selection.Range
    
    ' 2. 挿入直後は「枠全体」が選択されてしまい、入力すると枠が消えるため、
    ' カーソルの選択状態を解除し、右に1つ移動して「数式枠の内部」に入ります。
    Selection.Collapse Direction:=wdCollapseStart
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
SafeExit:
    On Error GoTo 0
End Sub

Sub InsertEquationSuperscript()
    On Error Resume Next
    Selection.TypeText Text:="^()"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    On Error GoTo 0
End Sub

Sub InsertEquationSubscript()
    On Error Resume Next
    Selection.TypeText Text:="_()"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    On Error GoTo 0
End Sub

Sub InsertSectionBreakNextPage()
    Selection.InsertBreak Type:=wdSectionBreakNextPage
End Sub


' ==========================================
' 8. マクロ＆ショートカット一覧表出力（未割り当て含む）
' ==========================================
Sub ExportCustomShortcutsCatalog()
    Dim kb As KeyBinding, doc As Document, tbl As Table, r As Long
    Dim macroName As String, cleanName As String, description As String
    Dim macroKeys As Object
    Dim m As Variant
    Dim currentKey As String
    Dim allMacros As Variant
    
    ' 1. 現在割り当てられているマクロのショートカットをスキャンして記憶
    Set macroKeys = CreateObject("Scripting.Dictionary")
    For Each kb In Application.KeyBindings
        If kb.KeyCategory = wdKeyCategoryMacro Then
            macroName = kb.Command
            cleanName = macroName
            If InStr(cleanName, ".") > 0 Then cleanName = Mid(cleanName, InStrRev(cleanName, ".") + 1)
            
            ' 同じマクロに複数キーがある場合はカンマで連結
            If macroKeys.Exists(cleanName) Then
                macroKeys(cleanName) = macroKeys(cleanName) & ", " & kb.KeyString
            Else
                macroKeys.Add cleanName, kb.KeyString
            End If
        End If
    Next kb
    
    ' 2. 新しい白紙の文書を作成
    Set doc = Documents.Add
    doc.Range.Text = "Word マクロ ＆ 便利ショートカット操作説明表" & vbCrLf & _
                     "作成日: " & Format(Date, "YYYY/MM/DD") & vbCrLf & vbCrLf
    
    ' 3列の表を作成（キー、コマンド名、機能・説明）
    Set tbl = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 1, 3)
    With tbl
        .Cell(1, 1).Range.Text = "ショートカットキー"
        .Cell(1, 2).Range.Text = "機能・マクロ名"
        .Cell(1, 3).Range.Text = "使ったときの状態（説明）"
        .Rows(1).Range.Bold = True
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray10
    End With
    
    r = 1
    
    ' 3. カタログに定義されているすべての自作マクロを網羅（未割り当て含む）
    allMacros = Array( _
        "RemoveLineBreaks", "FillGray_And_BottomDoubleLine_Word", "DeleteBorders_Word", _
        "AddBorders_Word", "AddBorders_OuterBold_InnerThin_Word", "SetNormalBorderAroundSelection_Word", _
        "TableAutoFitToContents", "TableAutoFitToWindow", "TableAutoFitFixed", _
        "InsertEmptyEquation", "InsertEquationSuperscript", "InsertEquationSubscript", _
        "FitImageToMargins", "FitImageToTableCell", "MaximizeImageSizeWithinPage", _
        "InsertSectionBreakNextPage", "ColorMagenta", "ColorRed", "ColorBlue", "ColorGreen", "ColorBlack", _
        "FillYellow", "FillGray", "FillLightGreen", "FillLightBlue", "FillLightRed", "FillAqua", "ClearFill", _
        "SetupWordShortcuts", "ExportCustomShortcutsCatalog" _
    )
    
    For Each m In allMacros
        cleanName = CStr(m)
        description = GetMacroDescription(cleanName)
        
        ' キーが割り当てられているかDictionaryから確認
        If macroKeys.Exists(cleanName) Then
            currentKey = macroKeys(cleanName)
        Else
            currentKey = "（未割り当て）"
        End If
        
        r = r + 1
        tbl.Rows.Add
        tbl.Cell(r, 1).Range.Text = currentKey
        tbl.Cell(r, 2).Range.Text = cleanName
        tbl.Cell(r, 3).Range.Text = description
        tbl.Cell(r, 1).Range.Bold = True
    Next m
    
    ' 4. マクロ以外のカスタム標準機能（表の行削除、行挿入など）を抽出し、説明文を紐付け
    For Each kb In Application.KeyBindings
        If kb.KeyCategory <> wdKeyCategoryMacro Then
            cleanName = kb.Command
            If InStr(cleanName, ".") > 0 Then cleanName = Mid(cleanName, InStrRev(cleanName, ".") + 1)
            description = GetMacroDescription(cleanName)
            
            r = r + 1
            tbl.Rows.Add
            tbl.Cell(r, 1).Range.Text = kb.KeyString
            tbl.Cell(r, 2).Range.Text = cleanName
            tbl.Cell(r, 3).Range.Text = description
            tbl.Cell(r, 1).Range.Bold = True
        End If
    Next kb
    
    ' 表の見栄えを整える
    With tbl
        With .Borders
            .InsideLineStyle = wdLineStyleSingle: .InsideLineWidth = wdLineWidth050pt
            .OutsideLineStyle = wdLineStyleSingle: .OutsideLineWidth = wdLineWidth050pt
        End With
        .Rows(1).Select
        Selection.Rows.HeadingFormat = True
        .Rows.AllowBreakAcrossPages = False
        .Columns(1).PreferredWidthType = wdPreferredWidthPoints: .Columns(1).PreferredWidth = CentimetersToPoints(3.5)
        .Columns(2).PreferredWidthType = wdPreferredWidthPoints: .Columns(2).PreferredWidth = CentimetersToPoints(4.5)
        .Columns(3).PreferredWidthType = wdPreferredWidthPoints: .Columns(3).PreferredWidth = CentimetersToPoints(8.5)
    End With
    
    doc.Range(0, 0).Select
    MsgBox "未割り当てマクロと標準機能の説明を含めた操作説明表を作成しました！", vbInformation
End Sub

Private Function GetMacroDescription(ByVal mName As String) As String
    Select Case mName
       ' --- 自作マクロ ---
        Case "RemoveLineBreaks": GetMacroDescription = "選択した範囲内にある邪魔な改行（^p）をすべて一瞬で消去します。"
        Case "FillGray_And_BottomDoubleLine_Word": GetMacroDescription = "セルを薄いグレーで塗りつぶし、下部に美しい二重線を引きます（見出し用）。"
        Case "DeleteBorders_Word": GetMacroDescription = "選択した表、または段落のすべての罫線を完全に削除してまっさらにします。"
        Case "AddBorders_Word": GetMacroDescription = "表や段落に、標準的な黒の0.5ptの格子罫線を引きます。"
        Case "AddBorders_OuterBold_InnerThin_Word": GetMacroDescription = "表の外枠を「太線（1.5pt）」、内側の区切り線を「細線（0.5pt）」にしてメリハリをつけます。"
        Case "SetNormalBorderAroundSelection_Word": GetMacroDescription = "内側の罫線は付けず、選択した範囲の「外周のみ」に0.5ptの細線を引きます。"
        Case "TableAutoFitToContents": GetMacroDescription = "選択した表の横幅を、中に入っている文字列の長さに合わせてキュッと自動縮小します。"
        Case "TableAutoFitToWindow": GetMacroDescription = "選択した表の横幅を、ページの余白（左右幅）いっぱいに綺麗に広げます。"
        Case "TableAutoFitFixed": GetMacroDescription = "表の自動変形を解除し、現在の列幅でがっちり固定します。"
        Case "InsertEmptyEquation": GetMacroDescription = "現在カーソルがある場所に、新しく空の数式ブロックを挿入します。"
        Case "InsertEquationSuperscript": GetMacroDescription = "数式内で上付き文字「 ^() 」を入力し、自動でカッコの間にカーソルを移動します。"
        Case "InsertEquationSubscript": GetMacroDescription = "数式内で下付き文字「 _() 」を入力し、自動でカッコの間にカーソルを移動します。"
        Case "FitImageToMargins": GetMacroDescription = "選択した画像のサイズを、ページの印刷余白いっぱいに拡大・縮小します（縦横比固定）。"
        Case "FitImageToTableCell": GetMacroDescription = "表の中にある画像のサイズを、セルの横幅ぴったりに自動調整します（行内配置用）。"
        Case "MaximizeImageSizeWithinPage": GetMacroDescription = "画像のサイズを、現在のページ内の残りスペースにギリギリ収まる最大サイズにします。"
        Case "InsertSectionBreakNextPage": GetMacroDescription = "カーソル位置に「次のページから開始するセクション区切り」をパッと挿入します。"
        Case "ColorMagenta": GetMacroDescription = "選択した文字の色を、鮮やかなマゼンタ（ピンク）に変更します。"
        Case "ColorRed": GetMacroDescription = "選択した文字の色を、パッと目立つ赤色に変更します。"
        Case "ColorBlue": GetMacroDescription = "選択した文字の色を、鮮やかな青色に変更します。"
        Case "ColorGreen": GetMacroDescription = "選択した文字の色を、鮮やかな緑色に変更します。"
        Case "ColorBlack": GetMacroDescription = "選択した文字の色を、標準の黒色に戻します。"
        Case "FillYellow": GetMacroDescription = "選択範囲をパステル調の明るい「黄色」で塗りつぶします。"
        Case "FillGray": GetMacroDescription = "選択範囲をパステル調の「グレー」で塗りつぶします。"
        Case "FillLightGreen": GetMacroDescription = "選択範囲をパステル調の目に優しい「薄緑」で塗りつぶします。"
        Case "FillLightBlue": GetMacroDescription = "選択範囲をパステル調の爽やかな「薄青」で塗りつぶします。"
        Case "FillLightRed": GetMacroDescription = "選択範囲をパステル調の「薄赤（ピンクがかった赤）」で塗りつぶします。"
        Case "FillAqua": GetMacroDescription = "選択範囲をパステル調の綺麗な「水色」で塗りつぶします。"
        Case "ClearFill": GetMacroDescription = "設定されている背景の塗りつぶし（網かけ）を完全に解除して透明に戻します。"
        Case "SetupWordShortcuts": GetMacroDescription = "Wordのショートカットキーを一括で登録・設定します。"
        Case "ExportCustomShortcutsCatalog": GetMacroDescription = "現在設定されている自作マクロとショートカットキーの一覧表を作成します。"
        
        ' --- Word標準コマンド（カスタム割り当て分） ---
        Case "InsertEquation": GetMacroDescription = "標準の数式ブロックを現在のカーソル位置に挿入します。"
        Case "CopyFormat": GetMacroDescription = "選択した文字列の書式（フォント、色、段落スタイルなど）をクリップボードにコピーします。"
        Case "PasteFormat": GetMacroDescription = "コピーした書式を、現在選択している文字列に適用（ペースト）します。"
        Case "TableInsertRowAbove": GetMacroDescription = "表の中で、現在カーソルがある行の「上」に新しい行を1行挿入します。"
        Case "TableSplitCells": GetMacroDescription = "選択した表のセルを、指定した行数・列数に細かく分割します。"
        Case "TableMergeCells": GetMacroDescription = "選択した複数のセルを結合して、1つの大きなセルにまとめます。"
        Case "FormatBordersAndShading": GetMacroDescription = "「線と網かけの設定」画面を開き、罫線の種類や背景の塗りつぶしを詳細に設定します。"
        Case "FormatParagraph": GetMacroDescription = "「段落」の設定画面を開き、行間、インデント、前後の間隔などを詳細に設定します。"
        Case "TableDeleteColumn": GetMacroDescription = "表の中で、現在カーソルがある「列」を丸ごと削除します。"
        Case "TableDeleteRow": GetMacroDescription = "表の中で、現在カーソルがある「行」を丸ごと削除します。"
        Case "PasteTextOnly": GetMacroDescription = "元の書式（色やサイズ）を完全に無視し、純粋な「テキスト（文字）のみ」として貼り付けます。"
        Case "EditPasteSpecial": GetMacroDescription = "「形式を選択して貼り付け」画面を開き、貼り付け形式（HTMLや画像など）を選んで実行します。"
        
        Case Else: GetMacroDescription = "（Word標準機能、または詳細説明のない機能です）"
    End Select
End Function