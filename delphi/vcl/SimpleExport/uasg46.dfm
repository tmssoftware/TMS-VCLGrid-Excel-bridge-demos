object Form1: TForm1
  Left = 414
  Top = 109
  Caption = 'TAdvExcelExport simple demo'
  ClientHeight = 553
  ClientWidth = 744
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  DesignSize = (
    744
    553)
  PixelsPerInch = 96
  TextHeight = 13
  object asg: TAdvStringGrid
    Left = 160
    Top = 8
    Width = 578
    Height = 506
    Cursor = crDefault
    Anchors = [akLeft, akTop, akRight, akBottom]
    DefaultRowHeight = 23
    DefaultDrawing = True
    DrawingStyle = gdsClassic
    FixedColor = clActiveBorder
    RowCount = 5
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goEditing]
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 0
    GridLineColor = 15527152
    GridFixedLineColor = 13947601
    HoverRowCells = [hcNormal, hcSelected]
    OnGetCellGradient = asgGetCellGradient
    OnCanEditCell = asgCanEditCell
    OnAnchorClick = asgAnchorClick
    OnGetEditorType = asgGetEditorType
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'MS Sans Serif'
    ActiveCellFont.Style = [fsBold]
    ActiveCellColor = 16575452
    ActiveCellColorTo = 16571329
    ControlLook.FixedGradientMirrorFrom = 16049884
    ControlLook.FixedGradientMirrorTo = 16247261
    ControlLook.FixedGradientHoverFrom = 16710648
    ControlLook.FixedGradientHoverTo = 16446189
    ControlLook.FixedGradientHoverMirrorFrom = 16049367
    ControlLook.FixedGradientHoverMirrorTo = 15258305
    ControlLook.FixedGradientDownFrom = 15853789
    ControlLook.FixedGradientDownTo = 15852760
    ControlLook.FixedGradientDownMirrorFrom = 15522767
    ControlLook.FixedGradientDownMirrorTo = 15588559
    ControlLook.FixedGradientDownBorder = 14007466
    ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownHeader.Font.Color = clWindowText
    ControlLook.DropDownHeader.Font.Height = -11
    ControlLook.DropDownHeader.Font.Name = 'Tahoma'
    ControlLook.DropDownHeader.Font.Style = []
    ControlLook.DropDownHeader.Visible = True
    ControlLook.DropDownHeader.Buttons = <>
    ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownFooter.Font.Color = clWindowText
    ControlLook.DropDownFooter.Font.Height = -11
    ControlLook.DropDownFooter.Font.Name = 'Tahoma'
    ControlLook.DropDownFooter.Font.Style = []
    ControlLook.DropDownFooter.Visible = True
    ControlLook.DropDownFooter.Buttons = <>
    EnhRowColMove = False
    Filter = <>
    FilterDropDown.Font.Charset = DEFAULT_CHARSET
    FilterDropDown.Font.Color = clWindowText
    FilterDropDown.Font.Height = -11
    FilterDropDown.Font.Name = 'Tahoma'
    FilterDropDown.Font.Style = []
    FilterDropDownClear = '(All)'
    FilterEdit.TypeNames.Strings = (
      'Starts with'
      'Ends with'
      'Contains'
      'Not contains'
      'Equal'
      'Not equal'
      'Clear')
    FixedRowHeight = 23
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWindowText
    FixedFont.Height = -11
    FixedFont.Name = 'MS Sans Serif'
    FixedFont.Style = []
    FloatFormat = '%.2f'
    GridImages = ImageList1
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    Look = glWin7
    Multilinecells = True
    PrintSettings.DateFormat = 'dd/mm/yyyy'
    PrintSettings.Font.Charset = ANSI_CHARSET
    PrintSettings.Font.Color = clWindowText
    PrintSettings.Font.Height = -11
    PrintSettings.Font.Name = 'Tahoma'
    PrintSettings.Font.Style = []
    PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
    PrintSettings.FixedFont.Color = clWindowText
    PrintSettings.FixedFont.Height = -11
    PrintSettings.FixedFont.Name = 'Tahoma'
    PrintSettings.FixedFont.Style = []
    PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
    PrintSettings.HeaderFont.Color = clWindowText
    PrintSettings.HeaderFont.Height = -11
    PrintSettings.HeaderFont.Name = 'MS Sans Serif'
    PrintSettings.HeaderFont.Style = []
    PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
    PrintSettings.FooterFont.Color = clWindowText
    PrintSettings.FooterFont.Height = -11
    PrintSettings.FooterFont.Name = 'MS Sans Serif'
    PrintSettings.FooterFont.Style = []
    PrintSettings.PageNumSep = '/'
    PrintSettings.NoAutoSize = True
    PrintSettings.PrintGraphics = True
    ScrollWidth = 16
    SearchFooter.Color = 16645370
    SearchFooter.ColorTo = 16247261
    SearchFooter.FindNextCaption = 'Find &next'
    SearchFooter.FindPrevCaption = 'Find &previous'
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'Tahoma'
    SearchFooter.Font.Style = []
    SearchFooter.HighLightCaption = 'Highlight'
    SearchFooter.HintClose = 'Close'
    SearchFooter.HintFindNext = 'Find next occurence'
    SearchFooter.HintFindPrev = 'Find previous occurence'
    SearchFooter.HintHighlight = 'Highlight occurences'
    SearchFooter.MatchCaseCaption = 'Match case'
    SortSettings.DefaultFormat = ssAutomatic
    SortSettings.HeaderColor = 16579058
    SortSettings.HeaderColorTo = 16579058
    SortSettings.HeaderMirrorColor = 16380385
    SortSettings.HeaderMirrorColorTo = 16182488
    URLShow = True
    Version = '7.4.2.1'
    ExplicitHeight = 498
    ColWidths = (
      64
      64
      64
      64
      64)
    RowHeights = (
      23
      23
      23
      23
      23)
  end
  object BtnExportExcel: TButton
    Left = 8
    Top = 521
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'Export to XLS/XLSX'
    TabOrder = 1
    OnClick = BtnExportExcelClick
    ExplicitTop = 513
  end
  object BtnAbout: TButton
    Left = 440
    Top = 520
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'About'
    TabOrder = 2
    OnClick = BtnAboutClick
    ExplicitTop = 512
  end
  object Panel1: TPanel
    Left = 8
    Top = 8
    Width = 146
    Height = 506
    Anchors = [akLeft, akTop, akBottom]
    TabOrder = 3
    ExplicitHeight = 498
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 144
      Height = 24
      Align = alTop
      BevelOuter = bvLowered
      Caption = 'Export Options'
      Color = clBtnShadow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindow
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentBackground = False
      ParentFont = False
      TabOrder = 0
    end
    object cbFormatting: TCheckBox
      Left = 8
      Top = 40
      Width = 97
      Height = 17
      Caption = 'Formatting'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object cbCellMargins: TCheckBox
      Left = 8
      Top = 63
      Width = 97
      Height = 17
      Caption = 'Cell Margins'
      Checked = True
      State = cbChecked
      TabOrder = 2
    end
    object cbCellsAsStrings: TCheckBox
      Left = 8
      Top = 86
      Width = 97
      Height = 17
      Caption = 'Cells as strings'
      TabOrder = 3
    end
    object cbCellSizes: TCheckBox
      Left = 8
      Top = 109
      Width = 97
      Height = 17
      Caption = 'Cell Sizes'
      Checked = True
      State = cbChecked
      TabOrder = 4
    end
    object cbFormulas: TCheckBox
      Left = 8
      Top = 132
      Width = 97
      Height = 17
      Caption = 'Formulas'
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object cbHardBorders: TCheckBox
      Left = 8
      Top = 155
      Width = 97
      Height = 17
      Caption = 'Hard borders'
      TabOrder = 6
    end
    object cbHiddenColumns: TCheckBox
      Left = 8
      Top = 178
      Width = 97
      Height = 17
      Caption = 'Hidden Columns'
      TabOrder = 7
    end
    object cbHiddenRows: TCheckBox
      Left = 8
      Top = 201
      Width = 97
      Height = 17
      Caption = 'Hidden Rows'
      TabOrder = 8
    end
    object cbImages: TCheckBox
      Left = 8
      Top = 224
      Width = 97
      Height = 17
      Caption = 'Images'
      Checked = True
      State = cbChecked
      TabOrder = 9
    end
    object cbOutlines: TCheckBox
      Left = 8
      Top = 271
      Width = 97
      Height = 17
      Caption = 'Outlines'
      Checked = True
      State = cbChecked
      TabOrder = 10
    end
    object cbPrintOptions: TCheckBox
      Left = 8
      Top = 294
      Width = 97
      Height = 17
      Caption = 'Print Options'
      Checked = True
      State = cbChecked
      TabOrder = 11
    end
    object cbRawHTML: TCheckBox
      Left = 8
      Top = 317
      Width = 97
      Height = 17
      Caption = 'Raw HTML'
      TabOrder = 12
    end
    object cbRawRTF: TCheckBox
      Left = 8
      Top = 340
      Width = 97
      Height = 17
      Caption = 'Raw RTF'
      TabOrder = 13
    end
    object cbReadOnlyToLocked: TCheckBox
      Left = 8
      Top = 363
      Width = 129
      Height = 17
      Caption = 'ReadOnly -> Locked'
      TabOrder = 14
    end
    object cbShowGridlines: TCheckBox
      Left = 8
      Top = 386
      Width = 97
      Height = 17
      Caption = 'Show Gridlines'
      Checked = True
      State = cbChecked
      TabOrder = 15
    end
    object cbSummaryRowsBelow: TCheckBox
      Left = 8
      Top = 409
      Width = 129
      Height = 17
      Caption = 'Summary rows below'
      TabOrder = 16
    end
    object cbStandardPalette: TCheckBox
      Left = 8
      Top = 432
      Width = 138
      Height = 17
      Caption = 'Excel standard palette'
      Checked = True
      State = cbChecked
      TabOrder = 17
    end
    object cbWordWrapped: TCheckBox
      Left = 8
      Top = 455
      Width = 138
      Height = 17
      Caption = 'Word wrapped'
      Checked = True
      State = cbChecked
      TabOrder = 18
    end
    object cbCheckboxes: TCheckBox
      Left = 8
      Top = 247
      Width = 97
      Height = 17
      Caption = 'Checkboxes'
      Checked = True
      State = cbChecked
      TabOrder = 19
    end
    object cbHyperlinks: TCheckBox
      Left = 8
      Top = 478
      Width = 138
      Height = 17
      Caption = 'Hyperlinks'
      Checked = True
      State = cbChecked
      TabOrder = 20
    end
  end
  object BtnExportPdf: TButton
    Left = 135
    Top = 520
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'Export to PDF'
    TabOrder = 4
    OnClick = BtnExportPdfClick
    ExplicitTop = 512
  end
  object BtnExportHTML: TButton
    Left = 262
    Top = 520
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'Export to HTML'
    TabOrder = 5
    OnClick = BtnExportHTMLClick
    ExplicitTop = 512
  end
  object btnPreviewGrid: TButton
    Left = 567
    Top = 520
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'Preview Grid'
    TabOrder = 6
    OnClick = btnPreviewGridClick
    ExplicitTop = 512
  end
  object ImageList1: TImageList
    Left = 368
    Top = 416
    Bitmap = {
      494C010102000400900010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000001000000001002000000000000010
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000808080000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000C3B5B500BA8A8A0096969600B9B9B9000000000000000000000000000000
      00000000000000000000000000000000000000000000000000000000FF000000
      8000000000008080800000000000000000000000000000000000808080000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000000000C3B5
      B500CB716800C21E0000916E6E0097979700B8B8B80000000000000000000000
      00000000000000000000000000000000000000000000000000000000FF000000
      8000000000008080800000000000000000000000000000000000808080008080
      8000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000B5C3B50093A8
      6800813E000000800000AE3C2700936C6C0097979700B8B8B800000000000000
      00000000000000000000000000000000000000000000000000000000FF000000
      8000000080000000000080808000000000000000FF0000008000000000008080
      8000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000067D4670000C0
      00000080000000800000485B0000AA3E2800946B6B0098989800000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      FF000000800000008000808080000000FF000000800000008000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000E600000080
      00001F7000005A731200009F0000008000004C590000A6412900AAAAAA000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000FF000000800000008000000080000000800000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000E600000080
      0000485B0000D42A2A0034D13400009F0000008000009136000096969600B7B7
      B700000000000000000000000000000000000000000000000000000000000000
      0000000000000000FF0000008000000080000000000080808000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000051D8510000B2
      0000485B0000DF36360098CC980034D03400009F0000455D0000966969009595
      9500B8B8B8000000000000000000000000000000000000000000000000000000
      0000000000000000FF0000008000000080000000800000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000B0C5B00053D6
      53001FB9000088A140000000000098CC980034D13400009F0000AA3C2400946B
      6B0094949400B9B9B90000000000000000000000000000000000000000000000
      00000000FF000000800000000000000080000000800000000000808080000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000009BCC9B00009C00003C61
      0000B6342000906F6F0000000000000000000000000000000000000000000000
      FF00000080000000000000000000000000000000FF0000008000000000008080
      8000808080000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000002DD72D000099
      0000C21E00009867670000000000000000000000000000000000000000000000
      FF0000008000000000000000000000000000000000000000FF00000080000000
      0000808080000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000A1C9A10029DB
      2900C2350000BF8F8F0000000000000000000000000000000000000000000000
      00000000FF000000000000000000000000000000000000000000000080000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000000000A6C8
      A600C0A79B00C2B9B90000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000000000FF000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000100000000100010000000000800000000000000000000000
      000000000000000000000000FFFFFF00FFFFFFFF00000000FFFFF7FF00000000
      F0FFC3DF00000000E07FC38F00000000C03FC10F00000000C03FE01F00000000
      C01FF03F00000000C00FF83F00000000C007F83F00000000C203F01F00000000
      FF83E30700000000FFC3E38700000000FFC3F7CF00000000FFE3FFDF00000000
      FFFFFFFF00000000FFFFFFFF0000000000000000000000000000000000000000
      000000000000}
  end
  object AdvGridExcelExport: TAdvGridExcelExport
    AdvStringGrid = asg
    Version = '1.0'
    Left = 280
    Top = 416
  end
  object SaveDialogExcel: TSaveDialog
    DefaultExt = 'xlsx'
    Filter = 'Excel 2007|*.xlsx|Excel 97/2003|*.xls'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 464
    Top = 320
  end
  object SaveDialogPdf: TSaveDialog
    DefaultExt = 'pdf'
    Filter = 'Pdf files|*.pdf'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 464
    Top = 368
  end
  object SaveDialogHtml: TSaveDialog
    DefaultExt = 'htm'
    Filter = 'Html files|*.htm'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 464
    Top = 416
  end
  object AdvPreviewDialog1: TAdvPreviewDialog
    CloseAfterPrint = False
    DialogCaption = 'Preview'
    DialogPrevBtn = 'Previous'
    DialogNextBtn = 'Next'
    DialogPrintBtn = 'Print'
    DialogCloseBtn = 'Close'
    Grid = asg
    PreviewFast = False
    PreviewWidth = 350
    PreviewHeight = 300
    PreviewLeft = 100
    PreviewTop = 100
    PreviewCenter = False
    Left = 368
    Top = 280
  end
end
