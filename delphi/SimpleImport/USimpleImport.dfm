object FSimpleImport: TFSimpleImport
  Left = 0
  Top = 0
  Caption = 'Simple Import Demo'
  ClientHeight = 506
  ClientWidth = 758
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    758
    506)
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 8
    Top = 8
    Width = 146
    Height = 457
    Anchors = [akLeft, akTop, akBottom]
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 357
      Width = 26
      Height = 13
      Caption = 'Zoom'
    end
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 144
      Height = 24
      Align = alTop
      BevelOuter = bvLowered
      Caption = 'Import Options'
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
      Top = 100
      Width = 97
      Height = 17
      Caption = 'Formatting'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object cbCellSizes: TCheckBox
      Left = 8
      Top = 123
      Width = 97
      Height = 17
      Caption = 'Cell Sizes'
      Checked = True
      State = cbChecked
      TabOrder = 2
    end
    object cbFormulas: TCheckBox
      Left = 8
      Top = 31
      Width = 97
      Height = 17
      Caption = 'Formulas'
      Checked = True
      State = cbChecked
      TabOrder = 3
    end
    object cbImages: TCheckBox
      Left = 8
      Top = 54
      Width = 97
      Height = 17
      Caption = 'Images'
      Checked = True
      State = cbChecked
      TabOrder = 4
    end
    object cbOutlines: TCheckBox
      Left = 8
      Top = 215
      Width = 97
      Height = 17
      Caption = 'Outlines'
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object cbPrintOptions: TCheckBox
      Left = 8
      Top = 192
      Width = 97
      Height = 17
      Caption = 'Print Options'
      Checked = True
      State = cbChecked
      TabOrder = 6
    end
    object cbHTML: TCheckBox
      Left = 8
      Top = 169
      Width = 97
      Height = 17
      Caption = 'HTML'
      TabOrder = 7
    end
    object cbActiveSheet: TCheckBox
      Left = 8
      Top = 308
      Width = 97
      Height = 17
      Caption = 'Active Sheet'
      TabOrder = 8
    end
    object cbReadOnlyToLocked: TCheckBox
      Left = 8
      Top = 146
      Width = 129
      Height = 17
      Caption = 'ReadOnly -> Locked'
      TabOrder = 9
    end
    object cbClearCells: TCheckBox
      Left = 8
      Top = 262
      Width = 97
      Height = 17
      Caption = 'Clear Cells'
      Checked = True
      State = cbChecked
      TabOrder = 10
    end
    object cbResizeGrid: TCheckBox
      Left = 8
      Top = 285
      Width = 138
      Height = 17
      Caption = 'Resize Grid'
      Checked = True
      State = cbChecked
      TabOrder = 11
    end
    object cbComments: TCheckBox
      Left = 8
      Top = 77
      Width = 97
      Height = 17
      Caption = 'Comments'
      Checked = True
      State = cbChecked
      TabOrder = 12
    end
    object edZoom: TEdit
      Left = 8
      Top = 376
      Width = 121
      Height = 21
      TabOrder = 13
      Text = '0'
      OnChange = edZoomChange
    end
  end
  object BtnAbout: TButton
    Left = 144
    Top = 473
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'About'
    TabOrder = 1
    OnClick = BtnAboutClick
  end
  object BtnImport: TButton
    Left = 8
    Top = 473
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'Import Excel File'
    TabOrder = 2
    OnClick = BtnImportClick
  end
  object AdvGridWorkbook1: TAdvGridWorkbook
    Left = 159
    Top = 8
    Width = 591
    Height = 457
    ActiveSheet = 0
    Sheets = <
      item
        Name = 'Sheet 1'
        Tag = 0
      end
      item
        Name = 'Sheet 2'
        Tag = 0
      end
      item
        Name = 'Sheet 3'
        Tag = 0
      end>
    TabLook.Font.Charset = DEFAULT_CHARSET
    TabLook.Font.Color = clWindowText
    TabLook.Font.Height = -11
    TabLook.Font.Name = 'Tahoma'
    TabLook.Font.Style = []
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 3
    Version = '3.3.2.0'
    object TAdvStringGrid
      Left = 0
      Top = 0
      Width = 587
      Height = 432
      Cursor = crDefault
      Align = alClient
      BorderStyle = bsNone
      DrawingStyle = gdsClassic
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      ParentFont = False
      ScrollBars = ssBoth
      TabOrder = 0
      HoverRowCells = [hcNormal, hcSelected]
      ActiveCellFont.Charset = DEFAULT_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -11
      ActiveCellFont.Name = 'Tahoma'
      ActiveCellFont.Style = [fsBold]
      ControlLook.FixedGradientHoverFrom = clGray
      ControlLook.FixedGradientHoverTo = clWhite
      ControlLook.FixedGradientDownFrom = clGray
      ControlLook.FixedGradientDownTo = clSilver
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
      FixedRowHeight = 22
      FixedFont.Charset = DEFAULT_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -11
      FixedFont.Name = 'Tahoma'
      FixedFont.Style = [fsBold]
      FloatFormat = '%.2f'
      HoverButtons.Buttons = <>
      HoverButtons.Position = hbLeftFromColumnLeft
      PrintSettings.DateFormat = 'dd/mm/yyyy'
      PrintSettings.Font.Charset = DEFAULT_CHARSET
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
      PrintSettings.HeaderFont.Name = 'Tahoma'
      PrintSettings.HeaderFont.Style = []
      PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
      PrintSettings.FooterFont.Color = clWindowText
      PrintSettings.FooterFont.Height = -11
      PrintSettings.FooterFont.Name = 'Tahoma'
      PrintSettings.FooterFont.Style = []
      PrintSettings.PageNumSep = '/'
      SearchFooter.FindNextCaption = 'Find &next'
      SearchFooter.FindPrevCaption = 'Find &previous'
      SearchFooter.Font.Charset = DEFAULT_CHARSET
      SearchFooter.Font.Color = clWindowText
      SearchFooter.Font.Height = -11
      SearchFooter.Font.Name = 'Tahoma'
      SearchFooter.Font.Style = []
      SearchFooter.HighLightCaption = 'Highlight'
      SearchFooter.HintClose = 'Close'
      SearchFooter.HintFindNext = 'Find next occurrence'
      SearchFooter.HintFindPrev = 'Find previous occurrence'
      SearchFooter.HintHighlight = 'Highlight occurrences'
      SearchFooter.MatchCaseCaption = 'Match case'
      SortSettings.DefaultFormat = ssAutomatic
      Version = '7.4.5.0'
    end
  end
  object AdvGridExcelImport: TAdvGridExcelImport
    AdvGridWorkbook = AdvGridWorkbook1
    Version = '1.0'
    Left = 592
    Top = 88
  end
  object OpenDialog: TOpenDialog
    Filter = 
      'Excel Files|*.xlsx;*.xls|Xls files (Excel 2003 or older)|*.xls|X' +
      'lsx files (Excel 2007 or newer)|*.xlsx|All files|*.*'
    Left = 592
    Top = 136
  end
end
