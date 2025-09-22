{*************************************************************************}
{ TAdvStringGrid demo unit                                                }
{                                                                         }
{ written by TMS Software                                                 }
{            copyright (c)1998-2012                                       }
{            Email : info@tmssoftware.Com                                 }
{            Web : https://www.tmssoftware.com                             }
{                                                                         }
{ The source code is given as is. The author is not responsible           }
{ for any possible damage done due to the use of this code.               }
{ The component can be freely used in any application. The complete       }
{ source code remains property of the author and may not be distributed,  }
{ published, given or sold in any form as such. No parts of the source    }
{ code can be included in any other component or application without      }
{ written authorization of the author.                                    }
{*************************************************************************}
unit uasg46;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, AdvGrid, StdCtrls, BaseGrid, ComCtrls, ShellAPI, JPEG,
  ImgList, AsgPrev, AdvObj, ExtCtrls, AdvUtil,
  {$if CompilerVersion >= 23.0} System.UITypes, {$IFEND}
  UAdvGridExcelExport;

type
  TForm1 = class(TForm)
    ImageList1: TImageList;
    asg: TAdvStringGrid;
    BtnExportExcel: TButton;
    BtnAbout: TButton;
    AdvGridExcelExport: TAdvGridExcelExport;
    Panel1: TPanel;
    Panel2: TPanel;
    cbFormatting: TCheckBox;
    cbCellMargins: TCheckBox;
    cbCellsAsStrings: TCheckBox;
    cbCellSizes: TCheckBox;
    cbFormulas: TCheckBox;
    cbHardBorders: TCheckBox;
    cbHiddenColumns: TCheckBox;
    cbHiddenRows: TCheckBox;
    cbImages: TCheckBox;
    cbOutlines: TCheckBox;
    cbPrintOptions: TCheckBox;
    cbRawHTML: TCheckBox;
    cbRawRTF: TCheckBox;
    cbReadOnlyToLocked: TCheckBox;
    cbShowGridlines: TCheckBox;
    cbSummaryRowsBelow: TCheckBox;
    cbStandardPalette: TCheckBox;
    cbWordWrapped: TCheckBox;
    SaveDialogExcel: TSaveDialog;
    cbCheckboxes: TCheckBox;
    BtnExportPdf: TButton;
    BtnExportHTML: TButton;
    SaveDialogPdf: TSaveDialog;
    SaveDialogHtml: TSaveDialog;
    AdvPreviewDialog1: TAdvPreviewDialog;
    btnPreviewGrid: TButton;
    cbHyperlinks: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure mgridGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var AAlignment: TAlignment);
    procedure asgAnchorClick(Sender: TObject; ARow, ACol: Integer;
      Anchor: string; var AutoHandle: Boolean);
    procedure asgCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure asgGetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure BtnExportExcelClick(Sender: TObject);
    procedure BtnAboutClick(Sender: TObject);
    procedure asgGetCellGradient(Sender: TObject; ARow, ACol: Integer;
      var Color, ColorTo, ColorMirror, ColorMirrorTo: TColor;
      var GD: TCellGradientDirection);
    procedure BtnExportPdfClick(Sender: TObject);
    procedure BtnExportHTMLClick(Sender: TObject);
    procedure btnPreviewGridClick(Sender: TObject);
  private
    procedure SetExportOptions;
    { Private declarations }
  public
    { Public declarations }
    picture: tpicture;
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

procedure TForm1.FormCreate(Sender: TObject);
begin
  Picture := TPicture.Create;
  Picture.LoadFromFile('..\..\slr.jpg');

  asg.RowCount := 25;
  asg.ColCount := 12;

  asg.MergeCells(1, 1, 2, 2);
  asg.MergeCells(1, 0, 2, 1);
  asg.MergeCells(0, 2, 1, 2);


  asg.Cells[1, 1] := '<font color="#FF0000">This</font> is <b>bold</b>';
  asg.Cells[1, 0] := '<img src="idx:0"><font face="tahoma">This is <b>Tahoma</b></font>';

  asg.Cells[0, 2] := 'Fixed col'#13'merged';

  asg.MergeCells(2, 3, 2, 2);
  asg.Cells[2, 3] := 'Here we have the <a href="action">link</a>';
  asg.Colors[2, 3] := clYellow;

  asg.MergeCells(0, 8, asg.ColCount - 1, 1);
  asg.Cells[0, 8] := 'This is a fixed merged cell across the grid';

  asg.MergeCells(4, 2, 4, 6);


  asg.AddPicture(4, 2, Picture, True, ShrinkWithAspectRatio, 5, haLeft, vaTop);

  asg.MergeCells(1, 6, 2, 2);
  asg.Cells[1, 6] := 'Wordwrapped text in a merged cell displayed here';

  asg.MergeCells(1, 10, 3, 1);
  asg.AddCheckBox(1, 10, False, False);
  asg.Cells[1, 10] := 'Checkbox 1';
  asg.RowHeights[10] := 50;
  asg.MergeCells(1, 11, 3, 1);
  asg.AddCheckBox(1, 11, True, False);
  asg.Cells[1, 11] := 'Checkbox 2';

  asg.Cells[0, 12] := 'Combo';

  asg.MergeCells(1, 12, 3, 1);
  asg.AddComboString('BMW');
  asg.AddComboString('Audi');
  asg.AddComboString('Porsche');
  asg.AddComboString('Ferrari');

  asg.AddComment(1,13,'This comment is at the end of the grid!');

  asg.AddNode(3,5);
  asg.AddNode(5,2);

  asg.ColWidths[5] := 100;

  asg.AddMultiImage(5, 10, 0, haBeforeText, vaTop);
  asg.CellImages[5, 10].Add(0);
  asg.CellImages[5, 10].Add(0);
  asg.Cells[5, 10] := 'Info';

  asg.AddMultiImage(5, 11, 0, haBeforeText, vaTop);
  asg.CellImages[5, 11].Add(1);
  asg.Cells[5, 11] := 'Warning';

  asg.AddShape(2, 9, csArrowDown, clRed, clBlack, haCenter, vaTop);
  asg.AddShape(1, 1, csCircle, clBlue, clBlack, haRight, vaCenter);

  asg.Colors[2, 5] := clRed;
  asg.ColorsTo[2, 5] := clNavy;
  asg.Gradients[3, 5] := TCellGradientDirection.GradientVertical;
  asg.Colors[3, 5] := clYellow;
  asg.ColorsTo[3, 5] := clGreen;
  asg.Gradients[3, 5] := TCellGradientDirection.GradientHorizontal;

  asg.RowHeights[15] := 60;

  asg.Cells[5, 12] := 'A cell <a href="https://www.tmssoftware.com" title="TMS software">hyperlink</a>';
  asg.Cells[5, 13] := 'https://www.tmssoftware.com';

  asg.PrintSettings.Title := TPrintPosition.ppTopLeft;
  asg.PrintSettings.TitleText := 'AdvStringGrid & Exporting';
  asg.PrintSettings.PageNr := TPrintPosition.ppTopRight;
  asg.PrintSettings.PagePrefix := 'Page';
  asg.PrintSettings.PageNumSep := ' of ';
  asg.PrintSettings.PageSuffix := 'Total';
  asg.PrintSettings.Time := TPrintPosition.ppBottomLeft;
  asg.PrintSettings.Date := TPrintPosition.ppBottomCenter;
  asg.PrintSettings.HeaderFont.Name := 'Times New Roman';
  asg.PrintSettings.HeaderFont.Color := clNavy;
  asg.PrintSettings.FooterFont.Name := 'Calibri';
  asg.PrintSettings.FooterFont.Color := clMaroon;
  asg.PrintSettings.FooterFont.Size := 14;
  asg.PrintSettings.FooterFont.Style := [fsBold];

end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  Picture.Free;
end;

procedure TForm1.mgridGetAlignment(Sender: TObject; ARow, ACol: Integer;
  var AAlignment: TAlignment);
begin
  AAlignment := taCenter;
end;

procedure TForm1.asgAnchorClick(Sender: TObject; ARow, ACol: Integer;
  Anchor: string; var AutoHandle: Boolean);
begin
  ShowMessage('Anchor clicked');
end;

procedure TForm1.asgCanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  CanEdit := (ARow > 5) and (ACol < 5);
end;

procedure TForm1.asgGetCellGradient(Sender: TObject; ARow, ACol: Integer;
  var Color, ColorTo, ColorMirror, ColorMirrorTo: TColor;
  var GD: TCellGradientDirection);
begin
  if (ARow = 15) then
  begin
    Color := clBlue;
    ColorTo := clWhite;
    ColorMirror := clGreen;
    ColorMirrorto := clRed;
    GD := TCellGradientDirection.GradientVertical;
  end;
end;

procedure TForm1.asgGetEditorType(Sender: TObject; ACol, ARow: Integer;
  var AEditor: TEditorType);
begin
  if (ACol = 1) and (ARow = 12) then
  begin
    AEditor := edComboList;
  end;

end;

procedure TForm1.SetExportOptions;
begin
    AdvGridExcelExport.ExportOptions.CellSizes := cbCellSizes.Checked;
    AdvGridExcelExport.ExportOptions.Formulas := cbFormulas.Checked;
    AdvGridExcelExport.ExportOptions.CellsAsStrings := cbCellsAsStrings.Checked;
    AdvGridExcelExport.ExportOptions.CellFormatting := cbFormatting.Checked;
    AdvGridExcelExport.ExportOptions.WordWrapped := cbWordWrapped.Checked;
    AdvGridExcelExport.ExportOptions.RawHTML := cbRawHTML.Checked;
    AdvGridExcelExport.ExportOptions.RawRTF := cbRawRTF.Checked;
    AdvGridExcelExport.ExportOptions.Images := cbImages.Checked;
    AdvGridExcelExport.ExportOptions.Checkboxes := cbCheckboxes.Checked;
    AdvGridExcelExport.ExportOptions.HiddenColumns := cbHiddenColumns.Checked;
    AdvGridExcelExport.ExportOptions.HiddenRows := cbHiddenRows.Checked;
    AdvGridExcelExport.ExportOptions.HardBorders := cbHardBorders.Checked;
    AdvGridExcelExport.ExportOptions.ShowGridLines := cbShowGridLines.Checked;
    AdvGridExcelExport.ExportOptions.CellMargins := cbCellMargins.Checked;
    AdvGridExcelExport.ExportOptions.ReadonlyCellsAsLocked := cbReadOnlyToLocked.Checked;
    AdvGridExcelExport.ExportOptions.PrintOptions := cbPrintOptions.Checked;
    AdvGridExcelExport.ExportOptions.Outlines := cbOutlines.Checked;
    AdvGridExcelExport.ExportOptions.SummaryRowsBelowDetail := cbSummaryRowsBelow.Checked;
    AdvGridExcelExport.ExportOptions.UseExcelStandardColorPalette := cbStandardPalette.Checked;
    if cbHyperlinks.Checked then
      AdvGridExcelExport.ExportOptions.Hyperlinks := THyperlinkExportMode.FirstLink
      else AdvGridExcelExport.ExportOptions.Hyperlinks := THyperlinkExportMode.None;
end;

procedure TForm1.BtnExportExcelClick(Sender: TObject);
begin
  if not SaveDialogExcel.Execute then exit;

  SetExportoptions;
  AdvGridExcelExport.Export(SaveDialogExcel.FileName, 'Result');

  if MessageDlg('Do you want to open the generated file?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    ShellExecute(0, 'open', PCHAR(SaveDialogExcel.FileName), nil, nil, SW_SHOWNORMAL);
  end;

end;

procedure TForm1.BtnExportHTMLClick(Sender: TObject);
begin
  if not SaveDialogHtml.Execute then exit;

  SetExportoptions;
  AdvGridExcelExport.ExportHtml(SaveDialogHtml.FileName, THtmlExportMode.SingleSheet);

  if MessageDlg('Do you want to open the generated file?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    ShellExecute(0, 'open', PCHAR(SaveDialogHtml.FileName), nil, nil, SW_SHOWNORMAL);
  end;
end;

procedure TForm1.BtnExportPdfClick(Sender: TObject);
begin
  if not SaveDialogPdf.Execute then exit;

  SetExportoptions;
  AdvGridExcelExport.ExportPdf(SaveDialogPdf.FileName);

  if MessageDlg('Do you want to open the generated file?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    ShellExecute(0, 'open', PCHAR(SaveDialogPdf.FileName), nil, nil, SW_SHOWNORMAL);
  end;
end;

procedure TForm1.btnPreviewGridClick(Sender: TObject);
begin
  AdvPreviewDialog1.Execute;
end;

procedure TForm1.BtnAboutClick(Sender: TObject);
begin
  ShowMessage('This is a demo based in Demo 46 of TAdvStringGrid, showing the basic functionality of TAdvExcelExport');
end;

end.
