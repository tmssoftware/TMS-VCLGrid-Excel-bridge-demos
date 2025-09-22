unit USimpleImport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, Grids, AdvObj, BaseGrid, AdvGrid,
  {$if CompilerVersion >= 23.0} System.UITypes, {$IFEND}
  StdCtrls, ExtCtrls, UAdvGridExcelImport, AdvGridWorkbook;

type
  TFSimpleImport = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    cbFormatting: TCheckBox;
    cbCellSizes: TCheckBox;
    cbFormulas: TCheckBox;
    cbImages: TCheckBox;
    cbOutlines: TCheckBox;
    cbPrintOptions: TCheckBox;
    cbHTML: TCheckBox;
    cbActiveSheet: TCheckBox;
    cbReadOnlyToLocked: TCheckBox;
    cbClearCells: TCheckBox;
    cbResizeGrid: TCheckBox;
    cbComments: TCheckBox;
    edZoom: TEdit;
    Label1: TLabel;
    AdvGridExcelImport: TAdvGridExcelImport;
    BtnAbout: TButton;
    BtnImport: TButton;
    OpenDialog: TOpenDialog;
    AdvGridWorkbook1: TAdvGridWorkbook;
    procedure BtnAboutClick(Sender: TObject);
    procedure BtnImportClick(Sender: TObject);
    procedure edZoomChange(Sender: TObject);
  private
    procedure SetImportOptions;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FSimpleImport: TFSimpleImport;

implementation

{$R *.dfm}

procedure TFSimpleImport.BtnAboutClick(Sender: TObject);
begin
  ShowMessage('This example shows how to use TAdvGridExcelImport to import an Excel file into a grid');
end;

procedure TFSimpleImport.SetImportOptions;
begin
  AdvGridExcelImport.ImportOptions.Formulas := cbFormulas.Checked;
    AdvGridExcelImport.ImportOptions.Images := cbImages.Checked;
    AdvGridExcelImport.ImportOptions.Comments := cbComments.Checked;
    AdvGridExcelImport.ImportOptions.CellFormatting := cbFormatting.Checked;
    AdvGridExcelImport.ImportOptions.CellSizes := cbCellSizes.Checked;
    AdvGridExcelImport.ImportOptions.LockedCellsAsReadonly := cbReadOnlyToLocked.Checked;
    AdvGridExcelImport.ImportOptions.Html := cbHtml.Checked;
    AdvGridExcelImport.ImportOptions.PrintOptions := cbPrintOptions.Checked;
    AdvGridExcelImport.ImportOptions.Outlines := cbOutlines.Checked;
    AdvGridExcelImport.ImportOptions.ClearCells := cbClearCells.Checked;
    AdvGridExcelImport.ImportOptions.ResizeGrid := cbResizeGrid.Checked;
    AdvGridExcelImport.ImportOptions.ActiveSheet := cbActiveSheet.Checked;

    AdvGridExcelImport.Zoom := StrToInt(edZoom.Text);
end;

procedure TFSimpleImport.BtnImportClick(Sender: TObject);
begin
  if not OpenDialog.Execute then exit;

  SetImportOptions;
  AdvGridExcelImport.Import(OpenDialog.FileName);
end;

procedure TFSimpleImport.edZoomChange(Sender: TObject);
var
  v: integer;
begin
  if not TryStrToInt(edZoom.Text, v) then
  begin
     ShowMessage('Zoom must be an integer number');
     exit;
  end;

  if (v <> 0) and ((v < 10) or (v > 400)) then
  begin
    ShowMessage('Zoom must be 0, or a number between 10 and 400');
    exit;
  end;

end;

end.
