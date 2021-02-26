unit UOnDemandFormatting;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, UAdvGridExcelExport, StdCtrls,
  ExtCtrls, Grids, AdvObj, BaseGrid, AdvGrid, ShellAPI,
  {$if CompilerVersion >= 23.0} System.UITypes, {$IFEND}
  VCL.FlexCel.Core, FlexCel.XlsAdapter, AdvUtil;

type
  TFMainForm = class(TForm)
    AdvGridExcelExport1: TAdvGridExcelExport;
    AdvStringGrid1: TAdvStringGrid;
    Panel1: TPanel;
    btnExportFtExcel: TButton;
    btnAboutFtExcel: TButton;
    SaveDialog: TSaveDialog;
    btnExportAsStringsFtExcel: TButton;
    Splitter1: TSplitter;
    Panel2: TPanel;
    AdvStringGrid2: TAdvStringGrid;
    AdvGridExcelExport2: TAdvGridExcelExport;
    PanelButton1: TPanel;
    btnExport: TButton;
    btnAbout: TButton;
    btnExportAsStrings: TButton;
    procedure AdvGridExcelExport1ExportCell(Sender: TObject;
      var Args: TExportCellEventArgs);
    procedure AdvGridExcelExport1ExportComment(Sender: TObject;
      var Args: TExportCommentEventArgs);
    procedure btnAboutFtExcelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure AdvStringGrid1GetFloatFormat(Sender: TObject; ACol, ARow: Integer;
      var IsFloat: Boolean; var FloatFormat: string);
    procedure btnExportFtExcelClick(Sender: TObject);
    procedure btnExportAsStringsFtExcelClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure btnExportAsStringsClick(Sender: TObject);
    procedure btnAboutClick(Sender: TObject);
    procedure AdvStringGrid2GetFloatFormat(Sender: TObject; ACol, ARow: Integer;
      var IsFloat: Boolean; var FloatFormat: string);
  private
    procedure DoExport(const AdvGridExcelExport: TAdvGridExcelExport; const AsString: boolean);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FMainForm: TFMainForm;

implementation

{$R *.dfm}

procedure TFMainForm.AdvGridExcelExport1ExportCell(Sender: TObject;
  var Args: TExportCellEventArgs);
var
  Fm: TFlxFormat;
begin
  //We can't modify Args.CellFormat directly
  //So we assign it to a variable.
  Fm := Args.CellFormat;
  if Args.GridRow = 1 then Fm.Format := '$ 0.000'
  else if Args.GridRow = 2 then Fm.Format := '"Pi is "0.000';

  Args.CellFormat := Fm;
end;

procedure TFMainForm.AdvGridExcelExport1ExportComment(Sender: TObject;
  var Args: TExportCommentEventArgs);
var
  Anchor: TClientAnchor;
begin
 if Args.GridRow = 3 then
 begin
   Args.CommentProperties.ShapeFill := TShapeFill_Create(true, TSolidFill_Create(Colors.Red));

   //Make the box a little bigger
   Anchor := Args.CommentProperties.Anchor;
   Args.CommentProperties.Anchor := TClientAnchor.Create(Anchor.AnchorType, Anchor.Row1, Anchor.Dy1, Anchor.Col1, Anchor.Dx1, Anchor.Row2 + 1, Anchor.Dy2, Anchor.Col2, Anchor.Dx2);
   //We might had done instead: Args.CommentProperties.AutoSize := true; //but it won't resize the way we want.
 end;
end;

procedure TFMainForm.btnAboutClick(Sender: TObject);
begin
  ShowMessage('This example shows how to convert formatting between the Grid and Excel. '+
    'By default, Excel and AdvStringGrid use incompatible format strings, so you need to tell AdvExcelExport ' +
    'how the formats map.' +#10#10+ 'We also show how you can use an event to format a comment.');

end;

procedure TFMainForm.btnAboutFtExcelClick(Sender: TObject);
begin
  ShowMessage(
    'In this pane we show how to use the "FormatType" property in the grid ' +
    'to format with the same strings in both Excel and the grid.');
end;

procedure TFMainForm.btnExportAsStringsClick(Sender: TObject);
begin
  //When exporting as strings, numbers will display exactly as they do in the grid,
  //but they won't be "real" numbers, they will be strings. (that can't be used in formulas)
  DoExport(AdvGridExcelExport1, true);

end;

procedure TFMainForm.btnExportAsStringsFtExcelClick(Sender: TObject);
begin
  //When exporting as strings, numbers will display exactly as they do in the grid,
  //but they won't be "real" numbers, they will be strings. (that can't be used in formulas)
  DoExport(AdvGridExcelExport2, true);
end;


procedure TFMainForm.btnExportClick(Sender: TObject);
begin
  DoExport(AdvGridExcelExport1, false);
end;

procedure TFMainForm.btnExportFtExcelClick(Sender: TObject);
begin
  DoExport(AdvGridExcelExport2, false);
end;

procedure TFMainForm.DoExport(const AdvGridExcelExport: TAdvGridExcelExport; const AsString: boolean);
begin
  if not SaveDialog.Execute then exit;
  AdvGridExcelExport.ExportOptions.CellsAsStrings := AsString;
  AdvGridExcelExport.Export(SaveDialog.FileName);
  if MessageDlg('Do you want to open the generated file?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    ShellExecute(0, 'open', PCHAR(SaveDialog.FileName), nil, nil, SW_SHOWNORMAL);
  end;
end;


procedure TFMainForm.FormCreate(Sender: TObject);
begin
  //First grid will be formatted with standard Delphi formatting.
  AdvStringGrid1.FloatFormat := '%g';
  AdvStringGrid1.Floats[1, 1] := 2.7128;
  AdvStringGrid1.Floats[1, 2] := 3.1416;
  AdvStringGrid1.Cells[1,3] := '=A2 + 1'; //This formula won't work when exporting as strings.

  AdvStringGrid1.AddComment(1,2, 'This cell is formatted with delphi formatting strings');
  AdvStringGrid1.AddComment(1,3, 'This formula won''t work when exporting as strings, since you can''t add 1 to a string.');

  //Second Grid is formatted with ftExcelFormatting.
  //We can't modify formats here, we need to modify it in the OnGetFloatFormat event.
  AdvStringGrid2.FormatType := ftExcel;
  AdvStringGrid2.FloatFormat := '%g';
  AdvStringGrid2.Floats[1, 1] := 2.7128;
  AdvStringGrid2.Floats[1, 2] := 3.1416;
  AdvStringGrid2.Cells[1,3] := '=A2 + 1'; //This formula won't work when exporting as strings.

  AdvStringGrid2.AddComment(1,2, 'This was formatted with the same format string in both Excel and the grid.');
  AdvStringGrid2.AddComment(1,3, 'This formula won''t work when exporting as strings, since you can''t add 1 to a string.');

end;

procedure TFMainForm.AdvStringGrid1GetFloatFormat(Sender: TObject; ACol,
  ARow: Integer; var IsFloat: Boolean; var FloatFormat: string);
begin
  //Uses delphi formats, requires a different ExportCell event in the AdvGridExcelExport to convert to Excel formatting.
  if ARow = 1 then FloatFormat := '%.3m'
  else if ARow = 2 then FloatFormat := 'Pi is %.3f';

end;

procedure TFMainForm.AdvStringGrid2GetFloatFormat(Sender: TObject; ACol,
  ARow: Integer; var IsFloat: Boolean; var FloatFormat: string);
begin
  //Uses Excel formatting, will be converted without futher events to Excel.
  if ARow = 1 then FloatFormat := '$ 0.000'
  else if ARow = 2 then FloatFormat := '"Pi is "0.000';

end;

end.
