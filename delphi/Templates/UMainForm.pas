unit UMainForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, BaseGrid, AdvGrid, ShellAPI,
  {$if CompilerVersion >= 23.0} System.UITypes, {$IFEND}
  UAdvGridExcelExport, AdvObj, FlexCel.VCLSupport, FlexCel.Core, FlexCel.XlsAdapter;

type
  TMainForm = class(TForm)
    AdvStringGrid1: TAdvStringGrid;
    Panel1: TPanel;
    Button1: TButton;
    Panel2: TPanel;
    AdvGridExcelExport1: TAdvGridExcelExport;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure AdvGridExcelExport1ExportCell(Sender: TObject;
      var Args: TExportCellEventArgs);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.AdvGridExcelExport1ExportCell(Sender: TObject;
  var Args: TExportCellEventArgs);
var
  fm: TFlxFormat;
begin
  if (Args.GridCol = 3) then
  begin
    fm := Args.CellFormat;
    fm.Format:='##.00';
    Args.CellFormat := fm;
  end;
end;

procedure TMainForm.Button1Click(Sender: TObject);
var
  Path: string;
  Xls: TXlsFile;
begin
  Path:= ExtractFilePath(Application.ExeName) + '..\..\';

  {NOTE: This example uses a "demoTemplate.xls" file that is used as a starting file.
         If you want to embed demoTemplate in your exe (in order to not distribute
		 extra files) you could embed it as a resource into the exe.
	}

  //Open the template.
  Xls := TXlsFile.Create(Path + 'demotemplate.xls', true); //Open the template into an XlsFile object that we will use to export the grid.
  try

    //Export into a an existing sheet, moving rows down.
    Xls.ActiveSheetByName := 'Rows Down';
    AdvGridExcelExport1.LocationOptions.XlsStartRow :=9;
    AdvGridExcelExport1.Export(Xls, TInsertInSheet.InsertRows);

    //Export into a an existing sheet, moving rows down and deleting a row, to be able to use the chart.
    xls.ActiveSheetByName := 'With Chart';
    AdvGridExcelExport1.LocationOptions.XlsStartRow :=11;
    AdvGridExcelExport1.Export(Xls, TInsertInSheet.InsertRowsExceptFirstAndSecond);

    //Finally save the file. We could do extra manipulation here.
    Xls.Save(Path + 'result.xls');
  finally
    Xls.Free;
  end;

  if MessageDlg('Do you want to open the generated file?', mtConfirmation, mbOKCancel, 0) = mrOk then
   ShellExecute(0,'open', PCHAR(Path + 'result.xls'), NIL,NIL, SW_SHOW);
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  AdvStringGrid1.SaveFixedCells := false;
  AdvStringGrid1.LoadFromCSV('..\..\populations.txt');
  AdvStringGrid1.AutoSizeColumns(false, 10);
end;

end.
