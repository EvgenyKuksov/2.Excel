unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, ComObj, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Excel_TLB, VBIDE_TLB,
  Math, Graph_TLB, Vcl.ExtDlgs, Vcl.ExtCtrls, Vcl.Imaging.jpeg;

type
  TForm1 = class(TForm)
    Image1: TImage;
    Button1: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    {Private declarations}
  public
    {Public declarations}
  end;

var
  Form1: TForm1;
  ExcelApp: ExcelApplication;
  Sheet: ExcelWorksheet;
  mchart: ExcelChart;
  mshape: Shape;
  Col: Char;
  Row: Integer;
  mAxis:Axis;
  MyDisp: IDispatch;
  y, x, Xn, Xk, Shag: Extended;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
  Xn:=StrToFloat(Edit1.Text);
  Xk:=StrToFloat(Edit2.Text);
  Shag:=StrToFloat(Edit3.Text);

//  Проверка диапазона
if Xn < Xk then
  begin
    ExcelApp := CreateOleObject('Excel.Application') as ExcelApplication; //Создаие документа Excel
    ExcelApp.Workbooks.add(xlWBatWorkSheet,0);  //Создание рабочего листа
    ExcelApp.Visible[0] := True;  //Открытие документа
    Sheet := ExcelApp.Workbooks[1].WorkSheets[1] as ExcelWorksheet;
    ExcelApp.Application.ReferenceStyle[0] := xlA1;


    //  Задаю значения "X"
    col:='A';
    x:=Xn;
    Sheet.Range[col+'1', col+'1'].Value[xlRangeValueDefault]:='x';
    row:=2;
    while (x<=Xk) and (x>=Xn) do
      begin
        Sheet.Range[col+IntToStr(row), col+IntToStr(row)].Value[xlRangeValueDefault]:=x;
        x:=x+Shag;
        row:=row+1;
      end;

  //  Задаю значения "Y"
    col:='B';
    x:=Xn;
    Sheet.Range[col+'1', col+'1'].Value[xlRangeValueDefault]:='y';
    row:=2;
    while (x<=Xk) and (x>=Xn) do
    begin
      if x>=1 then y:=power(cos(x),4);
      if x<=-1 then y:=power(cos(x),5);
      if (x>-2) and (x<2) then y:=1;
      Sheet.Range[col+IntToStr(row), col+IntToStr(row)].Value[xlRangeValueDefault]:=y;
      x:=x+Shag;
      row:=row+1;
    end;

  sheet.Range['A2','B'+inttostr(row)].Select;
  mshape:=Sheet.Shapes.AddChart(xlXYScatterSmoothNoMarkers,250,1,800,800);
  mchart:=(mshape.Chart as ExcelChart).Location(xlLocationAsNewSheet,EmptyParam);
  ExcelApp.Application.ActiveWorkbook.ActiveChart.SetElement(1);
  ExcelApp.Application.ActiveWorkbook.ActiveChart.ChartTitle[0].Text:='График функции';
  MyDisp:=mchart.Axes(xlValue, xlPrimary, 0);
  ExcelApp.Application.ActiveWorkbook.ActiveChart.Legend[0].Delete; //удаление легенды


     mAxis:=Axis(MyDisp);
     mAxis.HasTitle:=True;
     mAxis.AxisTitle.Caption:='х';

  MyDisp:=mchart.Axes(xlCategory, xlPrimary, 0);

     mAxis:=Axis(MyDisp);
     mAxis.HasTitle:=True;
     mAxis.AxisTitle.Caption:='Y';

  ExcelApp.Application.ActiveWorkbook.ActiveChart.SetElement(328);
  end
  else ShowMessage('Соблюдайте условие: "Xn < Xk"');
end;

end.
