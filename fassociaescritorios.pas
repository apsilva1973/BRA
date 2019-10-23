unit fassociaescritorios;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, datamodule_honorarios, Buttons, Grids, DBGrids;

type
  TfrmAssociaEscritorios = class(TForm)
    Label1: TLabel;
    tpNaoAtualizar: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    DBGrid1: TDBGrid;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    DBGrid2: TDBGrid;
    BitBtn1: TBitBtn;
    pesqEscritorio: TEdit;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure pesqEscritorioChange(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure DBGrid2DblClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAssociaEscritorios: TfrmAssociaEscritorios;

implementation

{$R *.dfm}

procedure TfrmAssociaEscritorios.SpeedButton1Click(Sender: TObject);
var
   bMark : TBookmarkList;
   i : integer;
begin
   bMark := dbGrid2.SelectedRows;

   for i := 0 to bMark.Count - 1 do
   begin
      Application.ProcessMessages;
      dbGrid2.DataSource.DataSet.BookMark := bMark.items[i];

      dmHonorarios.MarcaEscritorioIn(StrToInt(tpNAoAtualizar.text), dmHonorarios.dtsEscritoriosOut.FieldByName('idescritorio').AsInteger);
   end;

   dmHonorarios.ObtemEscritoriosAssociadosAoTipo(StrToInt(tpNAoAtualizar.text));
   dmHonorarios.ObtemEscritoriosNaoAssociadosAoTipo(StrToInt(tpNAoAtualizar.text), '');
end;

procedure TfrmAssociaEscritorios.SpeedButton2Click(Sender: TObject);
begin
   DBGrid2.DataSource.DataSet.DisableControls;
   try
      DBGrid2.DataSource.DataSet.first;
      while not DBGrid2.DataSource.DataSet.eof do
      begin
         dbGrid2.SelectedRows.CurrentRowSelected := true;
         DBGrid2.DataSource.DataSet.next;
      end;
   finally
      DBGrid2.DataSource.DataSet.EnableControls;
   end;
   SpeedButton1Click(Sender);
end;

procedure TfrmAssociaEscritorios.SpeedButton3Click(Sender: TObject);
var
   bMark : TBookmarkList;
   i : integer;
begin
   bMark := dbGrid1.SelectedRows;

   for i := 0 to bMark.Count - 1 do
   begin
      Application.ProcessMessages;
      dbGrid1.DataSource.DataSet.BookMark := bMark.items[i];

      dmHonorarios.MarcaEscritorioOut(StrToInt(tpNAoAtualizar.text),dmHonorarios.dtsEscritoriosIn.FieldByName('idescritorio').AsInteger);
   end;

   dmHonorarios.ObtemEscritoriosAssociadosAoTipo(StrToInt(tpNAoAtualizar.text));
   dmHonorarios.ObtemEscritoriosNaoAssociadosAoTipo(StrToInt(tpNAoAtualizar.text), '');
end;

procedure TfrmAssociaEscritorios.SpeedButton4Click(Sender: TObject);
begin
   DBGrid1.DataSource.DataSet.DisableControls;
   try
      DBGrid1.DataSource.DataSet.first;
      while not DBGrid1.DataSource.DataSet.eof do
      begin
         dbGrid1.SelectedRows.CurrentRowSelected := true;
         DBGrid1.DataSource.DataSet.next;
      end;
   finally
      DBGrid1.DataSource.DataSet.EnableControls;
   end;
   SpeedButton3Click(Sender);
end;

procedure TfrmAssociaEscritorios.pesqEscritorioChange(Sender: TObject);
begin
   dmHonorarios.ObtemEscritoriosNaoAssociadosAoTipo(StrToInt(tpNaoAtualizar.Text), pesqEscritorio.text);
end;

procedure TfrmAssociaEscritorios.DBGrid1DblClick(Sender: TObject);
begin
   SpeedButton3Click(Sender);

end;

procedure TfrmAssociaEscritorios.DBGrid2DblClick(Sender: TObject);
begin
   SpeedButton1Click(Sender);
end;

end.
