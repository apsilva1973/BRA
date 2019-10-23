unit fAtosPendentes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, datamodule_honorarios, StdCtrls, Buttons, Grids, DBGrids;

type
  TfrmAtosPendentes = class(TForm)
    Label1: TLabel;
    DBGrid1: TDBGrid;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label2: TLabel;
    BitBtn3: TBitBtn;
    Edit1: TEdit;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAtosPendentes: TfrmAtosPendentes;

implementation

{$R *.dfm}

procedure TfrmAtosPendentes.BitBtn1Click(Sender: TObject);
var
   numNota, idescritorio, idplanilha, linhaplanilha, sequencia : integer;
   valCorrigido : double;
begin
   if MessageDlg('Confirma a exclusão ato ' + dmHonorarios.dtsAtosPendentes.FieldByName('gcpj').AsString + ' ?',
                 mtConfirmation, [mbYes, mbno], 0) <> mrYes then
      exit;

   numNota := dmHonorarios.dtsAtosPendentes.FieldByName('numeronota').AsInteger;
   valCorrigido := dmHonorarios.dtsAtosPendentes.FieldByName('valorcorrigido').AsFloat;
   idescritorio := dmHonorarios.dtsAtosPendentes.FieldByName('idescritorio').AsInteger;
   idplanilha := dmHonorarios.dtsAtosPendentes.FieldByName('idplanilha').AsInteger;
   linhaplanilha := dmHonorarios.dtsAtosPendentes.FieldByName('linhaplanilha').AsInteger;
   sequencia := dmHonorarios.dtsAtosPendentes.FieldByName('sequencia').AsInteger;

   dmHonorarios.adoCmd.Parameters.Clear;
   dmHonorarios.adoCmd.CommandText := 'delete from tbplanilhasatos ' +
                                      'where idescritorio = ' + IntToStr(idescritorio) + ' and ' +
                                      'idplanilha = ' + IntToStr(idplanilha)  + ' and ' +
                                      'linhaplanilha = ' + IntToStr(linhaplanilha) + ' and ' +
                                      'sequencia = ' + IntToStr(sequencia);
   dmHonorarios.adoCmd.Execute;

   dmHonorarios.dtsAtosPendentes.Close;
   dmHonorarios.dtsAtosPendentes.Open;


   if numNota <> 0 then
   begin
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select * from tbcontrolenotas where idescritorio = ' + IntToStr(idescritorio)  + ' and ' +
                                      'numeronota = ' + IntToStr(numNota);
      dmHonorarios.dts.Open;

      dmHonorarios.adoCmd.Parameters.Clear;
      dmHonorarios.adoCmd.CommandText := 'update tbcontrolenotas set totaldeatos = ' + IntToStr(dmHonorarios.dts.FieldByName('totaldeatos').AsInteger - 1) + ', valortotal = :valor ' +
                                         'where idescritorio = ' + IntToStr(idescritorio)  + ' and ' +
                                         'numeronota = ' + IntToStr(numNota);

      dmHonorarios.adoCmd.Parameters.ParamByName('valor').Value := (dmHonorarios.dts.FieldByName('valortotal').AsFloat - valCorrigido);
      dmHonorarios.adoCmd.Execute;
   end;
end;

procedure TfrmAtosPendentes.BitBtn2Click(Sender: TObject);
begin
   if Edit1.Text = '' then
   begin
      ShowMessage('Informe o motivo da inconsistência');
      edit1.setFocus;
      exit;
   end;

   if MessageDlg('Deseja realmente marcar o ato ' + dmHonorarios.dtsAtosPendentes.FieldByName('gcpj').AsString + ' como inconsitente?',
                 mtConfirmation, [mbYes, mbno], 0) <> mrYes then
      exit;

   dmHonorarios.GravaOcorrencia(dmHonorarios.dtsAtosPendentes.FieldByName('idplanilha').AsInteger,
                                dmHonorarios.dtsAtosPendentes.FieldByName('linhaplanilha').AsInteger,
                                2,
                                edit1.Text);

   dmHonorarios.MarcaProcessoCruzadoGcpj(9, dmHonorarios.dtsAtosPendentes.FieldByName('idescritorio').AsInteger,
                                         dmHonorarios.dtsAtosPendentes.FieldByName('idplanilha').AsInteger,
                                         dmHonorarios.dtsAtosPendentes.FieldByName('anomesreferencia').AsString,
                                         dmHonorarios.dtsAtosPendentes.FieldByName('sequencia').AsInteger,
                                         dmHonorarios.dtsAtosPendentes.FieldByName('linhaplanilha').AsInteger);
   dmHonorarios.dtsAtosPendentes.Close;
   dmHonorarios.dtsAtosPendentes.Open;
end;

end.
