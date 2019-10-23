unit fAtosDigitados;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids, DBGrids, datamodule_honorarios;

type
  TfrmAdosDigitados = class(TForm)
    DBGrid1: TDBGrid;
    lblNumeroNota: TLabel;
    BitBtn1: TBitBtn;
    BitBtn3: TBitBtn;
    Label2: TLabel;
    totalDigitado: TEdit;
    BitBtn2: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    Falterado: boolean;
    FvalorDigitado: double;
    Fnumeronota: string;
    procedure Setalterado(const Value: boolean);
    procedure SetvalorDigitado(const Value: double);
    procedure Setnumeronota(const Value: string);
    { Private declarations }
  public
    { Public declarations }
    property alterado : boolean read Falterado write Setalterado;
    property valorDigitado : double read FvalorDigitado write SetvalorDigitado;
    property numeronota : string read Fnumeronota write Setnumeronota;
  end;

var
  frmAdosDigitados: TfrmAdosDigitados;

implementation

{$R *.dfm}

{ TfrmAdosDigitados }

procedure TfrmAdosDigitados.Setalterado(const Value: boolean);
begin
  Falterado := Value;
end;

procedure TfrmAdosDigitados.FormCreate(Sender: TObject);
begin
   alterado := false;
end;

procedure TfrmAdosDigitados.BitBtn1Click(Sender: TObject);
begin
   if MessageDlg('Confirma marcar ato ' + dmHonorarios.dtsAtosDigitados.FieldByName('gcpj').AsString + ' como NÃO DIGITADO? Após esta alteração o valor total digitado da nota será de R$' + FormatFloat('#,##0.00', (valorDigitado - dmHonorarios.dtsAtosDigitados.FieldByName('valorcorrigido').AsFloat)),
                 mtConfirmation, [mbYes, mbno], 0) <> mrYes then
      exit;

   dmHonorarios.adoCmd.CommandText := 'update tbplanilhasatos set fgdigitado=0, datahoradigitacao=null ' +
                                      'where idescritorio = ' + dmHonorarios.dtsAtosDigitados.FieldByName('idescritorio').AsString + ' and ' +
                                      'idplanilha = ' + dmHonorarios.dtsAtosDigitados.FieldByName('idplanilha').AsString + ' and ' +
                                      'linhaplanilha = ' + dmHonorarios.dtsAtosDigitados.FieldByName('linhaplanilha').AsString + ' and ' +
                                      'sequencia = ' + dmHonorarios.dtsAtosDigitados.FieldByName('sequencia').AsString;
   dmHonorarios.adoCmd.Execute;

   dmHonorarios.dtsAtosDigitados.Close;
   dmHonorarios.dtsAtosDigitados.Open;


   dmHonorarios.dts.Close;
   dmHonorarios.dts.CommandText := 'select sum(valorcorrigido) as valordigitado ' +
                                   'from tbplanilhasatos where numeronota = ' + numeronota + ' and fgdigitado = 1 ';
   dmHonorarios.dts.Open;

   valorDigitado := dmHonorarios.dts.FieldByName('valordigitado').AsFloat;
   totalDigitado.Text  := FormatFloat('#,##0.00',dmHonorarios.dts.FieldByName('valordigitado').AsFloat);

   alterado := true;
end;

procedure TfrmAdosDigitados.BitBtn2Click(Sender: TObject);
begin
   if MessageDlg('Confirma marcar todos os atos da nota como NÃO DIGITADOS? Após esta alteração o valor total digitado da nota será de R$0,00',
                 mtConfirmation, [mbYes, mbno], 0) <> mrYes then
      exit;

   dmHonorarios.adoCmd.CommandText := 'update tbplanilhasatos set fgdigitado=0, datahoradigitacao=null ' +
                                      'where idescritorio = ' + dmHonorarios.dtsAtosDigitados.FieldByName('idescritorio').AsString + ' and ' +
                                      'idplanilha = ' + dmHonorarios.dtsAtosDigitados.FieldByName('idplanilha').AsString + ' and ' +
                                      'numeronota = ' + numeroNota;
   dmHonorarios.adoCmd.Execute;

   dmHonorarios.dtsAtosDigitados.Close;
   dmHonorarios.dtsAtosDigitados.Open;


   dmHonorarios.dts.Close;
   dmHonorarios.dts.CommandText := 'select sum(valorcorrigido) as valordigitado ' +
                                   'from tbplanilhasatos where numeronota = ' + numeronota + ' and fgdigitado = 1 ';
   dmHonorarios.dts.Open;

   valorDigitado := dmHonorarios.dts.FieldByName('valordigitado').AsFloat;
   totalDigitado.Text  := FormatFloat('#,##0.00',dmHonorarios.dts.FieldByName('valordigitado').AsFloat);

   alterado := true;
end;

procedure TfrmAdosDigitados.BitBtn3Click(Sender: TObject);
begin
      close;ModalResult := mrOk;
end;

procedure TfrmAdosDigitados.SetvalorDigitado(const Value: double);
begin
  FvalorDigitado := Value;
end;

procedure TfrmAdosDigitados.Setnumeronota(const Value: string);
begin
  Fnumeronota := Value;
end;

end.
