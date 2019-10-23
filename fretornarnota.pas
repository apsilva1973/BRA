unit fretornarnota;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, datamodule_honorarios;

type
  TfrmRetornarFinalizada = class(TForm)
    Label1: TLabel;
    numNota: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRetornarFinalizada: TfrmRetornarFinalizada;

implementation

{$R *.dfm}

procedure TfrmRetornarFinalizada.BitBtn1Click(Sender: TObject);
begin
   if numNota.Text = '' then
   begin
      ShowMessage('Informe o número danota');
      exit;
   end;

   case numNota.tag of
      0 : begin
         if messageDlg('Confirma o retorno da nota: ' + numNota.Text + '?',mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
            exit;

         dmHonorarios.adoCmd.CommandText := 'update tbcontrolenotas set fgstatus=0 where numeronota = ' + numNota.Text;
         dmHonorarios.adoCmd.Execute;
      end;
      1 : begin
         if messageDlg('Deseja realmente marcar a nota ' + numNota.Text + ' como finalizada?',mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
            exit;

         dmHonorarios.adoCmd.CommandText := 'update tbcontrolenotas set fgstatus=9 where numeronota = ' + numNota.Text;
         dmHonorarios.adoCmd.Execute;

      end;
   end;

   ModalResult := mrOk;
end;

end.
