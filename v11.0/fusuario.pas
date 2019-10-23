unit fusuario;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls;

type
  TfrmUsuario = class(TForm)
    Label1: TLabel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    cbxUsuario: TComboBox;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmUsuario: TfrmUsuario;

implementation

{$R *.dfm}

procedure TfrmUsuario.BitBtn1Click(Sender: TObject);
begin
   if cbxUsuario.ItemIndex = -1 then
   begin
      ShowMessage('Selecione um usuário válido');
      ModalResult := mrNone;
      exit;
   end;

   if MessageDlg('Confirma o acesso à base de dados honorarios_' + cbxUsuario.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      ModalResult := mrNone;
end;

end.
