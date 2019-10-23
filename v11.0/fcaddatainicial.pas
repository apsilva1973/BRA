unit fcaddatainicial;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls;

type
  TfrmCadDataInicial = class(TForm)
    Label1: TLabel;
    dataInicial: TDateTimePicker;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCadDataInicial: TfrmCadDataInicial;

implementation

{$R *.dfm}

procedure TfrmCadDataInicial.BitBtn1Click(Sender: TObject);
begin
   if DayOfWeek(dataInicial.Date) in [1, 7] then
   begin
      ShowMessage('Data inicial não pode ser Sábado nem Domingo');
      dataInicial.SetFocus;
      ModalResult := mrNone;
      exit;
   end;
   ModalResult := mrOk;
end;

end.
