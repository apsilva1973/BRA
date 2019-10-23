unit about;

interface

uses Windows, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, ToolWin, ComCtrls;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    OKButton: TButton;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Comments: TLabel;
    Memo1: TMemo;
    Timer1: TTimer;
    procedure OKButtonClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutBox: TAboutBox;
implementation

{$R *.dfm}

procedure TAboutBox.OKButtonClick(Sender: TObject);
begin
   Timer1.Enabled := False;
   Self.Close;
end;

procedure TAboutBox.Timer1Timer(Sender: TObject);
begin
     Memo1.Text := '';
     Memo1.lines.Add('Sistema de Honorarios - Bradesco');
     Sleep(300);
     Memo1.lines.Add('');
     Memo1.lines.Add('Digita Honorários - versão 11.0.002(a).');
     Memo1.lines.Add('Aplicativo voltado para a área juridica do banco');
     Memo1.lines.Add('afim de calcular e validar os honorarios dos processos,');
     Memo1.lines.Add('junto aos escritorios de advocacia.');
     Sleep(300);
     Memo1.lines.Add('---------------------------------------------------------------');
     Memo1.lines.Add('Agradecimentos a Equipe Presenta Sistemas.');

     Sleep(300);
     Memo1.lines.Add('---------------------------------------------------------------');
     Sleep(300);
     Memo1.lines.Add('Developer Master  - Maria do Carmo');
     Sleep(300);
     Memo1.lines.Add('Developer Partner - Alexandre Pereira');
     Sleep(300);
     Memo1.lines.Add('---------------------------------------------------------------');
     Sleep(300);
     Memo1.lines.Add('');
     Memo1.lines.Add('');
     Memo1.lines.Add('');
     Memo1.lines.Add('');
     Memo1.lines.Add('');
     Memo1.lines.Add('');
     Memo1.lines.Add('');
     Memo1.lines.Add('');

end;

procedure TAboutBox.FormShow(Sender: TObject);
begin
 Memo1.Text := '';
 Timer1.Enabled := True;
end;

end.

