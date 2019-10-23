unit fconfigDb;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, inifiles;

type
  TfrmConfigDb = class(TForm)
    Label1: TLabel;
    Edit1: TEdit;
    SpeedButton1: TSpeedButton;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    od: TOpenDialog;
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmConfigDb: TfrmConfigDb;

implementation

{$R *.dfm}

procedure TfrmConfigDb.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfrmConfigDb.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   Action := cafree;
end;

procedure TfrmConfigDb.SpeedButton1Click(Sender: TObject);
begin
   if od.Execute then
      Edit1.Text := od.FileName;
end;

procedure TfrmConfigDb.BitBtn1Click(Sender: TObject);
var
   ini : TInifile;
begin
   if edit1.Text = '' then
   begin
      ShowMessage('Informe o nome do banco de dados');
      edit1.SetFocus;
      exit;
   end;

   if Not FileExists(Edit1.text) then
   begin
      ShowMessage('Arquivo não encontrado');
      edit1.SetFocus;
      exit;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'config.ini');
   try
      ini.WriteString('mdb', 'caminho', edit1.text);
   finally
      ini.free;
   end;
   close;
end;

procedure TfrmConfigDb.FormCreate(Sender: TObject);
var
   ini : Tinifile;
begin
   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'config.ini');
   try
      edit1.Text := ini.ReadString('mdb', 'caminho', '');
   finally
      ini.free;
   end;

end;

end.
