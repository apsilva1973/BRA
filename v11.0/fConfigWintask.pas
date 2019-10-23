unit fConfigWintask;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, maindir, datamodule_honorarios;

type
  TfConfiguraWintask = class(TForm)
    Label1: TLabel;
    Edit1: TEdit;
    SpeedButton1: TSpeedButton;
    Label2: TLabel;
    Edit2: TEdit;
    SpeedButton2: TSpeedButton;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    od: TOpenDialog;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fConfiguraWintask: TfConfiguraWintask;

implementation

{$R *.dfm}

procedure TfConfiguraWintask.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   Action := caFree;
end;

procedure TfConfiguraWintask.BitBtn1Click(Sender: TObject);
begin
   if edit1.Text = '' then
   begin
      ShowMessage('Informe o caminho de instalação do Wintask');
      Edit1.SetFocus;
      exit;
   end;

   if edit2.Text = '' then
   begin
      ShowMessage('Informe o diretórios dos arquivos .ROB');
      Edit2.setfocus;
      exit;
   end;

   if Not FileExists(edit1.text) then
   begin
      ShowMessage('Arquivo não encontrado');
      edit1.SetFocus;
      exit;
   end;

   if Not DirectoryExists(edit2.text) then
   begin
      ShowMessage('Diretório não encontrado');
      edit1.SetFocus;
      exit;
   end;
   dmHonorarios.SalvaConfiguracaoWintask(Edit1.Text, edit2.Text);
   close;
end;

procedure TfConfiguraWintask.SpeedButton1Click(Sender: TObject);
begin
   if od.Execute then
      Edit1.Text := od.FileName;
end;

procedure TfConfiguraWintask.SpeedButton2Click(Sender: TObject);
begin
   frmDir := TfrmDir.Create(nil);
   try
      if frmDir.ShowModal = mrOk then
      begin
         Edit2.Text := frmDir.edNomeDir.Text;
      end;
   finally
      frmDir.free;
   end;
end;

procedure TfConfiguraWintask.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfConfiguraWintask.FormCreate(Sender: TObject);
begin
   dmHonorarios.ObtemConfiguracaoWintask;
   if Not dmHonorarios.dts.eof then
   begin
      Edit1.Text := dmHonorarios.dts.FieldByName('diretorioExe').AsString;
      Edit2.Text := dmHonorarios.dts.FieldByName('diretorioRob').AsString;
   end;
end;

end.
