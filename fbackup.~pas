unit fbackup;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,StdCtrls, Spin, ComCtrls, Gauges, ExtCtrls, Buttons,GifImage,Presenta,untThread;





type
  TfrmBackup = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Button1: TButton;
    Button2: TButton;
    Gauge1: TGauge;
    SpinEdit1: TSpinEdit;
    Image1: TImage;
    Label4: TLabel;
    ProgressBar1: TProgressBar;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure executabackup(users: string;tempo:integer;dataatual:Tdatetime;FormBackup:TForm);
    procedure Mostra;
    { Public declarations }
  end;


type
   TMyThread = class(TThread)
 private
  procedure Execute; override;
 public
    Constructor CreateT(FormChama: TForm);
 end;



var
  frmBackup: TfrmBackup;
  Thread1: TMinhaThread;
   lflag   : boolean;
implementation

uses main, datamodule_honorarios;

{$R *.dfm}

procedure TfrmBackup.FormCreate(Sender: TObject);
begin
// Animate1.Active := true;
 Label1.Caption :=  Label1.Caption +': honorarios_'+MainForm.usuario;
 SpinEdit1.Value := 3;

 Thread1 := TMinhaThread.Create;
 Thread1.FreeOnTerminate := True;
 Thread1.ProgressBar     := ProgressBar1;
 Thread1.Form            := frmBackup;
end;

procedure TfrmBackup.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 action := caFree;
end;

procedure TfrmBackup.Button2Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmBackup.Button1Click(Sender: TObject);
begin


   if MessageDlg('Confirma o backup da  base de dados honorarios_'+ MainForm.usuario + '?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    Begin


    if Thread1.Suspended then
    begin
      Thread1.Tempo := 10;
      Thread1.Resume;
//      btnThread1.Caption := 'Pausar';
    end
    // senão, a Thread é pausada
    else
    begin
      Thread1.Suspend;
//      btnThread1.Caption := 'Continuar';
    end;


      //Image1.Picture.LoadFromFile('C:\Presenta\projeto\delphi\utils\bradesco.honorarios.digitador\V11.0\progs\load.gif');
      Gauge1.Visible := True;
      Label4.Caption := 'Aguarde..';
//      MyThread.Execute;

      Gauge1.Progress  := 45;
      SpinEdit1.Visible := False;
      executabackup(MainForm.usuario,SpinEdit1.Value,Date,frmBackup);

      Gauge1.Progress  := 100;
      Gauge1.Visible := False;
      Image1.Picture.LoadFromFile('');
      Label4.Caption := '';
      SpinEdit1.Visible := True;
      MessageDlg('Backup efetuado com sucesso!', mtCustom, [mbOk], 0);
      Thread1.Suspend;

    end;


end;


procedure TfrmBackup.executabackup(users: string;tempo:integer;dataatual:Tdatetime;FormBackup:TForm);
begin

  dmHonorarios.executabackup(users,tempo,dataatual,frmBackup);

end;



procedure TfrmBackup.Button3Click(Sender: TObject);
begin
 Image1.Picture.LoadFromFile('C:\Presenta\projeto\delphi\utils\bradesco.honorarios.digitador\V11.0\progs\loading.gif');
end;


procedure TMyThread.Execute;

begin
//  for LI:=1 to 2 do
//   begin
    //I:=LI;
//    frmBackup.Mostra;
    { Inicia a thread }
//    lflag := True;
    While lflag do
      Application.ProcessMessages;

//   end;
end; { of procedure }

Constructor TMyThread.CreateT(FormChama: TForm);
Begin
{ Executa Heranca Criando a Thread Suspença }
Inherited Create(true);
//Priority := tpHighest;//  tpNormal;
//FreeOnTerminate := True;
//Painel := pnlDisp; // Painel do Form que chama a thread que recebe o
//frmBackup :=   FormChama; // Painel do Form que chama a thread que recebe o
// Resultado da thread

End;


procedure TfrmBackup.FormShow(Sender: TObject);
begin
  lflag := true;
//  MyThread := TMyThread.CreateT(Self);

//  MyThread.Suspended := False;
//  MyThread.Execute;
  Label4.Caption := '';
end;



procedure TfrmBackup.Mostra;
begin
//  frmBackup.Image1.Picture.LoadFromFile('C:\Presenta\projeto\delphi\utils\bradesco.honorarios.digitador\V11.0\progs\load.gif');
//  Application.ProcessMessages;
end;



procedure TfrmBackup.SpeedButton1Click(Sender: TObject);
Var
 cmd : String;
begin
WinExecAndWait32V2('C:\Presenta\dados\v60.013\kdb\Progs\executa_sql.exe', SW_SHOWMAXIMIZED, 30)
end;

end.
