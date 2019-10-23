unit fcadescritorio;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, inifiles, calccpfcgc, DBCtrls, Mask,
  ExtCtrls, Grids, DBGrids, datamodule_honorarios;

type
  TfrmCadEscritorio = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    DBGrid1: TDBGrid;
    Label4: TLabel;
    DBNavigator1: TDBNavigator;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    Label5: TLabel;
    DBEdit4: TDBEdit;
    DBCheckBox1: TDBCheckBox;
    DBCheckBox2: TDBCheckBox;
    DBEdit5: TDBEdit;
    Label6: TLabel;
    Bevel1: TBevel;
    BitBtn1: TBitBtn;
    nomePesq: TEdit;
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure nomePesqChange(Sender: TObject);
    procedure DBNavigator1Click(Sender: TObject; Button: TNavigateBtn);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCadEscritorio: TfrmCadEscritorio;

implementation

{$R *.dfm}

procedure TfrmCadEscritorio.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfrmCadEscritorio.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   Action := cafree;
end;


procedure TfrmCadEscritorio.FormCreate(Sender: TObject);
begin
   dmHonorarios.ObtemCadastroEscritorios('');
end;

procedure TfrmCadEscritorio.BitBtn1Click(Sender: TObject);
begin
   ModalResult := mrOk;
end;

procedure TfrmCadEscritorio.nomePesqChange(Sender: TObject);
begin
   dmHonorarios.ObtemCadastroEscritorios(nomePesq.text);
end;

procedure TfrmCadEscritorio.DBNavigator1Click(Sender: TObject;
  Button: TNavigateBtn);
begin
   if Button = nbInsert then
      DBEdit1.setfocus;
end;

end.
