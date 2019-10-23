unit MAIN;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, Menus,
  StdCtrls, Dialogs, Buttons, Messages, ExtCtrls, ComCtrls, StdActns,
  ActnList, ToolWin, ImgList, fconfigwintask, fconfigdb, fimportaplan,
  datamodule_honorarios, presenta, dateutils, uplanjbm, fusuario, inifiles,
  fcaddatainicial, fcadadvogadointerno, fenviagcpj, uPresentaFunc;

type
  TMainForm = class(TForm)
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Help1: TMenuItem;
    N1: TMenuItem;
    FileExitItem: TMenuItem;
    HelpAboutItem: TMenuItem;
    OpenDialog: TOpenDialog;
    StatusBar: TStatusBar;
    ActionList1: TActionList;
    FileExit1: TAction;
    HelpAbout1: TAction;
    ToolBar2: TToolBar;
    ToolButton3: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    actConfigWintask: TAction;
    N2: TMenuItem;
    ToolButton1: TToolButton;
    Wintask1: TMenuItem;
    actConfigDB: TAction;
    ToolButton2: TToolButton;
    ToolButton4: TToolButton;
    BancodeDados1: TMenuItem;
    ToolButton5: TToolButton;
    N3: TMenuItem;
    actImportaPlanilha: TAction;
    ImportaPlanilha1: TMenuItem;
    actCadEscritorio: TAction;
    ToolButton6: TToolButton;
    actDigita: TAction;
    DigitaNotas1: TMenuItem;
    actValidaPlan: TAction;
    Validaodaplanilha1: TMenuItem;
    ToolButton8: TToolButton;
    actRelatorios: TAction;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    actRelatorios1: TMenuItem;
    lblUsu: TLabel;
    actConfiDataInicial: TAction;
    ToolButton11: TToolButton;
    DataInicial1: TMenuItem;
    actCadastraAdvogadoInerno: TAction;
    ToolButton12: TToolButton;
    actCadastraAdvogadoInerno1: TMenuItem;
    ToolButton13: TToolButton;
    actEnviaGcpj: TAction;
    EnviaGCPJ1: TMenuItem;
    procedure HelpAbout1Execute(Sender: TObject);
    procedure FileExit1Execute(Sender: TObject);
    procedure actConfigWintaskExecute(Sender: TObject);
    procedure actConfigDBExecute(Sender: TObject);
    procedure actImportaPlanilhaExecute(Sender: TObject);
    procedure actValidaPlanExecute(Sender: TObject);
    procedure actDigitaExecute(Sender: TObject);
    procedure actRelatoriosExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure actConfigDataInicial(Sender: TObject);
    procedure actCadastraAdvogadoInterno(Sender: TObject);
    procedure actEnviaGcpjExecute(Sender: TObject);

  private
    Fusuario: string;
    procedure Setusuario(const Value: string);
    { Private declarations }
  public
    { Public declarations }
    property usuario : string read Fusuario write Setusuario;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

uses about, fdigitaplan, frelatorios, fvalidaplan;


procedure TMainForm.HelpAbout1Execute(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TMainForm.FileExit1Execute(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.actConfigWintaskExecute(Sender: TObject);
begin
   //configura wintask
   fConfiguraWintask :=  TfConfiguraWintask.Create(nil);
end;

procedure TMainForm.actConfigDBExecute(Sender: TObject);
begin
   //configurar banco de dados
   frmConfigDb := TfrmConfigDb.Create(Nil);
end;

procedure TMainForm.actImportaPlanilhaExecute(Sender: TObject);
var
   child : TfrmImpPlan;
   Escritorio: TIdObject;
begin
   child := TfrmImpPlan.Create(nil);
   dmHonorarios.ObtemEscritoriosCadastrados(true);
   while  Not dmHonorarios.dts.Eof do
   begin
      Escritorio :=TIdObject.Create;
      Escritorio.id := dmHonorarios.dts.FieldByName('idescritorio').AsInteger;
//      child.cbEscritorios.Items.AddObject(dmHonorarios.dts.FieldByName('cnpjescritorio').AsString + '-' + dmHonorarios.dts.FieldByName('nomeescritorio').AsString, escritorio);
      child.cbEscritorios.Items.AddObject(dmHonorarios.dts.FieldByName('nomeescritorio').AsString + '-' + dmHonorarios.dts.FieldByName('cnpjescritorio').AsString, escritorio);
      dmHonorarios.dts.next;
   end;

   if MonthOf(Date) = 01 then //é dezembro. Pode ser referente ao ano anterior
      child.cbAno.Items.Add(IntToStr(YearOf(IncYear(date, -1))));
   child.cbAno.Items.Add(IntToStr(YearOf(date)));
   if child.cbAno.Items.Count = 1 then
      child.cbAno.ItemIndex := 1;
end;

procedure TMainForm.actValidaPlanExecute(Sender: TObject);
begin
   frmValidaPlan := TFrmValidaPlan.Create(Nil);
end;

procedure TMainForm.actDigitaExecute(Sender: TObject);
begin
   frmDigitaPlan := TFrmDigitaPlan.Create(nil);
   frmDigitaPlan.planJbm.usuario := usuario;
end;

procedure TMainForm.actRelatoriosExecute(Sender: TObject);
begin
   frmRelatorios := TfrmRelatorios.Create(nil);
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
   srec : TSearchRec;
   ret : integer;
   i : integer;
   ini : TInifile;
   dataBaseName : string;
   perfisIncluir : TStringList;
begin
   for i := 0 to 2 do
   begin
      try
         case i  of
            0 : ret := FindFirst('c:\temp\$T*.*', faanyfile, sREc);
            1 : ret := FindFirst(GetEnvironmentVariable('TMP') + '\$T*.*', faanyfile, sREc);
            2 : ret := FindFirst(GetEnvironmentVariable('TeMP') + '\$T*.*', faanyfile, sREc);
         end;

         while ret = 0 do
         begin
            Application.ProcessMessages;
            try
               case i  of
                  0 : DeleteFile('c:\temp\'+sREc.name);
                  1 : DeleteFile(GetEnvironmentVariable('TMP') + '\' + sREc.name);
                  2 : DeleteFile(GetEnvironmentVariable('TEMP') + '\' + sREc.name);
               end;
            except
            end;
            ret := FindNext(sRec);
         end;
      finally
         FindClose(srec);
      end;
   end;

   perfisIncluir := TStringList.Create;
   try
      ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
      try
         //adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('honorarios', 'databasename', '') + ';Persist Security Info=True';
         dataBaseName := ExtractFilePath(ini.readstring('honorarios', 'databasename', ''));

         StrToList(ini.ReadString('perfis', 'nomes', ''), ';', perfisIncluir);
      finally
         ini.free;
      end;

      frmUsuario := TfrmUsuario.Create(nil);
      try
//         frmUsuario.cbxUsuario.AddItem('AMANDA', nil);
//        frmUsuario.cbxUsuario.AddItem('NILCE', nil);
//         frmUsuario.cbxUsuario.AddItem('SAMARA', nil);
//         frmUsuario.cbxUsuario.AddItem('ROSE', nil);
//         frmUsuario.cbxUsuario.AddItem('NELSON', nil);
//         frmUsuario.cbxUsuario.AddItem('ADRIANA', nil);
//         frmUsuario.cbxUsuario.AddItem('CRISTIANE', nil);
//         frmUsuario.cbxUsuario.AddItem('ALEXANDRE', nil);
         for i := 0 to perfisIncluir.Count - 1 do
            frmUsuario.cbxUsuario.AddItem(perfisIncluir.Strings[i], nil);

         if frmUsuario.ShowModal <> mrOk then
         begin
            Application.Terminate;
            exit;
         end;

         if Not FileExists(dataBaseName + 'honorarios_' + frmUsuario.cbxUsuario.Text + '.mdb') then
         begin
            ShowMessage('Erro no endereço do banco de dados. Contate a Presenta Sistemas');
            Application.Terminate;
            exit;
         end;
         Show;
         lblUsu.Caption := '        Banco de dados do(a) usuário(a): ' + frmUsuario.cbxUsuario.Text;
         Caption := 'Digita Honorários - versão 11.0.001(c)';
         usuario := frmUsuario.cbxUsuario.Text;
         dmHonorarios.adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + dataBaseName + 'honorarios_' + frmUsuario.cbxUsuario.Text + '.mdb' + ';Persist Security Info=True';
         dmHonorarios.dirLck := dmHonorarios.dirLck + '\' + frmUsuario.cbxUsuario.Text;
         dmHonorarios.IndexaTabelaReclamadas;
         if Not DirectoryExists(dmHonorarios.dirLck) then
            mkDir(dmHonorarios.dirLck);
      finally
         frmUsuario.Free;
      end;
   finally
      perfisIncluir.Free;
   end;
end;

procedure TMainForm.Setusuario(const Value: string);
begin
  Fusuario := Value;
end;

procedure TMainForm.actConfigDataInicial(Sender: TObject);
begin
   frmCadDataInicial := TfrmCadDataInicial.Create(nil);
   try
      frmCadDataInicial.dataInicial.Date := dmHonorarios.ObtemUltimaDataCadastrada;
      if frmCadDataInicial.ShowModal <> mrOk then
         exit;

      dmHonorarios.CadastraDataInicial(dmHonorarios.ObtemUltimoAnoMesCadastrado, frmCadDataInicial.dataInicial.Date);
   finally
      frmCadDataInicial.Free;
   end;
end;

procedure TMainForm.actCadastraAdvogadoInterno(Sender: TObject);
begin

   //verifica se a tabela tbadvogadosinternos existe;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 codigofuncional from tbadvogadosinternos';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O mecanismo de banco de dados', merr.Message) <> 0) and
            (Pos('não encontrou a tabela', merr.Message) <> 0) and
            (Pos('tbadvogadosinternos', merr.Message) <> 0) then
            dmHonorarios.CriaTabelaAdvogadosInternos
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   frmCadAdvogadoInterno := TfrmCadAdvogadoInterno.Create(nil);
   try
      dmHonorarios.ObtemAdvogadosCadastrados(0);
      frmCadAdvogadoInterno.ShowModal;
   finally
      frmCadAdvogadoInterno.Free;
   end;
end;

procedure TMainForm.actEnviaGcpjExecute(Sender: TObject);
begin
   frmEnviaGcpj := TfrmEnviaGcpj.Create(nil);
end;

end.
