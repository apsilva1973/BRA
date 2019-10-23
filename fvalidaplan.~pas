unit fvalidaplan;

interface

uses
  windows, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Buttons, Grids, DBGrids, uplanjbm, inifiles,
  datamodulegcpj_base_i, Func_Wintask_Obj, mgenlib,  Messages, datamodulegcpj_base_V, datamodulegcpj_base_ii,
  datamodulegcpj_base_vii, fAtosPendentes, datamodulegcpj_base_iii, datamodulegcpj_base_viii, dateutils, datamodulegcpj_compartilhado,
  datamodulegcpj_baixados, adodb, datamodulegcpj_base_ix, main, datamodulegcpj_trabalhistas;

type
  TfrmValidaPlan = class(TForm)
    Label1: TLabel;
    DBGrid1: TDBGrid;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    re: TRichEdit;
    BitBtn3: TBitBtn;
    BitBtn5: TBitBtn;
    processando: TEdit;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    cbSemVolumetria: TCheckBox;
    BitBtn4: TBitBtn;
    cbSemReajuste: TCheckBox;
    BitBtn9: TBitBtn;
    cbAjuizamento: TCheckBox;
    Label2: TLabel;
    numNota: TEdit;
    BitBtn10: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure BitBtn10Click(Sender: TObject);
  private
    FplanJbm: TPlanilhaJBM;
    procedure SetplanJbm(const Value: TPlanilhaJBM);
    { Private declarations }
  public
    { Public declarations }
    property planJbm : TPlanilhaJBM read FplanJbm write SetplanJbm;
    procedure ObtemCodigoReclamadas;
    procedure IndexaBasesGCPJ;
    procedure IndexaBaseHonorarios;
    procedure CopiaBasesGCPJ;
    procedure ChecaNomeAutor;
    procedure ObtemTipoAcao;
    procedure ObtemNomeEmpresaGrupo;
    procedure ConfereValoresContestacao(tipo: integer);
    procedure ValidaOutrosValores;
    procedure CalculaNumeroDaNota;
    procedure CorrigirNota;
    procedure VerificaDuplicidades;
    procedure RemoveEstadosDeVolumetria;
    procedure TrataProcessoTrabalhista(tipo: integer);
    procedure AlertaAtoNaoValidado;
    procedure ExibeMensagem(mensagem: string);
  end;

var
  frmValidaPlan: TfrmValidaPlan;
  ie8 : boolean;
  ultimaPagina, passagem : integer;
  ehPresenta : integer;
  dataContratoNovo : string;

implementation

uses datamodule_honorarios, datamodulegcpj_base_X, datamodulegcpj_base_xi,
  DB;


{$R *.dfm}

{ TfrmValidaPlan }

procedure TfrmValidaPlan.SetplanJbm(const Value: TPlanilhaJBM);
begin
  FplanJbm := Value;
end;

procedure TfrmValidaPlan.FormCreate(Sender: TObject);
begin
   planJbm := TPlanilhaJBM.Create(ExibeMensagem);
   DBGrid1.DataSource := planJbm.dsPlan;
   planJbm.ObtemPlanilhasValidar;
   if planJbm.dtsPlanilha.RecordCount = 0 then
   begin
      BitBtn1.Enabled := false;
      BitBtn10.Enabled := false;
   end;

   if IsParam('/ie8') then
      ie8 := true
   else
      ie8 := false;
end;

procedure TfrmValidaPlan.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfrmValidaPlan.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   planJbm.Finalizar;
   planJbm.Free;
   action := CaFree;
end;

procedure TfrmValidaPlan.BitBtn1Click(Sender: TObject);
var
   ini : TIniFile;
   entrada, saida : array [0..255] of char;
   volumetriaContratoNovo : integer;
   indexar : boolean;
begin
   ShortDateFormat := 'dd/mm/yyyy';

   if Not DirectoryExists('c:\presenta') then
   begin
      ExibeMensagem('Criando o diretório: c:\presenta');
      MkDir('c:\presenta');
   end;

   if Not DirectoryExists('c:\presenta\basediaria') then
   begin
      ExibeMensagem('Criando o diretório: c:\presenta\basediaria');
      MkDir('c:\presenta\basediaria');
   end;

   if Not FileExists('c:\presenta\basediaria\CONFIG.INI') then
   begin
      ExibeMensagem('Criando o diretório: c:\presenta');
      StrPCopy(entrada, ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
      StrPcopy(saida, 'c:\presenta\basediaria\CONFIG.INI');
      CopyFile(entrada, saida, FALSE);
      Sleep(1000);
      BitBtn1Click(sender);
      exit;
   end;

   //verifica se a tabela tbtiposnaoatualizar existe;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 identificador from tbtiposnaoatualizar';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O mecanismo de banco de dados', merr.Message) <> 0) and
            (Pos('não encontrou a tabela', merr.Message) <> 0) and
            (Pos('tbtiposnaoatualizar', merr.Message) <> 0) then
            dmHonorarios.CriaTabelaTiposNaoAtualizar
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se o campo datadoato existe na tabela tbtiposnaoatualizar existe;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 datadoato from tbtiposnaoatualizar';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O parâmetro datadoato não tem valor padrão', merr.Message) <> 0) then
            dmHonorarios.CriaColunaDataDoAto_TiposNaoAtualizar
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se a tabela tbvaloresnaoatualizar existe;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 tipoandamento from tbvaloresnaoatualizar';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O mecanismo de banco de dados', merr.Message) <> 0) and
            (Pos('não encontrou a tabela', merr.Message) <> 0) and
            (Pos('tbvaloresnaoatualizar', merr.Message) <> 0) then
            dmHonorarios.CriaTabelaValoresNaoAtualizar
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se o campo fgenviadogcpj existe na tabela tbplanilhas;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 fgenviadogcpj from tbplanilhas';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O parâmetro fgenviadogcpj não tem valor padrão', merr.Message) <> 0) then
            dmHonorarios.CriaColunafgenviadogcpj_tbplanilhas
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se o campo dtenviogcpj existe na tabela tbplanilhas;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 dtenviogcpj from tbplanilhas';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O parâmetro dtenviogcpj não tem valor padrão', merr.Message) <> 0) then
            dmHonorarios.CriaColunadtenviogcpj_tbplanilhas
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se o campo fgretornogcpj existe na tabela tbplanilhas;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 fgretornogcpj from tbplanilhas';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O parâmetro fgretornogcpj não tem valor padrão', merr.Message) <> 0) then
            dmHonorarios.CriaColunafgretornogcpj_tbplanilhas
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   dmHonorarios.InsereTiposNaoAtualizarPadrao;
   dmHonorarios.InsereValoresNaoAtualizarPadrao;

   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      ehPresenta := ini.ReadInteger('gcpj', 'presenta', 0);
      indexar := ini.ReadBool('gcpj', 'indexar', true);
      if  ehPresenta = 0 then
      begin
         if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'basescopiadas','') then
         begin
            CopiaBasesGCPJ;
            ini.WriteString('gcpj', 'basescopiadas',FormatDateTime('dd/mm/yyyy', date));
         end;
      end;

      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         if indexar then
         begin
            IndexaBasesGCPJ;
            ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
         end;
      end;

      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'copiaDePara','') then
      begin
         ExibeMensagem('Criando tabela de DE/PARA');
         Application.ProcessMessages;
         planJbm.CadastraDeParaEmpresasGrupo;
         ini.WriteString('gcpj', 'copiaDePara',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   ini := TiniFile.Create(ExtractFilePath(Application.Exename) + 'CONFIG.INI');
   try
      dataContratoNovo := FormatDateTime('yyyy/mm/dd', StrToDate(ini.ReadString('ContratoNovo', 'Data', '01/11/2015')));
      volumetriaContratoNovo := ini.ReadInteger('Volumetria', 'ContratoNovo', 0);
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').Asinteger;
   planJbm.CarregaFiliais;

   ChecaNomeAutor;
   ObtemTipoAcao;
   ObtemCodigoReclamadas;
   ObtemNomeEmpresaGrupo;
   VerificaDuplicidades;
   ValidaOutrosValores;

   case volumetriaContratoNovo of
      0 : begin
         ConfereValoresContestacao(1);
         ConfereValoresContestacao(2);
      end;
   else
      ConfereValoresContestacao(3);
   end;

   //alterado em 03/04/2014
   AlertaAtoNaoValidado;
end;

procedure TfrmValidaPlan.ObtemCodigoReclamadas;
var
   reclamada : TProcessoReclamdas;
   ret : integer;
begin
   re.lines.Clear;

   processando.Text := 'Capturando dados das reclamadas';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(2);
   planJbm.dtsAtos.tag := 1;

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem(IntToStr(planJbm.dtsAtos.tag) + '-Cruzando processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString + ' com dados do GCPJ');
      planJbm.dtsAtos.tag := planJbm.dtsAtos.tag + 1;
      Application.ProcessMessages;


      reclamada := TProcessoReclamdas.Create(planJbm.idEscritorio,
                                             planJbm.idplanilha,
                                             planJbm.anomesreferencia,
                                             planJbm.sequencia,
                                             planJbm.linhaPlanilha,
                                             planJbm.gcpj,
                                             planJbm.ExibeMensagem);
      try
         reclamada.RemoveReclamadas;

         ExibeMensagem('   Obtendo dados nas bases I e IV  - Reclamadas');
         ret := reclamada.ObtemEnvolvidosDoProcesso(planjbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger, false);
         if ret = 0 then //erro. Não existe no GCPJ
         begin
            planJbm.GravaOcorrencia(2, 'Processo não cadastrado no GCPJ (Não achou envolvidos 1)');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if ret < 0 then
         begin
            ExibeMensagem('   Gravando ocorrência de erro');
            planJbm.GravaOcorrencia(2, reclamada.errorMessage);
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         ExibeMensagem('   Obtendo dados nas bases I e IV  - Reclamantes');
         ret := reclamada.ObtemEnvolvidosDoProcesso(planjbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger, true);
         if ret = 0 then //erro. Não existe no GCPJ
         begin
            planJbm.GravaOcorrencia(2, 'Processo não cadastrado no GCPJ  (Não achou envolvidos 2)');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if ret < 0 then
         begin
            ExibeMensagem('   Gravando ocorrência de erro');
            planJbm.GravaOcorrencia(2, reclamada.errorMessage);
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         //verifica se tem data da baixa
         planJbm.VerificaDataDaBaixa(planJbm.dtsAtos.FieldByName('gcpj').AsString);
         planJbm.MarcaProcessoCruzadoGcpj(3);
      finally
         reclamada.Free;
      end;
      planJbm.dtsAtos.Next;
   end;
end;


procedure TfrmValidaPlan.IndexaBasesGCPJ;
begin
   processando.Text := 'Indexando as bases do GCPJ';
   Application.ProcessMessages;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base I. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseI;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base II. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseII;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base III. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseIII;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base IV. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseIV;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base V. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseV;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base VIII. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseVIII;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base IX. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseIX;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base X. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseX;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base XI. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseXI;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base Compartilhada. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseCompartilhada;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base Trabalhistas. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseTrabalhistas;


   ExibeMensagem('Criando índice na base de dados do GCPJ Base Recuperados. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseRecuperados;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base Migrados. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseMigrados;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base VII. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseVII;

   ExibeMensagem('Criando índice na base de dados do GCPJ Base Baixados. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseBaixados;
   Sleep(100);
end;

procedure TfrmValidaPlan.ChecaNomeAutor;
begin
   re.lines.Clear;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').Asinteger;
   planJbm.CarregaFiliais;


   processando.Text := 'Verificando nome do autor';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(0);
   while not planJbm.dtsAtos.Eof do
   begin
       {planJbm.linhaPlanilha := planjbm.dtsatos.FieldByName('linhaPlanilha').AsInteger;
       planJbm.idEscritorio := planjbm.dtsatos.FieldByName('idescritorio').AsInteger;
       planJbm.idplanilha := planjbm.dtsatos.FieldByName('idplanilha').AsInteger;
       planJbm.sequencia :=  planjbm.dtsatos.FieldByName('sequencia').AsInteger;
       planJbm.anomesreferencia := planjbm.dtsatos.FieldByName('anomesreferencia').AsString;}
       planJbm.CarregaCamposDaTabela;
      //E-mail Carlos - 27/05/2019
       if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < '2018/01/01') then
       begin
          planJbm.GravaOcorrencia(2, 'Data do Ato menor que 01/01/2018, verificar.');
          planJbm.MarcaProcessoCruzadoGcpj(9);
          planJbm.dtsAtos.Next;
          continue;
       end;

      planJbm.CarregaCamposDaTabela;
      ExibeMensagem('Verificando nome do Autor do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;


      ExibeMensagem('Verificando se número do GCPJ é válido');
      if planJbm.dtsAtos.FieldByName('gcpj').IsNull then
      begin
         planJbm.GravaOcorrencia(2, 'Sem número do GCPJ');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if Length(planJbm.dtsAtos.FieldByName('gcpj').AsString) > 11 then
      begin
         planJbm.GravaOcorrencia(2, 'Número do GCPJ inválido');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      ExibeMensagem('Verificando se valor é difereente de zero');
      if planJbm.dtsAtos.FieldByName('valor').AsFloat = 0 then
      begin
         planJbm.GravaOcorrencia(2, 'Valor do ato zerado');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      //processo existe no gcpj?
      ExibeMensagem('Verificando se GCPJ existe');
      if planJbm.ProcessoExisteNoGcpj = 0 then
      begin
         planJbm.GravaOcorrencia(2, 'Processo não cadastrado no GCPJ');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if planJbm.NomeEnvolvidoOk then
         planJbm.MarcaNomeEnvolvidoOk;

      planJbm.MarcaProcessoCruzadoGcpj(1);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ObtemTipoAcao;
var
   ret : integer;
begin
   re.lines.Clear;

   processando.Text := 'Capturando dados do Tipo/Subtipo Ação';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(1);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem('Verificando tipo acao do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      planJbm.GravaTipoProcesso;
      ret := planJbm.GravaTipoSubtipo;
      if ret <> 1 then
      begin
         case ret of
            0 : planJbm.GravaOcorrencia(2, 'Processo sem TIPO AÇÃO');
            2 : planJbm.GravaOcorrencia(2, 'TIPO AÇÃO Inválido');
            3 : planJbm.GravaOcorrencia(2, 'TIPO AÇÃO não compatível com DEJUR 4429/4799/8288');
            4 : planJbm.GravaOcorrencia(2, 'Tipo de ato devido somente para ações Ativas');
         end;
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;
      planJbm.MarcaProcessoCruzadoGcpj(2);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ObtemNomeEmpresaGrupo;
var
  ret :integer;
begin
   re.lines.Clear;

   processando.Text := 'Capturando dados da Empresa Grupo';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(3);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem('Verificando Empresa Ligada: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      if planJbm.GravaEmpresaGrupo(0) = 1 then
         planJbm.MarcaProcessoCruzadoGcpj(4);

      //para laudo pericial não precisa verificar para quem foi distribuído o processo
      if planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'LAUDO PERICIAL' then
      begin
         //se for avulso trabalhista ou avulso civel, não pode estar distribuído para nenhum escritório

         if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'SUSCITACAO DE DIVIDA' then
         begin
            ret := planJbm.AtoDistribuidoParaEscritorio;
            if ret = 9 then //está distribuído para o escritório. Não pode
            begin
               planJbm.GravaOcorrencia(2, 'Ato distribuído para o escritório que está cobrando. Não pode ser pago');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO TRABALHISTA') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIÊNCIA TRABALHISTA') or
            (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'AVULSO CÍVEL (AUDITORIA)') or
            (Pos('AVULSO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0) then
         begin
            ret := planJbm.AtoDistribuidoParaEscritorio;
            if ret <> 0 then //está distribuído para um escritório. Não pode
            begin
               planJbm.GravaOcorrencia(2, 'Ato distribuído para escritório externo não pode ser pago');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end
         else
         begin
            if planJbm.tipoProcesso = 'CI' then
            begin
               if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AJUIZAMENTO') and
                  (planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger = 0) then
               begin
                  planJbm.GravaOcorrencia(2, 'Tipo de andamento não pode ser pago para ações contrárias');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               //valida se o ato foi distribuido para este excritorio.
               if (planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger <> 1) or
                  ((planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger = 1) and
                   (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CARTA PRECATORIA') and
                   (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'SUSCITACAO DE DIVIDA') and
                   (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CUMPRIMENTO DE CARTA PRECATORIA') and
                   (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'DESBLOQUEIOS') and
                   (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'ACOMPANHAMENTO') and
                   (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'SERVICO DE TIPIFICACAO') and
                   (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'SERVIÇO DE TIPIFICACAO')) then
               begin
                  If ((planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'SERVICO DE TIPIFICACAO') and
                     (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'SERVIÇO DE TIPIFICACAO')) then
                  begin //ale
                      ret := planJbm.AtoDistribuidoParaEscritorio;
                      if ret <> 9 then
                      begin
                         if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'LAUDO PERICIAL') then
                         begin
                            if ret = 0 then
                              planJbm.GravaOcorrencia(2, 'Processo não distribuído')
                            else
                              planJbm.GravaOcorrencia(2, 'Processo distribuído para outro escritório');

                            planJbm.MarcaProcessoCruzadoGcpj(9);
                            planJbm.dtsAtos.Next;
                            continue;
                         END;
                      end;
                  end; //ale
               end;
            end;
         end;
      end;
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ConfereValoresContestacao(tipo: integer);
var
   faixa1, faixa2,faixa3,faixa4,faixa5,faixa6,
   faixa7,faixa8,faixa9 : integer;
   total, faixa : integer;
   valorCorreto : double;
begin
   re.lines.Clear;

   processando.Text := 'Conferindo valores de Contestação';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(4);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
      exit;

   planJbm.cnpjescritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').AsInteger;

   case tipo of
      1 : planJbm.ObtemProcessosCruzarGcpj(5);
      2 : planJbm.ObtemProcessosCruzarGcpj(10);
      3 : planJbm.ObtemProcessosCruzarGcpj(11);
   end;

   if planJbm.dtsAtos.RecordCount = 0 then
      exit;

   //total de atos de contestação com a nova regra de cálculo
   total := planJbm.dtsAtos.RecordCount;



// regra antiga pela média
{   faixa1 := 0;
   faixa2 := 0;
   faixa3 := 0;
   faixa4 := 0;
   faixa5 := 0;
   faixa6 := 0;
   faixa7 := 0;
   faixa8 := 0;
   faixa9 := 0;
   faixa := 1;
   while total > 0 do
   begin
      case faixa of
         1 : begin
            if total >= 50 then
               faixa1 := 50
            else
               faixa1 := total;
            total := total - faixa1;
         end;
         2 : begin
            if total >= 50 then
               faixa2 := 50
            else
               faixa2 := total;
            total := total - faixa2;
         end;
         3 : begin
            if total >= 150 then
               faixa3 := 150
            else
               faixa3 := total;
            total := total - faixa3;
         end;
         4 : begin
            if total >= 250 then
               faixa4 := 250
            else
               faixa4 := total;
            total := total - faixa4;
         end;
         5 : begin
            if total >= 250 then
               faixa5 := 250
            else
               faixa5 := total;
            total := total - faixa5;
         end;
         6 : begin
            if total >= 250 then
               faixa6 := 250
            else
               faixa6 := total;
            total := total - faixa6;
         end;
         7 : begin
            if total >= 500 then
               faixa7 := 500
            else
               faixa7 := total;
            total := total - faixa7;
         end;
         8 : begin
            if total >= 500 then
               faixa8 := 500
            else
               faixa8 := total;
            total := total - faixa8;
         end;
         9 : begin
            faixa9 := total;
            total := total - faixa9;
         end;
      end;
      Inc(faixa);
   end;  }

   valorCorreto := 0;
   case tipo of
      1 : begin //contrarias
//         valorCorreto := ((faixa1*541) + (faixa2*525) + (faixa3*512) + (faixa4*500) + (faixa5*484) + (faixa6*471) + (faixa7*455) + (faixa8*441) + (faixa9*428));
         valorCorreto := ((faixa1*495.20) + (faixa2*480.80) + (faixa3*468) + (faixa4*457.60) + (faixa5*443.20) + (faixa6*431.20) + (faixa7*416) + (faixa8*404) + (faixa9*392));
         valorCorreto := valorcorreto / planJbm.dtsAtos.RecordCount;
      end;
      2 : begin //ativas
//         valorCorreto := ((faixa1*503) + (faixa2*488) + (faixa3*477) + (faixa4*465) + (faixa5*450) + (faixa6*438) + (faixa7*424) + (faixa8*411) + (faixa9*400));
         valorCorreto := ((faixa1*495.20) + (faixa2*480.80) + (faixa3*468) + (faixa4*457.60) + (faixa5*443.20) + (faixa6*431.20) + (faixa7*416) + (faixa8*404) + (faixa9*392));
         valorCorreto := valorcorreto / planJbm.dtsAtos.RecordCount;
      end;
      3 : begin

   //maio de 2016
//         valorCorreto := ((faixa1*541) + (faixa2*525) + (faixa3*512) + (faixa4*500) + (faixa5*484) + (faixa6*471) + (faixa7*455) + (faixa8*441) + (faixa9*428));
//         valorCorreto := valorcorreto / planJbm.dtsAtos.RecordCount;
   //maio 2017
//         valorCorreto := ((faixa1*591) + (faixa2*574) + (faixa3*559) + (faixa4*546) + (faixa5*529) + (faixa6*515) + (faixa7*497) + (faixa8*482) + (faixa9*468));
//         valorCorreto := valorcorreto / planJbm.dtsAtos.RecordCount;
//         valorCorreto := ((faixa1*619) + (faixa2*601) + (faixa3*585) + (faixa4*572) + (faixa5*554) + (faixa6*539) + (faixa7*520) + (faixa8*505) + (faixa9*490));

//         valorCorreto := ((faixa1*495.20) + (faixa2*480.80) + (faixa3*468) + (faixa4*457.60) + (faixa5*443.20) + (faixa6*431.20) + (faixa7*416) + (faixa8*404) + (faixa9*392));
//         valorCorreto := valorcorreto / planJbm.dtsAtos.RecordCount;

         // regra NOVA pela faixa.

          //faixa 1
             if ( total >= 0) and ( total <= 50) then
               valorCorreto := 495.20;
          //faixa 2
             if ( total >= 51) and ( total <= 100) then
               valorCorreto := 480.80;
          //faixa 3
             if ( total >= 101) and ( total <= 250) then
               valorCorreto := 468.00;
          //faixa 4
             if ( total >= 251) and ( total <= 500) then
               valorCorreto := 457.60;
          //faixa 5
             if ( total >= 501) and ( total <= 750) then
               valorCorreto := 443.20;
          //faixa 6
             if ( total >= 751) and ( total <= 1000) then
               valorCorreto := 431.20;
          //faixa 7
             if ( total >= 1001) and ( total <= 1500) then
               valorCorreto := 416;
          //faixa 8
             if ( total >= 1501) and ( total <= 2000) then
               valorCorreto := 404;
          //faixa 9
             if ( total >= 2001) then
               valorCorreto := 392;
      end;
   end;


   while not planjbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem('Verificando valores de contestação: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      if Not cbSemVolumetria.Checked then
      begin
         if planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorCorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
      end
      else
         planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
      planJbm.MarcaProcessoCruzadoGcpj(6);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ValidaOutrosValores;
var
   valorcorreto, valorCorreto10p, valorCorreto5p: double;
   listaTarifas, listaEscritorios : TStringList;
   nomeOrgaoJulgador : string;
   ret : integer;
   somaAtosDescontar, somaAtosDescontar2, descontar : double;
   tiposAtos : TStringList;
   valorConferir, valorConferirAte : double;
   valorPagar : double;
   andamento : string;
   temInicial : boolean;
//   lctconvite : boolean;
   temRequerimento : boolean;
   dtAjuizamento : TDateTime;
   dtCadastro : TDateTime;
   dtDistribuicao : TDateTime;
   totalPago : integer;
   pdts : TAdoDataset;

   function TemDuplicidade : boolean;
   var
      checouDb : boolean;
   begin
      result := true;
      checouDb := false;

      if (planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger <> 1) then
      begin
         if (planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 1) and (Pos('XITO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0)  then   //acordo drc contrárias
            ret := 1
         else
         begin
            ExibeMensagem('   Verificando se tem duplicidade');
            Application.ProcessMessages;

            checouDb := true;

            ret :=  dmgcpcj_base_XI.ObtemOutrosPagamentosDoProcesso(planJbm.gcpj,
                                                                   planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                                   planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                                   planjbm.tipoProcesso,
                                                                   IntToStr(planJbm.codGcpjEscritorio),
                                                                   planJbm.lstCnpjs);
         end
      end
      else
      begin
         ExibeMensagem('   Verificando se tem duplicidade (ATIVAS)');
         Application.ProcessMessages;

         if planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ACORDO' then
         begin
            checouDb := true;

            ret := dmgcpcj_base_XI.ObtemOutrosPagamentosDoProcessoAtivas(planJbm.gcpj,
                                                                   planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                                   planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                                   planjbm.tipoProcesso,
                                                                   IntToStr(planJbm.codGcpjEscritorio),
                                                                   planJbm.lstCnpjs);
         end
         else
            ret := 1;
      end;

      if ret <> 1 then
      begin
         if (planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger <> 1) then
            planJbm.GravaOcorrencia(2, 'Ato não devido para Ações Contrárias')
         else
            planJbm.GravaOcorrencia(2, 'Tipo de ato não reconhecido pelo sistema');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         exit;
      end;

      if checouDb then
      begin
         if dmgcpcj_base_XI.dtsNfiscdt.RecordCount > 0 then
         begin
            if (planJbm.tipoProcesso = 'TR') OR (planJbm.tipoProcesso = 'TO') OR (planJbm.tipoProcesso = 'TA') then
            begin
               if dmgcpcj_base_XI.dtsNfiscdt.RecordCount >= 2 then
               begin
                  planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente (Localizado em notas pagas - Base V - Trabalhista)');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  exit;
               end
            end
            else
            begin
               if Pos('TUTELA', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0 then
               begin
                  if dmgcpcj_base_XI.dtsNfiscdt.RecordCount >= 1 then
                  begin
                     planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente  (Localizado em notas pagas - Base V - Avulso)');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     exit;
                  end
               end
               else
               begin
                  if Pos('AVULSO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0 then
                  begin
                     if dmgcpcj_base_XI.dtsNfiscdt.RecordCount >= 2 then
                     begin
                        planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente  (Localizado em notas pagas - Base V - Avulso)');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        exit;
                     end
                  end
                  else
                  begin
                      if Pos('RECURSO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0 then
                      begin
                         if dmgcpcj_base_XI.dtsNfiscdt.RecordCount >= 1 then
                         begin
                            planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente  (Localizado em notas pagas)');
                            planJbm.MarcaProcessoCruzadoGcpj(9);
                            exit;
                         end
                      end
                   else
                  begin
                     if dmgcpcj_base_XI.dtsNfiscdt.RecordCount >= 5 then // mudança solicitada pela Elaine - 5 item da planilha de melhoria.
                     begin
//                        planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente  (Localizado em notas pagas - Base V - Civel)');
                        planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente.');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        exit;
                     end;
                  end;
                end;
               end;
            end;
         end;
      end
      else
      begin
         //verifica se foi digitado no GCPJ e está aguardando aprovação

         (***revisar
         tiposAtos := TStringList.Create;
         try
            ExibeMensagem('   Verificando se tem duplicidade (BASE_III)');
            Application.ProcessMessages;

            if FileExists(ExtractFilePath(Application.ExeName) + 'TIPOS_ATOS.TXT') then
               tiposAtos.LoadFromFile(ExtractFilePath(Application.ExeName) + 'TIPOS_ATOS.TXT');

            if dmgcpj_base_iii.ObtemPagamentosCadastradosGCPJ(planJbm.gcpj,
                                                              planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                              planjbm.tipoProcesso,
                                                              tiposAtos) <> 1 then
            begin
               planJbm.GravaOcorrencia(2, 'Tipo de ato não reconhecido pelo sistema');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               exit;
            end;

            if dmgcpj_base_iii.dts.RecordCount > 0 then
            begin
               //verifica se o mês e ano do pagamento são iguais
               if FormatDateTime('mmyyyy', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) = FormatDateTime('mmyyyy', dmgcpj_base_iii.dts.FieldByName('DATO_DESP_HONRS').AsDateTime)then
               begin
                  planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente  (Localizado em notas pagas-Base III)');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  exit;
               end;
            end;
         finally
            tiposAtos.Free;
         end;       ***)
      end;
      //verifica se está duplicado no próprio sistema HONORARIOS_DIGITA

      ExibeMensagem('   Verificando se tem duplicidade (BASE DO SISTEMA)');
      Application.ProcessMessages;

      if ((planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger <> 1) and (planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 0)) or
         ((planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 1) and (Pos('XITO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) = 0)) or  //acordo drc contrárias
         ((planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger = 1)  and (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ACORDO')) then
      begin
         ret := planJbm.ObtemOutrosPagamentosDoProcesso(planJbm.gcpj,
                                                        planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                        planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                        planjbm.tipoProcesso,
                                                        IntToStr(planJbm.codGcpjEscritorio),
                                                        planJbm.lstCnpjs);
         case ret of
            0 : begin
               planJbm.GravaOcorrencia(2, 'Tipo de ato não reconhecido pelo sistema');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               exit;
            end;

            2 : begin
               planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente (Localizado no sistema de validação)');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               exit;
            end;
         end;
      end;
      result := false;
   end;

begin

   re.lines.Clear;

   processando.Text := 'Validando valores';
   Application.ProcessMessages;

   if cbSemReajuste.Checked then
   begin
      if Not planJbm.ConfiguracaoSemReajusteEstaOk then
      begin
         ShowMessage('Para utilizar a funcionalidade de cálculo sem reajuste, é preciso alterar as datas dos atos na tela de configuração do sistema');
         exit;
      end;
   end;

   planJbm.CarregaCamposDaTabela;
   ExibeMensagem('Indexando tabelas internas');

   planJbm.CriaIndicesDigitaHonorarios;   //mexer alexandre voltar

   planJbm.ObtemProcessosCruzarGcpj(4);

   planJbm.cnpjescritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').AsInteger;

   planJbm.dtsAtos.Tag := 1;

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem(IntToStr(planJbm.dtsAtos.Tag) + '-Verificando valores: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      planJbm.dtsAtos.Tag := planJbm.dtsAtos.Tag + 1;
      Application.ProcessMessages;

      ExibeMensagem('   Obtendo valor do cálculo');
      Application.ProcessMessages;

      planJbm.GravaValorCalculo;

      ExibeMensagem('   Verificando tipo de ato');
      Application.ProcessMessages;

      ExibeMensagem('   1');
      Application.ProcessMessages;



//PEREIRA
      if ((planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'SERVIÇO DE TIPIFICACAO') or
         (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'SERVICO DE TIPIFICACAO')) then
      begin

        //valorCorreto := 30;
        planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
//        planJbm.GravaValorCorrigido(valorcorreto);
//        planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');

//         planJbm.GravaOcorrencia(2, 'Para tipo e Subtipo AGUARDANDO TIPIFICACAO, não é devido pagamento');
//         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;
//PEREIRA

//PEREIRA

      if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'DUPLICIDADE') or
         (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'EXTINCAO ANTES DA CITACAO') or
         (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'PRE-CADASTRO EM AREA INCORRETA') or
         (planJbm.dtsAtos.FieldByName('motivobaixa').AsString = 'PROCESSO ENVOLVE AREA DE OFICIOS') or
         (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'PROCESSO ENVOLVE GRUPO SEGURADOR') then
      begin
         planJbm.GravaOcorrencia(2, 'Processo não pode ser pago por causa do motivo da baixa');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

//PEREIRA


//PEREIRA
      if ((planJbm.dtsAtos.FieldByName('tipoacao').AsString = 'AGUARDANDO TIPIFICACAO') or
         (planJbm.dtsAtos.FieldByName('subtipoacao').AsString = 'AGUARDANDO TIPIFICACAO'))  then
      begin
         planJbm.GravaOcorrencia(2, 'Para tipo e Subtipo AGUARDANDO TIPIFICACAO, não é devido pagamento');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;
//PEREIRA



      if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'SUSCITACAO DE DIVIDA' then
      begin
//PEREIRA
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
         begin
            if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 557 then
            begin
               planJbm.GravaValorCorrigido(557.00);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         end
         else
//PEREIRA
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
         begin
            if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 619 then
            begin
               planJbm.GravaValorCorrigido(619.00);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
            begin
               if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 591 then
               begin
                  planJbm.GravaValorCorrigido(591.00);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
            end
            else
            begin
               if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 541 then
               begin
                  planJbm.GravaValorCorrigido(541.00);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
            end;
         end;

         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      ExibeMensagem('   2');
      Application.ProcessMessages;

      if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECUPERACAO JUDICIAL' then
      begin
         pdts := TADODataSet.Create(nil);
         try
            dmgcpj_compartilhado.ObtemTipoSubtipo(planJbm.dtsAtos.FieldByName('gcpj').AsString, pdts);

            //alterado em 17/04/2017 - sem e-mail com Samara
            if (pdts.FieldByName('coddejur').AsInteger <> 4799) and (pdts.FieldByName('coddejur').AsInteger <> 4429) and (pdts.FieldByName('coddejur').AsInteger <> 8288) then
            begin
               planJbm.GravaOcorrencia(2, 'Ato devido somente para Dejur 4799/4429/8288');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         finally
            pdts.free;
         end;

         if TemDuplicidade = true then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end; 

//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
         begin
            if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 296 then
            begin
               planJbm.GravaValorCorrigido(296.00);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         end
         else
//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
         begin
            if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 329 then
            begin
               planJbm.GravaValorCorrigido(329.00);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
            begin
               if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 314 then
               begin
                  planJbm.GravaValorCorrigido(314.00);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
           end
           else
           begin
             if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 287 then
             begin
               planJbm.GravaValorCorrigido(287.00);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
             end
             else
               planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
           end;
         end;

         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      ExibeMensagem('   3');
      Application.ProcessMessages;

      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADIANTAMENTO RESP EMPRESA') or
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADIANTAMENTO TRANSITO JULGADO') then
      begin
         pdts := TADODataSet.Create(nil);
         try
            dmgcpj_compartilhado.ObtemTipoSubtipo(planJbm.dtsAtos.FieldByName('gcpj').AsString, pdts);

            //alterado em 17/04/2017 - sem e-mail com Samara
            if (pdts.FieldByName('coddejur').AsInteger <> 4799) and (pdts.FieldByName('coddejur').AsInteger <> 4429) and (pdts.FieldByName('coddejur').AsInteger <> 8288) then
            begin
               planJbm.GravaOcorrencia(2, 'Ato devido somente para Dejur 4799/4429/8288');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         finally
            pdts.free;
         end;

         if TemDuplicidade = true then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end;
//PEREIRA
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
         begin
            valorConferir := 557;
            valorConferirAte := 5149;
         end
         else
//PEREIRA
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
         begin
            valorConferir := 619;
            valorConferirAte := 5722;
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
            begin
               valorConferir := 591;
               valorConferirAte := 5464;
            end
            else
            begin
               valorConferir := 541;
               valorConferirAte := 5000;
            end;
         end;
         
         valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.01);

         if (valorcorreto < valorConferir) then
         begin
            planJbm.GravaOcorrencia(2, 'Valor calculado inferior ao mínimo permitido: ' + FormatFloat('0.00#,##', valorcorreto));
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if valorcorreto > valorConferirAte then
            valorcorreto := valorConferirAte;

         if planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorCorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;


      ExibeMensagem('   4');
      Application.ProcessMessages;

      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO RECUPERACAO JUDICIAL') then
      begin
        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
        begin
          valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
        end
        else
         begin
          valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
         end;
         somaAtosDescontar := 0;
         descontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, 0, totalPago, 0, 'ADIANTAMENTO RESP EMPRESA');
         if descontar > 0 then
            somaAtosDescontar := somaAtosDescontar + descontar;

         Descontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, 0, totalPago, 0, 'ADIANTAMENTO TRANSITO JULGADO');
         if Descontar > 0 then
            somaAtosDescontar := somaAtosDescontar + descontar;

         Descontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, 0, totalPago, 0, 'AJUIZAMENTO');
         if Descontar > 0 then
            somaAtosDescontar := somaAtosDescontar + descontar;


//alexandre 25/05
        If (somaAtosDescontar > 101772) {and ((pdts.FieldByName('coddejur').AsInteger = 4429) or (pdts.FieldByName('coddejur').AsInteger = 4799))} then
         begin
           planJbm.GravaOcorrencia(2, 'Valor de teto excedido');
           planJbm.MarcaProcessoCruzadoGcpj(9);
           planJbm.dtsAtos.Next;
           continue;
         end;
//alexandre 25/05

         valorcorreto := valorcorreto - somaAtosDescontar;

         if valorcorreto <= 0 then
         begin
            planJbm.GravaOcorrencia(2, 'Valor calculado é menor ou igual a 0');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;


//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
         begin
            if valorcorreto > 101772 then
            begin
               valorCorreto := 101772 - somaAtosDescontar;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end;
         end
         else
//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
         begin
            if valorcorreto > 113080 then
            begin
               valorCorreto := 113080 - somaAtosDescontar;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end;
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
            begin
               if valorcorreto > 107986 then
               begin
                  valorCorreto := 107986 - somaAtosDescontar;
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end;
               end
            else
            begin
               if valorcorreto > 98817 then
               begin
                  valorCorreto := 98817 - somaAtosDescontar;
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end;
            end;
         end;

         planJbm.GravaValorCorrigido(valorcorreto);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      ExibeMensagem('   5');
      Application.ProcessMessages;

      if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADIANTAMENTO FALENCIA' then
      begin
         pdts := TADODataSet.Create(nil);
         try
            dmgcpj_compartilhado.ObtemTipoSubtipo(planJbm.dtsAtos.FieldByName('gcpj').AsString, pdts);

            //alterado em 17/04/2017 - sem e-mail com Samara
            if (pdts.FieldByName('coddejur').AsInteger <> 4799) and (pdts.FieldByName('coddejur').AsInteger <> 4429) and (pdts.FieldByName('coddejur').AsInteger <> 8288) then
            begin
               planJbm.GravaOcorrencia(2, 'Ato devido somente para Dejur 4799/4429/8288');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         finally
            pdts.Free;
         end;

         if TemDuplicidade = true then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end;

//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
         begin
            if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 557 then
            begin
               planJbm.GravaValorCorrigido(557);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         end
         else
//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
         begin
            if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 619 then
            begin
               planJbm.GravaValorCorrigido(619);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
            begin
               if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 591 then
               begin
                  planJbm.GravaValorCorrigido(591);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
            end
            else
            begin
               if planjbm.dtsAtos.FieldByName('valor').AsFloat <> 541 then
               begin
                  planJbm.GravaValorCorrigido(541.00);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
            end;
         end;
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      ExibeMensagem('   6');
      Application.ProcessMessages;


      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO FALENCIA') then
      begin

        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
        begin
          valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
        end
        else
         begin
          valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
         end;
         somaAtosDescontar := 0;
         Descontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, 0, totalPago, 0, 'ADIANTAMENTO FALENCIA');
         if Descontar > 0 then
            somaAtosDescontar := somaAtosDescontar + descontar;

         Descontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, 0, totalPago, 0, 'AJUIZAMENTO');
         if Descontar > 0 then
            somaAtosDescontar := somaAtosDescontar + descontar;



        //alexandre 25/05
                If somaAtosDescontar > 101772 then
                 begin
                   planJbm.GravaOcorrencia(2, 'Valor de teto excedido');
                   planJbm.MarcaProcessoCruzadoGcpj(9);
                   planJbm.dtsAtos.Next;
                   continue;
                 end;
        //alexandre 25/05

         valorcorreto := valorcorreto - somaAtosDescontar;

         if valorcorreto <= 0 then
         begin
            planJbm.GravaOcorrencia(2, 'Valor calculado é menor ou igual a 0');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
         begin
            if valorcorreto > 101772 then
            begin
               valorCorreto := 101772 - somaAtosDescontar;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end;
         end
         else
//pereira
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
         begin
            if valorcorreto > 113080 then
            begin
               valorCorreto := 113080 - somaAtosDescontar;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end;
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
            begin
               if valorcorreto > 107986 then
               begin
                  valorCorreto := 107986 - somaAtosDescontar;
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end;
            end
            else
            begin
               if valorcorreto > 98817 then
               begin
                  valorCorreto := 98817 - somaAtosDescontar;
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end;
            end;
         end;

         planJbm.GravaValorCorrigido(valorcorreto);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      ExibeMensagem('   7');
      Application.ProcessMessages;

      if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'LAUDO PERICIAL' then
      begin
         //utiliza o valor da planilha
         //verifica se já foi pago
         if planJbm.LaudoPericialJaFoiPago then
         begin
            planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      ExibeMensagem('   8');
      Application.ProcessMessages;

      //nao pode pagar ibi
      if (planjbm.tipoProcesso <> 'TR') and
         (planjbm.tipoProcesso <> 'TO') and
         (planjbm.tipoProcesso <> 'TA') and
         ((planjbm.dtsatos.FieldByName('empresaligadaagrupar').AsInteger = 63) OR
          (planjbm.dtsatos.FieldByName('empresaligadaagrupar').AsInteger = 7459) or
          (planjbm.dtsatos.FieldByName('empresaligadaagrupar').AsInteger = 5404)) then
      begin
         if planJbm.dtsPlanilha.FieldByName('fgpagaribi').AsInteger <> 1 then
         begin
            planJbm.GravaOcorrencia(2, 'Processo Banco IBI Civel');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if Not planJbm.dtsPlanilha.FieldByName('dataatosibi').IsNull then
         begin
            if FormatDateTime('yyyymmdd', dmgcpcj_base_V.ObtemDataCadastroGCPJ(planJbm.dtsAtos.FieldByName('gcpj').AsString)) < FormatDateTime('yyyymmdd', planJbm.dtsPlanilha.FieldByName('dataatosibi').AsDateTime) then
            begin
               planJbm.GravaOcorrencia(2, 'Processo Banco IBI Civel');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;
      end;

      ExibeMensagem('   9');
      Application.ProcessMessages;

//ALEXANDRE AQUI
      if (planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger = 0) and
         (planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 0) and
         (planjbm.tipoProcesso <> 'TR') and
         (planjbm.tipoProcesso <> 'TO') and
         (planjbm.tipoProcesso <> 'TA') then
      begin
         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO') OR
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO NA FASE DE EXECUCAO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE EXITO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'TUTELA')) and
            (dmgcpcj_base_I.EhProcessoPreocupante(planjbm.gcpj)) then
         begin
            ExibeMensagem('   Verificando se tem duplicidade - ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString);
            if TemDuplicidade = true then
            begin
               planJbm.dtsAtos.Next;
               continue;
            end;

            ExibeMensagem('   Marcando o valor correto - ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString);
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO') OR
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO NA FASE DE EXECUCAO') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'TUTELA') then
            begin
//PEREIRA
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
                  valorcorreto := 1030
               else
//PEREIRA
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                  valorcorreto := 1145
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
                  valorcorreto := 1093
               else
                  valorcorreto := 1000;
            end
            else
            begin
               if planjbm.dtsatos.FieldByName('motivobaixa').IsNull then
               begin
                  planJbm.GravaOcorrencia(2, 'Para este tipo de andamento, o processo tem que estar baixado no GCPJ');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               valorCorreto := 0;

               if Pos('ACORDO',  planjbm.dtsatos.FieldByName('motivobaixa').AsString) <> 0 then
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
                  begin
                     planJbm.GravaOcorrencia(2, 'Ato anterior à data do contrato');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;
//PEREIRA
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
                  begin
                     if planJbm.valorpedido <= 5150 then
                        valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.08))
                     else
                        valorcorreto := StrToFloat(FormatFloat('0.00', (5150 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.08));
                  end
                  else
//PEREIRA

                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                  begin
                     if planJbm.valorpedido <= 5722 then
                        valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.10))
                     else
                        valorcorreto := StrToFloat(FormatFloat('0.00', (5722 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.10));
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
                     begin
                        if planJbm.valorpedido <= 5464 then
                           valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.10))
                        else
                           valorcorreto := StrToFloat(FormatFloat('0.00', (5464 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.10));
                     end
                     else
                     begin
                        if planJbm.valorpedido <= 5000 then
                           valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.10))
                        else
                           valorcorreto := StrToFloat(FormatFloat('0.00', (5000 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.10));
                     end
                  end
               end
               else
               begin
                  if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'IMPROCEDENCIA') then
                  begin
//PEREIRA
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
                        valorcorreto := 5150
                     else
//PEREIRA
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                        valorcorreto := 5722
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
                        valorcorreto := 5464
                     else
                        valorcorreto := 5000;
                  end;
               end;

               if valorCorreto <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor calculado é menor ou igual a zero');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorcorreto then
               begin
                  planJbm.GravaValorCorrigido(valorCorreto);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
            end;

            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;
      end;

      ExibeMensagem('   10');
//      Application.ProcessMessages;
// mostro

         ret := 0;
         if  (((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PREPOSTO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString = 'EMBARGOS DE TERCEIROS')) or
             ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString = 'EMBARGOS DE TERCEIROS')) or
             ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONTESTACAO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString = 'EMBARGOS DE TERCEIROS')) or
             ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO NA FASE DE EXECUCAO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString = 'EMBARGOS DE TERCEIROS')) or
             ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE EXITO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString = 'EMBARGOS DE TERCEIROS'))) then
         begin
           ret := 1;
         end;


      if ((planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger = 1) and (ret = 0)) then
      begin
         ExibeMensagem('   11');
         ExibeMensagem('   É DRC Ativas');
//         Application.ProcessMessages;
         valorCorreto := 0;

         ExibeMensagem('      Verificando tipo de andamento: ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString);
         //alterado em 31/07/2014
         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PREPOSTO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIROS')) or
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIROS')) or
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONTESTACAO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIROS')) or
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO NA FASE DE EXECUCAO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIROS')) or
             ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE EXITO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIROS')) then
         begin
            ExibeMensagem('      Gravando ocorrência de erro');
            planJbm.GravaOcorrencia(2, 'Ato devido somente para ações Contrárias');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

      //alterado em 30/10/2013
         if (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0) and
               (Pos('XITO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0) and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'PENHORA') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'BUSCA E APREENSÃO/REINTEGRAÇÃO DE POSSE') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ACORDO') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ADJUDICACAO/ARREMATACAO') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ADJUDICAÇÃO/ARREMATAÇÃO') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ENTREGA AMIGAVEL') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'IRRECUPERABILIDADE ATE 1 ANO') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'IRRECUPERABILIDADE ATÉ 1 ANO') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'IRRECUPERABILIDADE APOS 1 ANO') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'IRRECUPERABILIDADE APÓS 1 ANO') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CARTA PRECATORIA') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CUMPRIMENTO DE CARTA PRECATORIA') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'DESBLOQUEIOS') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CONSOLIDACAO DE PROPRIEDADE') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CONSOLIDAÇÃO DE PROPRIEDADE') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CONSOLIDAÇÃO PROPRIEDADE (2 P/C ADIC)') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'HONORARIOS DE VENDA') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'LEVANTAMENTO DE ALVARA') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CONSIGNAÇÃO BANCÁRIA') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CONSIGNACAO BANCARIA') and
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PREPOSTO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIRO')) and
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIRO')) and
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONTESTACAO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIRO')) and
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'RECURSO NA FASE DE EXECUCAO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIRO')) and
            ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE EXITO') and (planjbm.dtsatos.FieldByName('tipoacao').AsString <> 'EMBARGOS DE TERCEIRO')) then
         begin
            ExibeMensagem('      Gravando ocorrência de erro');
            planJbm.GravaOcorrencia(2, 'Tipo de andamento: ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString + ' não reconhecido para AÇÕES ATIVAS');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         //regra enviada em 19/03/2015 às 12:43
         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ENTREGA AMIGAVEL') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO DE PROPRIEDADE') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO DE PROPRIEDADE') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO PROPRIEDADE (2 P/C ADIC)') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE VENDA') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADJUDICACAO/ARREMATACAO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADJUDICAÇÃO/ARREMATAÇÃO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'LEVANTAMENTO DE ALVARA') then
         begin
            //verifica se já foi pago ato de recuperação final
            ExibeMensagem('      ObtemSomaJaPagaADescontar');
            SomaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, 0, planjbm.dtsatos.FieldByName('tipoandamento').AsString, false, true);

             //alexandre 25/05
            If somaAtosDescontar > 101772 then
             begin
               planJbm.GravaOcorrencia(2, 'Valor de teto excedido');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
             end;
             //alexandre 25/05


            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'LEVANTAMENTO DE ALVARA') then
            begin
               if totalpago >= 2 then
               begin
                  ExibeMensagem('      Gravando ocorrência de erro');
                  planJbm.GravaOcorrencia(2, 'Já foram pagos 2 atos de Recuperação Final (localizado no GCPJ)');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end
            else
            begin
               if somaAtosDescontar > 0 then
               begin
                  ExibeMensagem('      Gravando ocorrência de erro');
                  planJbm.GravaOcorrencia(2, 'Já foi pago ato de Recuperação Final (localizado no GCPJ)');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AJUIZAMENTO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ENTREGA AMIGAVEL') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO DE PROPRIEDADE') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO DE PROPRIEDADE') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO PROPRIEDADE (2 P/C ADIC)') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE VENDA') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADJUDICACAO/ARREMATACAO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADJUDICAÇÃO/ARREMATAÇÃO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSÃO/REINTEGRAÇÃO DE POSSE') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'LEVANTAMENTO DE ALVARA') then
         begin
            //verifica se na própia planilha tem
            ExibeMensagem('      ObtemSomaJaPagaADescontar');
            somaAtosDescontar := dmHonorarios.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('tipoandamento').AsString, planJbm.dtsAtos.FieldByName('idplanilha').AsInteger, true);


             //alexandre 25/05
            If somaAtosDescontar > 101772 then
             begin
               planJbm.GravaOcorrencia(2, 'Valor de teto excedido');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
             end;
             //alexandre 25/05



            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'LEVANTAMENTO DE ALVARA') then
            begin
               if totalpago >= 2 then
               begin
                  ExibeMensagem('      Gravando ocorrência de erro');
                  planJbm.GravaOcorrencia(2, 'Já foram pagos 2 atos de Recuperação Final (localizado na planilha)');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end
            else
            begin
               if somaAtosDescontar > 0 then
               begin
                  if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AJUIZAMENTO') then
                     planJbm.GravaOcorrencia(2, 'Não pode pagar AJUIZAMENTO (consta ACORDO ou CONSOLIDACAO DE PROPRIEDADE ou ENTREGA AMIGAVEL ou  LEVANTAMENTO DE ALVARA ou ADJUDICACAO/ARREMATACAO na mesma planilha)')
                  else
                  begin
                  //revisar
                     ExibeMensagem('      Gravando ocorrência de erro');
                     if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE') or
                        (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSÃO/REINTEGRAÇÃO DE POSSE') then
                        planJbm.GravaOcorrencia(2, 'Não pode pagar BUSCA E APREENSÃO (consta ACORDO na mesma planilha)')
                     else
                        planJbm.GravaOcorrencia(2, 'Já foi pago ato de Recuperação Final (localizado na planilha)');
                  end;
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

         //regra enviada em 19/03/2015 as 15:54
         if planJbm.dtsAtos.FieldByName('tipoacao').AsString = 'AGUARDANDO AJUIZAMENTO DA ACAO' then
         begin
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ACORDO') and
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ENTREGA AMIGAVEL') and
               (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0) then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2, 'Para tipo de ação ' + planJbm.dtsAtos.FieldByName('tipoacao').AsString + ', só pode pagar ACORDO e ENTREGA AMIGAVEL');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;

         if (Pos('XITO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
         begin
            ExibeMensagem('      Gravando ocorrência de erro');
            planJbm.GravaOcorrencia(2, 'Não pode pagar êxito DRC ações Ativas');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;


         //regra incluída com e-mail de 18/06/2015 16:45
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/22' then //utlizar a rotina de novo cálculo
         begin
            ExibeMensagem('      Gravando ocorrência de erro');
            planJbm.GravaOcorrencia(2, 'Data do ato anterior a 22/04/2013');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         //Verificar se consta a inicial anexada no GCPJ (caminho abaixo descrito pelo Edson),
         //caso contrário, inconsistir pelo motivo NÃO CONSTA INICIAL NO GCPJ.
         //Banco de Dados: BaseDiaria_GCPJ_II.mdb
         //Tabela: GCPJB0K7
         //Campo: IANEXO_PROCS_JUDIC
         //Filtro: <expressão contendo a palavra "INICIAL">
         if Not cbAjuizamento.Checked then
         begin
            if (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
            begin
               ExibeMensagem('   AJUIZAMENTO - Tipo ação: ' + planJbm.dtsAtos.FieldByName('codtipoacao').AsString);
               Application.ProcessMessages;

               ExibeMensagem('   Verificando se tem inicial/requerimento');
               Application.ProcessMessages;

        //       dmgcpj_base_vii.VerificaInicialERequerimento(planJbm.dtsAtos.FieldByName('gcpj').AsString, temInicial, temRequerimento);
               ShowMessage('Não é possível verificar Inicial/Requerimento!!!');
               exit;

               if planJbm.dtsAtos.FieldByName('codtipoacao').AsInteger <> 8915 then
               begin
                  //if Not dmgcpj_base_vii.ProcessoTemInicial(planJbm.dtsAtos.FieldByName('gcpj').AsString) then
                  if Not temInicial then
                  begin
                     //checar se tem requerimento
                     //tem requerimento?
                     ExibeMensagem('   Não tem inicial');
                     Application.ProcessMessages;

                     //if Not dmgcpj_base_vii.ProcessoTemRequerimento(planJbm.dtsAtos.FieldByName('gcpj').AsString) then
                     if Not temRequerimento then
                     begin
                        ExibeMensagem('   Não tem requerimento');
                        Application.ProcessMessages;
                        planJbm.GravaOcorrencia(2, 'Falta anexar inicial/requerimento no GCPJ');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        planJbm.dtsAtos.Next;
                        continue;
                     end;

                     ExibeMensagem('   Verificando advogado interno');
                     Application.ProcessMessages;
                     //07/07/2014
                     //verifica se foi distribuido para o advogado interno
                     if not planJbm.ProcessoDistribuidoParaAdvogadoInterno then
                     begin
                        ExibeMensagem('      Gravando ocorrência de erro');
                        planJbm.GravaOcorrencia(2, 'Processo não distribuído para Advogado da Área responsável');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        planJbm.dtsAtos.Next;
                        continue;
                     end;
                  end;
               end
               else
               begin
                  ExibeMensagem('   Verificando se tem requerimento');
                  Application.ProcessMessages;

                  //if Not dmgcpj_base_vii.ProcessoTemRequerimento(planJbm.dtsAtos.FieldByName('gcpj').AsString) then
                  if not temRequerimento then
                  begin
                     ExibeMensagem('      Gravando ocorrência de erro');
                     planJbm.GravaOcorrencia(2, 'Falta anexar o requerimento no GCPJ');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;
               end;
            end;
         end;

         //incluido em 03/04/2014 - e-mail da Nilce de 25/03/2014
         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PENHORA') then
         begin
            //PENHORA: 1,5% sobre o valor da causa (valor base), limitado a R$11.000,00; É devido somente um ato de Penhora;
            valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.015);

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               if valorcorreto > 11000 then
                  valorcorreto := 11000;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 11690 then
                     valorcorreto := 11690;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 12645 then
                        valorcorreto := 12645;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 20000 then
                           valorcorreto := 20000;
                     end
                     else
                     begin
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                        begin
                           if valorcorreto > 21856 then
                              valorcorreto := 21856;
                        end
                        else
                        begin
//pereira
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                        begin
                           if valorcorreto > 22888 then
                              valorcorreto := 22888;
                        end
                        else
//pereira
                           if valorcorreto > 20600 then
                              valorcorreto := 20600;
                        end;
                     end;
                  end;
               end;
            end;


            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;
// mexer aqui Alexandre
         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE') or
             (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSÃO/REINTEGRAÇÃO DE POSSE')) then
         begin
            //BUSCA E APREENSAO/REINTEGRACAO DE POSSE: Deprecia o valor base em 20%, depois aplica-se 5% (anexa, tabela de cálculo em excel),
            //limitado a R$11.690,00. É devido somente 1 ato. Este ato deverá ser validado se tiver a referência "RC - bem apreendido", caso
            //ontrário recusar.
            ExibeMensagem('      Verificando se processo tem referência');

// mexer alexandre
            //só pode pagar 1 honorarios de venda - confirmar com a Maria.
            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'HONORARIOS DE VENDA');
            if somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar ato porque já pagou HONORARIOS DE VENDA');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
// mexer alexandre

            if (Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '285,236')) and (ehPresenta=0) then
            begin
               planJbm.GravaOcorrencia(2, 'Falta referência RC BEM APREENDIDO ou RC BEM REINTEGRADO no GCPJ');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //alterado em 13/12/2016 - Reunião Maria, Elaine e Adriana
            ExibeMensagem('      Verificando se processo está distribuído para o escritório');

            ret := planJbm.AtoDistribuidoParaEscritorio;
            if ret <> 9 then //não está distribuído para escritório. Não pode
            begin
               planJbm.GravaOcorrencia(2, 'Ato não está distribuído para o escritório que está cobrando. Não pode ser pago');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //BUSCA E APREENSAO/REINTEGRACAO DE POSSE: Deprecia o valor base em 20%, depois aplica-se 5% (anexa, tabela de cálculo em excel),
            //limitado a R$11.000,00. É devido somente 1 ato.
            //o valor base contem o valor do bem e não o valor a receber
            //depreciando o valor do veículo em 20%
            valorCorreto := planJbm.dtsAtos.FieldByName('valorbase').AsFloat - (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.20);

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
            begin
               //aplicando 5% sobre o valor depreciado
               valorcorreto := (valorCorreto * 0.05);

               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 11000 then
                     valorcorreto := 11000;

               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 11690 then
                        valorcorreto := 11690;
                  end
                  else
                  begin
                     if valorcorreto > 12645 then
                        valorcorreto := 12645;
                  end;
               end;
            end
            else
            begin
               dtDistribuicao := dmgcpcj_base_I.ObtemDataDaDistribuicao(planJbm.gcpj);
               if dtDistribuicao = 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Erro obtendo data de distribuicao do processo');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
//pereira2
              if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
              begin
               if DaysBetween(dtDistribuicao, planJbm.dtsatos.FieldByName('datadoato').AsDateTime) <= 90 then
                  valorcorreto := (valorCorreto * 0.075)
               else
                  valorcorreto := (valorCorreto * 0.050);
              end
              else
              begin
               if DaysBetween(dtDistribuicao, planJbm.dtsatos.FieldByName('datadoato').AsDateTime) <= 90 then
                  valorcorreto := (valorCorreto * 0.10)
               else
                  valorcorreto := (valorCorreto * 0.075);
              end;
//pereira2

               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 20000 then
                     valorcorreto := 20000;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 21856 then
                        valorcorreto := 21856;
                  end
                  else
//pereira
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 22888 then
                        valorcorreto := 22888;
                  end
                  else
//pereira
                  begin
                     if valorcorreto > 20600 then
                        valorcorreto := 20600;
                  end;
               end;
            end;

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;
// final alexandre

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE VENDA') then
         begin
            ExibeMensagem('   Obtendo valor já pago por ENTREGA AMIGAVEL');
            Application.ProcessMessages;

            //Nova Regra: HONORARIOS DE VENDA: 10% sobre o valor base, descontando-se todos os valores pagos, limitado a R$85.955,00.
            //Caso tenha pagamento de ENTREGA AMIGAVEL, recusar HONORARIOS DE VENDA.
            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'ENTREGA AMIGAVEL');

             //alexandre 25/05
            If somaAtosDescontar > 101772 then
             begin
               planJbm.GravaOcorrencia(2, 'Valor de teto excedido');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
             end;
             //alexandre 25/05



            if somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar ato porque já pagou ENTREGA AMIGAVEL');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;


            //só pode pagar 1 honorarios de venda
            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'HONORARIOS DE VENDA');
            if somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar ato porque já pagou HONORARIOS DE VENDA');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            ExibeMensagem('   Obtendo valor já pago para o processo');
            Application.ProcessMessages;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, 0 );

           if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
           begin
             valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
           end
           else
            begin
             valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
            end;
            
            if somaAtosDescontar > 0 then
               valorcorreto := valorCorreto - somaAtosDescontar;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               if valorcorreto > 85955 then
                  valorcorreto := 85955-somaAtosDescontar;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 91352 then
                     valorcorreto := 91352-somaAtosDescontar;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 98817 then
                        valorcorreto := 98817-somaAtosDescontar;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 107986 then
                           valorcorreto := 107986-somaAtosDescontar;
                     end
                     else
                     begin
//pereira
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 113080 then
                           valorcorreto := 113080-somaAtosDescontar;
                     end
                     else
//pereira
                        if valorcorreto > 101772 then
                           valorcorreto := 101772-somaAtosDescontar;
                     end;
                  end;
               end;
            end;

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'LEVANTAMENTO DE ALVARA') then
         begin
            ExibeMensagem('   Obtendo valor já pago para o processo');
            Application.ProcessMessages;

            //LEVANTAMENTO DE ALVARA: 10% sobre o valor base, descontando-se todos os valores pagos, limitado a R$85.955,00.


            // Verifica se ja foi pago alguma coisa para este escritorio. -- ver com a Maria.
            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, planJbm.codGcpjEscritorio);

              // Verifica se ja foi pago alguma coisa para qualquer escritorio.
//            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, 0);

             if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
             begin
               valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
             end
             else
              begin
               valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
              end;


             if (somaAtosDescontar > 0)  then
               valorcorreto := valorCorreto - somaAtosDescontar;


            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               if valorcorreto > 85955 then
                  valorcorreto := 85955-somaAtosDescontar;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 91352 then
                     valorcorreto := 91352-somaAtosDescontar;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 98817 then
                        valorcorreto := 98817-somaAtosDescontar;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 107986 then
                           valorcorreto := 107986-somaAtosDescontar;
                     end
                     else
//pereira
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 113080 then
                           valorcorreto := 113080-somaAtosDescontar;
                     end
                     else
//pereira
                     begin
                        if valorcorreto > 101772 then
                           valorcorreto := 101772-somaAtosDescontar;
                     end;
                  end;
               end;
            end;

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO') then
         begin
            //ACORDO: 10% Sobre o valor base, descontando-se todos os valores pagos (Ajuizamento, penhora, busca e apreensão/reintegração de posse, irrecuperabilidade,
            // etc), limitado a R$91.352,00. Pode ser mais de um pagamento, devido a acordos parcelados, porém, nunca poderá ultrapassar o teto máximo de R$91.352,00.
            //Verificar também a data do ato do acordo parcelado, tendo em vista que não poderá coincidir com atos de acordos  já pagos. Este ato deverá ser validado
            //se tiver a referência "RC - acordo formalizado/suspensão", caso contrário recusar.
            ExibeMensagem('      Verificando se processo tem referência');

//voltar Alexandre !!!!!

            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACORDO') then
            begin
                if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '201, 200, 350') then
                begin
                   planJbm.GravaOcorrencia(2, 'Falta referência RC ACORDO/RC ACORDO FORMALIZADO no GCPJ/RC ACORDO CONCILIANDO ESCRIT');
                   planJbm.MarcaProcessoCruzadoGcpj(9);
                   planJbm.dtsAtos.Next;
                   continue;
                end;
            end;


            //27/05/2019 - Solicitação da Samara
            if ((planJbm.dtsAtos.FieldByName('tipoacao').AsString = 'ALIENACAO FIDUCIARIA DE BENS IMOVEIS')) then
            begin
              If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'HONORARIOS DE VENDA') > 0 then
              begin
                 planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato HONORARIOS DE VENDA');
                 planJbm.MarcaProcessoCruzadoGcpj(9);
                 planJbm.dtsAtos.Next;
                 continue;
              end;

              If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'ENTREGA AMIGAVEL') > 0 then
              begin
                 planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato ENTREGA AMIGAVEL');
                 planJbm.MarcaProcessoCruzadoGcpj(9);
                 planJbm.dtsAtos.Next;
                 continue;
              end;

              If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'ACORDO') > 0 then
              begin
                 planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato ACORDO');
                 planJbm.MarcaProcessoCruzadoGcpj(9);
                 planJbm.dtsAtos.Next;
                 continue;
              end;


              If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'CONSOLIDACAO DE PROPRIEDADE') > 0 then
              begin
                 planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato CONSOLIDACAO DE PROPRIEDADE');
                 planJbm.MarcaProcessoCruzadoGcpj(9);
                 planJbm.dtsAtos.Next;
                 continue;
              end;
            end;

            //ACORDO: 10% Sobre o valor base, descontando-se todos os valores pagos (Ajuizamento, penhora, busca e apreensão/reintegração de posse, irrecuperabilidade,  etc),
            //limitado a R$85.955,00. Pode ser mais de uma pagamento, devido a acordos parcelados, porém, nunca poderá ultrapassar o teto máximo de R$85.955,00.
            //Verificar também a data do ato do acordo parcelado, tendo em vista que não poderá coincidir com atos de acordos  já pagos.
            //Quais atos deverão ser descontados para pagamento de Acordo, Adjudicação/Arrematação, Entrega Amigável e Honorários de Venda?: Ajuizamento,
            //Penhora, Busca e Apreensão/Reintegração de Posse e Irrecuperabilidade até e após 1 ano.

            ExibeMensagem('   Obtendo valor já pago para o processo');
//            Application.ProcessMessages;

            //buscar no gcpj só valor pagor por motivo ACORDO. Se tiver não entrar entrar nesta rotina de desconto de valores adiantados

            ExibeMensagem('   Verificando se já houve pagamento de acordo');
            Application.ProcessMessages;

//negão
            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'ACORDO');


            if somaAtosDescontar = 0 then
            begin
               ExibeMensagem('   SomaAtosDescontar(XI) maior do que 0');
               //verifica se na mesma planilha tem acordo validado (ou seja pode pagar)
               somaAtosDescontar := dmHonorarios.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('tipoandamento').AsString, planJbm.dtsAtos.FieldByName('idplanilha').AsInteger);
               if somaAtosDescontar = 0 then
               begin
                  //se não pagou acordo ainda, descontar outros adiantamentos
                  somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj,
                                                                                 planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger,
                                                                                 totalPago,
                                                                                 planJbm.codGcpjEscritorio);





                 if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
                 begin
                   valorcorreto10p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
                   valorcorreto5p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.05);
                 end
                 else
                  begin
                   valorcorreto10p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
                   valorcorreto5p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.05);
                  end;


                  if somaAtosDescontar > 0 then
                  begin
                     valorcorreto10p := valorCorreto10p - somaAtosDescontar;
                     valorcorreto5p := valorCorreto5p - somaAtosDescontar;
                  end;
               end
               else
               begin

                 if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
                 begin
                   valorcorreto10p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
                   valorcorreto5p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.05);
                 end
                 else
                  begin
                   valorcorreto10p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
                   valorcorreto5p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.05);
                  end;
               end;
            end
            else
            begin
               ExibeMensagem('   Verificando se já foi pago acordo para a mesma data');
//verifica se já foi pago acordo para a mesma data
               if dmgcpcj_base_XI.PagouAcordoNestaData(planJbm.gcpj,planjbm.dtsatos.FieldByName('datadoato').AsDateTime) then
               begin
                  planJbm.GravaOcorrencia(2, 'Já foi pago acordo para a data informada');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               //verifica se já pagou acordo no mês
               ExibeMensagem('   verifica se já foi pago acordo no mesmo mês');
               dmgcpcj_base_XI.ObtemACordosPagos(planJbm.gcpj);
               if dmgcpcj_base_XI.dtsNfiscdt.RecordCount > 0 then
               begin
                  dmgcpcj_base_XI.dtsNfiscdt.Tag := 0;
                  while Not dmgcpcj_base_XI.dtsNfiscdt.Eof do
                  begin
                     Application.ProcessMessages;
                     if FormatDateTime('yyyymm', dmgcpcj_base_XI.dtsNfiscdt.FieldByName('datOcorr').AsDateTime) = FormatDateTime('yyyymm', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) then
                     begin
                        ExibeMensagem('   Já foi pago acordo no mesmo mês');
                        planJbm.GravaOcorrencia(2, 'Já foi pago acordo no mês informado');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        planJbm.dtsAtos.Next;
                        dmgcpcj_base_XI.dtsNfiscdt.tag := 1;
                        break;
                     end;
                     dmgcpcj_base_XI.dtsNfiscdt.Next;
                  end;
                  if dmgcpcj_base_XI.dtsNfiscdt.Tag = 1 then
                     continue;
               end;


               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
               begin
                 valorcorreto10p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
                 valorcorreto5p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.05);
               end
               else
                begin
                 valorcorreto10p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
                 valorcorreto5p := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.05);
                end;
            end;


            //27/05/2019 - Solicitação da Samara
            if ((planJbm.dtsAtos.FieldByName('tipoacao').AsString = 'ALIENACAO FIDUCIARIA DE BENS IMOVEIS')) then
            begin
              If valorcorreto10p > 15448 then
              begin
                valorcorreto := 15448 - somaAtosDescontar;
              end
              else
              begin
                if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto10p))) and
                   (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto5p))) then
                begin
                   if (((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto10p) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto10p) > 0.55)) or
                      ((valorCorreto10p > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto10p - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55))) then
                   begin
                      if (((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto5p) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto5p) > 0.55)) or
                         ((valorCorreto5p > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto5p - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55))) then
                      begin
                         planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                         planJbm.MarcaProcessoCruzadoGcpj(9);
                         planJbm.dtsAtos.Next;
                         continue;
                      end
                      else
                         valorcorreto := valorCorreto5p;
                   end
                   else
                      valorcorreto := valorCorreto10p;
                end
                else
                begin
                   if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) = StrToFloat(FormatFloat('0.00', valorCorreto10p))) then
                      valorcorreto := valorCorreto10p
                   else
                      valorcorreto := valorCorreto5p;
                end;
             end;
            end
            else
            begin
             //alexandre 25/05
              If somaAtosDescontar > 101772 then
               begin
                 planJbm.GravaOcorrencia(2, 'Valor de teto excedido');
                 planJbm.MarcaProcessoCruzadoGcpj(9);
                 planJbm.dtsAtos.Next;
                 continue;
               end;
               //alexandre 25/05

              if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
              begin
                 if valorcorreto10p > 85955 then
                    valorcorreto10p := 85955-somaAtosDescontar;
              end
              else
              begin
                 if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
                 begin
                    if valorcorreto10p > 91352 then
                       valorcorreto10p := 91352-somaAtosDescontar;

                    if valorCorreto5p > 11690 then
                       valorCorreto5p := 11690-somaAtosDescontar;
                 end
                 else
                 begin
                    if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= dataContratoNovo)  and //utlizar a rotina de novo cálculo
                       (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30') then //utlizar a rotina de novo cálculo
                    begin
                       if valorcorreto10p > 98817 then
                          valorcorreto10p := 98817-somaAtosDescontar;

                       if valorCorreto5p >  20000 then
                          valorCorreto5p := 20000-somaAtosDescontar;
                    end
                    else
                    begin
                       if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                       begin
                          if valorcorreto10p > 98817 then
                             valorcorreto10p := 98817-somaAtosDescontar;

                          if valorCorreto5p > 12645 then
                             valorCorreto5p := 12645-somaAtosDescontar;
                       end
                       else
                       begin                                                                                   //alterei aqui  - '2016/05/01'
                          if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/05/01' then //utlizar a rotina de novo cálculo
                          begin
                             if valorcorreto10p > 107986 then
                                valorcorreto10p := 107986-somaAtosDescontar;

                             if valorCorreto5p >  21856 then
                                valorCorreto5p := 21856-somaAtosDescontar;
                          end
                          else
                          begin
                             if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                             begin
                                if valorcorreto10p > 113080 then
                                   valorcorreto10p := 113080-somaAtosDescontar;

                                if valorCorreto5p >  22888 then
                                   valorCorreto5p := 22888-somaAtosDescontar;
                             end
                             else
  //pereira
                             begin
                                if valorcorreto10p > 101772 then
                                   valorcorreto10p := 101772-somaAtosDescontar;

                                if valorCorreto5p >  20600 then
                                   valorCorreto5p := 20600-somaAtosDescontar;
                             end

  //pereira

                          end;
                       end;
                    end;
                 end;
              end;


              if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto10p))) and
                 (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto5p))) then
              begin
                 if (((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto10p) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto10p) > 0.55)) or
                    ((valorCorreto10p > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto10p - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55))) then
                 begin
                    if (((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto5p) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto5p) > 0.55)) or
                       ((valorCorreto5p > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto5p - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55))) then
                    begin
                       planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                       planJbm.MarcaProcessoCruzadoGcpj(9);
                       planJbm.dtsAtos.Next;
                       continue;
                    end
                    else
                       valorcorreto := valorCorreto5p;
                 end
                 else
                    valorcorreto := valorCorreto10p;
              end
              else
              begin
                 if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) = StrToFloat(FormatFloat('0.00', valorCorreto10p))) then
                    valorcorreto := valorCorreto10p
                 else
                    valorcorreto := valorCorreto5p;
              end;
            end;
         end;


         // pereirafull
         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PURGAÇÃO DE MORA') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PURGACAO DE MORA')) then
         begin
            //ADJUDICACAO/ARREMATACAO: 10% Sobre o valor base, descontando-se todos os valores pagos (Ajuizamento, penhora, busca e apreensão/reintegração
            //de posse, irrecuperabilidade,  etc), limitado a R$85.955,00. Não poderá haver pagamento de Acordo e Adjudicação ou Entrega amigável,
            //deve ser um ou outro.
            //Quais atos deverão ser descontados para pagamento de Acordo, Adjudicação/Arrematação, Entrega Amigável e Honorários de Venda?: Ajuizamento,
            //Penhora, Busca e Apreensão/Reintegração de Posse e Irrecuperabilidade até e após 1 ano.
            ExibeMensagem('   Obtendo valor já pago para o processo');
            Application.ProcessMessages;


            if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '174, 176, 353') then
            begin
               planJbm.GravaOcorrencia(2, 'Falta referência AF PURGAÇÃO DA MORA no GCPJ');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;


            If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'HONORARIOS DE VENDA') > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato HONORARIOS DE VENDA');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'ENTREGA AMIGAVEL') > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato ENTREGA AMIGAVEL');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'ACORDO') > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato ACORDO');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;


            If dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'CONSOLIDACAO DE PROPRIEDADE') > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Já foi feito pagamento para o Ato CONSOLIDACAO DE PROPRIEDADE');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;


            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, 0);

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
            begin
              valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
            end
            else
             begin
              valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
             end;

            if somaAtosDescontar > 0 then
               valorcorreto := valorCorreto - somaAtosDescontar;


//alefull

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               if valorcorreto > 85955 then
                  valorcorreto := 85955-somaAtosDescontar;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 91352 then
                     valorcorreto := 91352-somaAtosDescontar;
               end
               else
               begin
                  if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= dataContratoNovo)  and //utlizar a rotina de novo cálculo
                     (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30') then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 98817 then
                        valorcorreto := 98817-somaAtosDescontar;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 98817 then
                           valorcorreto := 98817-somaAtosDescontar;
                     end
                     else
                     begin                                                                                   //alterei aqui  - '2016/05/01'
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/05/01' then //utlizar a rotina de novo cálculo
                        begin
                           if valorcorreto > 107986 then
                              valorcorreto := 107986-somaAtosDescontar;

                        end
                        else
                        begin
                           if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                           begin
                              if valorcorreto > 113080 then
                                 valorcorreto := 113080-somaAtosDescontar;
                           end
                           else
//pereira
                           begin
                              if valorcorreto > 15448 then
                                 valorcorreto := 15448-somaAtosDescontar;
                           end
//pereira

                        end;
                     end;
                  end;
               end;
            end;
//alefull
            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADJUDICACAO/ARREMATACAO') then
         begin
            //ADJUDICACAO/ARREMATACAO: 10% Sobre o valor base, descontando-se todos os valores pagos (Ajuizamento, penhora, busca e apreensão/reintegração
            //de posse, irrecuperabilidade,  etc), limitado a R$85.955,00. Não poderá haver pagamento de Acordo e Adjudicação ou Entrega amigável,
            //deve ser um ou outro.
            //Quais atos deverão ser descontados para pagamento de Acordo, Adjudicação/Arrematação, Entrega Amigável e Honorários de Venda?: Ajuizamento,
            //Penhora, Busca e Apreensão/Reintegração de Posse e Irrecuperabilidade até e após 1 ano.
            ExibeMensagem('   Obtendo valor já pago para o processo');
            Application.ProcessMessages;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, 0);

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
            begin
              valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
            end
            else
             begin
              valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
             end;

            if somaAtosDescontar > 0 then
               valorcorreto := valorCorreto - somaAtosDescontar;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               if valorcorreto > 85955 then
                  valorcorreto := 85955-somaAtosDescontar;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 91352 then
                     valorcorreto := 91352-somaAtosDescontar;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 98817 then
                        valorcorreto := 98817-somaAtosDescontar;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 107986 then
                           valorcorreto := 107986-somaAtosDescontar;
                     end
                     else
                     begin

//pereira
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 113080 then
                           valorcorreto := 113080-somaAtosDescontar;
                     end
                     else
//pereira
                        if valorcorreto > 101772 then
                           valorcorreto := 101772-somaAtosDescontar;
                     end;
                  end;
               end;
            end;

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ENTREGA AMIGAVEL') then
         begin
            //ENTREGA AMIGAVEL: Deprecia o valor base em 20%, depois aplica-se 10% ou 8%, descontando-se todos os valores pagos (Ajuizamento,
            //penhora, busca e apreensão/reintegração de posse, irrecuperabilidade,  etc), limitado a R$91.352,00. É devido somente 1 ato.
            //Este ato deverá ser validado se tiver a referência "RC - entrega amigável", caso contrário recusar.

            ExibeMensagem('      Verificando se processo tem referência');
            if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '286') then
            begin
               planJbm.GravaOcorrencia(2, 'Falta referência RC ENTREGA AMIGAL no GCPJ');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
            //quando tiver pagamento de Busca e Apreensão, não autorizar ENTREGA AMIGÁVEL
            ExibeMensagem('   Verificando se foi pago Busca e Apreensão para este processo');
            Application.ProcessMessages;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE');
            if somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar ENTREGA AMIGÁVEL porque já pagou BUSCA E APREENSÃO');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'ENTREGA AMIGAVEL');
            if somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar ENTREGA AMIGÁVEL porque já houve pagamento anterior');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            ExibeMensagem('   Obtendo valor já pago para o processo');
            Application.ProcessMessages;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger,totalPago, 0);


             //alexandre 25/05
            If somaAtosDescontar > 101772 then
             begin
               planJbm.GravaOcorrencia(2, 'Valor de teto excedido');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
             end;
             //alexandre 25/05


            valorCorreto := planJbm.dtsAtos.FieldByName('valorbase').AsFloat - (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.20);

            //aplicando 10% ou 8% sobre o valor depreciado

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
            begin
              valorcorreto := (valorCorreto * 0.08)
            end
            else
            begin
              valorcorreto := (valorCorreto * 0.10);
            end;

            if somaAtosDescontar <> 0 then
               valorcorreto := valorcorreto - somaAtosDescontar;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               if valorcorreto > 85955 then
                  valorcorreto := 85955-somaAtosDescontar;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 91352 then
                     valorcorreto := 91352-somaAtosDescontar;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 98817 then
                        valorcorreto := 98817-somaAtosDescontar;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 107986 then
                           valorcorreto := 107986-somaAtosDescontar;
                     end
                     else
                     begin
//pereira
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 113080 then
                           valorcorreto := 113080-somaAtosDescontar;
                     end
                     else
//pereira
                        if valorcorreto > 101772 then
                           valorcorreto := 101772-somaAtosDescontar;
                     end;
                  end;
               end;
            end;

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
//               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
//                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
//               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
//             end;                                                             
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE ATE 1 ANO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE ATÉ 1 ANO') then
         begin
            //IRRECUPERABILIDADE ATE 1 ANO: Será devido o valor de R$637,00 e mais 4 anuais de R$200,00. É devido no máximo 5 atos.
            //Para pagamento o processo deve estar ativo no GCPJ, caso contrário barrar. Verificar se a data do ato é inferior a 1 ano,
            //após a data do ajuizamento no GCPJ.

            //o ato está baixado?
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
            begin
               if Not planjbm.dtsatos.FieldByName('motivobaixa').IsNull then
               begin
                  planJbm.GravaOcorrencia(2, 'Para este tipo de andamento, o processo não pode estar baixado no GCPJ');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

            //alterado conforme e-mail da Nilce de 08/05/2015

            ExibeMensagem('      Verificando se processo tem referência');
            if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '223') then
            begin
               planJbm.GravaOcorrencia(2, 'Falta referência RC IRRECUPERAVEL no GCPJ');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //Quando tiver qualquer pagamento, exceto volumetria, não pode pagar IRRECUPERABILIDADE
            ExibeMensagem('   Verificando se já foi feito algum pagamento para o processo');
            Application.ProcessMessages;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj,planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, '', True);
            if somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar IRRECUPERABILIDADE porque já houve outro pagamento');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;


            //verifica se a data do ajuizamento é superior a um ano
            dtAjuizamento := dmgcpcj_base_V.ObtemDataAjuizamento(planJbm.gcpj);
            if dtAjuizamento = 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Processo sem data de ajuizamento');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtAjuizamento) > 365 then
            begin
               planJbm.GravaOcorrencia(2, 'Irrecuperabilidade após 1 ano');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               valorPagar := 200.00;
               valorConferir := 637.00;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  valorConferir := 677.00;
                  valorPagar := 212.00;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
                  begin
                     valorConferir := 732.00;
                     valorPagar := 229.00;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        valorConferir := 732.00;
                        valorPagar := 732.00;
                     end
                     else
                     begin
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                        begin
                           valorConferir := 800.00;
                           valorPagar := 800.00;
                        end
                        else
                        begin
//pereira
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                        begin
                           valorConferir := 838.00;
                           valorPagar := 838.00;
                        end
                        else
//pereira
                           valorConferir := 262.00;
                           valorPagar := 262.00;
                        end;
                     end;
                  end;
               end;
            end;

//alterado pelo e-mnail da Nilce de 26/01/2015
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               valorcorreto := valorConferir;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;

(**            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorPagar then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

            //tem que verificar se já foi paga 1 irrecuperabilidade
            //acertar depois
            valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;**)
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE APOS 1 ANO') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE APÓS 1 ANO') then
         begin
            //IRRECUPERABILIDADE APOS 1 ANO: Será devido 5 parcelas anuais  de R$200,00. É devido no máximo 5 atos. Para pagamento o processo deve
            //permanecer ativo no GCPJ, caso contrário barrar. Verificar se a data do 1º. Ato é superior a 1 ano, após a data do ajuizamento no GCPJ.

            //o ato está baixado?
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
            begin
               if Not planjbm.dtsatos.FieldByName('motivobaixa').IsNull then
               begin
                  ExibeMensagem('      Gravando ocorrência de erro');
                  planJbm.GravaOcorrencia(2, 'Para este tipo de andamento, o processo não pode estar baixado no GCPJ');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

            //alterado conforme e-mail da Nilce de 08/05/2015
            ExibeMensagem('      Verificando se processo tem referência');
            if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '223') then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2, 'Falta referência RC IRRECUPERAVEL no GCPJ');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //verifica se a data do ajuizamento é superior a um ano
            ExibeMensagem('      Obtendo data do Ajuizamento');
            dtAjuizamento := dmgcpcj_base_V.ObtemDataAjuizamento(planJbm.gcpj);
            if dtAjuizamento = 0 then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2, 'Processo sem data de ajuizamento');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtAjuizamento) <= 365 then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2,'Irrecuperabilidade até 1 ano');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= dataContratoNovo then //utlizar a rotina de novo cálculo
            begin
               if DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtAjuizamento) > 1825 then
               begin
                  ExibeMensagem('      Gravando ocorrência de erro');
                  planJbm.GravaOcorrencia(2,'Excedido o limite de 5 anos');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

            //verifica se já houve pagamento no ano
               ExibeMensagem('      ObtemOutrosPagamentosDoProcesso');
            ret :=  dmgcpcj_base_XI.ObtemOutrosPagamentosDoProcesso(planJbm.gcpj,
                                                                   planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                                   planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                                   planjbm.tipoProcesso,
                                                                   IntToStr(planJbm.codGcpjEscritorio),
                                                                   planJbm.lstCnpjs);

            if dmgcpcj_base_XI.dtsNfiscdt.RecordCount > 0 then
            begin
               dmgcpcj_base_XI.dtsNfiscdt.tag := 0;
               while Not dmgcpcj_base_XI.dtsNfiscdt.Eof do
               begin
                  if FormatDateTime('yyyy', dmgcpcj_base_XI.dtsNfiscdt.FieldByName('datOcorr').AsDateTime) = FormatDateTime('yyyy', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) then
                  begin
                     ExibeMensagem('      Gravando ocorrência de erro');
                     planJbm.GravaOcorrencia(2,'Já houve pagamento no ano');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     dmgcpcj_base_XI.dtsNfiscdt.tag := 1;
                     break;
                  end;

                  if DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime,dmgcpcj_base_XI.dtsNfiscdt.FieldByName('datOcorr').AsDateTime) <= 365 then
                  begin
                     ExibeMensagem('      Gravando ocorrência de erro');
                     planJbm.GravaOcorrencia(2,'Já houve pagamento anterior ao prazo de um ano');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     dmgcpcj_base_XI.dtsNfiscdt.tag := 1;
                     break;
                  end;

                  dmgcpcj_base_XI.dtsNfiscdt.Next;
               end;
               if dmgcpcj_base_XI.dtsNfiscdt.tag = 1 then
                  continue;
            end;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 200.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 212.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
               valorConferir := 229.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 229.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 250.00
            else
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
               valorConferir := 262.00
            else

//pereira
               valorConferir := 235.00;


//alterado em 26/11/2014
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               valorcorreto := valorConferir;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;

//            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
//            begin
//               planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
///               planJbm.MarcaProcessoCruzadoGcpj(9);
//               planJbm.dtsAtos.Next;
//               continue;
//            end;

//            valorCorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CARTA PRECATORIA') OR
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CUMPRIMENTO DE CARTA PRECATORIA') then

         begin
            //CARTA PRECATORIA: Será devido o valor de R$438,00 por carta precatória. Não a limite de atos. O processo não deve estar distribuído para o Escritório que está
            //cobrando, caso positivo, recusar.
            //valida se o ato foi distribuido para este excritorio. Se foi recusa

            ret := planJbm.AtoDistribuidoParaEscritorio;
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
            begin
               if ret = 9 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar porque o processo está distribuído para este escritório');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end
            else
            begin
             if ret <> 9 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar porque o processo não está distribuído para este escritório');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 438.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 465.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
               valorConferir := 503.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 541.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 591.00
            else
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
               valorConferir := 619.00
            else
//pereira
               valorConferir := 557.00;




//alterado em 26/11/2014
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               valorcorreto := valorConferir;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;

//            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
//            begin
//               planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
//               planJbm.MarcaProcessoCruzadoGcpj(9);
//               planJbm.dtsAtos.Next;
//               continue;
//            end;
//            valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'DESBLOQUEIOS') then
         begin
            //DESBLOQUEIOS: Será devido um único ato de R$438,00. O processo não deve estar distribuído para o Escritório que está cobrando, caso positivo,
            //recusar. Limitado a 1 pagamento.
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 438.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 465.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 541.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 591.00
            else
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
               valorConferir := 619.00
            else
//pereira
               valorConferir := 557.00;



            //valida se o ato foi distribuido para este excritorio. Se foi recusa
            ret := planJbm.AtoDistribuidoParaEscritorio;
            if ret = 9 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar porque o processo está distribuído para este escritório');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
//alterado em 26/11/2014
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               valorcorreto := valorConferir;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
//            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
//            begin
//               planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
//               planJbm.MarcaProcessoCruzadoGcpj(9);
//               planJbm.dtsAtos.Next;
//               continue;
//            end;
//            valorCorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
         end;

         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO DE PROPRIEDADE') or
             (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO DE PROPRIEDADE')) then
         begin
            //CONSOLIDACAO DE PROPRIEDADE: Conferir se o percentual utilizado para cálculo dos Honorários está entre 4% e 8% em relação ao valor base,
            //caso positivo, aprovar. Está limitado a 8%. É devido um único ato. Este ato deverá ser validado se tiver a referência "RC - Bem consolidado",
            //caso contrário recusar.
            ExibeMensagem('      Verificando se processo tem referência');
            //alterado em 04/01/2017
           // if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
           // begin
               if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '178, 176') then
               begin
                  planJbm.GravaOcorrencia(2, 'Falta referência AF CONSOLIDAÇÃO DA PROPRIEDADE ou PURGAÇÃO DA MORA no GCPJ');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
           (** end
            else
            begin
               if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '178') then
               begin
                  planJbm.GravaOcorrencia(2, 'Falta referência AF CONSOLIDAÇÃO DA PROPRIEDADE');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;     **)

            somaAtosDescontar :=  dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'AJUIZAMENTO');

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < dataContratoNovo then //utlizar a rotina de novo cálculo
            begin
               //CONSOLIDACAO DE PROPRIEDADE: Conferir se o percentual utilizado para cálculo dos Honorários é até 8 % em relação ao valor base, caso positivo,
               //aprovar. Está limitado a 8%. É devido um único ato.

             if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
             begin
               valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
             end
             else
              begin
               valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
              end;

               valorcorreto := valorcorreto - somaAtosDescontar;

               //Maria, o percentual válido é de 4 a 8%.
               //[Maria do Carmo Amaral] Apenas para confirmar, o sistema vai calcular o valor aplicando 4% sobre o valor base (valor A),
               //depois vai calcular 8% sobre o valor base (valor B) e o valor correto tem que ser maior ou igual ao valor A e menor ou igual ao valor B.
               //Se o valor informado na planilha estiver dentro deste intervalo valida, caso contrário rejeita. (Nilce) Maria, isso mesmo.

               valorConferirAte := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.04);
               valorConferirAte := valorConferirAte - somaAtosDescontar;


               if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) < StrToFloat(FormatFloat('0.00', valorConferirAte)))  or
                  (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) > StrToFloat(FormatFloat('0.00', valorCorreto))) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
            end
            else
            begin
               dtDistribuicao := dmgcpcj_base_I.ObtemDataDaDistribuicao(planJbm.gcpj);
               if dtDistribuicao = 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Erro obtendo data de distribuição do processo');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               // regra antiga
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < '2018/06/01' then
               begin
                   if DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtDistribuicao) <= 90 then
                   begin
                      valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.02);
    //                  valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.01);
    //                  valorcorreto := valorcorreto - somaAtosDescontar;
                   end
                   else
                   begin
                      if (DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtDistribuicao) > 90) and
                         (DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtDistribuicao) <= 365)then
                      begin
                         valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.015);
    //                     valorcorreto := valorcorreto - somaAtosDescontar;
                      end
                      else
                      begin
                         valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.01);
    //                     valorcorreto := valorcorreto - somaAtosDescontar;
                      end
                   end;
               end;


               // regra NOVA tudo 1%
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
               begin
                   if DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtDistribuicao) <= 90 then
                   begin
    //                  valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.02);
                      valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.01);
    //                  valorcorreto := valorcorreto - somaAtosDescontar;
                   end
                   else
                   begin
                      if (DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtDistribuicao) > 90) and
                         (DaysBetween(planjbm.dtsatos.FieldByName('datadoato').AsDateTime, dtDistribuicao) <= 365)then
                      begin
                         valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.01);
    //                     valorcorreto := valorcorreto - somaAtosDescontar;
                      end
                      else
                      begin
                         valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.01);
    //                     valorcorreto := valorcorreto - somaAtosDescontar;
                      end
                   end;
               end;
    //pereira
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
               begin
                  if valorcorreto > 15448 then
                     valorcorreto := 15448
                  else
                  if valorcorreto < 1030 then
                     valorcorreto := 1030;
               end
               else
//pereira
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
               begin
                  if valorcorreto > 17165 then
                     valorcorreto := 17165
                  else
                  if valorcorreto < 1145 then
                     valorcorreto := 1145;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
                  begin
                     if valorcorreto > 16392 then
                        valorcorreto := 16392
                     else
                     if valorcorreto < 1093 then
                        valorcorreto := 1093;
                  end
                  else
                  begin
                     if valorcorreto > 15000 then
                        valorcorreto := 15000
                     else
                     if valorcorreto < 1000 then
                        valorcorreto := 1000;
                  end;
               end;
               valorcorreto := valorcorreto - somaAtosDescontar;

               if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto))) then
               begin
//                  valorCorreto := valorCorreto;
//                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end
               else
                  valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
            end;
         end;

         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') OR
             (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO PROPRIEDADE (2 P/C ADIC)')) then
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) > dataContratoNovo then //utlizar a rotina de novo cálculo
            begin
               planJbm.GravaOcorrencia(2, 'Não há previsão contratual');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            ExibeMensagem('      Verificando se processo tem referência');
            if Not dmgcpj_base_X.ProcessoTemReferencia(planJbm.gcpj, '178, 176') then
            begin
               planJbm.GravaOcorrencia(2, 'Falta referência AF CONSOLIDAÇÃO DA PROPRIEDADE ou PURGAÇÃO DA MORA no GCPJ');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC): Será devido mais 2% sobre o valor base, somente após o pagamento de 1 ato de CONSOLIDAÇÃO DE PROPRIEDADE.
            //É devido um único ato.
            //Já foi pago o primeiro ato de CONSOLIDACAO DE PROPRIEDADE?
            ExibeMensagem('   Obtendo valor já pago de CONSOLIDACAO DE PROPRIEDADE');
            Application.ProcessMessages;

            somaAtosDescontar :=  dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'CONSOLIDACAO DE PROPRIEDADE');
            if somaAtosDescontar = 0.00 then
            begin
               //verifica se tem na mesma planilha
               if Not planJbm.TemConsolidacaoPagaNoSistema then
               begin
                  planJbm.GravaOcorrencia(2, 'Não encontrado pagamento de CONSOLIDACAO DE PROPRIEDADE');
                   planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

//            valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.02);
            valorcorreto := (planjbm.dtsatos.FieldByName('valorBase').AsFloat * 0.01);

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;


         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PREPOSTO') then
         begin
            //PREPOSTO: Será devido o valor de R$60,00. Não há limite.
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 60.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 63.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 74.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 81.00
            else
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
               valorConferir := 85.00
            else
//pereira
               valorConferir := 85.00;


//alterado pelo e-mnail da Nilce de 26/01/2015
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               valorcorreto := valorConferir;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;

(**
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
            valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;**)
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSIGNACAO BANCARIA') then
         begin
            //CONSIGNACAO BANCARIA: Será devido um único pagamento de R$250,00. Limitado a 1 pagamento.
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 250.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 265.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 287.00
            else
            begin
               planJbm.GravaOcorrencia(2, 'Não pode mais pagar Consignação Bancária');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //[Maria do Carmo Amaral] Nilce, no caso da CONSIGNACAO BANCARIA, se a data do ato fosse anterior ou igual a 30/04/2014 utilizar o valor de 250,00,
            //limitado a um pagamento, se posterior a 30/04/2014, utilizar o valor de R$265,00. Esta regra está errada? (Nilce) Maria, neste caso corrigir para o
            //valor correto R$265,00 e aprovar.
            //[Maria do Carmo Amaral] Nilce, tínhamos acertado que no caso das ações ativas não íamos fazer correção de valor mas somente aceitar ou rejeitar.
            //É para passar a corrigir para todos os casos ou somente para CONSIGNACAO BANCARIA? (Nilce) Maria, por enquanto somente para Consignação Bancária.

            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               valorcorreto := valorConferir;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE EXITO') then
         begin
            //HONORARIOS DE EXITO (AÇÕES CONTRÁRIAS DRC): 10% sobre o valor base. Descontando-se valores já adiantados e não poderá ultrapassar o valor máximo de
            //R$85.955,00. Poderá haver mais que um ato pago, devido a acordo parcelado.
            //Verificar também a data do ato do acordo parcelado, tendo em vista que não poderá coincidir com atos de acordos  já pagos.
            ExibeMensagem('   Obtendo valor já pago para o processo');
            Application.ProcessMessages;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj,
                                                                           planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger,
                                                                           totalPago, planJbm.codGcpjEscritorio);

             if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
             begin
               valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
             end
             else
              begin
               valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
              end;

            if somaAtosDescontar > 0 then
               valorcorreto := valorCorreto - somaAtosDescontar;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            begin
               if valorcorreto > 85955 then
                  valorcorreto := 85955-somaAtosDescontar;
            end
            else
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 91352 then
                     valorcorreto := 91352-somaAtosDescontar;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 98817 then
                        valorcorreto := 98817-somaAtosDescontar;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 107986 then
                           valorcorreto := 107986-somaAtosDescontar;
                     end
                     else
//pereira
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 113080 then
                           valorcorreto := 113080-somaAtosDescontar;
                     end
                     else
//pereira
                     begin
                        if valorcorreto > 101772 then
                           valorcorreto := 101772-somaAtosDescontar;
                     end;
                  end;
               end;
            end;

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
            begin
               if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                  ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

        if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADESAO POUPADOR') then
        begin
           //ADESAO POUPADOR (AÇÕES CONTRÁRIAS DRC ou Ativas): pagar 200 por poupador
           ExibeMensagem('   Obtendo Poupador para pagamento por Ato');
           Application.ProcessMessages;
           ret := planJbm.ObtemEnvolvidosDoProcesso_2(planjbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger, false);

           if ret = 1 then
           begin
             valorCorreto := 200;
             planJbm.GravaOcorrencia(2, 'Valor pago por poupador R$ 200,00');
             planJbm.GravaValorCorrigido(valorcorreto);
             planJbm.MarcaProcessoCruzadoGcpj(6);
             planJbm.dtsAtos.Next;
             continue;
           end
           else
           begin
             planJbm.GravaOcorrencia(9, 'GCPJ não encontrado na tabela BaseDiaria_GCPJB076, verificar.');
             planJbm.dtsAtos.Next;
             continue;
           end;
        end;



         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ACORDO') and
//            (Pos('PREPOSTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0) and
            (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CARTA PRECATORIA') and
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CUMPRIMENTO DE CARTA PRECATORIA') then
         begin
            //se chegou aqui é porque o valor está correto
            //agora verifica as duplicidades
            dmgcpcj_base_XI.ObtemOutrosPagamentosDoProcessoAtivas(planJbm.gcpj,
                                                                 planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                                 planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                                 planjbm.tipoProcesso,
                                                                 IntToStr(planJbm.codGcpjEscritorio),
                                                                 planJbm.lstCnpjs);

            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PENHORA') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'BUSCA E APREENSÃO/REINTEGRAÇÃO DE POSSE') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'DESBLOQUEIOS') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO DE PROPRIEDADE') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO DE PROPRIEDADE') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSOLIDAÇÃO PROPRIEDADE (2 P/C ADIC)') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSIGNACAO BANCARIA') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CONSIGNAÇÃO BANCÁRIA') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AJUIZAMENTO') or
               (Pos('XITO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
               (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADJUDICACAO/ARREMATACAO') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ENTREGA AMIGAVEL') then
            begin
               //para estes atos deve-se somente 1 pagamento
               if dmgcpcj_base_XI.dtsNfiscdt.RecordCount > 0 then
               begin
                  if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AJUIZAMENTO') or
                     (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
                     planJbm.GravaOcorrencia(2, 'Ato já pago anteriormente')
                  else
                     planJbm.GravaOcorrencia(2, 'Excedido limite de pagamentos para este ato');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE ATE 1 ANO') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE APOS 1 ANO') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE ATÉ 1 ANO') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'IRRECUPERABILIDADE APÓS 1 ANO') then
            begin
               //para estes atos deve-se somente 1 pagamento
               if dmgcpcj_base_XI.dtsNfiscdt.RecordCount > 5 then
               begin
                  planJbm.GravaOcorrencia(2, 'Excedido limite de pagamentos para este ato');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end
         else
         begin
            //inserido em 03/04/2014 - e-mail da Nilce de 25/03/2014
            if (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0)then
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < '2013/04/22' then
               begin
                  //texto alterado a pedido da Nilce
                  //planJbm.GravaOcorrencia(2, 'Não pode pagar AJUIZAMENTO anterior a 22/04/2013');
                  planJbm.GravaOcorrencia(2, 'Fato Gerador anterior a vigência do Contrato, clausula 17.21.2');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;

            //if (Pos('PREPOSTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0) and
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ACOMPANHAMENTO') and
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CARTA PRECATORIA') and
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CUMPRIMENTO DE CARTA PRECATORIA') then
            begin
               if TemDuplicidade = true then
               begin
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;

         if (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2011/11/01' then //utlizar a rotina de novo cálculo
            begin
               planJbm.MarcaProcessoCruzadoGcpj(5);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;

         if valorcorreto <= 0 then
         begin
            planJbm.GravaOcorrencia(2, 'Erro inesperado na validação. Valor a ser pago inferior ou igual a zero');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         planJbm.GravaValorCorrigido(valorCorreto);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end //drcativas
      else
      begin
         ExibeMensagem('   Verificando se é processo trabalhista');
         Application.ProcessMessages;
                  //se for ACOMPANHAMENTO não TRABALHISTA - recusar - E-mail da Amanda de 21/03/2014
         if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ACOMPANHAMENTO' then
         begin
            if (planjbm.tipoProcesso <> 'TR') and
               (planjbm.tipoProcesso <> 'TO') and
               (planjbm.tipoProcesso <> 'TA') then
            begin
               planJbm.GravaOcorrencia(2, 'Só pode pagar ACOMPANHAMENTO para processo trabalhista');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;

         if (planjbm.tipoProcesso = 'TR') OR (planjbm.tipoProcesso = 'TO') OR (planjbm.tipoProcesso = 'TA') then
         begin
            //verifica se a reclamada é BVP (autonomos) ou ex-funcionário
            //regra alterada de acordo com o e-mail do Alexandre de 05/04/2017
(**            planJbm.ObtemReclamadasDoProcesso(planjbm.tipoProcesso);
            planJbm.dtsReclamadas.Tag := 0;

            while Not planJbm.dtsReclamadas.Eof do
            begin
               if (planJbm.dtsReclamadas.FieldByName('codempresa').AsInteger = 5310) then //bvp
                  planJbm.dtsreclamadas.Tag := planJbm.dtsreclamadas.Tag + 1
               else
               if (planJbm.dtsReclamadas.FieldByName('codigopessoaexterna').AsInteger <> 0) then
                  planJbm.dtsreclamadas.Tag := planJbm.dtsreclamadas.Tag + 1
               else
               if (planJbm.dtsReclamadas.FieldByName('codexfuncionario').AsInteger <> 0) then
               begin
                  planJbm.dtsreclamadas.Tag := 3;
                  break;
               end;
               planJbm.dtsReclamadas.Next;
            end;
            **)



                     //regra inserida em 03/04/2014 - email da Amanda de 21/03/2014
            if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'ACOMPANHAMENTO') then
            begin
               //- O processo deverá estar distribuído para o Escritório que está cobrando o ACOMPANHAMENTO, caso contrário
               //inconsistir pelo motivo: Processo não distribuído para o Escritório.
               ret := planJbm.AtoDistribuidoParaEscritorio;
               if ret <> 9 then
               begin
                  planJbm.GravaOcorrencia(2, 'Processo não distribuído para o Escritório');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then
               begin
                  if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 146.00) and
                     (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 58.00) then
                  begin
                     planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then
                  begin
                     //alterado em 096/08/2014
                     if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 155.00) and
                        (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 62.00) then
                     begin
                        planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        planJbm.dtsAtos.Next;
                        continue;
                     end;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then
                     begin
                        //alterado em 096/08/2014
                        if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 168.00) and
                           (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 67.00) then
                        begin
                           planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                           planJbm.MarcaProcessoCruzadoGcpj(9);
                           planJbm.dtsAtos.Next;
                           continue;
                        end;
                     end
                     else
                     begin
                        //alterado em 096/08/2014
                        if (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 184.00) and
                           (StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> 73.00) then
                        begin
                           planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                           planJbm.MarcaProcessoCruzadoGcpj(9);
                           planJbm.dtsAtos.Next;
                           continue;
                        end;
                     end;
                  end;
               end;

               planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end;

            planJbm.dtsReclamadas.Tag := dmgcpj_trabalhistas.ObtemTipoTrabalhista(planJbm.gcpj);

            if (planJbm.dtsReclamadas.Tag = 2) or (planJbm.dtsReclamadas.Tag = 3) then
            begin
               ExibeMensagem('   Tratando processo trabalhista (BVP/Ex-Funcionario)');
               TrataProcessoTrabalhista(1);
               Application.ProcessMessages;
               planJbm.dtsAtos.Next;
               continue;
            end;

            if (planJbm.dtsReclamadas.Tag = 1) then
            begin
               ExibeMensagem('   Tratando processo trabalhista (Empresa Terceira)');
               TrataProcessoTrabalhista(2);
               Application.ProcessMessages;
               planJbm.dtsAtos.Next;
               continue;
            end;

            if (planJbm.dtsReclamadas.Tag = 0) then
            begin
               planJbm.GravaOcorrencia(2, 'Não encontrou o processo na tabela Relatorios_Gerenciais_DADOS_TA_TO_TR');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if (planJbm.dtsReclamadas.Tag = 99) then
            begin
               planJbm.GravaOcorrencia(2, 'Encontrou mais de um registro na tabela Relatorios_Gerenciais_DADOS_TA_TO_TR');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            planJbm.GravaOcorrencia(2, 'Não encontrou regra para eset tipo de processo: ' + planJbm.tipoProcesso);
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'LEVANTAMENTO DE ALVARA') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ENTREGA AMIGAVEL') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'DESBLOQUEIOS') then
         begin
            planJbm.GravaOcorrencia(2, 'Tipo de ato devido somente para ações ativas');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         END;
      end;

      ExibeMensagem('   Não é DRC Ativas');
      Application.ProcessMessages;

      //alterado conforme solicitado pela Samara e Elaine em 17/01/2017

      if (planJbm.dtsAtos.FieldByName('tipoacao').AsString = 'CARTA CONVITE') and (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'CARTA CONVITE') then
      begin
         planJbm.GravaOcorrencia(2, 'Não pode pagar tipo de andamento para ação do tipo CARTA CONVITE');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      //verifica duplicidade
      //ACOMPANHAMENTO trabalhista não é preciso checar duplicidade de pagamento
      //regra alterada em 03/04/2014 - email da Amanda de 21/03/2014
      //if (Pos('PREPOSTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0) and
      if (Pos('ACORDO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0) and
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'ACOMPANHAMENTO') then
      begin
         ExibeMensagem('   Verificando se tem duplicidade');
         Application.ProcessMessages;

         if TemDuplicidade = true then
         begin
            ExibeMensagem('   Duplicidade encontrada');
            Application.ProcessMessages;
            planJbm.dtsAtos.Next;
            continue;
         end;
      end;

      if (planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger <> 1) then
      begin
         if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PREPOSTO' then
         begin
            if (planjbm.tipoProcesso = 'TR') or
               (planjbm.tipoProcesso = 'TO') or
               (planjbm.tipoProcesso = 'TA') then //se for processo trabalhista não pode pagar PREPOSTO - E-mail da Nilce de 11/03/2014
            begin
               planJbm.GravaOcorrencia(2, 'Não pode cobrar PREPOSTO em processo trabalhista');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;

         //carta precaotoria

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CARTA PRECATORIA') OR
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CUMPRIMENTO DE CARTA PRECATORIA') then
         begin
            //CARTA PRECATORIA: Será devido o valor de R$438,00 por carta precatória. Não a limite de atos. O processo não deve estar distribuído para o Escritório que está
            //cobrando, caso positivo, recusar.
            //valida se o ato foi distribuido para este excritorio. Se foi recusa
            ret := planJbm.AtoDistribuidoParaEscritorio;
            if ret = 9 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar porque o processo está distribuído para este escritório');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 667.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 721.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
               valorConferir := 591.00
            else
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
               valorConferir := 619.00
            else
//pereira
               valorConferir := 557.00;


//alterado em 26/11/2014
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorConferir then
            begin
               valorcorreto := valorConferir;
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               valorcorreto := planjbm.dtsatos.FieldByName('valor').AsFloat;
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.MarcaProcessoCruzadoGcpj(6);
            planJbm.dtsAtos.Next;
            continue;
         end;

         //tem recurso em juizado
         if planjbm.dtsatos.FieldByName('fgjuizado').AsInteger = 1 then
         begin
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIêNCIA AVULSA') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIENCIA AVULSA') then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode cobrar AUDIENCIA AVULSA em processo juizado');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
               (Pos('AGRAVO FAT', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
            begin

               //alterado de acordo com e-mail da nilce de 10/03/2015
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2015/02/18' then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode cobrar RECURSO em processo juizado');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               listaEscritorios := TStringList.Create;
               try
                  listaEscritorios.LoadFromFile(ExtractFilePath(Application.ExeName) + 'ESCRITORIOS_TARIFAS.TXT');

                  if Not planJbm.EhEscritorioTarifa(listaEscritorios) then
                  begin
                     nomeOrgaoJulgador := planJbm.ObtemNomeOrgaoJulgador(planJbm.ObtemCodOrgProcesso);
                     if (Pos('JEC', nomeOrgaoJulgador) <> 0) or
                        (Pos('JUIZADO ESPECIAL', nomeOrgaoJulgador) <> 0)  then
                        planJbm.GravaOcorrencia(2, 'Órgão julgador Juizado')
                     else
                        planJbm.GravaOcorrencia(2, 'Não pode cobrar RECURSO em processo juizado');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;

                  //se for JEC para Bradesco financiamentos, para alguns escritorios pode deixar pagar.
                  listaTarifas := TstringList.Create;
                  try
                     listaTarifas.LoadFromFile(ExtractFilePath(Application.ExeName) + 'SUBTIPOS_TARIFAS.TXT');
                     if Not planJbm.EhSubtipoTarifa(listaTarifas) then
                     begin
                        planJbm.GravaOcorrencia(2, 'Não pode cobrar RECURSO em processo juizado');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        planJbm.dtsAtos.Next;
                        continue;
                     end;
                  finally
                     listaTarifas.Free;
                  end;
               finally
                  listaEscritorios.free;
               end;
            end;
         end;

         if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORARIOS INICIAIS') or
            (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORÁRIOS INICIAIS') then
         begin
(*** Removido - e-mail da Nilce de 19/02/2016
            if (planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 1) and (FormatDateTime('yyyymmdd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) > '20130421') then
            begin
               planJbm.GravaOcorrencia(2, 'Ato não devido após o dia 21/04/2013');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
            **)

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, 0, totalPago, planJbm.codGcpjEscritorio, 'CONTESTACAO');
            if somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'Já houve pagamento de CONTESTACAO');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;


            //verifica se já houve pagamento para o mesmo escritório de honorários iniciais - 19/12/2016
            ExibeMensagem('     ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString + ' - ' + IntToStr(planJbm.codGcpjEscritorio));
            SomaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger,
                                                                           totalPago, planJbm.codGcpjEscritorio,
                                                                           planjbm.dtsatos.FieldByName('tipoandamento').AsString, false, false);

            if (totalpago >= 1) or  (somaAtosDescontar > 0) then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2, 'Já houve pagamento de HONORÁRIOS INICIAIS para este escritório (Base XI)');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if cbSemReajuste.Checked then
            begin
               valorCorreto := planJbm.ObtemValorSemReajuste('HONORARIOS INICIAIS');
               if valorCorreto = 0 then
               begin
                  ExibeMensagem('   Programa está calculando valor igual a zero');
                  Application.ProcessMessages;
                  exit;
               end;
            end
            else
            begin
//pereira
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
                  valorCorreto := 65.00
               else
//pereira
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                  valorCorreto := 73.00
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
                  valorCorreto := 70.00
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2015/05/01' then
                  valorCorreto := 64.00
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2014/05/01' then
                  valorCorreto := 59.00
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2013/05/01' then
                   valorCorreto := 56.00
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01' then
                  valorCorreto := 52.00
               else
                  valorCorreto := 50.00;
            end;

            if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
            begin
               planJbm.GravaValorCorrigido(valorcorreto);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
             planJbm.MarcaProcessoCruzadoGcpj(6);
             planJbm.dtsAtos.Next;
             continue;
         end;

         if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORARIOS FINAIS') or
            (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORÁRIOS FINAIS') then
         begin
             //removido conforme solicitação da Samara e Elaine em 17/01/2017
             //if (planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 1) and (FormatDateTime('yyyymmdd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) > '20130421') then
             //begin
             //   planJbm.GravaOcorrencia(2, 'Ato não devido após o dia 21/04/2013');
             //   planJbm.MarcaProcessoCruzadoGcpj(9);
             //   planJbm.dtsAtos.Next;
             //   continue;
             //end;

            //verifica se já houve pagamento para o mesmo escritório de honorários finais - 19/12/2016
            SomaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, planJbm.codGcpjEscritorio,
                                                                           planjbm.dtsatos.FieldByName('tipoandamento').AsString, false, false);

            if totalpago >= 1 then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2, 'Já houve pagamento de HONORÁRIOS FINAIS para este escritório (Base XI)');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

             if cbSemReajuste.Checked then
             begin
                valorCorreto := planJbm.ObtemValorSemReajuste('HONORARIOS FINAIS'); //56.00
               if valorCorreto = 0 then
               begin
                  ExibeMensagem('   Programa está calculando valor igual a zero');
                  Application.ProcessMessages;
                  exit;
               end;
             end
             else
             begin
//pereira
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
                  valorCorreto := 65.00
               else
//pereira
                if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                   valorCorreto := 73.00
                else
                if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
                   valorCorreto := 70.00
                else
                if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2015/05/01' then
                   valorCorreto := 64.00
                else
                if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2014/05/01' then
                   valorCorreto := 59.00
                else
                if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2013/05/01' then
                   valorCorreto := 56.00
                else
                if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01' then
                   valorCorreto := 52.00
                else
                   valorCorreto := 50.00;
             end;

             //19/04/2017
             (**
             if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'EXTINTO COM JULGAMENTO DE MERITO') or
                (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO COM CUSTOS') or
                (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'IMPROCEDENCIA') then
             begin
               if planJbm.valorpedido <= 5000 then
                  valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.15))
               else
                  valorcorreto := StrToFloat(FormatFloat('0.00', (5000 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.15));
             end;

             if valorcorreto <= 0 then
             begin
               planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
             end;                        **)

             if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
             begin
                planJbm.GravaValorCorrigido(valorcorreto);
                planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
             end
             else
                planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
             planJbm.MarcaProcessoCruzadoGcpj(6);
             planJbm.dtsAtos.Next;
             continue;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIêNCIA AVULSA') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIENCIA AVULSA') then
         begin
            if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2015/05/01') then
               valorcorreto := 357.00
            else
            if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2014/05/01') then
               valorcorreto := 330.00
            else
            if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2013/05/01') then
               valorcorreto := 311.00
            else
            if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01') then
               valorcorreto := 290.00
            else
               valorCorreto := 279.00;

            if (planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto) then
            begin
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               planJbm.GravaValorCorrigido(valorcorreto);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;

         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'CARTA CONVITE') then
         begin
            if planJbm.dtsAtos.FieldByName('tipoacao').AsString <> 'CARTA CONVITE' then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar CARTA CONVITE para tipo de ação: '  + planJbm.dtsAtos.FieldByName('tipoacao').AsString);
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;


            //distribuido para este escritorio?
            ret := planJbm.AtoDistribuidoParaEscritorio;
            if ret <> 9 then
            begin
               if ret = 0 then
                  planJbm.GravaOcorrencia(2, 'Processo não distribuído')
               else
                  planJbm.GravaOcorrencia(2, 'Processo distribuído para outro escritório');

               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then
               valorCorreto := 340.00
            else
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
               valorCorreto := 378.00
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2016/05/01' then
               valorCorreto := 361.00
            else
               valorcorreto := 330.00;


            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj,
                                                                           planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger,
                                                                           totalPago,
                                                                           planJbm.codGcpjEscritorio,
                                                                           'CARTA CONVITE',
                                                                           false,
                                                                           false);
            IF somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'CARTA CONVITE já foi paga para este escritório');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj,
                                                                           planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger,
                                                                           totalPago,
                                                                           planJbm.codGcpjEscritorio,
                                                                           'CONTESTACAO',
                                                                           false,
                                                                           false);
            IF somaAtosDescontar > 0 then
            begin
               planJbm.GravaOcorrencia(2, 'CARTA CONVITE não pode ser paga. Já houve pagamento de volumetria');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj,
                                                                           planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger,
                                                                           totalPago,
                                                                           planJbm.codGcpjEscritorio,
                                                                           'PREPOSTO',
                                                                           false,
                                                                           false);
            IF (totalPago > 2) then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar CARTA CONVITE. Já foram pagos mais de 2 prepostos');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if (planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto) then
            begin
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               planJbm.GravaValorCorrigido(valorcorreto);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end
            else
             begin
              //alexandre 26/10/2017
              If valorcorreto <> 0 then
               begin
                planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                planJbm.GravaValorCorrigido(valorcorreto);
               end;
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
              //alexandre 26/10/2017
             end;

         end;

      end; //DRCATIVAS <> 1


      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'ADESAO POUPADOR') then
      begin
         //ADESAO POUPADOR (AÇÕES CONTRÁRIAS DRC ou Ativas): pagar 200 por poupador
         ExibeMensagem('   Obtendo Poupador para pagamento por Ato');
         Application.ProcessMessages;
         ret := planJbm.ObtemEnvolvidosDoProcesso_2(planjbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger, false);

         if ret = 1 then
         begin
           valorCorreto := 200;
           planJbm.GravaOcorrencia(2, 'Valor pago por poupador R$ 200,00');
           planJbm.GravaValorCorrigido(valorcorreto);
           planJbm.MarcaProcessoCruzadoGcpj(6);
           planJbm.dtsAtos.Next;
           continue;
         end
         else
         begin
           planJbm.GravaOcorrencia(2, 'GCPJ não encontrado na tabela BaseDiaria_GCPJB076, verificar.');
           planJbm.MarcaProcessoCruzadoGcpj(9);
           planJbm.dtsAtos.Next;
           continue;
         end;
      end;

      if (Pos('EXITO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
         (Pos('ÊXITO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         if planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 0 then
         begin
            if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'EXTINTO SEM JULGAMENTO DE MERITO') or
               (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO SEM CUSTOS') or
               (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'LIQUIDACAO') or
               (planJbm.dtsAtos.FieldByName('motivobaixa').AsString = 'DESISTENCIA DA ACAO') or
               (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'PROCEDENTE SEM CUSTOS') or
               (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'APENSADO') or
               (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'TRANSFERENCIA PARA DRC') or
               (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ARQUIVADO') or
               (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'DUPLICIDADE') or
               (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'SOLUCIONADO') or
               (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'PROCEDENTE SEM CUSTOS') or
               (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'EXTINCAO DO PROCESSO') or
               (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'IMPLANTACAO DE SENTENCA') or
               (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'CADASTRO INDEVIDO PELA DEPENDENCIA') or
               (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'CADASTRO IRREGULAR PELA DEPENDENCIA') then
            begin
               planJbm.GravaOcorrencia(2, 'EXITO não pode ser pago por causa do motivo da baixa');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;

         //se for do DRC tem uma regra especial para tratar
         if planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 1 then
         begin
            //tem que estar ajuizado
(** retirada por instrução da Nilce - Telefone - 22/01/2015
            if planjbm.dtsatos.FieldByName('fgjuizado').AsInteger <> 1 then
            begin
               //texto alterado - Email da Nilce de 19/05/2014
//               planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) para processo não ajuizado');
               planJbm.GravaOcorrencia(2, 'Não pode pagar êxito DRC ações Contrárias');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;                                           **)

            if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORARIOS DE EXITO') or
                (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'HONORÁRIOS DE ÊXITO')) then
            begin
               if (Not planjbm.dtsatos.FieldByName('motivobaixa').IsNull) And
                  (Pos('ACORDO', planjbm.dtsatos.FieldByName('motivobaixa').AsString) = 0) then
               begin
                  planJbm.GravaOcorrencia(2, 'Não há previsão de pagamento para o motivo da baixa');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               //HONORARIOS DE EXITO (AÇÕES CONTRÁRIAS DRC): 10% sobre o valor base. Descontando-se valores já adiantados e não poderá ultrapassar o valor máximo de
               //R$85.955,00. Poderá haver mais que um ato pago, devido a acordo parcelado.
               //Verificar também a data do ato do acordo parcelado, tendo em vista que não poderá coincidir com atos de acordos  já pagos.
               ExibeMensagem('   Obtendo valor já pago para o processo');
               Application.ProcessMessages;

               (**Tratar igual ACORDO ativas
               somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj);
               valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);

               if somaAtosDescontar > 0 then
                  valorcorreto := valorCorreto - somaAtosDescontar;***)

               somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger , totalPago, 0, 'HONORARIOS DE EXITO');
               if somaAtosDescontar = 0 then
               begin
                  //verifica se na mesma planilha tem acordo validado (ou seja pode pagar)
                  somaAtosDescontar := dmHonorarios.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('tipoandamento').AsString, planJbm.dtsAtos.FieldByName('idplanilha').AsInteger);
                  if somaAtosDescontar = 0 then
                  begin
                     //se não pagou acordo ainda, descontar outros adiantamentos
                     somaAtosDescontar := dmgcpcj_base_XI.ObtemSomaJaPagaADescontar(planJbm.gcpj, planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger, totalPago, 0);


                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
                     begin
                       valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
                     end
                     else
                      begin
                       valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
                      end;

                     if somaAtosDescontar > 0 then
                        valorcorreto := valorCorreto - somaAtosDescontar;
                  end
                  else

                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
                  begin
                    valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
                  end
                  else
                   begin
                    valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
                   end;
               end
               else

               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2018/06/01' then //utlizar a rotina de novo cálculo
               begin
                 valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.08);
               end
               else
                begin
                 valorcorreto := (planJbm.dtsAtos.FieldByName('valorbase').AsFloat * 0.10);
                end;

               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 85955 then
                     valorcorreto := 85955-somaAtosDescontar;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 91352 then
                        valorcorreto := 91352-somaAtosDescontar;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 98817 then
                           valorcorreto := 98817-somaAtosDescontar;
                     end
                     else
                     begin
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                        begin
                           if valorcorreto > 107986 then
                              valorcorreto := 107986-somaAtosDescontar;
                        end
                        else
                        begin
//pereira
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                        begin
                           if valorcorreto > 113080 then
                              valorcorreto := 113080-somaAtosDescontar;
                        end
                        else
//pereira
                           if valorcorreto > 101772 then
                              valorcorreto := 101772-somaAtosDescontar;
                        end;
                     end;
                  end;
               end;

               if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorCorreto)) then
               begin
                  if ((planjbm.dtsatos.FieldByName('valor').AsFloat > valorCorreto) and ((planjbm.dtsatos.FieldByName('valor').AsFloat - valorCorreto) > 0.55)) or
                     ((valorCorreto > planjbm.dtsatos.FieldByName('valor').AsFloat) and ((valorCorreto - planjbm.dtsatos.FieldByName('valor').AsFloat) > 0.55)) then
                  begin
                     planJbm.GravaOcorrencia(2, 'Valor não compatível com tipo de andamento');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;
               end;
               planJbm.GravaValorCorrigido(valorcorreto);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //verifica se foi acordo parcelado
(**            if (planjbm.dtsatos.FieldByName('motivobaixa').IsNull) then //não está baixado. verificamos se é ACORDO PARCELADO
            begin
               ExibeMensagem('   Obtendo eventos do processo');
               Application.ProcessMessages;

               //obtem os eventos para verificar se é parcelamento
               dmgcpj_base_VIII.ObtemEventosDoProcesso(planJbm.gcpj);

               if dmgcpj_base_VIII.dts.RecordCount = 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) - Não baixado e sem evento para verificar parcelamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               //verificar como tratar parcelado
               ExibeMensagem('Encontrou eventos de parcela - comunicar presenta sistemas');
               planJbm.dtsAtos.Next;
               continue;
            end
            else
            begin**)
               //está baixado verificamos se é acordo a vista
               if (Pos('ACORDO A VISTA', planjbm.dtsatos.FieldByName('motivobaixa').AsString) <> 0) then
               begin
                  if (planjbm.dtsatos.FieldByName('motivobaixa').IsNull) and
                     (planjbm.dtsatos.FieldByName('databaixa').IsNull) then
                  begin
                     planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) - ACORDO A VISTA para processo não baixado');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;

                  if planjbm.dtsatos.FieldByName('valorbaixa').AsFloat <= 0.00 then
                  begin
                     planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) - ACORDO A VISTA com valor de cálculo zerado');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;

                  valorcorreto := (planjbm.dtsatos.FieldByName('valorbaixa').AsFloat * 10) / 100;
                  if valorcorreto > 50000.00 then
                     valorCorreto := 50000.00;

                  if valorcorreto <= 0 then
                  begin
                     planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;

                  if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                  begin
                     planJbm.GravaValorCorrigido(valorcorreto);
                     planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                  end
                  else
                     planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                  planJbm.MarcaProcessoCruzadoGcpj(6);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            //end;
         end //drccontrarias = 1
         else
         begin
            if planJbm.dtsAtos.FieldByName('fgDrcAtivas').AsInteger = 1 then
            begin
               if planjbm.dtsatos.FieldByName('fgjuizado').AsInteger <> 1 then
               begin
               //texto alterado - Email da Nilce de 19/05/2014
//               planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) para processo não ajuizado');
                  planJbm.GravaOcorrencia(2, 'Não pode pagar êxito DRC ações Contrárias');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if (Not planjbm.dtsatos.FieldByName('motivobaixa').IsNull) then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) - para processo baixado');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if planJbm.valorpedido <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) - com valor de cálculo zerado');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               valorcorreto := (planJbm.valorpedido * 10) / 100;
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               begin
                  if valorcorreto > 85955 then
                     valorcorreto := 85955;
               end
               else
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
                  begin
                     if valorcorreto > 91352 then
                        valorcorreto := 91352;
                  end
                  else
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                     begin
                        if valorcorreto > 98817 then
                           valorcorreto := 98817;
                     end
                     else
                     begin
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                        begin
                           if valorcorreto > 107986 then
                              valorcorreto := 107986;
                        end
                        else
                        begin
//pereira
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                        begin
                           if valorcorreto > 113080 then
                              valorcorreto := 113080;
                        end
                        else
//pereira

                           if valorcorreto > 101772 then
                              valorcorreto := 101772;
                        end
                     end;
                  end;
               end;

               if valorcorreto <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário (DRC)');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
               begin
                  planJbm.GravaValorCorrigido(valorcorreto);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end
            else
            begin
               if ((planjbm.dtsatos.FieldByName('motivobaixa').IsNull) and
                  (planjbm.dtsatos.FieldByName('databaixa').IsNull) and
                  (Pos('PLANOS ECONOMICOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) = 0)) then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO para processo não baixado');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               //se for acord com custo verifica se o benefício é menor do que 5000
               if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO COM CUSTOS') or
                  (Pos('PLANOS ECONOMICOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0) then
               begin
                 if planJbm.valorpedido <= 5000 then
                    valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.08))
                 else
                    // solicitação da SAMARA email de 06/08/19
                    valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.08));
                    //valorcorreto := StrToFloat(FormatFloat('0.00', (5000 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.08));

                 if (valorcorreto <= 199.99) and (Pos('PLANOS ECONOMICOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0) then
                 Begin
                   valorCorreto := 200;
                   planJbm.GravaValorCorrigido(valorcorreto);
                   planJbm.MarcaProcessoCruzadoGcpj(6);
                   planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                   planJbm.dtsAtos.Next;
                   continue;
                 end;

                 if valorcorreto <= 0 then
                 begin
                    planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                    planJbm.MarcaProcessoCruzadoGcpj(9);
                    planJbm.dtsAtos.Next;
                    continue;
                 end;

                 if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                 begin
                  If valorCorreto > 400 then
                      valorCorreto := 400;
                   planJbm.GravaValorCorrigido(valorcorreto);
                   planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                 end
                 else
                   planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                   planJbm.MarcaProcessoCruzadoGcpj(6);
                   planJbm.dtsAtos.Next;
                   continue;
               end;

{               if (planjbm.dtsatos.FieldByName('motivobaixa').AsString <> 'ACORDO COM CUSTOS') then  //VERIFICAR
               begin
                if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de antigo cálculo
                  valorCorreto := 750
                else
                  valorCorreto := 400;


                 if valorcorreto <= 0 then
                 begin
                    planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                    planJbm.MarcaProcessoCruzadoGcpj(9);
                    planJbm.dtsAtos.Next;
                    continue;
                 end;

                 if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                 begin
                   planJbm.GravaValorCorrigido(valorcorreto);
                   planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                 end
                 else
                   planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                   planJbm.MarcaProcessoCruzadoGcpj(6);
                   planJbm.dtsAtos.Next;
                   continue;
               end;}


               if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'IMPROCEDENCIA') or
                  (planJbm.dtsAtos.FieldByName('motivobaixa').AsString = 'EXTINTO COM JULGAMENTO DE MERITO') then
               begin
                  if Pos('PLANOS ECONOMICOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0 then
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2011/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 309.14
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
                        valorCorreto := 342
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 355
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 381
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 405
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 438
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 479
                     else
//pereira
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                        valorcorreto := 502
                     else
//pereira
                        valorcorreto := 452;

                     if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                     begin
                        planJbm.GravaValorCorrigido(valorcorreto);
                        planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                     end
                     else
                        planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                     planJbm.MarcaProcessoCruzadoGcpj(6);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;

                  if Pos('ACAO DE REPARACAO DE DANOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0  then
                  begin
                     if planJbm.dtsAtos.FieldByName('subtipoacao').AsString = 'OBRIGACAO DE FAZER' then
                     begin
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
                           valorCorreto := 342
                        else
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
                           valorCorreto := 355
                        else
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
                           valorCorreto := 381
                        else
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
                           valorCorreto := 405
                        else
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                           valorCorreto := 438
                        else
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                           valorCorreto := 479
                        else
    //pereira
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                            valorcorreto := 502
                        else
    //pereira
                           valorcorreto := 452;



                        if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                        begin
                           planJbm.GravaValorCorrigido(valorcorreto);
                           planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                        end
                        else
                           planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                        planJbm.MarcaProcessoCruzadoGcpj(6);
                        planJbm.dtsAtos.Next;
                        continue;
                     end;
                  end;

                  if planjbm.valorpedido <= 5000 then
                     valorCorreto := StrToFloat(FormatFloat('0.00', planjbm.valorpedido * 0.08))
                  else
//pereira200
                    if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de antigo cálculo
                      valorCorreto := 750
                    else
                      valorCorreto := 400;


                  if valorcorreto <= 0 then
                  begin
                     planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;

                  if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                  begin
                     planJbm.GravaValorCorrigido(valorcorreto);
                     planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                  end
                  else
                     planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                  planJbm.MarcaProcessoCruzadoGcpj(6);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'EXTINTO COM JULGAMENTO DE MERITO' then
               begin
                  if (Pos('PLANOS ECONOMICOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0) or
                     (planjbm.dtsatos.FieldByName('tipoacao').AsString = 'OBRIGACAO DE FAZER') then
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
                        valorCorreto := 342
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 355
                     else
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
                        valorcorreto := 381
                     else
                        valorcorreto := 405;
                  end
                  else
                  begin
                     if planjbm.valorpedido <= 5000 then
                        valorCorreto := StrToFloat(FormatFloat('0.00', planjbm.valorpedido * 0.08))
                     else
                       //pereira2
                      if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de antigo cálculo
                        valorCorreto := 750
                      else
                        valorCorreto := 400;
                  end;

                  if valorcorreto <= 0 then
                  begin
                     planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;

                  if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                  begin
                     planJbm.GravaValorCorrigido(valorcorreto);
                     planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                  end
                  else
                     planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                  planJbm.MarcaProcessoCruzadoGcpj(6);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;
      end;

      if (Pos('ACOMPAN', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) and
         (Pos('DELEGACIA', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         planJbm.GravaOcorrencia(2, 'Não pode pagar ACOMPANHAMENTO EM DELEGACIA');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'PREPOSTO' then
      begin
         if cbSemReajuste.Checked then
         begin
            valorCorreto := planJbm.ObtemValorSemReajuste('PREPOSTO');
            if valorCorreto = 0 then
            begin
               ExibeMensagem('   Programa está calculando valor igual a zero');
               Application.ProcessMessages;
               exit;
            end;
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 57
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 60
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 64
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 68
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 74
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 81
            else
//pereira
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
               valorCorreto := 85
            else
//pereira
               valorCorreto := 85;
         end;

         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      //inserido em 03/04/2014 - e-mail da Nilce de 25/03/2014
      if (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0)then
      begin
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < '2013/04/22' then
         begin
            //texto alterado a pedido da Nilce
            //planJbm.GravaOcorrencia(2, 'Não pode pagar AJUIZAMENTO anterior a 22/04/2013');
            planJbm.GravaOcorrencia(2, 'Fato Gerador anterior a vigência do Contrato, clausula 17.21.2');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;
      end;

      if Not cbSemReajuste.checked then
      begin
         if (Pos('TUTELA', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
         //alterado em 30/10/2013
            (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0)then
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2011/11/01' then //utlizar a rotina de novo cálculo
            begin
               planJbm.MarcaProcessoCruzadoGcpj(5);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;
      end;

      if (Pos('TUTELA',planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
       //alterado em 30/10/2013
         (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
         (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) OR
         (Pos('AGRAVO FAT', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         if cbSemReajuste.Checked then
         begin
            if (Pos('TUTELA',planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
               andamento := 'TUTELA'
            else
            if (Pos('AJUIZAMENTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
               andamento := 'AJUIZAMENTO'
            ELSE
            if (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
               andamento := 'RECURSO'
            else
               andamento := '';
            valorCorreto := planJbm.ObtemValorSemReajuste(andamento); //470

            if valorCorreto = 0 then
            begin
               ExibeMensagem('   Programa está calculando valor igual a zero');
               Application.ProcessMessages;
               exit;
            end;

            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
            begin
               planJbm.GravaValorCorrigido(valorcorreto);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
            planJbm.MarcaProcessoCruzadoGcpj(6);
            planJbm.dtsAtos.Next;
            continue;
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2011/05/01' then //utlizar a rotina de novo cálculo
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
                  valorCorreto := 422
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
                  valorCorreto := 438
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
                  valorCorreto := 470
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
                  valorCorreto := 500
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
                  valorCorreto := 541
               else
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
                  valorCorreto := 591
               else
//PEREIRA
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2018/05/31' then //utlizar a rotina de novo cálculo
                  valorCorreto := 619
               else
                  valorCorreto := 557;
//PEREIRA

               if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
               begin
                  planJbm.GravaValorCorrigido(valorcorreto);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            End;

            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2010/05/01' then //utlizar a rotina de novo cálculo
            begin
               if planjbm.dtsatos.FieldByName('valor').AsFloat <> 382 then
               begin
                  planJbm.GravaValorCorrigido(382);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if planjbm.dtsatos.FieldByName('valor').AsFloat <> 370 then
            begin
               planJbm.GravaValorCorrigido(370);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
            planJbm.MarcaProcessoCruzadoGcpj(6);
            planJbm.dtsAtos.Next;
            continue;
         end;
      end;

      if (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'AVULSO TRABALHISTA') or
         (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'AUDIÊNCIA TRABALHISTA') or
         (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'AUDIÊNCIA AVULSA') or
         (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'INSTRUCAO') then
      begin
         if (planJbm.tipoProcesso <> 'TR') AND (planJbm.tipoProcesso <> 'TO') AND (planJbm.tipoProcesso <> 'TA') then
         begin
            planJbm.GravaOcorrencia(2, 'Tipo de ato somente permitido para processo TRABALHISTA');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 564
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 585
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 628
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 667
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 721
         else
            valorCorreto := 788;


         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CÍVEL (AUDITORIA)') or
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CÍVEL') or
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CIVEL') OR
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CIVEL (AUDITORIA)')then
      begin
         if planJbm.tipoProcesso <> 'CI' then
         begin
            planJbm.GravaOcorrencia(2, 'Tipo de ato somente permitido para processo CIVEL');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 279
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 290
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 311
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 330
         else
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 357
         else
            valorCorreto := 390;



         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if (planJbm.dtsAtos.FieldByName('fgDrcContrarias').AsInteger = 1) and (Pos('XITO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         planJbm.GravaOcorrencia(2, 'Para ações Revisionais são devidos êxito somente em acordo');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      planJbm.GravaOcorrencia(2, 'Não achou regra no programa para validar o pagamento deste ato');
      planJbm.MarcaProcessoCruzadoGcpj(9);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.CalculaNumeroDaNota;
var
   valorCobrar : double;
   index : integer;
   limiteAtos : Integer;
   ini : TInifile;
begin
   re.lines.Clear;

   processando.Text := 'Calculando valor das notas';
   Application.ProcessMessages;

   //verifica o limite de atos por nota
   ini := TiniFile.Create(ExtractFilePath(Application.Exename) + 'CONFIG.INI');
   try
      limiteAtos := ini.ReadInteger('Notas', 'LimiteDeAtos', 899);
   finally
      ini.Free;
   end;


   planJbm.ObtemProcessosCruzarGcpj(3);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
   begin
      ExibeMensagem('Não foi possível calcular as notas. Ainda existe processo não validado. Motivo: 3');
      exit;
   end;

   planJbm.ObtemProcessosCruzarGcpj(4);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
   begin
      ExibeMensagem('Não foi possível calcular as notas. Ainda existe processo não validado. Motivo: 4');
      exit;
   end;

   planJbm.ObtemProcessosCruzarGcpj(5);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
   begin
      ExibeMensagem('Não foi possível calcular as notas. Ainda existe processo não validado. Motivo: 5');
      exit;
   end;

   for index := 1 to 2 do
   begin
      planJbm.LimpaDadosNota;
      planJbm.ObtemProcessosCruzarGcpj(6);

      while not planJbm.dtsAtos.Eof do
      begin
         planJbm.CarregaCamposDaTabela;
         planJbm.GuardaTipoProcesso;

         //index=1 processos trabalhistas
         if ((planJbm.tipoProcesso = 'TR') OR (planJbm.tipoProcesso = 'TO') OR (planJbm.tipoProcesso = 'TA')) and (index = 2) then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end;

         //index=2 processos civeis
         if (planJbm.tipoProcesso = 'CI') and (index = 1) then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end;

         ExibeMensagem('Calculando número da nota para o processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
         if (planJbm.empresaLigada <> planjbm.dtsAtos.FieldByName('empresaligadaagrupar').ASInteger) and
            (planJbm.empresaLigada <> 0) then
         begin
            planJbm.GravaValorTotalDaNota;
            planJbm.LimpaDadosNota;
         end;

         if planJbm.empresaLigada = 0 then
         begin
            planJbm.empresaLigada := planjbm.dtsAtos.FieldByName('empresaligadaagrupar').ASInteger;
            planJbm.ObtemNumeroNotaEmpresa;
         end;

         valorCobrar := planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat;
         if valorCobrar = 0 then
         begin
            ExibeMensagem('Erro inesperado: valor do ato ficou zerado');
            planJbm.GravaOcorrencia(2, 'Erro inesperado: valor do ato ficou zerado');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            continue;
         end;

//regra alterada pelo e-mail da Samara em 05/10/2016
//         if ((planJbm.valorTotalNota + valorCobrar) > 130000) or
//            (planjbm.totalAtos > 150) then

         if (planJbm.tipoProcesso = 'TR') OR (planJbm.tipoProcesso = 'TO') OR (planJbm.tipoProcesso = 'TA') then
         begin
            if ((planJbm.valorTotalNota + valorCobrar) > 200000) or
               (planjbm.totalAtos > limiteAtos) then
            begin
               //finaliza a nota por causa do valor
               planJbm.GravaValorTotalDaNota;
               planjbm.EncerraNotaPeloValor;
               planJbm.LimpaDadosNota;
               planJbm.empresaLigada := planjbm.dtsAtos.FieldByName('empresaligadaagrupar').ASInteger;
               planJbm.ObtemNumeroNotaEmpresa;
            end;
         end
         else
         begin
            if ((planJbm.valorTotalNota + valorCobrar) > 200000) or
                (planjbm.totalAtos > limiteAtos) then
            begin
               //finaliza a nota por causa do valor
               planJbm.GravaValorTotalDaNota;
               planjbm.EncerraNotaPeloValor;
               planJbm.LimpaDadosNota;
               planJbm.empresaLigada := planjbm.dtsAtos.FieldByName('empresaligadaagrupar').ASInteger;
               planJbm.ObtemNumeroNotaEmpresa;
            end;
         end;

         planJbm.valorTotalNota := planJbm.valorTotalNota + valorCobrar;
         planJbm.totalAtos := planJbm.totalAtos + 1;
         planJbm.GravaNotaNoProcesso;

         planJbm.dtsAtos.Next;

         Application.ProcessMessages;
      end;

      if (planJbm.empresaLigada <> 0) then
         planJbm.GravaValorTotalDaNota;
   end;

   //já fiinalizado?
   planJbm.ObtemProcessosCruzarGcpj(9);
   if planJbm.dtsAtos.RecordCount = 0 then
      planJbm.MarcaPlanilhaValidada;

end;

procedure TfrmValidaPlan.BitBtn3Click(Sender: TObject);
var
   arquivo : TstringList;
   linha : string;
   ini : TIniFile;
   contador : integer;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   re.lines.Clear;

   planJbm.ObtemProcessosCruzarGcpj(7);

   arquivo := TStringList.Create;
   try
      arquivo.add('Ano/Mês;Escritório;Cód. Empresa Grupo;Nome Empresa Grupo;Nota;GCPJ;Parte Contrária;Tipo Ato;Tipo Ação;Sub Tipo Ação;Data Baixa;Motivo Baixa;Data do Ato;Valor do Pedido;Valor Pagar;'+
                  'Cod.Reclamada1;Nome Reclamada1;Cod.Reclamada2;Nome Reclamada2;Cod.Reclamada3;Nome Reclamada3;Cod.Reclamada4;Nome Reclamada4');
      while not planJbm.dtsAtos.Eof do
      begin
         planJbm.CarregaCamposDaTabela;
         ExibeMensagem('Exportando: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
         Application.ProcessMessages;

         linha := planJbm.anomesreferencia + ';' + planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString + ';' + planJbm.dtsAtos.FieldByName('nomeempresaligada').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('numeronota').AsString + ';' + planJbm.dtsAtos.FieldByName('gcpj').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('partecontraria').AsString + ';' + planJbm.dtsAtos.FieldByName('tipoandamento').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('tipoacao').AsString + ';' + planJbm.dtsAtos.FieldByName('subtipoacao').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('databaixa').AsString + ';' + planJbm.dtsAtos.FieldByName('motivobaixa').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('datadoato').AsString + ';' + FormatFloat('#,##0.00', planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat) + ';' +
                  FormatFloat('#,##0.00', planJbm.dtsAtos.FieldByName('valordopedido').AsFloat);

         planJbm.ObtemReclamadasDoProcesso('');
         if planJbm.dtsReclamadas.Eof then
            linha := linha + ';;;;;;;';

         contador := 0;
         while not planJbm.dtsReclamadas.Eof do
         begin
            Inc(contador);
            if contador > 4 then
               break;
            linha := linha + ';' + planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString + ';' + planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
            planjbm.dtsReclamadas.Next;
         end;
         arquivo.Add(linha);
         planJbm.dtsAtos.Next;
      end;
      arquivo.SaveToFile(ExtractFilePath(Application.ExeName) + 'NOTAS.TXT');
   finally
      arquivo.free;
   end;
end;

procedure TfrmValidaPlan.BitBtn5Click(Sender: TObject);
var
   ini : TIniFile;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;

   re.lines.Clear;

   planJbm.ObtemProcessosCruzarGcpj(-1);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem('Verificando Empresa Ligada: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;
      planJbm.GravaEmpresaGrupo(1);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.BitBtn6Click(Sender: TObject);
var
   ini : TIniFile;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');

   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;

   re.lines.Clear;

   planJbm.ObtemProcessosCruzarGcpj(-1);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem('Revisando duplicidades: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

(*      if (pos('TRABALHISTA', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
         (Pos('TRABALHO', planJbm.dtsatos.FieldByName('Vara').AsString) <> 0) or
         (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'INSTRUCAO') then
         planjbm.tipoprocesso := 'T'
      else
         planjbm.tipoprocesso := 'C';*)

//      if Pos('PREPOSTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0 then
//      begin
         dmgcpcj_base_XI.ObtemOutrosPagamentosDoProcesso(planJbm.gcpj,
                                                         planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                         planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                         planjbm.tipoprocesso,
                                                         IntToStr(planJbm.codGcpjEscritorio),
                                                         planJbm.lstCnpjs);
         if dmgcpcj_base_XI.dtsNfiscdt.RecordCount > 0 then
         begin
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO TRABALHISTA') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIÊNCIA TRABALHISTA') or
               (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'INSTRUCAO') then
            begin
               if dmgcpcj_base_XI.dtsNfiscdt.RecordCount >= 2 then
               begin
                  planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.RemoveNotaProcesso;
               end
            end
            else
            begin
               planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.RemoveNotaProcesso;
            end;
         end;
      //end;
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.BitBtn7Click(Sender: TObject);
var
   arquivo : Tstringlist;
   arquivoout: TstringList;
   i, p : integer;
   estagio : integer;
   linha, valor : string;
begin
   arquivo := TStringlist.Create;
   try
      arquivo.LoadFromFile('c:\presenta\temp\atos.txt');

      arquivoOut := TstringlIst.Create;
      try
         estagio := 0;
         for i := 0 to arquivo.Count - 1 do
         begin
            case estagio of
               0 : begin
                  if Pos('Processo Ato Data do Ato  Valor Dependência', arquivo.strings[i]) = 0 then
                     continue;
                  Inc(estagio);
               end;
               1 : begin
                  if Pos('[', arquivo.strings[i]) <> 0 then
                  begin
                     estagio := 0;
                     continue;
                  end;

                  linha := Trim(arquivo.strings[i]);
                  valor := trim(copy(linha, 1, 10));

                  linha := Copy(linha, 11, length(linha));
                  p := Pos('01/11/2011', linha);
                  valor := valor + ';' + Trim(Copy(linha, 1, p-1));

                  linha := Trim(Copy(linha, p+10, length(linha)));
                  p := pos(' ', linha);
                  valor := valor + ';' + Trim(Copy(linha, 1, p-1));
                  arquivoout.add(valor);
               end;

            end;
         end;
         arquivoout.SaveToFile('c:\presenta\temp\digitados.txt');
      finally
         arquivoout.free;
      end;
   finally
      arquivo.free;
   end;

end;

procedure TfrmValidaPlan.BitBtn8Click(Sender: TObject);
var
   ini : TIniFile;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
    ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
    try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   CalculaNumeroDaNota;
   ExibeMensagem('Cálculo das notas finalizado');
end;


procedure TfrmValidaPlan.CorrigirNota;
var
   wintask :  TWintask;
   fName, fname2 : string;
   valor : string;
   achou : boolean;
   i, estagio, index, p, ret : integer;
   arquivo : TStringList;
   dataDigitar : string;
   tipoato : string;
   ini : TiniFile;
   diretorio :string;

   function TemDestinatario : boolean;
   var
      i,p : integer;
      arquivo : TStringlist;
      codigo : string;
   begin
      result := false;
      arquivo := TStringList.Create;
      try
         arquivo.LoadFromFile(fname);

         for i := 0 to arquivo.count - 1 do
         begin
            p :=  pos('despesaHonorariosProcessoVO.composicaoDespHonorariosVO.dependenciaBradescoOnLineVO.cdDestinatario" maxlength="5" size="5" value="', arquivo.strings[i]);
            if p = 0 then
               continue;
            codigo := Copy(arquivo.strings[i], p+129, length(arquivo.strings[i]));
            p := Pos('" onkeydown="', codigo);
            if p = 0 then
               exit;

            codigo := Trim(Copy(codigo, 1, p-1));
            if codigo <> '' then
               result := true;
            exit;
         end;
      finally
         arquivo.free;
      end;
   end;

   function ObtemDependenciaDigitar:string;
   var
      i,p : integer;
      arquivo : TStringlist;
      estagio : integer;
      codigo, nome : string;
      valstr : string;
      dependencia : string;
      primeira : boolean;
   begin
      result := '';
      codigo := '';
      nome:= '';
      planJbm.ObtemReclamadasDoProcesso('');

      while Not planJbm.dtsReclamadas.Eof do
      begin
         if planJbm.dtsReclamadas.FieldByName('sequenciagcpj').AsInteger = 1 then //guarda a primeira por seguranca
         begin
            codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
            nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
         end;

         //é bradesco?
         if planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger = 237 then //procura uma agencia
         begin
            if planJbm.dtsReclamadas.FieldByName('codempresa').AsInteger = planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger then
            begin
               if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString = 'A' then //agencia
               begin
                  codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
            end;
         end
         else
         begin
            if (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger) or
               (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('codempresaligada').AsInteger) then
            begin
               if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A' then //não pode ser agencia
               begin
                  codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
            end;
         end;
         planJbm.dtsReclamadas.Next;
      end;

      if (codigo = '0') or (codigo = '') then
         primeira := true
      else
         primeira := false;

      arquivo := TStringList.Create;
      try
         arquivo.LoadFromFile(fname2);

         while true do
         begin
            estagio := 0;
            dependencia := '';

            for i := 0 to arquivo.count - 1 do
            begin
               case estagio of
                  0 : begin
                     if pos('<select name="selDependencia"', arquivo.strings[i]) = 0 then
                        continue;
                     Inc(estagio);
                  end;
                  1 : begin
                     p := Pos('<option value="', arquivo.strings[i]);
                     if p = 0 then
                        break;;

                     valstr := trim(Copy(arquivo.strings[i], p+15, length(arquivo.strings[i])));
                     p := Pos('">', valstr);
                     if p = 0 then
                        exit;

                     if Not primeira then
                     begin
                        if RemoveNaoNumericos(Trim(copy(valstr, 1, p-1))) <> codigo then
                           continue;
                     end;

                     valStr := Trim(Copy(valStr, p+2, length(valstr)));
                     p := pos('</option>', valStr);
                     if p = 0 then
                     begin
                        result := '';
                        exit;
                     end;

                     valStr := Copy(valStr, 1, p-1);
                     if valStr = '' then
                        ShowMessage('Erro');

                     p := Pos('&#39;', valStr);
                     if p <> 0 then
                        valstr := copy(valstr, 1, p-1) + '''' + copy(valstr, p+5, length(valstr));

                     Result := valstr;
                     exit;
                  end;
               end;
            end;
            //nao achou
            if primeira then
            begin
               Result := nome;
               exit;
            end;
            primeira := true;
         end;
      finally
         arquivo.free;
      end;
   end;

begin
   re.lines.Clear;

   if FormatDateTime('hh:nn', Now) > '22:00' then
      exit;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.cnpjescritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.nomeescritorio := planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString;
   planJbm.nomedigitar := planJbm.dtsPlanilha.FieldByName('nomedigitar').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').AsInteger;

   dataDigitar := planJbm.ObtemDataDigitar(0);
   if dataDigitar = '' then
   begin
      ShowMessage('Não encontrada a data inicial para o ano/mês de referência');
      exit;
   end;

   if numnota.text = '' then
   begin
      ShowMessage('Informe o número da nota para ser corrigida');
      numNota.SetFocus;
      exit;
   end;

   ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
   try
      diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_Programa\presenta\mdb');
   finally
      ini.Free;
   end;

   wintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                              diretorio + '\rob\',
                              '');

   wintask.FecharTodos('TaskExec.exe');
   wintask.FecharTodos('TaskLock.exe');
   wintask.FecharTodos('Tasksync.exe');
   try
      ret := planJbm.ObtemProcessosDigitar(numnota.text, true);
      if ret = -1 then
         exit;

      if ret = 0 then
      begin
         ShowMessage('Nenhum processo encontrado para digitacao');
         exit;
      end;

      planJbm.dtsAtos.tag := 0;

      if planJbm.dtsAtos.Eof then
      begin
         ShowMessage('Não encontrados atos para a nota');
         exit;
      end;

      planJbm.empresaLigada := planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger;
      planJbm.numerodanota :=  planJbm.dtsAtos.FieldByName('numeronota').AsInteger;

      //localizar a tela da nota
      fName := wintask.GetPagina_CopyText('IEXPLORE.EXE|Internet Explorer_Server|GCPJ - Windows Internet Explorer',
                                                   'Alteração');
      if fName = '' then //esta na pagina errada
      begin
         ShowMessage('Erro posicionando na tela de alteração');
         exit;
      end;

      try
         arquivo := TStringList.Create;
         try
            arquivo.LoadFromFile(fName);

            estagio := 0;

            Index := 0;


            for i := 0 to arquivo.Count - 1 do
            begin
               case estagio of
                  0 : begin
                     if Pos('Processo Ato Repasse Afeta Resultado Data do Ato Valor Dependência', arquivo.strings[i]) = 0 then
                        continue;
                     Inc(estagio);
                  end;
                  1 : begin
                     valor := Trim(arquivo.Strings[i]);
                     if valor = '' then
                        continue;

                     p := Pos(']', valor);
                     if p <> 0 then
                     begin
                        //é ultima página?
                        valor := Trim(Copy(valor, p+1, length(valor)));
                        if valor = '' then
                        begin
                           Inc(passagem);
                           if Passagem > 3 then
                           begin
                              wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                              Sleep(100);

                              wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                              Sleep(100);

                              ShowMessage('Fim');
                              exit;
                           end;

                           wintask.ClickHTMLElement('GCPJ', 'IMG[SRC= ''http://www4.net.bradesco.com.br/imagens/setaTabelaD'']');
                           Sleep(100);
                           CorrigirNota;
                           exit;
                        end;

                        passagem := 0;
                        //mudar de pagina
                        p := Pos(' ', valor);
                        if p <> 0 then
                           valor := Trim(copy(valor, 1, p-1));

                        wintask.ClickHTMLElement('GCPJ', 'A[INNERTEXT= ''' + valor + ''']');
                        ultimaPagina := StrToInt(valor);
                        sleep(100);
                        CorrigirNota;
                        exit;
                     end;

                     tipoato := Trim(Copy(valor, 11, Length(valor)));
                     p := Pos(' ', tipoato);
                     if p <> 0 then
                        tipoato := Trim(Copy(tipoato, 1, p-1));

                     valor := IntToStr(StrToInt(Copy(valor, 1, 10)));

                     (*                     if StrToInt(valor) < ultimoGcpj then
                     begin
                        wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                        Sleep(100);

                        wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                        Sleep(100);

                        ShowMessage('Fim');
                        exit;
                     end;*)

//                     ultimoGcpj := StrToInt(valor);

                     achou := false;
                     while Not planJbm.dtsAtos.Eof do
                     begin
                        if valor <> planJbm.dtsAtos.FieldByName('gcpj').AsString then
                        begin
                           planJbm.dtsAtos.Next;
                           continue;
                        end;
                        achou := true;
                        break;
                     end;

                     if not achou then
                     begin
                        ShowMessage('Erro inesperado');
                        exit;
                     end;

                     if planJbm.JaAlterado(tipoAto) then
                     begin
                        Inc(index);
                        continue;
                     end;

                     achou := false;

                     planJbm.CarregaCamposDaTabela;
                     ExibeMensagem('Digitando processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
                     Application.ProcessMessages;

                     planJbm.ObtemReclamadasDoProcesso('');
                     //tem agencia?
                     while not planJbm.dtsReclamadas.Eof do
                     begin
                        if planJbm.dtsReclamadas.FieldByName('codempresa').AsInteger <> 237 then
                        begin
                           planJbm.dtsReclamadas.next;
                           continue;
                        end;
                        if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A' then
                        begin
                           planJbm.dtsReclamadas.next;
                           continue;
                        end;
                        achou := true;
                        break;
                     end;

                     if Not achou then //não precisa alterar
                     begin
                        planJbm.SalvaAlterado(tipoato);
                        Inc(Index);
                        continue;
                     end;

                     //se já é a mesma não precisa alterar
                     if Pos(planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString + ' - ', arquivo.Strings[i]) <> 0 then
                     begin
                        planJbm.SalvaAlterado(tipoato);
                        Inc(Index);
                        continue;
                     end;


                     //altera
                     if index = 0 then
                        wintask.ClickHTMLElement('GCPJ', 'INPUT RADIO[NAME= ''composicoesSelecs'']')
                     else
                        wintask.ClickHTMLElement('GCPJ', 'INPUT RADIO[NAME= ''composicoesSelecs'', INDEX=''' + IntToStr(index+1) + ''']');
                     Sleep(100);

                     wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''alterar'']');
                     sleep(100);

                     fname2 := wintask.EstaNaPagina_NotepadCodePage('IEXPLORE.EXE|Internet Explorer_Server|GCPJ - Windows Internet Explorer|1',
                                                       '787','297',
                                                       '856','490',
                                                       'NOTEPAD.EXE|Notepad|finHonorariosNovoAction',
                                                       'Documento do Prestador: ',
                                                       false);
                     if fname2 = '' then
                     begin
                        ExibeMensagem('Erro');
                        exit;
                     end;

                     try
                        valor := ObtemDependenciaDigitar;
                        if valor = '' then
                        begin
                           ExibeMensagem('Erro');
                           exit;
                        end;

                        wintask.SelectHTMLItem('GCPJ', 'SELECT[NAME= ''selDependencia'']', Trim(valor));
                        Sleep(180);

                        if wintask.Verifica_Janela_PopUp('IEXPLORE.EXE|Static|Esta dependência foi encerrada, deseja alterá-la para', valor) = 1 then
                        begin
                           wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                           Sleep(20);
                        end;


                        wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''incluir'']');
                        Sleep(100);

                        wintask.ClickHTMLElement('GCPJ - Gestão e Controle de Processos Jurídicos', 'INPUT BUTTON[VALUE= ''confirmar inclusão'']');
                        Sleep(100);

                        planJbm.SalvaAlterado(tipoato);
                        Inc(index);
                     finally
                        DeleteFile(fname2);
                     end;
                  end;
               end;
            end;
         finally
            arquivo.Free;
         end;
      finally
         DeleteFile(fname);
      end;
   finally
      wintask.Free;
   end;
end;
procedure TfrmValidaPlan.VerificaDuplicidades;
begin
   re.lines.Clear;

   processando.Text := 'Verificando processos duplicados na planilha';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(4);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      ExibeMensagem('Verificando duplicidade do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      //if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'PREPOSTO')  and
      if   (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'ACOMPANHAMENTO MENSAL') and
           (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'PARCELA') and
           (planJbm.dtsAtos.FieldByName('tipoandamento').AsString <> 'ADESAO POUPADOR') then
      begin
         if planJbm.TemProcessoDuplicado then
         begin
            planJbm.GravaOcorrencia(2, 'Ato em duplicidade na planilha');
               planJbm.MarcaProcessoCruzadoGcpj(9);
         end;
      end
      else
      begin
         if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'PARCELA') and (planJbm.tipoProcesso = 'CI') then
         begin
            planJbm.GravaOcorrencia(2, 'Não pode pagar PARCELA para processo cível');
            planJbm.MarcaProcessoCruzadoGcpj(9);
         end
         else
         begin
            if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'PARCELA') then
            begin
               if planJbm.TemProcessoDuplicado then
               begin
                  planJbm.GravaOcorrencia(2, 'Ato em duplicidade na planilha');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
               end;
            end;
         end;
      end;
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.RemoveEstadosDeVolumetria;
var
   estado : string;
   orgaoJulgador : integer;
   semVolumetria : boolean;
   valorcorreto : double;
begin
   //rotina desabilitada a pedido da Nilce - E-mail recebido em 03/09/2012
   exit;

   planJbm.ObtemProcessosCruzarGcpj(5);
   if planJbm.dtsAtos.RecordCount = 0 then
      exit;
   re.lines.Clear;

   processando.Text := 'Excluindo estados do calculo de volumetria';
   Application.ProcessMessages;

   while not planJbm.dtsAtos.Eof do
   begin
      Application.ProcessMessages;
      ExibeMensagem('Verificando comarca do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      planjbm.CarregaCamposDaTabela;

      orgaoJulgador := planJbm.ObtemCodOrgProcesso;
      semvolumetria := false;

      if orgaoJulgador = 0 then
      begin
         planJbm.dtsAtos.Next;
         continue;
      end;

      estado := planJbm.ObtemUfProcesso(orgaoJulgador);

      if estado = '' then
      begin
         planjbm.dtsAtos.next;
         continue;
      end;

      //remove estados que não pode considerar volumetria
      if (planJbm.cnpjEscritorio = '08833747000132') or //sette camra
         (planJbm.cnpjEscritorio = '10508423000170')then //jbm
      begin
         if (estado = 'MG') or
            (estado = 'AM') OR
            (estado = 'PA') or
            (estado = 'RO') OR
            (estado = 'RR') or
            (estado = 'AP') OR
            (estado = 'AC') or
            (estado = 'AL') or
            (estado = 'BA') or
            (estado = 'CE') or
            (estado = 'MA') or
            (estado = 'PB') or
            (estado = 'PE') or
            (estado = 'PI') or
            (estado = 'RN') or
            (estado = 'SE') or
            (estado = 'TO') then
            semvolumetria := true;
      end
      else
      begin
         if (planJbm.cnpjEscritorio = '05159996000104') then //rocha marinho
         begin
            if (estado = 'AM') OR
               (estado = 'PA') or
               (estado = 'RO') OR
               (estado = 'RR') or
               (estado = 'AP') OR
               (estado = 'AC') or
               (estado = 'AL') or
               (estado = 'BA') or
               (estado = 'CE') or
               (estado = 'MA') or
               (estado = 'PB') or
               (estado = 'PE') or
               (estado = 'PI') or
               (estado = 'RN') or
               (estado = 'SE') or
               (estado = 'TO') then
               semvolumetria := true;
         end;
      end;

      if not semvolumetria then
      begin
         planjbm.dtsAtos.Next;
         continue;
      end;

      if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2011/05/01' then //utlizar a rotina de novo cálculo
      begin
         if cbSemReajuste.Checked then
         begin
            valorCorreto := planJbm.ObtemValorSemReajuste('TUTELA'); //470
            if valorCorreto = 0 then
            begin
               ExibeMensagem('   Programa está calculando valor igual a zero');
               Application.ProcessMessages;
               exit;
            end;
         end
         else
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 422
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2013/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 438
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2014/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 470
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2015/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 500
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2016/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 541
            else
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2017/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 591
            else
               valorCorreto := 619;
         end;

         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2010/05/01' then //utlizar a rotina de novo cálculo
      begin
         if planjbm.dtsatos.FieldByName('valor').AsFloat <> 382 then
         begin
            planJbm.GravaValorCorrigido(382);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if planjbm.dtsatos.FieldByName('valor').AsFloat <> 370 then
      begin
         planJbm.GravaValorCorrigido(370);
         planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
      end
      else
         planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
      planJbm.MarcaProcessoCruzadoGcpj(6);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.CopiaBasesGCPJ;
var
   ini : TiniFile;
   diretorio : string;
   arquivoOrigem : string;
   arquivoDestino : string;
begin
   //copia as bases para o sistema
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');

   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');

   try
      diretorio := ini.ReadString('honorarios', 'diretorioGcpj', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\_BaseDiaria');

      ExibeMensagem('Copiando a base de dados do GCPJ Base RELATORIOS DADOS PROCESSOS. Aguarde.....');
      Application.ProcessMessages;
      Sleep(500);

      arquivoOrigem := '\\mz-vv-fs-140\d4040_1\Publico\Sistemas_Compartilhados\Gerador_Relatorios\_DADOS\Relatorios_Gerenciais_DADOS_PROCESSOS.mdb';//'Q:\Publico\Backup_sistemas\Compartilhado\Gerador_Relatorios\_DADOS\Relatorios_Gerenciais_DADOS_PROCESSOS.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_compartilhada', 'databasename', 'C:\presenta\basediaria\Relatorios_Gerenciais_DADOS_PROCESSOS.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ Base TRABALHISTAS. Aguarde.....');
      Application.ProcessMessages;
      Sleep(500);

      arquivoOrigem := 'i:\publico\sistemas_compartilhados\gerador_relatorios\_dados\Relatorios_Gerenciais_DADOS_TA_TO_TR.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_trabalhistas', 'databasename', 'C:\presenta\basediaria\Relatorios_Gerenciais_DADOS_TA_TO_TR.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      ExibeMensagem('Copiando a base de dados do Honorarios. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_Programa\Honorarios_Programa_II.mdb';
      arquivoDestino := ini.ReadString('honorarios_programas', 'databasename', 'C:\presenta\basediaria\Honorarios_Programa_II.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(200);

      ExibeMensagem('Copiando a base de dados do GCPJ Base PROCESSOS BAIXADOS. Aguarde.....');
      Application.ProcessMessages;


      arquivoOrigem := 'i:\publico\sistemas_compartilhados\gerador_relatorios\_dados\Relatorios_Gerenciais_DADOS_CI_BAIXADOS.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_baixados', 'databasename', 'C:\presenta\basediaria\Relatorios_Gerenciais_DADOS_CI_BAIXADOS.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(200);


      ExibeMensagem('Copiando a base de dados do GCPJ Base PROCESSOS RECUPERADOS. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := '\\mz-vv-fs-140\d4040_1\Publico\Sistemas_Compartilhados\Gerador_Relatorios_Recuperacao_Creditos\_DADOS\Relatorios_Gerenciais_DADOS_PROCESSOS_RC.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_recuperados', 'databasename', 'C:\presenta\basediaria\Relatorios_Gerenciais_DADOS_PROCESSOS_RC.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(200);


      ExibeMensagem('Copiando a base de dados do GCPJ Base VII. Aguarde.....');
      Application.ProcessMessages;


      arquivoOrigem := diretorio + '\basediaria_gcpj_vii.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_vii', 'databasename', 'C:\presenta\basediaria\basediaria_gcpj_vii.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(2000);


      ExibeMensagem('Copiando a base de dados do GCPJ Base 56. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB056.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB056', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB056.mdb');

     if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(2000);


      ExibeMensagem('Copiando a base de dados do GCPJ Base 57. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB057.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB057', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB057.mdb');

     if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(2000);

//alexandre - 13/04
      ExibeMensagem('Copiando a base de dados do GCPJ Base 57_ate_2016. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB057_ate_2016.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB057_ate_2016', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB057_ate_2016.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false); 
//alexandre - 13/04

      ExibeMensagem('Copiando a base de dados do GCPJ Base 75. Aguarde.....');
      Application.ProcessMessages;

//      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB075.mdb';
//      arquivoDestino := ini.ReadString('gcpj_base_i', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB075.mdb');

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB075.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB075', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB075.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB024. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB024.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB024', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB024.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB088. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB088.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB088', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB088.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB0B7. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB0B7.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB0B7', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB0B7.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      ExibeMensagem('Copiando a base de dados do GCPJ Base III. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\basediaria_gcpj_iii.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_iii', 'databasename', 'C:\presenta\basediaria\basediaria_gcpj_iii.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB050. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB050.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB050', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB050.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB066. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB066.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB066', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB066.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB0J8_MESU9021_GCPJB0K1. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB0J8_MESU9021_GCPJB0K1.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB0J8_MESU9021_GCPJB0K1', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB0J8_MESU9021_GCPJB0K1.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB076. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB076.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB076', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB076.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB077. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB077.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB077', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB077.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);



      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB081. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB081.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB081', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB081.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      ExibeMensagem('Copiando a base de dados do GCPJ Base VIII. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\basediaria_gcpj_VIII.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_viii', 'databasename', 'C:\presenta\basediaria\basediaria_gcpj_viii.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      ExibeMensagem('Copiando a base de dados do GCPJ Base IX. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\basediaria_gcpj_ix.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_ix', 'databasename', 'C:\presenta\basediaria\basediaria_gcpj_ix.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(200);

      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB010. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB010.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB010', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB010.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(200);

//BaseDiaria_GCPJB010_de_2016_ate_2017
{      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB010_de_2016_ate_2017. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB010_de_2016_ate_2017.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB010_de_2016_ate_2017', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB010_de_2016_ate_2017.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(200);


//BaseDiaria_GCPJB010_ate_2015
      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJB10_de_2015. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJB10_de_2015.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJB10_de_2015', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB10_de_2015.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      Sleep(200);}






      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJ_NFISCDT. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJ_NFISCDT.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJ_NFISCAL', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJ_NFISCDT.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
      sLEEP(100);

      ExibeMensagem('Copiando a base de dados do GCPJ BaseDiaria_GCPJ_NFISCAL. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := diretorio + '\BaseDiaria_GCPJ_NFISCAL.mdb';
      arquivoDestino := ini.ReadString('BaseDiaria_GCPJ_NFISCAL', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJ_NFISCAL.mdb');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      Sleep(500);

   finally
      ini.Free;
   end;
end;

procedure TfrmValidaPlan.BitBtn4Click(Sender: TObject);
begin
   if MessageDlg('Tem certeza que deseja excluir a planilha?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      exit;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.RemovePlanilhas(planJbm.idplanilha, planJbm.idEscritorio, planJbm.anomesreferencia);
   planJbm.dtsPlanilha.Close;
   planJbm.dtsPlanilha.Open;
   ShowMessage('Planilha removida com sucesso');
end;

procedure TfrmValidaPlan.IndexaBaseHonorarios;
begin

end;

procedure TfrmValidaPlan.AlertaAtoNaoValidado;
begin
   ExibeMensagem('Verificando se sobrou ato não validado');

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').Asinteger;
   planJbm.CarregaFiliais;


   planJbm.ObtemProcessosCruzarGcpj(9);
   if planJbm.dtsAtos.RecordCount > 0 then
   begin
      while not planJbm.dtsAtos.Eof do
      begin
{        planJbm.linhaPlanilha := planjbm.dtsatos.FieldByName('linhaPlanilha').AsInteger;
        planJbm.idEscritorio := planjbm.dtsatos.FieldByName('idescritorio').AsInteger;
        planJbm.idplanilha := planjbm.dtsatos.FieldByName('idplanilha').AsInteger;
        planJbm.sequencia :=  planjbm.dtsatos.FieldByName('sequencia').AsInteger;
        planJbm.anomesreferencia := planjbm.dtsatos.FieldByName('anomesreferencia').AsString;}
        planJbm.CarregaCamposDaTabela;

        //E-mail Carlos - 27/05/2019
         if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < '2018/01/01') then
         begin
           ExibeMensagem(IntToStr(planJbm.dtsAtos.tag) + '-Marcando processo como inconsistente: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
           planJbm.CarregaCamposDaTabela;
           planJbm.GravaOcorrencia(2, 'Data do Ato menor que 01/01/2018, verificar.');
           planJbm.MarcaProcessoCruzadoGcpj(9);
           planJbm.dtsAtos.Next;
           Continue;
         end;

         ExibeMensagem(IntToStr(planJbm.dtsAtos.tag) + '-Marcando processo como inconsistente: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
         planJbm.CarregaCamposDaTabela;
         planJbm.GravaOcorrencia(2, 'Não encontrada regra de validação no sistema');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
      end;
        ExibeMensagem('Validação finalizada com sucesso');
//      ExibeMensagem('Ainda existe processo não validado. Sistema não pode prosseguir com o cálculo das notas')
   end
   else
      ExibeMensagem('Validação finalizada com sucesso');
end;

procedure TfrmValidaPlan.BitBtn9Click(Sender: TObject);
begin
   frmAtosPendentes := TfrmAtosPendentes.Create(NIl);

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.ObtemAtosPendentes;

   frmAtosPendentes.ShowModal;
end;

procedure TfrmValidaPlan.DBGrid1DblClick(Sender: TObject);
begin
   if MessageDlg('Tem certeza que deseja marcar todos os atos desta planilha como não validados?', mtConfirmation, [mbYes, mbNo],0) <> mrYes then
      exit;

   planJbm.ResetFlagsDaPlanilha;

   ShowMessage('Ok');
end;

procedure TfrmValidaPlan.ExibeMensagem(mensagem: string);
begin
   re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-' + mensagem);
   Application.ProcessMessages;
end;

procedure TfrmValidaPlan.TrataProcessoTrabalhista(tipo: integer);
var
   valorconferir : double;
begin
   if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) < '2016/08/01' then //utlizar a rotina de novo cálculo
   begin
      ExibeMensagem('      Gravando ocorrência de erro');
      planJbm.GravaOcorrencia(2, 'Data do ato anterior a 01/08/2016');
      planJbm.MarcaProcessoCruzadoGcpj(9);
      exit;
   end;

   if (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'TUTELA') or (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'PARCELA') then
   begin
      //se o processo estiver encerrado gera erro - Solicitado pela Elaine 20/01/2017
      //verificar 45 dias da data - solicitado pela Gabriela em 08/03/2017
      if (Not planjbm.dtsatos.FieldByName('databaixa').IsNull) then
      begin
         if (DaysBetween(planJbm.dtsatos.FieldByName('datadoato').AsDateTime, planjbm.dtsatos.FieldByName('databaixa').AsDateTime) > 45) then
         begin
            ExibeMensagem('      Gravando ocorrência de erro');
            planJbm.GravaOcorrencia(2, 'Não pode pagar ato para processo encerrado em ' + FormatDateTime('dd/mm/yyyy', planjbm.dtsatos.FieldByName('databaixa').AsDateTime));
            planJbm.MarcaProcessoCruzadoGcpj(9);
            exit;
         end;
      end;

      if dmgcpcj_base_XI.PagouTrabalhistaNestaData(planjbm.dtsatos.FieldByName('tipoandamento').AsString, planJbm.dtsAtos.FieldByName('gcpj').AsString,planJbm.dtsatos.FieldByName('datadoato').AsDateTime) then
      begin
         ExibeMensagem('      Gravando ocorrência de erro');
         planJbm.GravaOcorrencia(2, 'Este ato já foi pago anteriormente');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         exit;
      end;
   end;

   case tipo of
      1 : begin //BVP ou exfuncionario
         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'TUTELA') or
             (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'INSTRUCAO'))then
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
               valorconferir := 807.00
            else
               valorconferir := 788.00;

            if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2, 'Valor incompatível com o tipo de andamento');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               exit;
            end;
         end
         else
         begin
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'PARCELA') then
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                  valorconferir := 77.00
               else
                  valorconferir := 75.00;

               if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
               begin
                  ExibeMensagem('      Gravando ocorrência de erro');
                  planJbm.GravaOcorrencia(2, 'Valor incompatível com o tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  exit;
               end;
            end
            else
            begin
               if (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'EXITO') then
               begin
                  if planjbm.dtsatos.FieldByName('motivobaixa').IsNull then
                  begin
                     planJbm.GravaOcorrencia(2, 'Para este tipo de andamento, o processo tem que estar baixado no GCPJ');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     exit;
                  end;

                  if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO REALIZADO ANTES DA AUDIENCIA') then
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                        valorconferir := 3481.00
                     else
                        valorconferir := 3400.00;

                     if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
                     begin
                        ExibeMensagem('      Gravando ocorrência de erro');
                        planJbm.GravaOcorrencia(2, 'Valor não compatível com o tipo de andamento/motivo da baixa');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        exit;
                     end;
                  end
                  else
                  begin
                     if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'IMPROCEDENCIA') then
                     begin
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                           valorconferir := 4607.00
                        else
                           valorconferir := 4500.00;

                        if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
                        begin
                           ExibeMensagem('      Gravando ocorrência de erro');
                           planJbm.GravaOcorrencia(2, 'Valor não compatível com o tipo de andamento/motivo da baixa');
                           planJbm.MarcaProcessoCruzadoGcpj(9);
                           exit;
                        end;
                     end
                     else
                     begin
                        if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO REALIZADO ENTRE DA AUDIENCIA') then
                        begin
                           if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                              valorconferir := 2560.00
                           else
                              valorconferir := 2500.00;

                           if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
                           begin
                              ExibeMensagem('      Gravando ocorrência de erro');
                                 planJbm.GravaOcorrencia(2, 'Valor não compatível com o tipo de andamento/motivo da baixa');
                              planJbm.MarcaProcessoCruzadoGcpj(9);
                              exit;
                           end;
                        end;
                     end;
                  end;
               end
               else
               begin
                  //se chegou aqui é porque está com o nome errado do tipo de andamento
                  ExibeMensagem('      Gravando ocorrência de erro');
                  if planjbm.dtsatos.FieldByName('motivobaixa').AsString <> '' then
                     planJbm.GravaOcorrencia(2, 'Nomenclatura incorreta do ato  ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString + ' motivo da baixa: ' + planjbm.dtsatos.FieldByName('motivobaixa').AsString)
                  else
                     planJbm.GravaOcorrencia(2, 'Nomenclatura incorreta do ato  ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString);
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  exit;
               end;
            end;
         end;
      end;

      2 : begin //terceira
         if ((planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'TUTELA') or
             (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'INSTRUCAO'))then
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
               valorconferir := 512.00
            else
               valorconferir := 500.00;

            if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
            begin
               ExibeMensagem('      Gravando ocorrência de erro');
               planJbm.GravaOcorrencia(2, 'Valor incompatível com o tipo de andamento');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               exit;
            end;
         end
         else
         begin
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'PARCELA') then
            begin
               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                  valorconferir := 51.00
               else
                  valorconferir := 50.00;

               if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
               begin
                  ExibeMensagem('      Gravando ocorrência de erro');
                  planJbm.GravaOcorrencia(2, 'Valor incompatível com o tipo de andamento');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  exit;
               end
            end
            else
            begin
               if (planjbm.dtsatos.FieldByName('tipoandamento').AsString  = 'EXITO') then
               begin
                  if planjbm.dtsatos.FieldByName('motivobaixa').IsNull then
                  begin
                     planJbm.GravaOcorrencia(2, 'Para este tipo de andamento, o processo tem que estar baixado no GCPJ');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     exit;
                  end;

                  if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO REALIZADO ANTES DA AUDIENCIA') then
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                        valorconferir := 2048.00
                     else
                        valorconferir := 2000.00;

                     if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
                     begin
                        ExibeMensagem('      Gravando ocorrência de erro');
                        planJbm.GravaOcorrencia(2, 'Valor não compatível com o tipo de andamento/motivo da baixa');
                        planJbm.MarcaProcessoCruzadoGcpj(9);
                        exit;
                     end;
                  end
                  else
                  begin
                     if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'IMPROCEDENCIA') then
                     begin
                        if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                           valorconferir := 2560.00
                        else
                           valorconferir := 2500.00;

                        if (planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
                        begin
                           ExibeMensagem('      Gravando ocorrência de erro');
                           planJbm.GravaOcorrencia(2, 'Valor não compatível com o tipo de andamento/motivo da baixa');
                           planJbm.MarcaProcessoCruzadoGcpj(9);
                           exit;
                        end;
                     end
                     else
                     begin
                        if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO REALIZADO ENTRE DA AUDIENCIA') then
                        begin
                           if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2017/05/01' then
                              valorconferir := 1638.00
                           else
                              valorconferir := 1600.00;

                           if(planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorconferir) then
                           begin
                              ExibeMensagem('      Gravando ocorrência de erro');
                              planJbm.GravaOcorrencia(2, 'Valor não compatível com o tipo de andamento/motivo da baixa');
                              planJbm.MarcaProcessoCruzadoGcpj(9);
                              exit;
                           end;
                        END;
                     end;
                  end;
               end
               else
               begin
                  //se chegou aqui é porque está com o nome errado do tipo de andamento
                  ExibeMensagem('      Gravando ocorrência de erro');
                  if planjbm.dtsatos.FieldByName('motivobaixa').AsString <> '' then
                     planJbm.GravaOcorrencia(2, 'Nomenclatura incorreta do ato ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString + ' motivo da baixa: ' + planjbm.dtsatos.FieldByName('motivobaixa').AsString)
                  else
                     planJbm.GravaOcorrencia(2, 'Nomenclatura incorreta do ato  ' + planjbm.dtsatos.FieldByName('tipoandamento').AsString);
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  exit;
               end;
            end;
         end;
      end;
   end;

   planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
   planJbm.MarcaProcessoCruzadoGcpj(6);
end;
procedure TfrmValidaPlan.BitBtn10Click(Sender: TObject);
begin
   planJbm.ObtemPlanilhasValidar;
   if planJbm.dtsPlanilha.RecordCount = 0 then
   begin
      ShowMessage('Nenhuma planilha para ser validada');
      exit;
   end;

   while not planJbm.dtsPlanilha.Eof do
   begin
      BitBtn1.Click;
      Application.ProcessMessages;
      BitBtn8.Click;
      Application.ProcessMessages;
      planJbm.dtsPlanilha.Next;
   end;

   //envia para a plataforma em lote - solicitado por Gabriela em 08/03/2017
(**   mainForm.Tag := 1;
   try
      mainForm.actEnviaGcpjExecute(Sender);
   finally
      mainForm.Tag := 0;
   end; **)
end;

end.
