unit uPlanJBM;

interface
   uses
      datamodule_honorarios, adodb, datamodulegcpj_base_I, datamodulegcpj_base_IV, datamodulegcpj_compartilhado, datamodulegcpj_recuperados,
      datamodulegcpj_base_migrados, sysutils, strutils,mywin, mgenlib, db, datamodulegcpj_base_III,
      dialogs, datamodulegcpj_base_II, forms, classes,datamodulegcpj_base_v, comobj,ExtCtrls, Func_Wintask_Obj, dateutils,
      datamodulegcpj_base_vII, datamodulegcpj_base_vIII, datamodulegcpj_baixados, datamodulegcpj_base_ix, datamodulegcpj_base_x,
      datamodulegcpj_base_xi, datamodulegcpj_trabalhistas, inifiles;

type
  TProcExibeMensagem = Procedure(mensagem: string) of object;

  TClickOkThread = Class(TThread)
  private
    FpleaseClose: boolean;
    Fwintask: TWintask;
    procedure SetpleaseClose(const Value: boolean);
    procedure Setwintask(const Value: TWintask);
  protected
     procedure Execute; override;
  public
     property wintask : TWintask read Fwintask write Setwintask;
     property pleaseClose : boolean read FpleaseClose write SetpleaseClose;
     Constructor Create;
  end;

(**14/07/2017
  TOutlookThread = Class(TThread)
  private
    Fwintask: TWintask;
    FpleaseClose: boolean;
    procedure Setwintask(const Value: TWintask);
    procedure SetpleaseClose(const Value: boolean);
  protected
     procedure Execute;override;
  public
     property wintask : TWintask read Fwintask write Setwintask;
     property pleaseClose: boolean read FpleaseClose write SetpleaseClose;
     Constructor Create;
  end;


  TOutlookObj = Class(TObject)
  private
    FMailFolder: Olevariant;
    FSystem: Olevariant;
    FNameSpace: Olevariant;
    FMensagem: string;
    FSubject: string;
    FemailPara: TStringList;
    FWintask: TWintask;
    FoutlookThread: TOutlookThread;
    procedure SetMailFolder(const Value: Olevariant);
    procedure SetNameSpace(const Value: Olevariant);
    procedure SetSystem(const Value: Olevariant);
    procedure SetemailPara(const Value: TStringList);
    procedure SetMensagem(const Value: string);
    procedure SetSubject(const Value: string);
    procedure SetWintask(const Value: TWintask);
    procedure SetoutlookThread(const Value: TOutlookThread);
  public
     property System : Olevariant read FSystem write SetSystem;
     property NameSpace : Olevariant read FNameSpace write SetNameSpace;
     property MailFolder : Olevariant read FMailFolder write SetMailFolder;
     property Subject : string read FSubject write SetSubject;
     property emailPara : TStringList read FemailPara write SetemailPara;
     property Mensagem : string read FMensagem write SetMensagem;
     property Wintask : TWintask read FWintask write SetWintask;
     property outlookThread : TOutlookThread read FoutlookThread write SetoutlookThread;

     constructor Create; overload;
     procedure SendMail;
     procedure ClickYes;
  end;
            **)
  TIdObject = class(Tobject)
  private
    Fid: integer;
    procedure Setid(const Value: integer);
  public
     property id : integer read Fid write Setid;
  end;
  TProcessoBaixa = class(tobject)
  private
    FmotivoBaixa: string;
    FdataDaBaixa: TdateTime;
    Fgcpj: string;
    FNome_reu: string;
    Fvalorbaixa: double;
    procedure SetdataDaBaixa(const Value: TdateTime);
    procedure SetmotivoBaixa(const Value: string);
    procedure Setgcpj(const Value: string);
    procedure SetNome_reu(const Value: string);
    procedure Setvalorbaixa(const Value: double);
  public
     property dataDaBaixa : TdateTime read FdataDaBaixa write SetdataDaBaixa;
     property motivoBaixa : string read FmotivoBaixa write SetmotivoBaixa;
     property gcpj : string read Fgcpj write Setgcpj;
     property Nome_reu : string read FNome_reu write SetNome_reu;
     property valorbaixa : double read Fvalorbaixa write Setvalorbaixa;

     constructor Create(numerogcpj: string);
  end;

  TProcessoReclamdas = class(Tobject)
  private
    FehSgcp: boolean;
    Fidplanilha: integer;
    FcodReclamada: integer;
    Fidescritorio: integer;
    Flinhaplanilha: integer;
    FnomeReclamada: string;
    Fsequencia: integer;
    Fanomesreferencia: string;
    FtipoReclamada: string;
    Fgcpj: string;
    Fnome_reu : string;
    FerrorMessage: string;
    Fsequenciagcpj: integer;
    Fempresa: integer;
    Fcodpessoaexterna: integer;
    FExibeMensagem: TProcExibeMensagem;
    Fcodexfuncionario: integer;
    procedure Setanomesreferencia(const Value: string);
    procedure SetcodReclamada(const Value: integer);
    procedure SetehSgcp(const Value: boolean);
    procedure Setidescritorio(const Value: integer);
    procedure Setidplanilha(const Value: integer);
    procedure Setlinhaplanilha(const Value: integer);
    procedure SetnomeReclamada(const Value: string);
    procedure Setsequencia(const Value: integer);
    procedure SettipoReclamada(const Value: string);
    procedure Setgcpj(const Value: string);
    procedure SetNome_reu(const Value: string);
    procedure SeterrorMessage(const Value: string);
    procedure Setsequenciagcpj(const Value: integer);
    procedure Setempresa(const Value: integer);
    procedure Setcodpessoaexterna(const Value: integer);
    procedure SetExibeMensagem(const Value: TProcExibeMensagem);
    procedure Setcodexfuncionario(const Value: integer);
   public
      property codReclamada : integer read FcodReclamada write SetcodReclamada;
      property nomeReclamada: string read FnomeReclamada write SetnomeReclamada;
      property tipoReclamada : string read FtipoReclamada write SettipoReclamada;
      property ehSgcp: boolean read FehSgcp write SetehSgcp;
      property idescritorio: integer read Fidescritorio write Setidescritorio;
      property idplanilha: integer read Fidplanilha write Setidplanilha;
      property anomesreferencia: string read Fanomesreferencia write Setanomesreferencia;
      property sequencia: integer read Fsequencia write Setsequencia;
      property linhaplanilha: integer read Flinhaplanilha write Setlinhaplanilha;
      property gcpj : string read Fgcpj write Setgcpj;
      property Nome_reu : string read FNome_reu write SetNome_reu;
      property errorMessage: string read FerrorMessage write SeterrorMessage;
      property sequenciagcpj: integer read Fsequenciagcpj write Setsequenciagcpj;
      property empresa: integer read Fempresa write Setempresa;
      property codpessoaexterna : integer read Fcodpessoaexterna write Setcodpessoaexterna;
      property codexfuncionario : integer read Fcodexfuncionario write Setcodexfuncionario;
      property ExibeMensagem : TProcExibeMensagem read FExibeMensagem write SetExibeMensagem;

      constructor Create(pidescritorio, pidplanilha: integer; panomesreferencia: string; psequencia, plinhaplanilha: integer;
                     pgcpj: string; const PProcedure: TProcExibeMensagem = Nil);
      procedure RemoveReclamadas;

      function ObtemEnvolvidosDoProcesso(fgDrcAtivas:integer; buscarAutor: boolean):integer;
      procedure LimpaCampos;
   end;

   TPlanilhaJBM = class(tobject)
   private
    Fanomesreferencia: string;
    Fidplanilha: integer;
    FdtsPlanilha: TADODataSet;
    Fbaixados: TProcessoBaixa;
    FtipoProcesso: string;
    FehJec: boolean;
    Fanomescompetencia: string;
    FvalorTotalNota: double;
    FempresaLigada: integer;
    Fnumerodanota: integer;
    FtotalAtos: integer;
    FidEscritorio: integer;
    FValor: double;
    FlinhaPlanilha: integer;
    Fsequencia: integer;
    FCliente: string;
    Fdivisaodafilial: string;
    FComarca: string;
    FVara: string;
    FUF: string;
    Fpartecontraria: string;
    Ftipoandamento: string;
    Fidprocessounico: string;
    Fprocesso: string;
    Fgcpj: string;
    FNome_reu : String;
    FPorcentagem: string;
    FvalBeneficioEconomico: string;
    Fdatadoato: TDateTime;
    Fvalordesembolsado: string;
    Fnomeescritorio: string;
    FdsPlan: TDataSource;
    FdtsAtos: TAdoDataset;
    FdtsReclamadas: TAdoDataset;
    FcnpjEscritorio: string;
    Fvalorpedido: double;
    Fnomedigitar: string;
    FcodGcpjEscritorio: integer;
    FdtsNotasPendentes: TAdoDataset;
    FdsNotasPendentes: TDataSource;
    FdtsNotasDigitando: TAdoDataSet;
    FdsNotasDigitando: TDataSource;
    Ffgibiativo: integer;
    FdataIbi: Tdatetime;
    Fusuario: string;
    FdrcContrarias: boolean;
    FlstCnpjs: TStringList;
    FvalorBase: double;
    FdataPlanilha: TDateTime;
    FfgDrcContrarias: integer;
    FExibeMensagem: TProcExibeMensagem;
    procedure Setanomesreferencia(const Value: string);
    procedure Setidplanilha(const Value: integer);
    procedure SetdtsPlanilha(const Value: TADODataSet);
    procedure Setbaixados(const Value: TProcessoBaixa);
    procedure SettipoProcesso(const Value: string);
    procedure SetehJec(const Value: boolean);
    procedure Setanomescompetencia(const Value: string);
    procedure SetempresaLigada(const Value: integer);
    procedure Setnumerodanota(const Value: integer);
    procedure SetvalorTotalNota(const Value: double);
    procedure SettotalAtos(const Value: integer);
    procedure SetidEscritorio(const Value: integer);
    procedure SetCliente(const Value: string);
    procedure SetComarca(const Value: string);
    procedure Setdatadoato(const Value: TDateTime);
    procedure Setdivisaodafilial(const Value: string);
    procedure Setgcpj(const Value: string);
    procedure SetNome_reu(const Value: string);
    procedure Setidprocessounico(const Value: string);
    procedure SetlinhaPlanilha(const Value: integer);
    procedure Setpartecontraria(const Value: string);
    procedure SetPorcentagem(const Value: string);
    procedure Setprocesso(const Value: string);
    procedure Setsequencia(const Value: integer);
    procedure Settipoandamento(const Value: string);
    procedure SetUF(const Value: string);
    procedure SetvalBeneficioEconomico(const Value: string);
    procedure SetValor(const Value: double);
    procedure SetVara(const Value: string);
    procedure Setvalordesembolsado(const Value: string);
    procedure Setnomeescritorio(const Value: string);
    procedure SetdsPlan(const Value: TDataSource);
    procedure SetdtsAtos(const Value: TAdoDataset);
    procedure SetdtsReclamadas(const Value: TAdoDataset);
    procedure SetcnpjEscritorio(const Value: string);
    procedure Setvalorpedido(const Value: double);
    procedure Setnomedigitar(const Value: string);
    procedure SetcodGcpjEscritorio(const Value: integer);
    procedure SetdsNotasPendentes(const Value: TDataSource);
    procedure SetdtsNotasPendentes(const Value: TAdoDataset);
    procedure SetdsNotasDigitando(const Value: TDataSource);
    procedure SetdtsNotasDigitando(const Value: TAdoDataSet);
    procedure SetdataIbi(const Value: Tdatetime);
    procedure Setfgibiativo(const Value: integer);
    procedure Setusuario(const Value: string);
    procedure SetdrcContrarias(const Value: boolean);
    procedure SetlstCnpjs(const Value: TStringList);
    procedure SetvalorBase(const Value: double);
    procedure SetdataPlanilha(const Value: TDateTime);
    procedure SetfgDrcContrarias(const Value: integer);
    procedure SetExibeMensagem(const Value: TProcExibeMensagem);
   public
      property anomesreferencia : string read Fanomesreferencia write Setanomesreferencia;
      property idplanilha : integer read Fidplanilha write Setidplanilha;
      property dtsPlanilha : TADODataSet read FdtsPlanilha write SetdtsPlanilha;
      property baixados : TProcessoBaixa read Fbaixados write Setbaixados;
      property tipoProcesso : string read FtipoProcesso write SettipoProcesso;
      property ehJec : boolean read FehJec write SetehJec;
      property anomescompetencia : string read Fanomescompetencia write Setanomescompetencia;
      property valorTotalNota : double read FvalorTotalNota write SetvalorTotalNota;
      property empresaLigada: integer read FempresaLigada write SetempresaLigada;
      property numerodanota: integer read Fnumerodanota write Setnumerodanota;
      property totalAtos : integer read FtotalAtos write SettotalAtos;
      property idEscritorio : integer read FidEscritorio write SetidEscritorio;
      property sequencia: integer read Fsequencia write Setsequencia;
      property linhaPlanilha: integer read FlinhaPlanilha write SetlinhaPlanilha;
      property Cliente: string read FCliente write SetCliente;
      property processo: string read Fprocesso write Setprocesso;
      property partecontraria: string read Fpartecontraria write Setpartecontraria;
      property gcpj: string read Fgcpj write Setgcpj;
      property Nome_reu: string read FNome_reu write SetNome_reu;
      property Comarca: string read FComarca write SetComarca;
      property Vara: string read FVara write SetVara;
      property tipoandamento: string read Ftipoandamento write Settipoandamento;
      property UF: string read FUF write SetUF;
      property divisaodafilial: string read Fdivisaodafilial write Setdivisaodafilial;
      property idprocessounico: string read Fidprocessounico write Setidprocessounico;
      property Porcentagem: string read FPorcentagem write SetPorcentagem;
      property valBeneficioEconomico: string read FvalBeneficioEconomico write SetvalBeneficioEconomico;
      property Valor: double read FValor write SetValor;
      property datadoato: TDateTime read Fdatadoato write Setdatadoato;
      property valordesembolsado: string read Fvalordesembolsado write Setvalordesembolsado;
      property nomeescritorio: string read Fnomeescritorio write Setnomeescritorio;
      property dsPlan : TDataSource read FdsPlan write SetdsPlan;
      property dtsAtos : TAdoDataset read FdtsAtos write SetdtsAtos;
      property dtsReclamadas: TAdoDataset read FdtsReclamadas write SetdtsReclamadas;
      property cnpjEscritorio: string read FcnpjEscritorio write SetcnpjEscritorio;
      property valorpedido: double read Fvalorpedido write Setvalorpedido;
      property nomedigitar : string read Fnomedigitar write Setnomedigitar;
      property codGcpjEscritorio: integer read FcodGcpjEscritorio write SetcodGcpjEscritorio;
      property dtsNotasPendentes : TAdoDataset read FdtsNotasPendentes write SetdtsNotasPendentes;
      property dsNotasPendentes : TDataSource read FdsNotasPendentes write SetdsNotasPendentes;
      property usuario : string read Fusuario write Setusuario;
      property drcContrarias : boolean read FdrcContrarias write SetdrcContrarias;
      property lstCnpjs : TStringList read FlstCnpjs write SetlstCnpjs;
      property valorBase : double read FvalorBase write SetvalorBase;
      property dataPlanilha : TDateTime read FdataPlanilha write SetdataPlanilha;
      property fgDrcContrarias : integer read FfgDrcContrarias write SetfgDrcContrarias;
      property ExibeMensagem : TProcExibeMensagem read FExibeMensagem write SetExibeMensagem;

      procedure dsPlanDataChange(Sender: TObject; Field: TField);
      property dtsNotasDigitando : TAdoDataSet read FdtsNotasDigitando write SetdtsNotasDigitando;
      property dsNotasDigitando : TDataSource read FdsNotasDigitando write SetdsNotasDigitando;
      property fgibiativo : integer read Ffgibiativo write Setfgibiativo;
      property dataIbi : Tdatetime read FdataIbi write SetdataIbi;
      procedure dsNotasPendentesDataChange(Sender: TObject; Field: TField);

      constructor Create(const PProcedure: TProcExibeMensagem = Nil);overload;

      procedure ObtemProcessosCruzarGcpj(tipo: integer);
      procedure Finalizar;
      procedure GravaOcorrencia(tipoErro: integer; mensagem: string);
      procedure MarcaProcessoCruzadoGcpj(estagio: integer);
      procedure VerificaDataDaBaixa(gcpj: string);
      function NomeEnvolvidoOk : boolean;
      procedure MarcaNomeEnvolvidoOk;
      function GravaTipoSubtipo : integer;
      procedure GravaTipoProcesso;
      function ObtemProcessosDigitar(numnota: string; const alteracao: boolean=false):integer;
      procedure MarcaProcessoDigitado;
      procedure GravaValorCorrigido(valor: double);
      procedure GuardaTipoProcesso;
      procedure GravaValorTotalDaNota;
      procedure ObtemNumeroNotaEmpresa;
      procedure LimpaDadosNota;
      procedure EncerraNotaPeloValor;
      procedure GravaNotaNoProcesso;
      function NotaJaFoiDigitada : boolean;
      function ObtemValorTotalDaNota : double;
      procedure MarcaDocumentoFinalizado;
      function ObtemEnvolvidosDoProcesso_2(fgDrcAtivas:integer; buscarAutor: boolean):integer;
//      procedure NotificaNotaFinalizada;
//      procedure NotificaNotaIniciada;
//      procedure NotificaErro(mensagem: string);

      procedure CriaIndiceBaseI;
      procedure CriaIndiceBaseIII;
      procedure CriaIndiceBaseIV;
      procedure CriaIndiceBaseII;
      procedure CriaIndiceBaseCompartilhada;
      procedure CriaIndiceBaseTrabalhistas;
      procedure CriaIndiceBaseMigrados;
      procedure CriaIndiceBaseV;
      procedure CriaIndiceBaseVII;
      procedure CriaIndiceBaseVIII;
      procedure CriaIndiceBaseIX;
      procedure CriaIndiceBaseX;
      procedure CriaIndiceBaseXI;
      procedure CriaIndiceBaseBaixados;
      procedure CriaIndiceBaseRecuperados;

      procedure CadastraPlanilha;
      procedure CadastraAto;
      function JaExistePlanilhaDoMes: integer;
      procedure MarcaPlanilhaImportada;
      procedure ObtemPlanilhasValidar;
      procedure CarregaCamposDaTabela;
      function ProcessoExisteNoGcpj:integer;
      function GravaEmpresaGrupo(tipo: integer):integer;
      procedure ValidaValor;
      procedure GravaValorCalculo;
      procedure ObtemReclamadasDoProcesso(ptipoProcesso: string);
      procedure MarcaPlanilhaValidada;
      procedure RemoveNotaProcesso;
      function SomaTotalDigitadoDaNota:double;
      function ObtemDataDigitar(formato: integer):String;
      function ObtemTotalAtosDigitados:integer;
      procedure SalvaAlterado(tipoato: string);
      function JaAlterado(tipoato: string):boolean;
      procedure ObtemPlanilhasDigitar;
      procedure MarcaPlanilhaFinalizada;
      procedure MarcaPlanilhaNaoFinalizada;
      procedure CadastraDataInicial(dtInicial: TDateTime);
      procedure ObtemPlanilhasCadastradas(const order: integer=0);
      procedure ObtemInconsistenciasImportacao;
      procedure ObtemInconsistenciasPosImportacao;
      procedure ObtemValoresRecalculadosSistema;
      procedure ObtemNotasFinalizadas(numNota: string);
      function TemProcessoDuplicado:boolean;
      procedure ObtemPlanilhasEnviarGpj(const filtro: integer=0);

      function AtoDistribuidoParaEscritorio : integer;
      procedure RemovePlanilhas(idPlanilha, idEscritorio: integer; anomesreferencia: string);

      function ObtemCodOrgProcesso:integer;
      function ObtemUfProcesso(orgaoJulgador: integer):string;
      function ObtemNomeOrgaoJulgador(orgaoJulgador: integer) : string;

      function ObtemDirLck:string;

      function EhDrcContraria : boolean;

      function EhDrcAtiva : boolean;

      function EhSubtipoTarifa(lista: TStringList) : boolean;

      function EhEscritorioTarifa(lista: TStringList) : boolean;

      function ProcessoEhJuizado : boolean;

      procedure CarregaFiliais;

      function ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso, codescritorio: string; lstCnpjs: TStringList):integer;

      function LaudoPericialJaFoiPago : boolean;

      //inserido em 03/04/2014
      procedure ObtemAtosPendentes;

      function TemConsolidacaoPagaNoSistema: boolean;

      //inserido em 07/07/2014
      function ObtemValorSemReajuste(andamento: string) : double;
      function ProcessoDistribuidoParaAdvogadoInterno : boolean;

      //inserido em 31/07/2014
      function ConfiguracaoSemReajusteEstaOk : boolean;
      procedure CriaIndicesDigitaHonorarios;
      procedure ResetFlagsDaPlanilha;

      //incluido em 05/02/2015
      function ObtemTipoAndamento(andamento: string; tipoProcesso: string; motivoBaixa: string) : string;
      function ObtemDependenciaDigitar : string;

      procedure RemoveOcorrenciasDaPlanilha;

      procedure CadastraDeParaEmpresasGrupo;

   end;

implementation

{ TPlanilhaJBM }

function TPlanilhaJBM.AtoDistribuidoParaEscritorio: integer;
var
   lstCnpj : TStringList;
   strValues : Tstringlist;
   i, j : integer;
begin
   result := dmgcpcj_base_I.ObtemEscritorioDistribuidoPara(gcpj);
   if result = 0 then
      //ato não está distribuído
     exit;

   result := 1;

   if dmgcpcj_base_I.dtsXII.FieldByName('CPSSOA_EXTER_ADVOG').AsInteger = codGcpjEscritorio then
   begin
      //distribuído para o mesmo escritório
      result := 9;
      exit;
   end;

   //distribuído para outros escritórios
   lstCnpj := TStringList.Create;
   try
      lstCnpj.LoadFromFile(ExtractFilePath(Application.ExeName) + 'ESCRITORIOS_DISTRIBUICAO.TXT');

      for i := 0 to lstCnpj.Count - 1 do
      begin
         strValues := TstringList.Create;
         try
            strtolist(lstCnpj.Strings[i], ';', strValues);

            for j := 0 to (strvalues.count - 1) do
            begin
               if codGcpjEscritorio = StrToInt(strValues.Strings[j]) then
                  break;
            end;

            if j > (strValues.count - 1) then
               continue;

            //achou

            for j := 0 to (strValues.count - 1) do
            begin
               if codGcpjEscritorio = StrToInt(strValues.Strings[j]) then
                  continue;

               if  dmgcpcj_base_I.dtsXII.FieldByName('CPSSOA_EXTER_ADVOG').AsInteger <> StrToInt(strValues.Strings[j]) then
                  continue;

               result := 9;
               exit;
            end;
         finally
            strValues.free;
         end;
      end;
   finally
      lstCnpj.Free;
   end;

end;

procedure TPlanilhaJBM.CadastraAto;
begin
   dmHonorarios.CadastraAto(idescritorio,
                            idplanilha,
                            anomesreferencia,
                            sequencia,
                            linhaplanilha,
                            Cliente,
                            processo,
                            partecontraria,
                            RemoveNaoNumericos(gcpj),
                            Comarca,
                            Vara,
                            tipoandamento,
                            UF,
                            divisaodafilial,
                            idprocessounico,
                            Porcentagem,
                            valBeneficioEconomico,
                            Valor,
                            datadoato,
                            valorBase);
end;

procedure TPlanilhaJBM.CadastraDataInicial(dtInicial: TDateTime);
begin
   dmHonorarios.CadastraDataInicial(anomesreferencia, dtInicial);
end;

procedure TPlanilhaJBM.CadastraPlanilha;
begin
   idplanilha := dmHonorarios.CadastraPlanilha(idEscritorio, anomesreferencia, Fsequencia);
end;

procedure TPlanilhaJBM.CarregaCamposDaTabela;
begin
   idEscritorio := dtsAtos.FieldByName('idescritorio').AsInteger;
   idPlanilha := dtsAtos.FieldByName('idplanilha').AsInteger;
   anomesreferencia := dtsAtos.FieldByName('anomesreferencia').AsString;
   sequencia :=  dtsAtos.FieldByName('sequencia').AsInteger;
   linhaPlanilha := dtsAtos.FieldByName('linhaplanilha').AsInteger;
   gcpj := dtsAtos.FieldByName('gcpj').AsString;
   nome_reu := dtsAtos.FieldByName('cliente').AsString;
   tipoandamento := dtsAtos.FieldByName('tipoandamento').AsString;
   datadoato := dtsAtos.FieldByName('datadoato').AsDateTime;
   tipoProcesso := dtsAtos.FieldByName('fgtipoprocesso').AsString;
   valorBase := dtsAtos.FieldByName('valorbase').AsFloat;
   dataPlanilha := dmHonorarios.ObtemDataImportacaoPlanilha(idPlanilha);
   fgDrcContrarias := dtsAtos.FieldByName('fgDrcContrarias').AsInteger;
end;

constructor TPlanilhaJBM.Create(const PProcedure: TProcExibeMensagem = Nil);
begin
   inherited Create;

   anomesreferencia := '';
   idplanilha := 0;
   tipoProcesso := '';
   ehJec := false;
   drcContrarias := false;
   lstCnpjs := TStringList.Create;

   dtsPlanilha := TADODataSet.Create(nil);
   dtsPlanilha.Connection := dmHonorarios.adoConn;

   dtsAtos := TADODataSet.Create(nil);
   dtsAtos.Connection := dmHonorarios.adoConn;

   dsPlan := TDataSource.Create(nil);
   dsPlan.DataSet := dtsPlanilha;
   dsPlan.OnDataChange := dsPlanDataChange;;

   dtsReclamadas := TADODataSet.Create(nil);
   dtsReclamadas.Connection := dmHonorarios.adoConn;

   dtsNotasPendentes := TADODataSet.Create(nil);
   dtsNotasPendentes.Connection := dmHonorarios.adoConn;

   dsNotasPendentes := TDataSource.Create(nil);
   dsNotasPendentes.DataSet := dtsNotasPendentes;
   dsNotasPendentes.OnDataChange := dsNotasPendentesDataChange;

   dtsNotasDigitando := TADODataSet.Create(nil);
   dtsNotasDigitando.Connection := dmHonorarios.adoConn;

   dsNotasDigitando := TDataSource.Create(nil);
   dsNotasDigitando.DataSet := dtsNotasDigitando;

   ExibeMensagem := PProcedure;
end;

procedure TPlanilhaJBM.CriaIndiceBaseCompartilhada;
begin
   dmgcpj_compartilhado.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseI;
begin
   dmgcpcj_base_I.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseII;
begin
   dmgcpj_base_ii.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseIII;
begin
   dmgcpj_base_iii.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseIV;
begin
   dmgcpj_base_iv.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseMigrados;
begin
   dmGcpj_migrados.CreateIndex;
end;

procedure TPlanilhaJBM.dsNotasPendentesDataChange(Sender: TObject;
  Field: TField);
begin
   dtsNotasDigitando.Close;
   If dtsNotasPendentes.Eof then
      exit;

   dmHonorarios.ObtemNotasDigitando(dtsPlanilha.FieldByName('idplanilha').AsInteger,
                                    dtsNotasPendentes.FieldByName('numeronota').AsInteger,
                                    FdtsNotasDigitando);
end;

procedure TPlanilhaJBM.dsPlanDataChange(Sender: TObject; Field: TField);
begin
   dtsNotasPendentes.Close;
   If dtsPlanilha.Eof then
      exit;

   dmHonorarios.ObtemNotasPendentes(dtsPlanilha.FieldByName('idplanilha').AsInteger, fdtsNotasPendentes);
end;

procedure TPlanilhaJBM.EncerraNotaPeloValor;
begin
   dmHonorarios.MarcaNotaAtingiuValor(StrToInt(LeftStr(IntToStr(numerodanota),4)),
                                      StrToInt(RightStr(IntToStr(numerodanota),5)),
                                      anomesreferencia,
                                      empresaligada);
end;

procedure TPlanilhaJBM.Finalizar;
begin
   dtsAtos.Free;
   dtsPlanilha.Free;
   dsPlan.free;
   dtsReclamadas.Free;
end;


function TPlanilhaJBM.GravaEmpresaGrupo(tipo: integer):integer;
var
   ret : integer;
   codempresa, paraCodigo: integer;
   nomeempresa : string;
   codtratar : string;
   paranome : string;
   dts : TAdoDataset;
   IAGRUP_BDSCO : string;
begin
   result := 0;
   codempresa := 0;

   //cadastra DeParaEmpresasGrupo

//mexer alexandre    

(**   if ((Pos('TRABALHISTA', dtsAtos.FieldByName('tipoandamento').AsString) = 0) and
      (Pos('TRABALHO', dtsAtos.FieldByName('vara').AsString) = 0)) and
      (dtsAtos.FieldByName('tipoandamento').AsString <> 'INSTRUCAO')  then //civel*)
   dts := TADODataset.Create(Nil);
   try
      if (tipoProcesso <> 'TR') and (tipoProcesso <> 'TO') and (tipoProcesso <> 'TA') then
      begin
         if Assigned(ExibeMensagem) then
            ExibeMensagem('   ==> (1) Obtendo empresa grupo civel (Base Relatórios)');

         ret := dmgcpj_compartilhado.ObtemCodEmpresaGrupoCivel(dtsAtos.FieldByName('gcpj').AsString, dts);
         if ret = 0 then
            exit;

         if (Not dts.FieldByName('CodJunAgrup').IsNull) and
            (dts.FieldByName('CodJunAgrup').AsInteger <> 0) then
         begin
            if Assigned(ExibeMensagem) then
               ExibeMensagem('   ==> (2) Obtendo empresa grupo (Base IV)');

            ret := dmgcpj_base_iv.ObtemNomeEmpresaGrupo(dts.FieldByName('CodJunAgrup').AsInteger);
            if ret = 0 then
               exit;

            IAGRUP_BDSCO := dmgcpj_base_iv.dts3T.FieldByName('IAGRUP_BDSCO').AsString;

            if Assigned(ExibeMensagem) then
               ExibeMensagem('   ==> (3) Obtendo complemento da empresa grupo (Base IV)');

            ret := dmgcpj_base_iv.ObtemComplementoEmpresaGrupo(dts.FieldByName('CodJunEmp').AsInteger,
                                                               dts.FieldByName('CodJun').AsInteger);
            if ret = 0 then  //se não tem mesu não é bradesco
            begin
               if  IAGRUP_BDSCO = 'BRADESCO - AGENCIA' then
               begin
                  codempresa := 237;
                  nomeempresa := IAGRUP_BDSCO;
               end
               else
               begin
                  codempresa := dts.FieldByName('CodJun').AsInteger;
                  nomeempresa := dts.FieldByName('NomJun').AsString;
               end;
            end
            else
            begin
               if dts.FieldByName('CodJunEmp').AsInteger = 237 then
               begin
//                  if dmgcpj_base_iv.dtsMesu.FieldByName('ITPO_DEPDC').AsString = 'EMPRESAS LIGADAS' then
                  if dmgcpj_base_iv.dtsMesu.FieldByName('ITPO_DEPDC').AsString = 'EMPRESA LIGADA' then   // modificado Alexandre em 04/03/2018 - solicitado pela Maria do Carmo
                  begin
                     codempresa := dts.FieldByName('CodJun').AsInteger;
                     nomeempresa := IAGRUP_BDSCO;
                  end
                  else
                  begin
                     if tipo = 0 then
                     begin
                        if (dts.FieldByName('codjuncon').AsInteger) <>
                           (dts.FieldByName('codjun').AsInteger) then
                        begin
                           if Assigned(ExibeMensagem) then
                              ExibeMensagem('   ==> (4) Obtendo reclamadas no processo (Base Interna)');

                           //verifica se a primeira reclamada não é agencia e é o codjuncon
                           dmHonorarios.ObtemReclamadasDoProcesso(idEscritorio, idplanilha, anomesreferencia, sequencia, linhaPlanilha, fdtsreclamadas,tipoProcesso);

                           nomeempresa := '';
                           codempresa := 0;

                           while Not dtsReclamadas.Eof do
                           begin
                              if (dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A') and
                                 (dts.FieldByName('codjuncon').AsInteger = dtsReclamadas.FieldByName('codigoreclamada').AsInteger) and
                                 (dtsReclamadas.FieldByName('codempresa').AsInteger <> 237) then
                              begin
                                 codempresa := dtsReclamadas.FieldByName('codempresa').AsInteger;
                                 nomeempresa := dts.FieldByName('NomJunCon').AsString;
                                 break;
                              end;
                              if (dtsReclamadas.FieldByName('tiporeclamada').AsString = 'L') and
                                 (dts.FieldByName('codjuncon').AsInteger = dtsReclamadas.FieldByName('codigoreclamada').AsInteger) then
                              begin
                                 codempresa := dtsReclamadas.FieldByName('codigoreclamada').AsInteger;
                                 nomeempresa := dts.FieldByName('NomJunCon').AsString;
                                 break;
                              end;

                              if tipoProcesso <> 'CI' then
                              begin
                                 if dts.FieldByName('codjuncon').AsInteger = dtsReclamadas.FieldByName('codigoreclamada').AsInteger then
                                 begin
                                    codempresa := dtsReclamadas.FieldByName('codempresa').AsInteger;
                                    nomeempresa := dts.FieldByName('NomJunCon').AsString;
                                    break;
                                 end;
                              end;
                              dtsReclamadas.Next;
                           end;

                           if nomeempresa = '' then
                           begin
                              codempresa := dts.FieldByName('CodJunEmp').AsInteger;
                              nomeempresa := IAGRUP_BDSCO;
                           end;
                        end
                        else
                        begin
                           codempresa := dts.FieldByName('CodJunEmp').AsInteger;
                           nomeempresa := IAGRUP_BDSCO;
                        end;
                        end
                        else
                     begin
                        if IAGRUP_BDSCO <> dts.FieldByName('NomJunCon').AsString then
                        begin
                           codtratar := dts.FieldByName('CODJUNAGRUP').AsString;
                           if LeftStr(codtratar, 2) = '63' then //ibi
                              codempresa := 63
                           else
                           if RightStr(codtratar, 2) = '03' then
                              codempresa := StrToInt(Copy(codtratar, 1, Length(codtratar)-2))
                           else
                           if (RightStr(codtratar, 2) = '01') or (RightStr(codtratar, 2) = '02') then
                              codempresa := StrToInt(Copy(codtratar, 1, Length(codtratar)-3));
                              nomeempresa := IAGRUP_BDSCO;
                           end
                        else
                        begin
                           codempresa := dts.FieldByName('CodJunEmp').AsInteger;
                           nomeempresa := IAGRUP_BDSCO;
                        end;
                     end;
                  end
               end
               else
               begin
                  codempresa := dts.FieldByName('CodJunEmp').AsInteger;
                  nomeempresa := IAGRUP_BDSCO;
               end;
            end;
         end
         else
         begin
            if (dts.FieldByName('NomJunCon').AsString = 'BRADESCO DIA E NOITE') or
               (dts.FieldByName('NomJunCon').AsString = 'BRADESCO - AGENCIA') OR
               (dts.FieldByName('NomJunCon').AsString = 'BRADESCO - DEPARTAMENTO') or
               (dts.FieldByName('NomJunCon').AsString = 'EMPR.FINANCIAMENTO')  then
            begin
               codempresa := 237;
               nomeempresa := dts.FieldByName('NomJunCon').AsString;
            end
            else
            begin
               if (dts.FieldByName('CodJunCon').AsInteger <> 0) then
               begin
                  codempresa := dts.FieldByName('CodJunCon').AsInteger;
                  nomeempresa := dts.FieldByName('NomJunCon').AsString;
               end;
            end;
         end
      end
      else
      begin
         if Assigned(ExibeMensagem) then
            ExibeMensagem('   ==> (5) Obtendo Empresa Grupo Trabalhista (Base relatórios)');

         ret := dmgcpj_compartilhado.ObtemCodEmpresaGrupoTrabalhista(dtsAtos.FieldByName('gcpj').AsString, dts);
         if ret = 0 then
            exit;

         if dts.FieldByName('CodJunAgrup').AsInteger = 0 then
         begin
            GravaOcorrencia(2, 'Erro obtendo Empresa Grupo (falta CodJunAgrup)');
            MarcaProcessoCruzadoGcpj(9);
            result := -1;
            exit;
         end
         else
         begin
            if Assigned(ExibeMensagem) then
               ExibeMensagem('   ==> (6) Obtendo Nome Empresa Grupo (Base IV)');

            ret := dmgcpj_base_iv.ObtemNomeEmpresaGrupo(dts.FieldByName('CodJunAgrup').AsInteger);
         end;

         if ret = 0 then
         begin
            GravaOcorrencia(2, 'Erro obtendo Empresa Grupo (CodJunAgrup inválido)');
            MarcaProcessoCruzadoGcpj(9);
            result := -1;
            exit;
         end;

         IAGRUP_BDSCO := dmgcpj_base_iv.dts3T.FieldByName('IAGRUP_BDSCO').AsString;

         if dts.FieldByName('CodJunEmp').AsInteger = 0 then //trabalhista emp;terceira
         begin
            if Assigned(ExibeMensagem) then
               ExibeMensagem('   ==> (7) Obtendo Complemento Empresa Grupo(Base IV)');

            ret := dmgcpj_base_iv.ObtemComplementoEmpresaGrupo(dts.FieldByName('CodJunConEmp').AsInteger,
                                                               dts.FieldByName('CodJunCon').AsInteger);
            if ret = 0 then
               exit;

            if dts.FieldByName('CodJunConEmp').AsInteger = 237 then
            begin
//               if dmgcpj_base_iv.dtsMesu.FieldByName('ITPO_DEPDC').AsString = 'EMPRESAS LIGADAS' then
               if dmgcpj_base_iv.dtsMesu.FieldByName('ITPO_DEPDC').AsString = 'EMPRESA LIGADA' then  // modificado Alexandre em 04/03/2018 - solicitado pela Maria do Carmo
               begin
                  codempresa := dts.FieldByName('CodJunCon').AsInteger;
                  nomeempresa := IAGRUP_BDSCO;
               end
               else
               begin
                  if tipo = 0 then
                  begin
                     if (dts.FieldByName('codjuncon').AsInteger) <>
                        (dts.FieldByName('codjun').AsInteger) then
                     begin
                           //verifica se a primeira reclamada não é agencia e é o codjuncon
                        if Assigned(ExibeMensagem) then
                           ExibeMensagem('   ==> (8) Obtendo reclamadas do processo(Base Interna)');

                        dmHonorarios.ObtemReclamadasDoProcesso(idEscritorio, idplanilha, anomesreferencia, sequencia, linhaPlanilha, fdtsreclamadas, tipoProcesso);

                        nomeempresa := '';
                        codempresa := 0;

                        while Not dtsReclamadas.Eof do
                        begin
                           if (dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A') and
                              (dts.FieldByName('codjuncon').AsInteger = dtsReclamadas.FieldByName('codigoreclamada').AsInteger) and
                              (dtsReclamadas.FieldByName('codempresa').AsInteger <> 237) then
                           begin
                              codempresa := dtsReclamadas.FieldByName('codempresa').AsInteger;
                              nomeempresa := dts.FieldByName('NomJunCon').AsString;
                              break;
                           end;
                           dtsReclamadas.Next;
                        end;

                        if nomeempresa = '' then
                        begin
                           codempresa := dts.FieldByName('CodJunConEmp').AsInteger;
                           nomeempresa := IAGRUP_BDSCO;
                        end;
                     end
                     else
                     begin
                        codempresa := dts.FieldByName('CodJunConEmp').AsInteger;
                        nomeempresa := IAGRUP_BDSCO;
                     end;
                  end
                  else
                  begin
                     if IAGRUP_BDSCO <> dts.FieldByName('NomJunCon').AsString then
                     begin
                        codtratar := dts.FieldByName('CODJUNAGRUP').AsString;
                        if LeftStr(codtratar, 2) = '63' then //ibi
                           codempresa := 63
                        else
                        if RightStr(codtratar, 2) = '03' then
                           codempresa := StrToInt(Copy(codtratar, 1, Length(codtratar)-2))
                        else
                        if (RightStr(codtratar, 2) = '01') or (RightStr(codtratar, 2) = '02') then
                           codempresa := StrToInt(Copy(codtratar, 1, Length(codtratar)-3));
                        nomeempresa := IAGRUP_BDSCO;
                     end
                     else
                     begin
                        codempresa := dts.FieldByName('CodJunConEmp').AsInteger;
                        nomeempresa := IAGRUP_BDSCO;
                     end;
                  end;
               end
            end
            else
            begin
               codempresa := dts.FieldByName('CodJunConEmp').AsInteger;
               nomeempresa := IAGRUP_BDSCO;
            end;
         end
         else
         begin
            if Assigned(ExibeMensagem) then
               ExibeMensagem('   ==> (9) Obtendo complemento empresa grupo(Base IV)');

            ret := dmgcpj_base_iv.ObtemComplementoEmpresaGrupo(dts.FieldByName('CodJunEmp').AsInteger,
                                                               dts.FieldByName('CodJun').AsInteger);
            if ret = 0 then
               exit;

            if dts.FieldByName('CodJunEmp').AsInteger = 237 then
            begin
               if dmgcpj_base_iv.dtsMesu.FieldByName('CTPO_DEPDC').AsString = 'L' then
               begin
                  codempresa := dts.FieldByName('CodJun').AsInteger;
                  nomeempresa := IAGRUP_BDSCO;
               end
               else
               begin
                  if dmgcpj_base_iv.dtsMesu.FieldByName('CTPO_DEPDC').AsString = 'D' then
                  begin
                     codempresa := 237;
                     nomeempresa := 'BANCO BRADESCO S/A';
                  end
                  else
                  begin
                     if tipo = 0 then
                     begin
                        if (dts.FieldByName('codjuncon').AsInteger) <>
                           (dts.FieldByName('codjun').AsInteger) then
                        begin
                           //verifica se a primeira reclamada não é agencia e é o codjuncon

                           if Assigned(ExibeMensagem) then
                              ExibeMensagem('   ==> (10) Obtendo reclamadas do processo(Base Interna)');

                           dmHonorarios.ObtemReclamadasDoProcesso(idEscritorio, idplanilha, anomesreferencia, sequencia, linhaPlanilha, fdtsreclamadas, '');

                           nomeempresa := '';
                           codempresa := 0;

                           while Not dtsReclamadas.Eof do
                           begin
                              if (dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A') and
                                 (dts.FieldByName('codjuncon').AsInteger = dtsReclamadas.FieldByName('codigoreclamada').AsInteger) and
                                 (dtsReclamadas.FieldByName('codempresa').AsInteger <> 237) then
                              begin
                                 codempresa := dtsReclamadas.FieldByName('codempresa').AsInteger;
                                 nomeempresa := dts.FieldByName('NomJunCon').AsString;
                                 break;
                              end;
                              dtsReclamadas.Next;
                           end;

                           if nomeempresa = '' then
                           begin
                              codempresa := dts.FieldByName('CodJunEmp').AsInteger;
                              nomeempresa := IAGRUP_BDSCO;
                           end;
                        end
                        else
                        begin
                           codempresa := dts.FieldByName('CodJunEmp').AsInteger;
                           nomeempresa := IAGRUP_BDSCO;
                        end;
                     end
                     else
                     begin
                        if IAGRUP_BDSCO <> dts.FieldByName('NomJun').AsString then
                        begin
                           codtratar := dts.FieldByName('CODJUNAGRUP').AsString;
                           if LeftStr(codtratar, 2) = '63' then //ibi
                              codempresa := 63
                           else
                           if RightStr(codtratar, 2) = '03' then
                              codempresa := StrToInt(Copy(codtratar, 1, Length(codtratar)-2))
                           else
                           if (RightStr(codtratar, 2) = '01') or (RightStr(codtratar, 2) = '02') then
                              codempresa := StrToInt(Copy(codtratar, 1, Length(codtratar)-3));
                           nomeempresa := IAGRUP_BDSCO;
                        end
                        else
                        begin
                           codempresa := dts.FieldByName('CodJunEmpp').AsInteger;
                           nomeempresa := IAGRUP_BDSCO;
                        end;
                     end;
                  end
               end
            end
            else
            begin
               codempresa := dts.FieldByName('CodJunEmp').AsInteger;
               nomeempresa := IAGRUP_BDSCO;
            end;
         end;
      end;
   finally
      dts.free;
   end;

   if Pos('TEMPO SER', UpperCase(nomeempresa)) <> 0 then
   begin
      paraNome := 'TEMPO SERVICOS LTDA';
      paraCodigo := 5172;
   end
   else
   begin
      paraNome := nomeempresa;
      if Assigned(ExibeMensagem) then
         ExibeMensagem('   ==> (11) Obtendo DE/PARA empresa grupo(Base Interna)');
      paraCodigo := dmHonorarios.ObtemDeParaEmpresaGrupo(codempresa, tipoProcesso, paraNome);
   end;

   //alterado em 23/04/2015
   if (codempresa = 4510) or (codempresa = 4840) or (codempresa = 4130 ) or (codempresa = 4008) then
      codempresa := 237;

   if Assigned(ExibeMensagem) then
      ExibeMensagem('   ==> (12) Gravando Empresa Grupo(Base Interna)');

   dmHonorarios.GravaDadosEmpresaGrupo(idescritorio,
                                       idplanilha,
                                       anomesreferencia,
                                       sequencia,
                                       linhaplanilha,
                                       codempresa,
                                       nomeempresa,
                                       paracodigo,
                                       tipo,
                                       paranome);

   if codempresa = 0 then
   begin
      if tipoprocesso = 'PE' then
         GravaOcorrencia(2, 'Processo penal - sem envolvido')
      else
         GravaOcorrencia(2, 'Processo sem empresa ligada');
      MarcaProcessoCruzadoGcpj(9);
      result := -1;
      exit;
   end;

   result := 1;
end;


procedure TPlanilhaJBM.GravaNotaNoProcesso;
begin
   dmHonorarios.GravaNotaNoProcesso(idescritorio,
                                    idplanilha,
                                    anomesreferencia,
                                    sequencia,
                                    linhaplanilha,
                                    numeroDaNota);
end;

procedure TPlanilhaJBM.GravaOcorrencia(tipoErro: integer;
  mensagem: string);
begin
   dmHonorarios.GravaOcorrencia(idplanilha,
                                linhaPlanilha,
                                tipoErro,
                                mensagem);
end;

procedure TPlanilhaJBM.GravaTipoProcesso;
begin
   dmHonorarios.GravaDadosTipoProcesso(idescritorio,
                                       idplanilha,
                                       anomesreferencia,
                                       sequencia,
                                       linhaplanilha,
                                       dmgcpj_compartilhado.ObtemAreaProcesso(dtsAtos.FieldByName('gcpj').AsString));
end;

function TPlanilhaJBM.GravaTipoSubtipo : integer;
var
   ret : integer;
   juizado : integer;
   fgContraria : integer;
   fgAtiva : integer;
   orgaoJulgador : integer;
   nomeOrgaoJulgador : string;
   pdts : TAdoDataSet;
begin
   fgAtiva := 0;
   fgContraria := 0;
   Juizado := 0;

   result := 1;

   gcpj := dtsAtos.FieldByName('gcpj').AsString;
   orgaoJulgador := ObtemCodOrgProcesso;
   nomeOrgaoJulgador := ObtemNomeOrgaoJulgador(orgaoJulgador);

   pdts := TADODataSet.Create(nil);
   try
      ret := dmgcpj_compartilhado.ObtemTipoSubtipo(dtsAtos.FieldByName('gcpj').AsString, pdts);

      if (ret = 0) or (pdts.FieldByName('codacao').AsInteger = 0) then
      begin
         //se for DRC ativa pode estar zerado
         //alterado em 30/10/2013
         if (pdts.FieldByName('coddejur').AsInteger = 4429) or
         (pdts.FieldByName('coddejur').AsInteger = 4799) or
         (pdts.FieldByName('coddejur').AsInteger = 8288) then
         begin
            If (Pos('JEC', nomeOrgaoJulgador) <> 0) OR (Pos('JUIZADO ESPECIAL', nomeOrgaoJulgador) <> 0)then
               juizado := 1;

            fgAtiva := 1;
            fgContraria := 0;

            dmHonorarios.GravaDadosTipoSubtipo(idescritorio,
                                               idplanilha,
                                               anomesreferencia,
                                               sequencia,
                                               linhaplanilha,
                                               0,
                                               '',
                                               0,
                                               '',
                                               juizado,
                                               fgContraria,
                                               fgAtiva);
            exit;
         end;

         result := 0;
         exit;
      end;

      if (pdts.FieldByName('codacao').AsInteger = 8908) then
      begin
         result := 2;
         exit;
      end;

      if (pdts.FieldByName('codacao').AsInteger in [35,46,52 , 53 , 60 , 68 , 69 , 70 , 77 , 78 , 79 , 80 , 81 , 82 , 87 , 88 , 89 , 91, 93]) or
         (pdts.FieldByName('codacao').AsInteger = 8905) or
         (Pos('JEC', nomeOrgaoJulgador) <> 0) or
         (Pos('JUIZADO ESPECIAL', nomeOrgaoJulgador) <> 0)then
         juizado := 1
      else
      if (pdts.FieldByName('codacao').AsInteger = 8911) or
         (pdts.FieldByName('codacao').AsInteger = 8912) or
         (pdts.FieldByName('codacao').AsInteger = 8913) or
         (pdts.FieldByName('codacao').AsInteger = 8914) or
         (pdts.FieldByName('codacao').AsInteger = 8915) then
      begin
         If (Pos('JEC', nomeOrgaoJulgador) <> 0) OR (Pos('JUIZADO ESPECIAL', nomeOrgaoJulgador) <> 0)then
            juizado := 1;
      end;

      if //(pdts.FieldByName('coddejur').AsInteger = 4785) and
         ((pdts.FieldByName('codacao').AsInteger = 8911) or
         (pdts.FieldByName('codacao').AsInteger = 8912) or
         (pdts.FieldByName('codacao').AsInteger = 8913) or
         (pdts.FieldByName('codacao').AsInteger = 8914)) then
         fgContraria := 1;

      if (pdts.FieldByName('codacao').AsInteger = 8915) then
         fgAtiva := 1;


      if (pdts.FieldByName('coddejur').AsInteger = 4429) or (pdts.FieldByName('coddejur').AsInteger = 8288) or (pdts.FieldByName('coddejur').AsInteger = 4799) then
      begin
         if fgContraria = 1 then
         begin
            result := 3;
            exit;
         end;
         fgAtiva := 1;
      end;

      dmHonorarios.GravaDadosTipoSubtipo(idescritorio,
                                         idplanilha,
                                         anomesreferencia,
                                         sequencia,
                                         linhaplanilha,
                                         pdts.FieldByName('codacao').AsInteger,
                                         pdts.FieldByName('nomacao').AsString,
                                         pdts.FieldByName('codsub').ASInteger,
                                         pdts.FieldByName('nomsub').AsString,
                                         juizado,
                                         fgContraria,
                                         fgAtiva);

      if
         (pdts.FieldByName('coddejur').AsInteger <> 4799) and
         (pdts.FieldByName('coddejur').AsInteger <> 4429) and
         (pdts.FieldByName('coddejur').AsInteger <> 8288) and

         ((Pos('ACORDO', tipoandamento) <> 0) or (Pos('DESBLOQUEIO', tipoandamento) <> 0)) then
         result := 4
      else
      if
         (pdts.FieldByName('coddejur').AsInteger <> 4799) and
         (pdts.FieldByName('coddejur').AsInteger <> 4429) and
         (pdts.FieldByName('coddejur').AsInteger <> 8288) and
         (Pos('DESBLOQUEIO', tipoandamento) <> 0) then
         result := 4;
   finally
      pdts.Free;
   end;
end;

procedure TPlanilhaJBM.GravaValorCalculo;
begin
   if (Not dtsAtos.FieldByName('valordopedido').IsNull) and (dtsAtos.FieldByName('valordopedido').AsFloat <> 0) then
   begin
      valorpedido := dtsAtos.FieldByName('valordopedido').AsFloat;
      exit;
   end;

   dmgcpj_base_ii.ObtemValorCalculoOriginal(gcpj);
   if dmgcpj_base_ii.dtsb7.Eof then
      exit;

   dmHonorarios.GravaValorCalculo(idescritorio,
                                  idplanilha,
                                  anomesreferencia,
                                  sequencia,
                                  linhaplanilha,
                                  dmgcpj_base_ii.dtsB7.FieldByName('VORIGN_CALC_PROCS').AsFloat);
   valorpedido := dmgcpj_base_ii.dtsB7.FieldByName('VORIGN_CALC_PROCS').AsFloat;
end;

procedure TPlanilhaJBM.GravaValorCorrigido(valor: double);
begin
  dmHonorarios.GravaValorCorrigido(idescritorio,
                                   idplanilha,
                                   anomesreferencia,
                                   sequencia,
                                   linhaplanilha,
                                   valor);
end;

procedure TPlanilhaJBM.GravaValorTotalDaNota;
begin
   dmHonorarios.GravaValorTotalDaNota(StrToInt(LeftStr(IntToStr(numerodanota),4)),
                                      StrToInt(RightStr(IntToStr(numerodanota),5)),
                                      anomesreferencia,
                                      totalAtos,
                                      valorTotalNota);
end;

procedure TPlanilhaJBM.GuardaTipoProcesso;
begin
   tipoProcesso := dtsAtos.FieldByName('fgtipoprocesso').asString;
   ehJec := (dtsAtos.FieldByName('fgjuizado').AsInteger = 1);
end;

function TPlanilhaJBM.JaAlterado(tipoato: string): boolean;
begin
   dmHonorarios.dts.Close;
   dmHonorarios.dts.CommandText := 'select * from tbalterados where ' +
                                   'numeronota = ' + dtsAtos.FieldByName('numeronota').AsString + ' and ' +
                                   'gcpj = ' + dtsAtos.FieldByName('gcpj').AsString + ' and ' +
                                   'tipoandamento = ''' + tipoato + '''';
   dmHonorarios.dts.Open;
   Result := (dmHonorarios.dts.RecordCount > 0);
end;

function TPlanilhaJBM.JaExistePlanilhaDoMes: integer;
begin
   result := dmHonorarios.JaExistePlanilhaDoMes(idescritorio, anomesreferencia);
end;

procedure TPlanilhaJBM.LimpaDadosNota;
begin
   empresaLigada := 0;
   valorTotalNota := 0;
   totalAtos := 0;
   numerodanota := 0;
end;

procedure TPlanilhaJBM.MarcaDocumentoFinalizado;
begin
   dmHonorarios.MarcaDocumentoFinalizado(StrToInt(LeftStr(IntToStr(numerodanota),4)),
                                         StrToInt(RightStr(IntToStr(numerodanota),5)),
                                         anomesreferencia,
                                         empresaligada);
end;

procedure TPlanilhaJBM.MarcaNomeEnvolvidoOk;
begin
   dmHonorarios.MarcaNomeEnvolvidoOk(idescritorio,
                                     idplanilha,
                                     anomesreferencia,
                                     sequencia,
                                     linhaplanilha);
end;

procedure TPlanilhaJBM.MarcaPlanilhaFinalizada;
begin
   dmHonorarios.MarcaPlanilhaFinalizada(idEscritorio, idplanilha, anomesreferencia, sequencia);
end;

procedure TPlanilhaJBM.MarcaPlanilhaImportada;
begin
   dmHonorarios.MarcaPlanilhaImportada(idEscritorio, idplanilha, anomesreferencia, sequencia);
end;

procedure TPlanilhaJBM.MarcaPlanilhaValidada;
begin
   dmHonorarios.MarcaPlanilhaValidada(idEscritorio, idplanilha, anomesreferencia, sequencia);
end;

procedure TPlanilhaJBM.MarcaProcessoCruzadoGcpj(estagio: integer);
begin
   dmHonorarios.MarcaProcessoCruzadoGcpj(estagio,
                                         idescritorio,
                                         idplanilha,
                                         anomesreferencia,
                                         sequencia,
                                         linhaplanilha);
end;

procedure TPlanilhaJBM.MarcaProcessoDigitado;
begin
   dmHonorarios.MarcaProcessoDigitado(idescritorio,
                                         idplanilha,
                                         anomesreferencia,
                                         sequencia,
                                         linhaplanilha);
end;

function TPlanilhaJBM.NomeEnvolvidoOk: boolean;
var
   autor : string;
   partecontraria : string;
begin
   autor := dmgcpj_compartilhado.ObtemNomeEnvolvido(gcpj);
   partecontraria := dtsAtos.FieldByName('partecontraria').AsString;

   RemoveEsteCaracter('.', autor);
   RemoveEsteCaracter('.', partecontraria);

   RemoveEsteCaracter(' ', autor);
   RemoveEsteCaracter(' ', partecontraria);

   autor := XPTrimAcentuacao(autor);
   partecontraria := XPTrimAcentuacao(partecontraria);

   result := (autor = partecontraria);
end;

function TPlanilhaJBM.NotaJaFoiDigitada: boolean;
begin
   result := dmHonorarios.NotaJaFoiDigitada(dtsAtos.FieldByName('numeronota').AsInteger);
end;


function TPlanilhaJBM.ObtemDataDigitar(formato: integer): string;
var
   dataDigitar: TdateTime;
begin
   dataDigitar := dmHonorarios.ObtemDataDigitar(anomesreferencia);
   if dataDigitar = 0 then
      result := ''
   else
   begin
      case formato of
         0 : result := FormatDateTime('ddmmyyyy', dataDigitar);
         1 : result := FormatDateTime('dd/mm/yyyy', dataDigitar);
         2 : Result := FormatDateTime('dd.mm.yyyy', dataDigitar);
      end;
   end;
end;

procedure TPlanilhaJBM.ObtemInconsistenciasImportacao;
begin
   dmHonorarios.ObtemInconsistenciasImportacao(idescritorio,
                                               idplanilha,
                                               anomesreferencia,
                                               sequencia,
                                               fdtsAtos);
end;

procedure TPlanilhaJBM.ObtemInconsistenciasPosImportacao;
begin
   dmHonorarios.ObtemInconsistenciasPosImportacao(idescritorio,
                                                  idplanilha,
                                                  anomesreferencia,
                                                  sequencia,
                                                  fdtsAtos);
end;

procedure TPlanilhaJBM.ObtemNotasFinalizadas(numNota: string);
begin
   dmHonorarios.ObtemNotasFinalizadas(idescritorio,
                                      idplanilha,
                                      anomesreferencia,
                                      sequencia,
                                      fdtsAtos,
                                      numNota);
end;

procedure TPlanilhaJBM.ObtemNumeroNotaEmpresa;
begin
   numerodanota := dmHonorarios.ObtemNumeroNotaEmpresa(idescritorio,
                                                       idplanilha,
                                                       empresaLigada,
                                                       StrToInt(LeftStr(anomesreferencia,4)),
                                                       anomesreferencia);
   totalAtos := dmHonorarios.dts.FieldByname('totaldeatos').AsInteger;
   valorTotalNota := dmHonorarios.dts.FieldByname('valortotal').AsFloat;
end;

procedure TPlanilhaJBM.ObtemPlanilhasCadastradas(const order: integer=0);
begin
   dmHonorarios.ObtemPlanilhasCadastradas(fdtsPlanilha, order);
end;

procedure TPlanilhaJBM.ObtemPlanilhasDigitar;
begin
   dmHonorarios.ObtemPlanilhasDigitar(fdtsPlanilha);
end;

procedure TPlanilhaJBM.ObtemPlanilhasValidar;
begin
   dmHonorarios.ObtemPlanilhasValidar(fdtsPlanilha);
end;

procedure TPlanilhaJBM.ObtemProcessosCruzarGcpj(tipo: integer);
begin
   dmHonorarios.ObtemProcessosCruzarGcpj(tipo,
                                         idEscritorio,
                                         idplanilha,
                                         sequencia,
                                         anomesreferencia,
                                         fdtsAtos);
end;

function TPlanilhaJBM.ObtemProcessosDigitar(numnota: string; const alteracao: boolean=false):integer;
begin
   result := -1;

   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select empresa_ligada_agrupar, anodocumento, sequencia from  tbcontrolenotas ' +
                                      'where anomesreferencia = ''' + anomesreferencia + ''' and idescritorio = ' + inttostr(idescritorio) + ' and ' +
                                      'idplanilha = ' + IntToStr(idplanilha);
      if not alteracao then
         dmHonorarios.dts.CommandText := dmHonorarios.dts.CommandText + ' and fgstatus = 0';

      if numnota <> '' then
         dmHonorarios.dts.CommandText := dmHonorarios.dts.CommandText + ' and sequencia = ' + IntToStr(StrToInt(RightStr(numnota,5)));
      dmHonorarios.dts.CommandText := dmHonorarios.dts.CommandText + ' order by totaldeatos';
      dmHonorarios.dts.Open;
      if dmHonorarios.dts.eof then
      begin
         result := 0;
         exit;
      end;

      dtsAtos.Close;
      if numnota = '' then
      begin
         dtsAtos.CommandText := 'select * from tbplanilhasatos where fgcruzadogcpj >= 5 and fgcruzadogcpj <> 9 and fgliberadodigitacao=1 and fgdigitado=0 and ' +
                                ' idplanilha = ' + IntToStr(idplanilha) + ' ' +
                                'order by numeronota, gcpj '
      end
      else
      if Not alteracao then
         dtsAtos.CommandText := 'select * from tbplanilhasatos where fgcruzadogcpj >= 5 and fgcruzadogcpj <> 9 and fgliberadodigitacao=1 and fgdigitado=0 and ' +
                                'numeronota = ' + dmHonorarios.dts.FieldByName('anodocumento').AsString + LeftZeroStr('%05d', [dmHonorarios.dts.FieldByName('sequencia').Asinteger]) + ' ' +
                                'order by numeronota, gcpj '
      else
         dtsAtos.CommandText := 'select * from tbplanilhasatos where fgcruzadogcpj >= 5 and fgcruzadogcpj <> 9 and fgliberadodigitacao=1 and fgdigitado=1 and ' +
                                'numeronota = ' + dmHonorarios.dts.FieldByName('anodocumento').AsString + LeftZeroStr('%05d', [dmHonorarios.dts.FieldByName('sequencia').Asinteger]) + ' ' +
                                'order by numeronota, gcpj ';
      dtsAtos.Open;
      if not dtsAtos.Eof then
         result := 1
      else
         result := 0;
   except
      on merr : Exception do
         ShowMessage(merr.Message);
   end;
end;

procedure TPlanilhaJBM.ObtemReclamadasDoProcesso(ptipoProcesso: string);
begin
   dmHonorarios.ObtemReclamadasDoProcesso(idescritorio,
                                          idplanilha,
                                          anomesreferencia,
                                          sequencia,
                                          linhaplanilha,
                                          fdtsReclamadas,
                                          ptipoProcesso);

end;

function TPlanilhaJBM.ObtemTotalAtosDigitados: integer;
begin
   result := dmHonorarios.ObtemTotalAtosDigitados(IntToStr(numerodanota));
end;

function TPlanilhaJBM.ObtemCodOrgProcesso: integer;
begin
   result := dmgcpj_compartilhado.ObtemCodOrgProcesso(gcpj);
end;

procedure TPlanilhaJBM.ObtemValoresRecalculadosSistema;
begin
   dmHonorarios.ObtemValoresRecalculadosSistema(idescritorio,
                                                idplanilha,
                                                anomesreferencia,
                                                sequencia,
                                                fdtsAtos);
end;

function TPlanilhaJBM.ObtemValorTotalDaNota: double;
begin
   valorTotalNota := dmHonorarios.ObtemValorTotalDaNota(StrToInt(LeftStr(dtsAtos.FieldbyName('numeronota').AsString, 4)),
                                                        StrToInt(RightStr(dtsAtos.FieldbyName('numeronota').AsString, 5)),
                                                        ftotalatos);
   result := valorTotalNota;
end;

function TPlanilhaJBM.ProcessoExisteNoGcpj: integer;
begin
   ExibeMensagem('Verificando existencia na base de processos ativos');
   result := dmgcpj_compartilhado.ProcessoExisteGcpj(gcpj);
   if result = 0 then
   begin
      ExibeMensagem('Verificando existencia na base de processos recuperados');
      result := dmgcpj_recuperados.ProcessoExisteGcpj(gcpj);
   end;
end;

procedure TPlanilhaJBM.RemoveNotaProcesso;
begin
   dmHonorarios.RemoveNotaProcesso(idescritorio,
                                   idplanilha,
                                   anomesreferencia,
                                   sequencia,
                                   linhaplanilha);
end;

procedure TPlanilhaJBM.RemovePlanilhas(idPlanilha, idEscritorio: integer;
  anomesreferencia: string);
begin
   dmHonorarios.RemovePlanilha(idPlanilha, idEscritorio, anomesreferencia);
end;

procedure TPlanilhaJBM.SalvaAlterado(tipoato: string);
begin
   dmHonorarios.adoCmd.parameters.Clear;
   dmHonorarios.adoCmd.CommandText := 'insert into tbalterados(numeronota, gcpj, tipoandamento) ' +
                                      'values(' + dtsAtos.FieldByName('numeronota').AsString + ', ' +
                                       dtsAtos.FieldByName('gcpj').AsString + ', ''' +
                                       tipoato + ''')';
   dmHonorarios.adoCmd.Execute;
end;

procedure TPlanilhaJBM.Setanomescompetencia(const Value: string);
begin
  Fanomescompetencia := Value;
end;

procedure TPlanilhaJBM.Setanomesreferencia(const Value: string);
begin
  Fanomesreferencia := Value;
end;

procedure TPlanilhaJBM.Setbaixados(const Value: TProcessoBaixa);
begin
  Fbaixados := Value;
end;

procedure TPlanilhaJBM.SetCliente(const Value: string);
begin
  FCliente := Value;
end;

procedure TPlanilhaJBM.SetcnpjEscritorio(const Value: string);
begin
  FcnpjEscritorio := Value;
end;

procedure TPlanilhaJBM.SetcodGcpjEscritorio(const Value: integer);
begin
  FcodGcpjEscritorio := Value;
end;

procedure TPlanilhaJBM.SetComarca(const Value: string);
begin
  FComarca := Value;
end;

procedure TPlanilhaJBM.Setdatadoato(const Value: TDateTime);
begin
  Fdatadoato := Value;
end;

procedure TPlanilhaJBM.Setdivisaodafilial(const Value: string);
begin
  Fdivisaodafilial := Value;
end;

procedure TPlanilhaJBM.SetdsNotasDigitando(const Value: TDataSource);
begin
  FdsNotasDigitando := Value;
end;

procedure TPlanilhaJBM.SetdsNotasPendentes(const Value: TDataSource);
begin
  FdsNotasPendentes := Value;
end;

procedure TPlanilhaJBM.SetdsPlan(const Value: TDataSource);
begin
  FdsPlan := Value;
end;

procedure TPlanilhaJBM.SetdtsAtos(const Value: TAdoDataset);
begin
  FdtsAtos := Value;
end;

procedure TPlanilhaJBM.SetdtsNotasDigitando(const Value: TAdoDataSet);
begin
  FdtsNotasDigitando := Value;
end;

procedure TPlanilhaJBM.SetdtsNotasPendentes(const Value: TAdoDataset);
begin
  FdtsNotasPendentes := Value;
end;

procedure TPlanilhaJBM.SetdtsPlanilha(const Value: TADODataSet);
begin
  FdtsPlanilha := Value;
end;

procedure TPlanilhaJBM.SetdtsReclamadas(const Value: TAdoDataset);
begin
  FdtsReclamadas := Value;
end;

procedure TPlanilhaJBM.SetehJec(const Value: boolean);
begin
  FehJec := Value;
end;

procedure TPlanilhaJBM.SetempresaLigada(const Value: integer);
begin
  FempresaLigada := Value;
end;

procedure TPlanilhaJBM.Setgcpj(const Value: string);
begin
  Fgcpj := Value;
  RemoveEsteCaracter('''', fgcpj);
end;

procedure TPlanilhaJBM.SetNome_reu(const Value: string);
begin
  FNome_reu := Value;
end;

procedure TPlanilhaJBM.SetidEscritorio(const Value: integer);
begin
  FidEscritorio := Value;
end;

procedure TPlanilhaJBM.Setidplanilha(const Value: integer);
begin
  Fidplanilha := Value;
end;

procedure TPlanilhaJBM.Setidprocessounico(const Value: string);
begin
  Fidprocessounico := Value;
end;

procedure TPlanilhaJBM.SetlinhaPlanilha(const Value: integer);
begin
  FlinhaPlanilha := Value;
end;

procedure TPlanilhaJBM.Setnomedigitar(const Value: string);
begin
  Fnomedigitar := Value;
end;

procedure TPlanilhaJBM.Setnomeescritorio(const Value: string);
begin
  Fnomeescritorio := UpperCase(Value);
end;

procedure TPlanilhaJBM.Setnumerodanota(const Value: integer);
begin
  Fnumerodanota := Value;
end;

procedure TPlanilhaJBM.Setpartecontraria(const Value: string);
begin
  Fpartecontraria := Value;
end;

procedure TPlanilhaJBM.SetPorcentagem(const Value: string);
begin
  FPorcentagem := Value;
end;

procedure TPlanilhaJBM.Setprocesso(const Value: string);
begin
  Fprocesso := Value;
end;

procedure TPlanilhaJBM.Setsequencia(const Value: integer);
begin
  Fsequencia := Value;
end;

procedure TPlanilhaJBM.Settipoandamento(const Value: string);
begin
  if (UpperCase(value) = 'CONTESTAÇÃO') or (UpperCase(value) = 'CONTESTACAO') OR (UpperCase(value) = 'CONTESTACÃO') or
     (value = 'CONTESTAçãO') then
    Ftipoandamento := 'TUTELA'
  else
    Ftipoandamento := Value;
end;

procedure TPlanilhaJBM.SettipoProcesso(const Value: string);
begin
  FtipoProcesso := Value;
end;

procedure TPlanilhaJBM.SettotalAtos(const Value: integer);
begin
  FtotalAtos := Value;
end;

procedure TPlanilhaJBM.SetUF(const Value: string);
begin
  FUF := Value;
end;

procedure TPlanilhaJBM.SetvalBeneficioEconomico(const Value: string);
begin
  FvalBeneficioEconomico := Value;
end;

procedure TPlanilhaJBM.SetValor(const Value: double);
begin
  FValor := Value;
end;

procedure TPlanilhaJBM.Setvalordesembolsado(const Value: string);
begin
  Fvalordesembolsado := Value;
end;

procedure TPlanilhaJBM.Setvalorpedido(const Value: double);
begin
  Fvalorpedido := Value;
end;


procedure TPlanilhaJBM.SetvalorTotalNota(const Value: double);
begin
  FvalorTotalNota := Value;
end;

procedure TPlanilhaJBM.SetVara(const Value: string);
begin
  FVara := Value;
end;

function TPlanilhaJBM.SomaTotalDigitadoDaNota: double;
begin
   result := dmHonorarios.SomaTotalDigitadoDaNota(dtsAtos.FieldByName('numeronota').AsString);
end;

function TPlanilhaJBM.TemProcessoDuplicado: boolean;
begin
   dmHonorarios.ObtemProcessosIguais(idEscritorio, idplanilha, sequencia, gcpj, tipoandamento, anomesreferencia, datadoato, Valor, fgDrcContrarias);




   if (tipoandamento = 'PARCELA') and (tipoProcesso <> 'CI') then
   begin
      while Not dmHonorarios.adoDts.Eof do
      begin
         if  linhaPlanilha <> dmHonorarios.adoDts.FieldByName('linhaplanilha').AsInteger then
         begin
            if FormatDateTime('mm/yyyy', datadoato) = FormatDateTime('mm/yyyy', dmHonorarios.adoDts.FieldByName('datadoato').AsDateTime) then
            begin
               result := true;
               exit;
            end;
         end;
         dmHonorarios.adoDts.Next;
      end;
      Result := false;
   end
   else
      result := (dmHonorarios.adodts.RecordCount > 1);
end;

procedure TPlanilhaJBM.ValidaValor;
begin

end;

procedure TPlanilhaJBM.VerificaDataDaBaixa(gcpj: string);
begin
   baixados := TProcessoBaixa.Create(gcpj);
   try
      dmgcpj_compartilhado.ObtemDadosDaBaixa(gcpj, baixados.FmotivoBaixa, baixados.fdataDaBaixa, baixados.fvalorbaixa);
      if baixados.motivoBaixa <> '' then
         dmHonorarios.GravaDataDaBaixa(idEscritorio,
                                       idPlanilha,
                                       anomesreferencia,
                                       sequencia,
                                       linhaPlanilha,
                                       baixados.dataDaBaixa,
                                       baixados.motivoBaixa,
                                       baixados.valorbaixa);
   finally
      baixados.Free;
   end;

end;

function TPlanilhaJBM.ObtemUfProcesso(orgaoJulgador: integer): string;
begin
   result := dmgcpj_base_iv.ObtemUfProcesso(orgaoJulgador);
end;

(**
procedure TPlanilhaJBM.NotificaNotaFinalizada;
var
   outlookObj : TOutlookObj;
begin
   exit;
   try
      if Not dmhonorarios.NotificouNotaFinalizada(numerodanota) then
      begin
         outlookObj := TOutlookObj.Create;
         try
            outlookobj.emailPara.Add('4040.nilce@bradesco.com.br');
            outlookobj.emailPara.Add('4040.amandat@bradesco.com.br');
            outlookobj.emailPara.Add('maria@presenta.com.br');
            outlookObj.Subject := 'Nota: ' + IntToStr(numerodanota) + ' - Finalizada Digitação';
            outlookObj.Mensagem := 'Finalizada a digitação da nota: ' + IntToStr(numerodanota) + ' do escritório: ' + nomeescritorio + ', referência: ' + anomesreferencia;
            outlookObj.SendMail;
         finally
            outlookObj.emailPara.Free;
            outlookObj.Free;
         end;
         dmHonorarios.MarcaNotaFinalizadaNotificada(numerodanota);
      end;
   except
   end;
end;
   **)

(**
procedure TPlanilhaJBM.NotificaNotaIniciada;
var
   outlookObj : TOutlookObj;
begin
   exit;
   try
      if Not dmhonorarios.NotificouNotaIniciada(numerodanota) then
      begin
         outlookObj := TOutlookObj.Create;
         try
//            outlookobj.emailPara.Add('4040.nilce@bradesco.com.br');
//            outlookobj.emailPara.Add('4040.amandat@bradesco.com.br');
            outlookobj.emailPara.Add('maria@presenta.com.br');
            outlookObj.Subject := 'Nota: ' + IntToStr(numerodanota) + ' - Inicio da  Digitação';
            outlookObj.Mensagem := 'Iniciou a digitação da nota: ' + IntToStr(numerodanota) + ' do escritório: ' + nomeescritorio + ', referência: ' + anomesreferencia;
            outlookObj.SendMail;
         finally
            outlookObj.emailPara.Free;
            outlookObj.Free;
         end;
         dmHonorarios.MarcaNotaIniciadaNotificada(numerodanota);
      end;
   except
   end;
end;
**)

procedure TPlanilhaJBM.CriaIndiceBaseV;
begin
   dmgcpcj_base_v.CreateIndex;
end;

(**
procedure TPlanilhaJBM.NotificaErro(mensagem: string);
var
   outlookObj : TOutlookObj;
begin
   exit;
   try
      outlookObj := TOutlookObj.Create;
      try
         outlookobj.emailPara.Add('maria@presenta.com.br');
         if Uppercase(usuario) = 'SAMARA' then
            outlookobj.emailPara.Add('4040.samara@bradesco.com.br')
         else
         if Uppercase(usuario) = 'NILCE' then
            outlookobj.emailPara.Add('4040.nilce@bradesco.com.br')
         else
            outlookobj.emailPara.Add('4040.amandat@bradesco.com.br');;
         outlookObj.Subject := 'Erro no processamento da nota: ' + IntToStr(numerodanota) + ' na máquina: ' + usuario;
         outlookObj.Mensagem := mensagem;
         outlookObj.SendMail;
      finally
         outlookObj.emailPara.Free;
         outlookObj.Free;
      end;
      dmHonorarios.MarcaNotaFinalizadaNotificada(numerodanota);
   except
   end;
end;
   **)
procedure TPlanilhaJBM.SetdataIbi(const Value: Tdatetime);
begin
  FdataIbi := Value;
end;

procedure TPlanilhaJBM.Setfgibiativo(const Value: integer);
begin
  Ffgibiativo := Value;
end;

function TPlanilhaJBM.ObtemDirLck: string;
begin
   result := dmHonorarios.dirLck;
end;

procedure TPlanilhaJBM.Setusuario(const Value: string);
begin
  Fusuario := Value;
end;

procedure TPlanilhaJBM.SetdrcContrarias(const Value: boolean);
begin
  FdrcContrarias := Value;
end;

function TPlanilhaJBM.EhDrcContraria: boolean;
begin
   result := (dtsatos.FieldByName('codtipoacao').AsInteger = 8911) or
             (dtsatos.FieldByName('codtipoacao').AsInteger = 8912);
//             (dtsatos.FieldByName('codtipoacao').AsInteger = 8913);
//             (dtsatos.FieldByName('codtipoacao').AsInteger = 8914);
end;

function TPlanilhaJBM.EhSubtipoTarifa(lista: TStringList): boolean;
var
   i : integer;
   strValues : TStringList;
begin
   result := false;
   for i := 0 to lista.Count - 1 do
   begin
      strValues := TStringList.Create;
      try
         strToList(lista.Strings[i], ';', strValues);
         if ((dtsatos.FieldByName('codtipoacao').AsInteger = StrToInt(strValues.Strings[0])) and
             (dtsatos.FieldByName('codsubtipoacao').AsInteger = StrToInt(strValues.Strings[1]))) then
         begin
            result := true;
            exit;
         end;
      finally
         strValues.Free;
      end;
   end;
end;

function TPlanilhaJBM.EhEscritorioTarifa(lista: TStringList): boolean;
var
   i : integer;
   strValues : TStringList;
begin
   result := false;
   for i := 0 to lista.Count - 1 do
   begin
      strValues := TStringList.Create;
      try
         strToList(lista.Strings[i], ';', strValues);
         if cnpjEscritorio <> strValues.Strings[0] then
            continue;

         //verifica as reclamadas do processo são da empresa correta
(**         dmHonorarios.ObtemReclamadasDoProcesso(idEscritorio, idplanilha, anomesreferencia, sequencia, linhaPlanilha, fdtsreclamadas);
         while Not dtsReclamadas.Eof do
         begin
            Application.ProcessMessages;
            if dtsReclamadas.FieldByName('codigoreclamada').IsNull then
            begin
               dtsReclamadas.Next;
               continue;
            end;

            if dtsReclamadas.FieldByName('codigoreclamada').AsInteger = StrToInt(strValues.Strings[0]) then
            begin
               result := true;
               exit;
            end;
            dtsReclamadas.Next;
         end;

         if dtsatos.FieldByName('empresaligadaagrupar').AsInteger <> StrToInt(strValues.strings[0]) then
            continue;                            *)
         result := true;
         exit;
      finally
         strValues.Free;
      end;
   end;
end;


procedure TPlanilhaJBM.MarcaPlanilhaNaoFinalizada;
begin
   dmHonorarios.MarcaPlanilhaNaoFinalizada(idEscritorio, idplanilha, anomesreferencia, sequencia);
end;

function TPlanilhaJBM.EhDrcAtiva: boolean;
begin
   result := (dtsatos.FieldByName('codtipoacao').AsInteger = 8914) or
             (dtsatos.FieldByName('codtipoacao').AsInteger = 8913) or
             (dtsatos.FieldByName('codtipoacao').AsInteger = 8915);
end;

function TPlanilhaJBM.ProcessoEhJuizado: boolean;
begin
   if dmgcpj_compartilhado.ObtemNumeroDoProcesso(gcpj) = '' then
      result := false
   else
      result := true;
end;

function TPlanilhaJBM.ObtemNomeOrgaoJulgador(
  orgaoJulgador: integer): string;
begin
   result := dmgcpj_base_iv.ObtemNomeOrgaoJulgador(orgaoJulgador);
end;

procedure TPlanilhaJBM.SetlstCnpjs(const Value: TStringList);
begin
  FlstCnpjs := Value;
end;

procedure TPlanilhaJBM.CarregaFiliais;
var
   i, j : integer;
   strValues : TStringList;
   arqCnpjs : TStringList;
begin
   arqCnpjs := TStringList.Create;
   try
      arqCnpjs.LoadFromFile(ExtractFilePath(Application.ExeName) + 'ESCRITORIOS_DISTRIBUICAO.TXT');
      lstCnpjs.Clear;
      for i := 0 to arqCnpjs.Count - 1 do
      begin
         strValues := TstringList.Create;
         try
            strtolist(arqCnpjs.Strings[i], ';', strValues);

            for j := 0 to (strvalues.count - 1) do
            begin
               if codGcpjEscritorio = StrToInt(strValues.Strings[j]) then
                  break;
            end;

            if j > (strValues.count - 1) then
               continue;

            //achou
            for j := 0 to (strValues.count - 1) do
               lstCnpjs.Add(strValues.Strings[j]);
         finally
            strValues.free;
         end;
      end;

      if lstCnpjs.Count = 0 then
         lstCnpjs.Add(IntToStr(codGcpjEscritorio))
   finally
      arqCnpjs.Free;
   end;
end;

function TPlanilhaJBM.ObtemOutrosPagamentosDoProcesso(gcpj, tipoato,
  escritorio, tipoprocesso, codescritorio: string;
  lstCnpjs: TStringList): integer;
begin
   result := dmHonorarios.ObtemOutrosPagamentosDoProcesso(idEscritorio,
                                                          idplanilha,
                                                          gcpj,
                                                          dtsatos.FieldByName('tipoandamento').AsString,
                                                          dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                          lstCnpjs,
                                                          dataPlanilha);
   if result <> 1 then
   begin
      result := 0; //tipo de ato não reconhecido
      exit;
   end;

   if dmHonorarios.dts.RecordCount > 0 then
   begin
      if (tipoProcesso = 'TR') OR (tipoProcesso = 'TO') OR (tipoProcesso = 'TA') then
      begin
         if dmgcpcj_base_V.dts.RecordCount >= 2 then
         begin
            result := 2; //já foi pago
            exit;
         end;
      end
      else
      begin
         if Pos('AVULSO', dtsAtos.FieldByName('tipoandamento').AsString) <> 0 then
         begin
            if dmgcpcj_base_V.dts.RecordCount >= 2 then
            begin
               result := 2; //já foi pago
               exit;
            end;
         end
         else
         begin
            result := 2; //já foi pago
            exit;
         end;
      end;
   end;
end;

function TPlanilhaJBM.LaudoPericialJaFoiPago: boolean;
var
   mesPagto, mesAtual : integer;
begin
   result := false;

   dmgcpcj_base_xi.ObtemPagamentosLaudoPericial(gcpj, IntToStr(codGcpjEscritorio));
   while not dmgcpcj_base_xi.dtsNfiscdt.Eof do
   begin
      if StrToFloat(FormatFloat('0.00', dtsatos.FieldByName('valor').AsFloat)) = StrToFloat(FormatFloat('0.00', dmgcpcj_base_xi.dtsNfiscdt.FieldByName('VlrDespesa').AsFloat)) then
      begin
         //já foi pago valor igual
         //so pode pagar se o ultimo pagamento for superior a 2 meses
         mesPagto := MonthOf(dmgcpcj_base_xi.dtsNfiscdt.FieldByName('DatMovto').AsDateTime);
         mesAtual := MonthOf(Date);

         if (mesAtual - mesPagto) <= 2 then
         begin
            result := true;
            exit;
         end;
      end;
      dmgcpcj_base_xi.dtsNfiscdt.Next;
   end;
end;

procedure TPlanilhaJBM.CriaIndiceBaseVII;
begin
   dmgcpj_base_vii.CreateIndex;
end;

procedure TPlanilhaJBM.ObtemAtosPendentes;
begin
   dmHonorarios.ObtemAtosPendentes(idplanilha);
end;

procedure TPlanilhaJBM.SetvalorBase(const Value: double);
begin
  FvalorBase := Value;
end;

procedure TPlanilhaJBM.SetdataPlanilha(const Value: TDateTime);
begin
  FdataPlanilha := Value;
end;

procedure TPlanilhaJBM.CriaIndiceBaseVIII;
begin
   dmgcpj_base_viii.CreateIndex;
end;

function TPlanilhaJBM.TemConsolidacaoPagaNoSistema: boolean;
begin
   result := dmHonorarios.TemConsolidacaoPagaNoSistema(gcpj);
end;

function TPlanilhaJBM.ObtemValorSemReajuste(andamento: string): double;
begin
   result := dmHonorarios.ObtemValorSemReajuste(cnpjEscritorio, andamento, dtsatos.FieldByName('datadoato').AsDateTime);
end;

function TPlanilhaJBM.ProcessoDistribuidoParaAdvogadoInterno: boolean;
var
   codFuncional : integer;
begin
   if dtsAtos.FieldByName('gcpj').AsString = '' then
   begin
      result := false;
      exit;
   end;
   
   //obtem o código do advogado interno no processo
   if dtsAtos.FieldByName('gcpj').AsInteger <= 900000000 then
      codFuncional := dmgcpcj_base_I.ProcessoDistribuidoAdvogadoInterno_SGCP(dtsAtos.FieldByName('gcpj').AsString)
   else
      codFuncional := dmgcpcj_base_I.ProcessoDistribuidoAdvogadoInterno_GCPJ(dtsAtos.FieldByName('gcpj').AsString);
      
   if codFuncional = 0 then
   begin
      //não distribuido para advogado interno
      result := false;
      exit;
   end;

   dmHonorarios.ObtemAdvogadosCadastrados(codFuncional);
   result := (dmHonorarios.dtsAdvogados.RecordCount > 0);
end;

function TPlanilhaJBM.ConfiguracaoSemReajusteEstaOk: boolean;
begin
   result := dmHonorarios.ConfiguracaoSemReajusteEstaOK;
end;

procedure TPlanilhaJBM.CriaIndicesDigitaHonorarios;
begin
   dmHonorarios.CreateIndex;
end;

procedure TPlanilhaJBM.ResetFlagsDaPlanilha;
begin
   dmHonorarios.ResetFlagsDaPlanilha(dtsPlanilha.FieldByName('idplanilha').AsInteger);
end;

procedure TPlanilhaJBM.ObtemPlanilhasEnviarGpj(const filtro: integer);
begin
   dmHonorarios.ObtemPlanilhasEnviarGcpj(fdtsPlanilha, filtro);
end;

function TPlanilhaJBM.ObtemTipoAndamento(andamento, tipoProcesso,
  motivoBaixa: string): string;
begin
   result := '';
   if (UpperCase(andamento) = 'RECURSO EM FASE DE EXECUÇÃO') or
      (UpperCase(andamento) = 'RECURSO EM FASE DE EXECUCAO') OR
      (UpperCase(andamento) = 'RECURSO NA FASE DE EXECUÇÃO') or
      (UpperCase(andamento) = 'RECURSO NA FASE DE EXECUCAO') then
      result := 'RECURSO NA FASE DE EXECUCAO'
   else
   if (Pos('RECURSO', UpperCase(andamento)) <> 0) OR
      (Pos('AGRAVO FAT', UpperCase(andamento)) <> 0) then
      result := 'RECURSO'
   else
   if Pos('TUTELA', UpperCase(andamento)) <> 0 then
      result := 'CONTESTACAO'
   else
   if UpperCase(andamento) = 'PREPOSTO' then
      result := 'PREPOSTO'
   else
   if (Pos('AVULSO TRAB', UpperCase(andamento)) <> 0) or
      (((tipoprocesso = 'TR') OR (tipoprocesso = 'TO') OR (tipoprocesso = 'TA')) and
      (UpperCase(andamento) = 'AUDIÊNCIA TRABALHISTA')or (UpperCase(andamento) = 'AUDIENCIA TRABALHISTA') or
      (UpperCase(andamento) = 'AUDIêNCIA AVULSA') or (UpperCase(andamento) = 'AUDIÊNCIA AVULSA') or
      (UpperCase(andamento) = 'INSTRUCAO') or (UpperCase(andamento) = 'AUDIENCIA AVULSA')) then
      result := 'INSTRUCAO'
   else
   if (Pos('ÊXITO', UpperCase(andamento)) <> 0) or (Pos('EXITO', UpperCase(andamento)) <> 0) then
      result := 'HONORARIOS DE EXITO'
   else
   if (UpperCase(andamento) = 'AVULSO') or
      (UpperCase(andamento) = 'AVULSO CÍVEL (AUDITORIA)') or
      (UpperCase(andamento) = 'AVULSO CIVEL (AUDITORIA)') or
      (UpperCase(andamento) = 'AVULSO CIVEL') or
      (UpperCase(andamento) = 'AVULSO CÍVEL') then
      result := 'ASSESSORIA JURIDICA'
   else
   if (UpperCase(andamento) = 'AUDIêNCIA AVULSA') or
      (UpperCase(andamento) = 'AUDIENCIA AVULSA') then
      result := 'ACOMPANHAMENTO'
   else
   if (UpperCase(andamento) = 'HONORARIOS FINAIS') or
      (UpperCase(andamento) = 'HONORÁRIOS FINAIS') then
   begin
      if (motivobaixa = 'ACORDO COM CUSTOS') or
         (motivobaixa = 'IMPROCEDENCIA') or
         (motivoBaixa = 'EXTINTO COM JULGAMENTO DE MERITO') then
         result := 'HONORARIOS DE EXITO'
      else
         result := 'HONORARIOS FINAIS';
   end
   else
   if (UpperCase(andamento) = 'HONORARIOS INICIAIS') or
      (UpperCase(andamento) = 'HONORÁRIOS INICIAIS') then
      result := 'HONORARIOS INICIAIS'
   else
   if Pos('AJUIZAMENTO', UpperCase(andamento)) <> 0 then
      result := 'AJUIZAMENTO'
   else
   if UpperCase(andamento) = 'PARCELA' then
      result := 'ACOMPANHAMENTO'
   else
      result := UpperCase(andamento);
end;

function TPlanilhaJBM.ObtemDependenciaDigitar: string;
   var
      i,p : integer;
      arquivo : TStringlist;
      estagio : integer;
      codigo, nome : string;
      valstr : string;
      dependencia : string;
      primeira : boolean;
      sequencia : integer;
   begin
      result := '';
      codigo := '';
      nome:= '';
      ObtemReclamadasDoProcesso('');
      sequencia := 1;

      while Not dtsReclamadas.Eof do
      begin
         if dtsReclamadas.FieldByName('sequenciagcpj').AsInteger = sequencia then //guarda a primeira por seguranca
         begin
            if dtsReclamadas.FieldByName('codigoreclamada').AsInteger <> 0 then
            begin
               codigo := dtsReclamadas.FieldByName('codigoreclamada').AsString;
               nome := dtsReclamadas.FieldByName('nomereclamada').AsString;
            end
            else
            begin
               Inc(sequencia);
               dtsReclamadas.Next;
               continue;
            end;
         end;

         //é bradesco?
         if dtsAtos.FieldByName('empresaligadaagrupar').AsInteger = 237 then //procura uma agencia
         begin
            if dtsReclamadas.FieldByName('codempresa').AsInteger = dtsAtos.FieldByName('empresaligadaagrupar').AsInteger then
            begin
               if dtsReclamadas.FieldByName('tiporeclamada').AsString = 'A' then //agencia
               begin
                  codigo := dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
            end;
         end
         else
         begin
            if (dtsReclamadas.FieldByName('codigoreclamada').AsInteger = 4001) and (dtsAtos.FieldByName('empresaligadaagrupar').AsInteger = 5172) then
            begin
               codigo := dtsReclamadas.FieldByName('codigoreclamada').AsString;
               nome := dtsReclamadas.FieldByName('nomereclamada').AsString;
               break;
            end
            else
            if (dtsReclamadas.FieldByName('codigoreclamada').AsInteger = dtsAtos.FieldByName('empresaligadaagrupar').AsInteger) or
               (dtsReclamadas.FieldByName('codigoreclamada').AsInteger = dtsAtos.FieldByName('codempresaligada').AsInteger) then
            begin
               if dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A' then //não pode ser agencia
               begin
                  codigo := dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
            end;
         end;
         dtsReclamadas.Next;
      end;

      if (codigo = '0') or (codigo = '') then
         primeira := true
      else
         primeira := false;


      if not primeira then
      begin
         if StrToInt(codigo) = 5404 then
         begin
            //tem a dependência 4027?
            dtsReclamadas.First;

            while Not dtsReclamadas.Eof do
            begin
               Application.ProcessMessages;
               if dtsReclamadas.FieldByName('codigoreclamada').AsInteger = 4027 then
               begin
                  codigo := dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
               dtsReclamadas.Next;
            end;
         end;
      end;

      Result := codigo;

      (***
      arquivo := TStringList.Create;
      try
         arquivo.LoadFromFile(fname);

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
                        ShowMessage('Erro valor');

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
      end;             **)
end;

procedure TPlanilhaJBM.SetfgDrcContrarias(const Value: integer);
begin
  FfgDrcContrarias := Value;
end;

procedure TPlanilhaJBM.RemoveOcorrenciasDaPlanilha;
begin
   dmHonorarios.RemoveOcorrenciasDaPlanilha(idplanilha);
end;

procedure TPlanilhaJBM.SetExibeMensagem(const Value: TProcExibeMensagem);
begin
  FExibeMensagem := Value;
end;

procedure TPlanilhaJBM.CriaIndiceBaseBaixados;
begin
   dmgcpj_baixados.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseIX;
begin
   dmgcpj_base_IX.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseX;
begin
   dmgcpj_base_X.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseRecuperados;
begin
   dmgcpj_recuperados.CreateIndex;
end;

procedure TPlanilhaJBM.CriaIndiceBaseXI;
begin
   dmgcpcj_base_XI.CreateIndex;
end;


procedure TPlanilhaJBM.CriaIndiceBaseTrabalhistas;
begin
   dmgcpj_trabalhistas.CreateIndex;
end;

procedure TPlanilhaJBM.CadastraDeParaEmpresasGrupo;
var
   str, arquivo : TStringList;
   i : integer;
begin

   dmHonorarios.ReindexTabelaDePara(0);
   dmHonorarios.LimpaTabelaDePara;

   ExibeMensagem('Inserindo empresas DE/PARA');

   arquivo := TStringList.Create;
   try
      arquivo.LoadFromFile(ExtractFilePath(Application.ExeName) + 'de_para.txt');

      for i:=0 to arquivo.count - 1 do
      begin
         str := TStringList.Create;
         try
            StrToList(arquivo.strings[i], ';', str, true);

            ExibeMensagem('Incluindo: ' + str.Strings[0] + '/' + str.Strings[1]);
            dmHonorarios.CadastraDeParaEmpresaGrupo(StrToInt(str.Strings[0]),str.Strings[1],StrToInt(str.Strings[2]),str.Strings[3],str.Strings[4]);
         finally
            str.free;
         end;
      end;
   finally
      arquivo.free;
   end;
   dmHonorarios.ReindexTabelaDePara(1);
end;

{ TProcessoReclamdas }

constructor TProcessoReclamdas.Create(pidescritorio, pidplanilha: integer; panomesreferencia: string; psequencia, plinhaplanilha: integer; pgcpj: string; const PProcedure : TProcExibeMensagem = Nil);
begin
   inherited create;
   idEscritorio := pidescritorio;
   idplanilha := pidplanilha;
   anomesreferencia := panomesreferencia;
   sequencia := psequencia;
   linhaplanilha := plinhaplanilha;
   gcpj := pgcpj;
   ExibeMensagem := PProcedure;
end;



procedure TProcessoReclamdas.LimpaCampos;
begin
   nomeReclamada := '';
   codReclamada := 0;
   tipoReclamada:= '';
   codpessoaexterna := 0;
   empresa := 0;
   codexfuncionario := 0;
end;

function TProcessoReclamdas.ObtemEnvolvidosDoProcesso(fgDrcAtivas:integer; buscarAutor: boolean): integer;
begin

   if Assigned(ExibeMensagem) then
      ExibeMensagem('   ObtemEnvolvidosNoProcesso');
   //localiza os envolvidos - reus
   result := dmgcpj_base_IX.ObtemEnvolvidosNoProcesso(gcpj, fgDrcAtivas, buscarAutor);
   if result = 0 then
      exit;

   errorMessage := '';
   while Not dmgcpj_base_IX.dtsPesq.Eof do
   begin
      if Assigned(ExibeMensagem) then
         ExibeMensagem('   ObtemCodigoReclamada');
      LimpaCampos;
      nomeReclamada := dmgcpj_base_IX.dtsPesq.FieldByName('IENVDO_REFRD_PROCS').AsString;
      sequenciagcpj := dmgcpj_base_IX.dtsPesq.FieldByName('CSEQ_PROCS_REU').AsInteger;
      //obtem o tipo de dependencia
      result := dmgcpj_base_iv.ObtemCodigoReclamada(dmgcpj_base_IX.dtspesq.FieldByName('CENVDO_PROCS_REU').AsString,
                                                    fcodReclamada,
                                                    fempresa,
                                                    fcodpessoaexterna,
                                                    fcodexfuncionario,
                                                    buscarAutor,
                                                    gcpj);
      if result = 0 then
      begin
         errorMessage := 'Não encontrou o código da reclamada no GCPJ';
         result := -1;
         exit;
      end;
      if (codpessoaexterna = 0) and (codexfuncionario = 0) then
      begin
         if Assigned(ExibeMensagem) then
            ExibeMensagem('   ObtemTipoDependencia');

         tipoReclamada := dmgcpj_base_iv.ObtemTipoDependencia(codReclamada, empresa);
      end;

      if Assigned(ExibeMensagem) then
         ExibeMensagem('   CadastraReclamada');

      dmHonorarios.CadastraReclamada(idescritorio,
                                     idplanilha,
                                     anomesreferencia,
                                     sequencia,
                                     linhaplanilha,
                                     codReclamada,
                                     nomereclamada,
                                     tipoReclamada,
                                     sequenciagcpj,
                                     codpessoaexterna,
                                     empresa,
                                     codexfuncionario);


      dmgcpj_base_IX.dtsPesq.next;
   end;
end;


function TPlanilhaJBM.ObtemEnvolvidosDoProcesso_2(fgDrcAtivas:integer; buscarAutor: boolean): integer;
begin
   Result := 0;
   if Assigned(ExibeMensagem) then
      ExibeMensagem('   ObtemEnvolvidosNoProcesso');
   //localiza os envolvidos - reus
   result := dmgcpj_base_IX.ObtemEnvolvidosNoProcesso_2(gcpj, fgDrcAtivas, buscarAutor, Nome_reu);
//   if result = 0 then
//      exit;

end;


procedure TProcessoReclamdas.RemoveReclamadas;
begin
   dmHonorarios.RemoveReclamadas(idescritorio, idplanilha, anomesreferencia, sequencia, linhaplanilha);
end;

procedure TProcessoReclamdas.Setanomesreferencia(const Value: string);
begin
  Fanomesreferencia := Value;
end;

procedure TProcessoReclamdas.Setcodexfuncionario(const Value: integer);
begin
  Fcodexfuncionario := Value;
end;

procedure TProcessoReclamdas.Setcodpessoaexterna(const Value: integer);
begin
  Fcodpessoaexterna := Value;
end;

procedure TProcessoReclamdas.SetcodReclamada(const Value: integer);
begin
  FcodReclamada := Value;
end;

procedure TProcessoReclamdas.SetehSgcp(const Value: boolean);
begin
  FehSgcp := Value;
end;

procedure TProcessoReclamdas.Setempresa(const Value: integer);
begin
  Fempresa := Value;
end;

procedure TProcessoReclamdas.SeterrorMessage(const Value: string);
begin
  FerrorMessage := Value;
end;

procedure TProcessoReclamdas.SetExibeMensagem(
  const Value: TProcExibeMensagem);
begin
  FExibeMensagem := Value;
end;

procedure TProcessoReclamdas.Setgcpj(const Value: string);
begin
  Fgcpj := Value;
end;

procedure TProcessoReclamdas.SetNome_reu(const Value: string);
begin
  FNome_reu := Value;
end;

procedure TProcessoReclamdas.Setidescritorio(const Value: integer);
begin
  Fidescritorio := Value;
end;

procedure TProcessoReclamdas.Setidplanilha(const Value: integer);
begin
  Fidplanilha := Value;
end;

procedure TProcessoReclamdas.Setlinhaplanilha(const Value: integer);
begin
  Flinhaplanilha := Value;
end;

procedure TProcessoReclamdas.SetnomeReclamada(const Value: string);
begin
  FnomeReclamada := Value;
end;

procedure TProcessoReclamdas.Setsequencia(const Value: integer);
begin
  Fsequencia := Value;
end;

procedure TProcessoReclamdas.Setsequenciagcpj(const Value: integer);
begin
  Fsequenciagcpj := Value;
end;

procedure TProcessoReclamdas.SettipoReclamada(const Value: string);
begin
  FtipoReclamada := Value;
end;

{ TProcessoBaixa }

constructor TProcessoBaixa.Create(numeroGcpj: string);
begin
   inherited Create;
   motivoBaixa := '';
   dataDaBaixa := 0;
   gcpj := numeroGcpj;
end;

procedure TProcessoBaixa.SetdataDaBaixa(const Value: TdateTime);
begin
  FdataDaBaixa := Value;
end;

procedure TProcessoBaixa.Setgcpj(const Value: string);
begin
  Fgcpj := Value;
end;

procedure TProcessoBaixa.SetNome_reu(const Value: string);
begin
  FNome_reu := Value;
end;

procedure TProcessoBaixa.SetmotivoBaixa(const Value: string);
begin
  FmotivoBaixa := Value;
end;

procedure TProcessoBaixa.Setvalorbaixa(const Value: double);
begin
  Fvalorbaixa := Value;
end;

{ TIdObject }

procedure TIdObject.Setid(const Value: integer);
begin
  Fid := Value;
end;

{ TOutlookObj }
       (**
procedure TOutlookObj.ClickYes;
begin

end;

constructor TOutlookObj.Create;
begin
   inherited Create;
   FSystem := CreateOleObject('Outlook.Application');
   FNameSpace := Fsystem.GetNameSpace('MAPI');
   FMailFolder := FNameSpace.GetDefaultFolder(4);
   FEmailPara := TStringList.Create;
end;

procedure TOutlookObj.SendMail;
var
   mailItem : Olevariant;
   i : integer;
begin
   outlookThread := TOutlookThread.Create;
   try
      outlookThread.Resume;

      mailItem := Self.System.CreateItem(0);
      for i := 0 to Self.emailPara.Count - 1 do
         mailItem.Recipients.Add(Self.emailPara.Strings[i]);
      mailItem.Subject := Self.Subject;
      mailItem.body := Self.Mensagem;
      mailitem.send;
   finally
      outlookThread.pleaseClose := true;
   end;
end;

procedure TOutlookObj.SetemailPara(const Value: TStringList);
begin
  FemailPara := Value;
end;

procedure TOutlookObj.SetMailFolder(const Value: Olevariant);
begin
  FMailFolder := Value;
end;

procedure TOutlookObj.SetMensagem(const Value: string);
begin
  FMensagem := Value;
end;

procedure TOutlookObj.SetNameSpace(const Value: Olevariant);
begin
  FNameSpace := Value;
end;

procedure TOutlookObj.SetoutlookThread(const Value: TOutlookThread);
begin
  FoutlookThread := Value;
end;

procedure TOutlookObj.SetSubject(const Value: string);
begin
  FSubject := Value;
end;

procedure TOutlookObj.SetSystem(const Value: Olevariant);
begin
  FSystem := Value;
end;


procedure TOutlookObj.SetWintask(const Value: TWintask);
begin
  FWintask := Value;
end;

{ TOutlookThread }

constructor TOutlookThread.Create;
begin
   pleaseClose := false;
   Fwintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                               'Q:\Publico\Backup_sistemas\Compartilhado\Honorarios_Programa\presenta\rob\',
                               '');
   FreeOnTerminate := true;
   inherited Create(true);
end;

procedure TOutlookThread.Execute;
var
   valor : string;

begin
  inherited;

  Sleep(500);

  while true do
  begin
     if pleaseClose then
        break;

     if wintask.Verifica_Janela_PopUp('OUTLOOK.EXE|Static|Progr. tentando acessar inform. de endereço de email armazenados no Outlook',
                                       valor) = 1 then
     begin
        wintask.Click_Button('OUTLOOK.EXE|#32770|Microsoft Outlook', 'Permitir');
        sleep(100);
        continue;
     end;

     if wintask.Verifica_Janela_PopUp('OUTLOOK.EXE|Static|Programa tentando enviar email em seu nome.',
                                      valor) = 1 then
     begin
        wintask.Click_Button('OUTLOOK.EXE|#32770|Microsoft Outlook', 'Permitir');
        sleep(100);
        continue;
     end;

     Sleep(300);
  end;
end;

procedure TOutlookThread.SetpleaseClose(const Value: boolean);
begin
  FpleaseClose := Value;
end;

procedure TOutlookThread.Setwintask(const Value: TWintask);
begin
  Fwintask := Value;
end;
          **)
{ TClickOkThread }

constructor TClickOkThread.Create;
var
   ini : TiniFile;
   diretorio : string;
begin
   ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
   try
      diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_Programa\presenta');
   finally
      ini.Free;
   end;

   pleaseClose := false;
   Fwintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                               diretorio + '\rob\',
                               '');
   FreeOnTerminate := true;
   wintask.ie8 := true;
   inherited Create(true);
end;

procedure TClickOkThread.Execute;
var
   popUp, valor : string;
   i : integer;
begin
  inherited;
  Sleep(60000);
  while true do
  begin
     try
        if pleaseClose then
           break;

        popup := 'IEXPLORE.EXE|Static|CODIGO PRESTADOR INVALIDO';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if WINTASK.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        popup := 'IEXPLORE.EXE|Static|CODIGO DO TIPO DOCUMENTO HONORARIOS INVALIDO';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        popup := 'IEXPLORE.EXE|Static|DATA DO ATO INVALIDO';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        if wintask.ie8 then
           popup := 'IEXPLORE.EXE|Static|Esta dependência foi encerrada, deseja alterá-la para'
        else
           popup := 'IEXPLORE.EXE|Static|Esta dependência foi encerrada, deseja alterá-la para';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        popup := 'TASKEXEC.EXE|Static|Error at line 4 : Page time out in UsePage!';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           wintask.Click_Button('TASKEXEC.EXE|#32770|Execution error in fc_click_HTMLItem.rob', 'OK');
           Sleep(400);
        end;

        popup := 'IEXPLORE.EXE|Static|CODIGO DO TIPO DOCUMENTO HONORARIOS INVALIDO';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        popUp := 'IEXPLORE.EXE|Static|O campo CNPJ é obrigatório';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        popUp := 'IEXPLORE.EXE|Static|O campo Nº do Processo Bradesco é obrigatório';
        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        popUp := 'IEXPLORE.EXE|Static|O campo Data do Documento é obrigatório';
        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da página da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
           Sleep(400);
        end;

        for i:= 0 to 10 do
        begin
           Application.ProcessMessages; 
           Sleep(6000);
        end;
     except
     end;
  end;
end;

procedure TClickOkThread.SetpleaseClose(const Value: boolean);
begin
  try
     FpleaseClose := Value
  except
  end;
end;

procedure TClickOkThread.Setwintask(const Value: TWintask);
begin
  Fwintask := Value;
end;

end.
