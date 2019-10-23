unit datamodule_honorarios;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, strutils, mywin, mgenlib, dialogs, dateutils,comobj,ComCtrls,Gauges,Graphics,ExtCtrls,GifImage;

type
  TdmHonorarios = class(TDataModule)
    adoConn: TADOConnection;
    adoCmd: TADOCommand;
    dts: TADODataSet;
    dtsOrg: TADODataSet;
    adoDts: TADODataSet;
    dtsEscritorios: TADODataSet;
    dsEscritorios: TDataSource;
    dtsAtosDigitados: TADODataSet;
    dsAtosDigitados: TDataSource;
    adoConnRpt: TADOConnection;
    dtsRpt: TADODataSet;
    dtsRptTotal: TADODataSet;
    dtsAtosPendentes: TADODataSet;
    dsAtosPendentes: TDataSource;
    dtsTiposNaoAtualizar: TADODataSet;
    dsTiposNaoAtualizar: TDataSource;
    dtsValoresNaoAtualizar: TADODataSet;
    dsValoresNaoAtualizar: TDataSource;
    dtsEscritoriosIn: TADODataSet;
    dtsEscritoriosOut: TADODataSet;
    dsEscritoriosIn: TDataSource;
    dsEscritoriosOut: TDataSource;
    dtsAdvogados: TADODataSet;
    dsAdvogados: TDataSource;
    procedure DataModuleCreate(Sender: TObject);
    procedure dtsEscritoriosNewRecord(DataSet: TDataSet);
    procedure dtsTiposNaoAtualizarBeforeDelete(DataSet: TDataSet);
    procedure dsTiposNaoAtualizarDataChange(Sender: TObject;
      Field: TField);
    procedure dtsTiposNaoAtualizarNewRecord(DataSet: TDataSet);
    procedure dtsValoresNaoAtualizarNewRecord(DataSet: TDataSet);
    procedure dtsValoresNaoAtualizarBeforePost(DataSet: TDataSet);
    procedure dtsTiposNaoAtualizarBeforePost(DataSet: TDataSet);
    procedure Timer1Timer(Sender: TObject);
  private
    FdirLck: string;
    Fhandle: integer;
    procedure SetdirLck(const Value: string);
    procedure Sethandle(const Value: integer);
    { Private declarations }
  public
    { Public declarations }
    property dirLck : string read FdirLck write SetdirLck;
    property handle : integer read Fhandle write Sethandle;

    function ObtemAnoMesReferenciaProcessar(tipoPlanilha: integer) : string;
    function PlanilhaJaCruzadaGcpj(tipoplan: integer; anomesref: string) : integer;
    procedure ObtemProcessosCruzarGcpj(tipoDados, idescritorio, idplanilha, sequencia : integer; anomesreferencia: string; VAR dtsRetorno: TadoDataset);
    procedure CadastraReclamada(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha, codigoreclamada: integer; nomereclamada,
          tiporeclamada: string; sequenciagcpj, codpessoaexterna,empresa, codexfuncionario: integer);
    procedure ObtemProcessosCruzadosGcpj;
    procedure ObtemPlanilhaConvertida(gcpj: string);
    procedure InverteReclamadas(id: integer; codReclamada, nomeReclamada: string);
    function ObtemIdPlanilhaProcessar(tipoPlanilha: integer; anomesreferencia: string):integer;
    procedure GravaOcorrencia(idplanilha, linhaplanilha: integer; tipoErro: integer; mensagem: string);
    procedure MarcaProcessoCruzadoGcpj(estagio, idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer);
    procedure GravaDataDaBaixa(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; dataDaBaixa: Tdatetime; motivoBaixa: string;valorbaixa: double);
    procedure MarcaNomeEnvolvidoOk(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer);
    procedure GravaDadosEmpresaGrupo(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; codempresa: Integer; nomeEmpresa: string; empresaPara, tipo: integer; nomeempresapara: string);
    procedure GravaDadosTipoSubtipo(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha, codacao: integer; nomacao: string; codsubtipo: integer; nomsub: string; juizado, fgContrarias, fgAtivas: integer);
    procedure MarcaProcessoDigitado(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer);
    procedure GravaDadosTipoProcesso(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; tipoprocesso:string);
    function ObtemDeParaEmpresaGrupo(empresaDe: integer; tipoProcesso: string; var paranome: string):integer;
    procedure GravaValorCorrigido(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; valor: double);
    procedure GravaValorTotalDaNota(ano, sequencia: integer; anomesreferencia: string; totalAtos: integer; valorTotalNota: double);
    function ObtemNumeroNotaEmpresa(idescritorio, idplanilha, empresaLigada: integer; anodocumento: integer; aomesreferencia: string):integer;
    procedure MarcaNotaAtingiuValor(anodocumento, sequencia: integer; anomesreferencia: string; empresaligada:integer);
    procedure GravaNotaNoProcesso(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; numeroDaNota: integer);
    function NotaJaFoiDigitada(numerodocumento : Integer):boolean;
    function ObtemValorTotaldaNota(anodocumento, sequencia: integer; var totalAtos: integer):double;
    procedure MarcaDocumentoFinalizado(anodocumento, sequencia: integer; anomesreferencia: string; empresaligada:integer);
    procedure SalvaConfiguracaoWintask(diretorioExe, diretorioRob: string);
    procedure ObtemConfiguracaoWintask;
    procedure ObtemEscritoriosCadastrados(const soAtivos: boolean = false);
    procedure ObtemCadastroEscritorios(nome: string);
    function JaExistePlanilhaDoMes(idescritorio: integer; anomesreferencia: string) : integer;
    function CadastraPlanilha(idEscritorio: integer; anomesreferencia: string; var sequencia: integer):integer;
    procedure  CadastraAto(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; Cliente, processo,
                            partecontraria, gcpj, Comarca, Vara, tipoandamento, UF, divisaodafilial, idprocessounico, Porcentagem,
                            valBeneficioEconomico: string; Valor: double; datadoato: TdateTime; valorBase: double);
    procedure MarcaPlanilhaImportada(idEscritorio, idplanilha: integer; anomesreferencia:string;sequencia:integer);
    procedure ObtemPlanilhasValidar(var dtsRetorno: TAdoDataset);
    procedure RemoveReclamadas(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia: integer; linhaplanilha:integer);
    procedure GravaValorCalculo(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; valor: double);
    procedure ObtemReclamadasDoProcesso(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia,linhaplanilha: integer;
                              var dtsRetorno: TAdoDataset; tipoProcesso: string);
    procedure MarcaPlanilhaValidada(idEscritorio, idplanilha: integer; anomesreferencia:string;sequencia:integer);
    procedure RemoveNotaProcesso(idescritorio,idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha:integer);
    function SomaTotalDigitadoDaNota(numeroNota: string):double;
    function ObtemDataDigitar(anomesreferencia: string):TDateTime;
    function ObtemTotalAtosDigitados(numeroNota: string):integer;
    procedure ObtemPlanilhasDigitar(var dtsRetorno: TAdoDataset);
    procedure MarcaPlanilhaFinalizada(idEscritorio, idplanilha: integer; anomesreferencia:string;sequencia:integer);
    procedure CadastraDataInicial(anomesReferencia: string; dtInicial: Tdatetime);
    procedure ObtemPlanilhasCadastradas(var dtsRetorno: TAdoDataset; const order: integer=0);
    procedure ObtemPlanilhasEnviarGcpj(var dtsRetorno: TAdoDataset; const filtro: integer=0);
    procedure ObtemInconsistenciasImportacao(idescritorio,idplanilha: integer; anomesreferencia: string; sequencia: integer; var dtsRetorno: TAdoDataset);
    procedure ObtemInconsistenciasPosImportacao(idescritorio,idplanilha: integer; anomesreferencia: string; sequencia: integer; var dtsRetorno: TAdoDataset);
    procedure ObtemValoresRecalculadosSistema(idescritorio,idplanilha: integer; anomesreferencia: string; sequencia: integer; var dtsRetorno: TAdoDataset);
    procedure ObtemNotasFinalizadas(idescritorio,idplanilha: integer; anomesreferencia: string; sequencia: integer; var dtsRetorno: TAdoDataset; numNota:string);
    procedure ObtemProcessosIguais(idEscritorio, idplanilha, sequencia: integer; gcpj, tipoandamento, anomesreferencia: string; datadoato: TdateTime; valorAto: double; fgdrccontrarias: integer);
    procedure RemovePlanilha(idPlanilha, idEscritorio: integer; anomesreferencia: string);
    procedure ObtemNotasPendentes(idplanilha: integer; var dtsNotasPendentes: TAdoDataset);
    procedure ObtemNotasDigitando(idPlanilha, numeronota: integer; var dtsNotasDigitando: TAdoDataSet);
    procedure CadastraEscritorio(cnpj, nome, codigo, apelido: string);
    function NotificouNotaFinalizada(numeronota: integer):boolean;
    procedure MarcaNotaFinalizadaNotificada(numeronota: integer);
    function NotificouNotaIniciada(numeronota: integer):boolean;
    procedure MarcaNotaIniciadaNotificada(numeronota: integer);
    procedure EsperaLiberacaoGravacao;
    procedure LiberaGravacao;
    procedure MarcaPlanilhaNaoFinalizada(idEscritorio, idplanilha: integer; anomesreferencia:string;sequencia:integer);
    function ObtemOutrosPagamentosDoProcesso(idEscritorio: integer; idplanilha: integer; gcpj: string; tipoAto: string;
                  escritorio: string; lstCnpjs: TStringList; dtPlanilha: TdateTime) : integer;
    //inserido em 03/04/2014
    function ObtemUltimaDataCadastrada : TDateTime;
    function ObtemUltimoAnoMesCadastrado : string;
    procedure ObtemAtosPendentes(idplanilha: integer);
    procedure CriaColunaValorBase;
    function ObtemDataImportacaoPlanilha(idPlanilha: integer):TDateTime;
    function TemConsolidacaoPagaNoSistema(gcpj: string):boolean;

    //inserido em 07/07/2014
    procedure CriaTabelaTiposNaoAtualizar;
    procedure CriaTabelaValoresNaoAtualizar;
    procedure InsereTiposNaoAtualizarPadrao;
    procedure InsereValoresNaoAtualizarPadrao;
    procedure ObtemTiposNaoAtualizarCadastrados;
    procedure ExcluiValoresNaoAtualizarFK(identificador: Integer);
    procedure ExcluiTiposNaoAtualizar(identificador: Integer);
    procedure ExcluiValoresNaoAtualizarPK(identificador: Integer; tipoandamento: string);
    procedure CriaColunaIdentificador;
    procedure ObtemEscritoriosAssociadosAoTipo(identificador: integer);
    procedure ObtemEscritoriosNaoAssociadosAoTipo(identificador: integer; nomeescritorio: string);
    procedure MarcaEscritorioIn(identificador, idEscritorio: integer);
    procedure MarcaEscritorioOut(identificador, idEscritorio: integer);
    function ObtemValorSemReajuste(cnpjEscritorio, tipoAndamento: string; dataDoAto: TDateTime):double;
    procedure CriaTabelaAdvogadosInternos;
    procedure ObtemAdvogadosCadastrados(codfuncional: integer);


    //inserido em 31/07/2014
    procedure CriaColunaDataDoAto_TiposNaoAtualizar;
    function ConfiguracaoSemReajusteEstaOK : boolean;
    procedure CreateIndex;
    procedure ResetFlagsDaPlanilha(idPlanilha: integer);

    //inserido em 01/08/2014
    procedure ExcluiColunaIdentificador;
    procedure CriaTabelaLinkEscritoriosValores;

    //inserido em 04/02/2015
    procedure CriaColunafgenviadogcpj_tbplanilhas;
    procedure CriaColunadtenviogcpj_tbplanilhas;
    procedure CriaColunafgretornogcpj_tbplanilhas;

    //insserido em 10/03/2015
    function ObtemSomaJaPagaADescontar(gcpj, tipodeandamento : string; idplanilha : integer; const recuperacaoFinal : boolean = false):double;

    procedure RemoveOcorrenciasDaPlanilha(idplanilha: integer);

    procedure CriaColunaCodExFuncionario;

    procedure IndexaTabelaReclamadas;


    procedure CadastraDeParaEmpresaGrupo(empresaDe: integer; nomeEmpresaDe: string; empresaPara: integer; nomeEmpresaPara: string; tipoProcesso: string);

    procedure LimpaTabelaDePara;

    procedure ReindexTabelaDePara(tipo: integer);

    procedure executabackup(users: string;tempo:integer;dataatual:Tdatetime;FormBackup:TForm);

  end;

var
  dmHonorarios: TdmHonorarios;

implementation

uses
   fcadastravaloresnaoatualizar;
{$R *.dfm}

{ TdmHonorarios }

procedure TdmHonorarios.CadastraReclamada(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha, codigoreclamada: integer; nomereclamada,
          tiporeclamada: string; sequenciagcpj, codpessoaexterna, empresa, codexfuncionario: integer);
begin
   adoDts.Close;
   adoDts.CommandText := 'select idplanilha from tbplanilhasreclamadas ' +
                         'where idescritorio=' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and anomesreferencia= ''' + anomesreferencia + ''' ' +
                         'and sequencia = ' + IntToStr(sequencia) + ' and linhaplanilha = ' + IntToStr(linhaplanilha) + ' and  '  +
                         'codigoreclamada = ' + IntToStr(codigoreclamada);
   adoDts.Open;
   if adodts.eof then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'insert into tbplanilhasreclamadas(idescritorio, idplanilha, anomesreferencia, sequencia, linhaplanilha, codigoreclamada, nomereclamada, tiporeclamada, sequenciagcpj, codigopessoaexterna, codempresa, codexfuncionario) ' +
                             'values(' + IntToStr(idescritorio) + ', ' + IntToStr(idplanilha) + ', ''' + anomesreferencia + ''', ' + IntToStr(sequencia) + ', ' + IntToStr(linhaplanilha) + ', '  +
                             IntToStr(codigoreclamada) + ', :nomereclamada, ''' + tiporeclamada + ''', ' + IntToStr(sequenciagcpj) + ', ' + IntToStr(codpessoaexterna) + ', ' +
                             IntToStr(empresa) + ',' + Inttostr(codexfuncionario) + ')';
      adoCmd.Parameters.ParamByName('nomereclamada').Value := nomereclamada;
      adoCmd.Execute;
   end;
end;

procedure TdmHonorarios.InverteReclamadas(id: integer; codReclamada,
  nomeReclamada: string);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'Update tbjbm set cod_reclamada_1 = ' + codReclamada + ', reclamada_1 = "' + nomeReclamada + '", ' +
                            'cod_reclamada_2 = 0, reclamada_2 = null, fgreclamada1encontrada=1, fgcruzadogcpj=2 ' +
                            'where id=' + IntToStr(id);
      adocmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

function TdmHonorarios.ObtemAnoMesReferenciaProcessar(tipoPlanilha: integer): string;
begin
   dts.Close;
   dts.CommandText := 'Select anomesreferencia from tbplanilhasprocessamento where fgimportada = 1 and fgcruzadagcpj=0 and tipoplanilha=' + IntToStr(tipoPlanilha) +
                      ' order by anomesreferencia';
   dts.Open;
   if dts.Eof then
   begin
      Result := '';
      exit;
   end;
   dts.First;
   result := dts.FieldByName('anomesreferencia').AsString;
end;

procedure TdmHonorarios.ObtemPlanilhaConvertida(gcpj: string);
begin
   dtsOrg.Close;
   dtsOrg.CommandText := 'SELECT sgcp, COD_RECLAMADA_1, COD_RECLAMADA_2, RECLAMADA_1, RECLAMADA_2 ' +
                         'FROM principa ' +
                         'where sgcp = ' + gcpj;
   dtsOrg.open;
end;

procedure TdmHonorarios.ObtemProcessosCruzadosGcpj;
begin
   dts.Close;
   dts.CommandText := 'SELECT idplanilha, GCPJ, COD_RECLAMADA_1, COD_RECLAMADA_2, RECLAMADA_1, RECLAMADA_2, ' +
                      'tipoandamento ' +
                      'FROM TBJBM ' +
                      'ORDER BY GCPJ';
   dts.open;
end;

procedure TdmHonorarios.ObtemProcessosCruzarGcpj(tipoDados, idescritorio, idplanilha, sequencia: integer; anomesreferencia: string; var dtsRetorno: TadoDataset);
begin
   dtsRetorno.Close;
   dtsRetorno.Parameters.Clear;
   dtsRetorno.CommandText := '';
   case tipoDados of
      0 : dtsRetorno.CommandText := 'Select idescritorio, idplanilha, sequencia, anomesreferencia, linhaplanilha, valor, gcpj, partecontraria,  ' +
                                    'fgcruzadogcpj, tipoandamento, vara, datadoato, fgtipoprocesso, valorbase, codtipoacao,fgDrcAtivas, fgDrcContrarias, codtipoacao,cliente ' +
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj = ' + IntToStr(tipoDados) + ' order by linhaplanilha ';
      1, 2, 3 : dtsRetorno.CommandText := 'Select idescritorio, idplanilha, sequencia, anomesreferencia, linhaplanilha, gcpj, partecontraria, tipoandamento, vara, datadoato, fgtipoprocesso, valorbase, fgDrcAtivas, fgDrcContrarias, codtipoacao,cliente '+
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj = ' + IntToStr(tipoDados) + ' order by gcpj ';
      4 : dtsRetorno.CommandText := 'Select * '+
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj = ' + IntToStr(tipoDados) + ' ' +
                                    'order by linhaplanilha ';
      5 : begin
         dtsRetorno.CommandText := 'Select idescritorio, idplanilha, sequencia, anomesreferencia, linhaplanilha, gcpj, tipoandamento, vara, valor, motivobaixa, tipoacao, codtipoacao, codsubtipoacao, subtipoacao, fgjuizado, '+
                                    'datadoato, motivobaixa, databaixa, valorcorrigido, datadoato, fgtipoprocesso, fgDrcContrarias, fgDrcAtivas, valorbase,cliente ' +
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj = 5 ' +
                                    'and datadoato >= :data and tipoandamento = ''TUTELA'' and fgDrcAtivas = 0 ' +
                                    'order by linhaplanilha ';
         dtsRetorno.Parameters.ParamByName('data').Value := EncodeDate(2011,11,01);
      end;
      //nota
      6 : dtsRetorno.CommandText := 'Select * '+
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj in (5,6) and numeronota = 0 ' +
                                    'order by empresaligadaagrupar, gcpj ';
      //relatorio
      7 : dtsRetorno.CommandText := 'Select * '+
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and numeronota <> 0 ' +
                                    'order by numeronota ';
      //validação
     -1 : dtsRetorno.CommandText := 'Select * '+
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj <> 9 ' +
                                    'order by numeronota ';
      //checa se finalizou
      9 : dtsRetorno.CommandText := 'Select * '+
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj <> 9 and fgcruzadogcpj <> 6 ' +
                                    'order by numeronota ';
      10 : begin
         dtsRetorno.CommandText := 'Select idescritorio, idplanilha, sequencia, anomesreferencia, linhaplanilha, gcpj, tipoandamento, vara, valor, motivobaixa, tipoacao, codtipoacao, codsubtipoacao, subtipoacao, fgjuizado, '+
                                    'datadoato, motivobaixa, databaixa, valorcorrigido, datadoato, fgtipoprocesso, fgDrcContrarias, fgDrcAtivas, valorbase,cliente ' +
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj = 5 ' +
                                    'and datadoato >= :data and tipoandamento = ''AJUIZAMENTO'' and fgDrcAtivas = 1 ' +
                                    'order by linhaplanilha ';
         dtsRetorno.Parameters.ParamByName('data').Value := EncodeDate(2011,11,01);
      end;
      11 : begin
         dtsRetorno.CommandText := 'Select idescritorio, idplanilha, sequencia, anomesreferencia, linhaplanilha, gcpj, tipoandamento, vara, valor, motivobaixa, tipoacao, codtipoacao, codsubtipoacao, subtipoacao, fgjuizado, '+
                                    'datadoato, motivobaixa, databaixa, valorcorrigido, datadoato, fgtipoprocesso, fgDrcContrarias, fgDrcAtivas, valorbase,cliente ' +
                                    'from tbplanilhasatos ' +
                                    'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                                    'sequencia = ' + IntToStr(sequencia) + ' and fgcruzadogcpj = 5 ' +
                                    'and datadoato >= :data and (tipoandamento = ''TUTELA'' or tipoandamento = ''AJUIZAMENTO'') and fgDrcAtivas in(0, 1) ' +
                                    'order by linhaplanilha ';
         dtsRetorno.Parameters.ParamByName('data').Value := EncodeDate(2011,11,01);
      end;
   end;
   dtsRetorno.Open;
end;

function TdmHonorarios.PlanilhaJaCruzadaGcpj(tipoplan: integer;
  anomesref: string): integer;
begin
   result := 0;
   dts.Close;
   dts.CommandText := 'select tipoplanilha, fgimportada, fgcruzadagcpj, tipoandamento from tbplanilhasprocessamento where tipoplanilha = ' + IntToStr(tipoplan) + ' and '+
                      'anomesreferencia = ''' + anomesref + '''';
   dts.Open;
   if dts.Eof then
   begin
      result := 1;
      exit;
   end;

   if dts.FieldByName('fgimportada').AsInteger = 0 then
   begin
      result := 1;
      exit;
   end;

   if dts.FieldByName('fgcruzadagcpj').AsInteger = 1 then
   begin
      result := 2;
      exit;
   end;
end;

procedure TdmHonorarios.DataModuleCreate(Sender: TObject);
//var
//   ini : TiniFile;
begin
(**   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('honorarios', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;**)
   dirLck := ExtractFilePath(Application.ExeName) + '..\dirlck';
end;

function TdmHonorarios.ObtemIdPlanilhaProcessar(tipoPlanilha: integer;
  anomesreferencia: string): integer;
begin
   dts.Close;
   dts.CommandText := 'Select idplanilha ' +
                      'from tbplanilhasprocessamento ' +
                      'where fgimportada = 1 and fgcruzadagcpj=0 and tipoplanilha=' + IntToStr(tipoPlanilha) + ' and ' +
                      'anomesreferencia = ''' + anomesreferencia + '''';
   dts.Open;
   if dts.Eof then
   begin
      Result := 0;
      exit;
   end;
   result := dts.FieldByName('idplanilha').AsInteger;
end;

procedure TdmHonorarios.GravaOcorrencia(idplanilha, linhaplanilha: integer; tipoErro: integer; mensagem: string);
var
   idTipoOcorrencia : integer;
   erro : string;
begin
   dts.Close;
   dts.CommandText := 'select idtipoocorrencia from tbtiposocorrencia where ocorrencia = ''' + mensagem + '''';
   dts.Open;
   if dts.eof then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'insert into tbtiposocorrencia(ocorrencia) ' +
                            'values(''' + mensagem + ''')';
      try
         adoCmd.Execute;
         Sleep(500);
      except
         on merr : Exception do
            erro := merr.Message;
      end;
      GravaOcorrencia(idplanilha, linhaplanilha, tipoErro, mensagem);
      exit;
   end;

   idTipoOcorrencia := dts.FieldByName('idtipoocorrencia').AsInteger;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'delete from tbocorrenciasprocessamento where idtipoerro = ' + IntToStr(tipoErro) + ' and ' +
                         'idplanilha = ' + IntToStr(idplanilha) + ' and linhaplanilha = ' + IntToStr(linhaplanilha);
   adoCmd.Execute;

(**   dts.Close;
   dts.CommandText := 'select idplanilha from tbocorrenciasprocessamento where idtipoerro = ' + IntToStr(tipoErro) + ' and ' +
                      'idplanilha = ' + IntToStr(idplanilha) + ' and idtipoocorrencia = ' + IntToStr(idTipoOcorrencia) + ' and ' +
                      'linhaplanilha = ' + IntToStr(linhaplanilha);
   dts.Open;
   if dts.eof then
   begin**)
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'insert into tbocorrenciasprocessamento(idplanilha, linhaplanilha, idtipoerro, datahoraocorrencia, idtipoocorrencia) ' +
                            'values(' + IntToStr(idplanilha) + ', ' + IntToStr(linhaplanilha) + ', ' + IntToStr(tipoErro) + ', :hoje, ' + IntToStr(idTipoOcorrencia) + ')';
      adoCmd.Parameters.ParamByName('hoje').Value := Now;
      try
         adoCmd.Execute;
         Sleep(500);
      except
         on merr : Exception do
            erro := merr.Message;
      end;
//   end;
end;

procedure TdmHonorarios.MarcaProcessoCruzadoGcpj(estagio, idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhasatos set fgcruzadogcpj = ' + IntToStr(estagio) + ' ' +
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.GravaDataDaBaixa(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; dataDaBaixa: Tdatetime; motivoBaixa: string;valorbaixa: double);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhasatos set databaixa = :dtBaixa, motivobaixa = ''' + motivoBaixa + ''', valorbaixa = :valor ' +
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Parameters.ParamByName('dtBaixa').Value := dataDaBaixa;
      adoCmd.Parameters.ParamByName('valor').Value := valorbaixa;
      adoCmd.execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.MarcaNomeEnvolvidoOk(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhasatos set fgautorok = 1 ' +
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.GravaDadosEmpresaGrupo(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; codempresa: Integer; nomeEmpresa: string; empresaPara, tipo: integer; nomeempresapara: string);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      case tipo of
         0 : adoCmd.CommandText := 'update tbplanilhasatos set nomeempresaligada = ''' + nomeEmpresa + ''', codempresaligada = ' +  IntToStr(codempresa) + ', '+
                                   'empresaligadaagrupar = '+ IntToStr(empresaPara) + ', nomeempresaagrupar = ''' + nomeempresapara + ''' '+
                                   'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                                   'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                                   'linhaplanilha = ' + IntToStr(linhaplanilha);
         1 : adoCmd.CommandText := 'update tbplanilhasatos set checarnomeempresagrupo = ''' + nomeEmpresa + ''', checarcodempresagrupo = ' +  IntToStr(codempresa) + ', '+
                                   'checarempresaagrupar = '+ IntToStr(empresaPara) + ' ' +
                                   'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                                   'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                                   'linhaplanilha = ' + IntToStr(linhaplanilha);
      end;
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;

end;

procedure TdmHonorarios.GravaDadosTipoSubtipo(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha, codacao: integer; nomacao: string; codsubtipo: integer; nomsub: string; juizado, fgContrarias, fgAtivas: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhasatos set codtipoacao = ' +IntToStr(codacao) + ', tipoacao = ''' + nomacao + ''', codsubtipoacao = ' + IntToStr(codsubtipo) + ', subtipoacao = ''' + nomsub + ''', ' +
                            ' fgjuizado = ' + IntToStr(juizado) + ', fgDrcContrarias = ' + IntToStr(fgContrarias) + ', fgDrcAtivas = ' + IntToStr(fgAtivas) + ' ' +
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.MarcaProcessoDigitado(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'UPDATE tbplanilhasatos set fgdigitado=1, datahoradigitacao = :agora ' +
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Parameters.ParamByName('agora').Value := Now;
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

function TdmHonorarios.ObtemDeParaEmpresaGrupo(empresaDe: integer;
  tipoProcesso: string; var paranome: string): integer;
var
   validarTipo : string;
begin
   validarTipo := tipoProcesso;
   if (validarTipo='TR') or (validarTipo='TO') or (validarTipo='TA') then
      validarTipo := 'T'
   else
      if (validarTipo='CI') then
         validarTipo := 'C';

   dts.Close;
   dts.CommandText := 'SELECT * FROM TBDEPARAEMPRESAGRUPO WHERE CODIGO_EMPRESA_GRUPO_DE = ' + IntToStr(empresaDe);
   dts.Open;

   result := empresaDe;
   while not dts.Eof do
   begin
      if dts.FieldByName('tipo_processo').AsString = 'Z' then //serve para qualquer tipo de processso
      begin
         paranome := dts.FieldByName('nome_empresa_grupo_para').AsString;
         result := dts.FieldByName('codigo_empresa_grupo_para').AsInteger;
         exit;
      end;

      if dts.FieldByName('tipo_processo').AsString = validarTipo then
      begin
         paranome := dts.FieldByName('nome_empresa_grupo_para').AsString;
         result := dts.FieldByName('codigo_empresa_grupo_para').AsInteger;
         exit;
      end;
      dts.Next;
   end;
end;

procedure TdmHonorarios.GravaValorCorrigido(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer;
  valor: double);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'UPDATE tbplanilhasatos SET valorcorrigido = :valor ' +
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Parameters.ParamByName('valor').Value := StrToFloat(FormatFloat('0.00',valor));
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.GravaValorTotalDaNota(ano, sequencia: integer;
  anomesreferencia: string; totalAtos: integer; valorTotalNota: double);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.parameters.clear;
      adoCmd.CommandText := 'UPDATE tbcontrolenotas SET totalDeAtos='+ IntToStr(totalAtos) + ', '+
                            'valortotal=:valor '+
                            'where anodocumento = ' + IntToStr(ano) + ' and sequencia = '+ IntToStr(sequencia) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + '''';
      adoCmd.Parameters.ParamByName('valor').value := valorTotalNota;
      adocmd.execute;
   finally
//      LiberaGravacao;
   end;
end;

function TdmHonorarios.ObtemNumeroNotaEmpresa(idescritorio, idplanilha, empresaLigada,
  anodocumento: integer; aomesreferencia: string):integer;
var
   sequencia : integer;
begin
   dts.Close;
   dts.CommandText := 'select anodocumento, anomesreferencia, sequencia, totaldeatos, valortotal from tbcontrolenotas '+
                      'where anodocumento = ' + IntToStr(anodocumento) + ' and anomesreferencia = ''' + aomesreferencia + ''' and ' +
                      'empresa_ligada_agrupar = ' + IntToStr(empresaLigada) + ' and fgatingiuvalorlimite = 0 and fgstatus = 0 and ' +
                      'idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha);
   dts.Open;
   if dts.eof then
   begin
      dts.Close;
      dts.CommandText := 'select max(sequencia) as ultimo from  tbcontrolenotas ' +
                         'where anodocumento = ' + IntToStr(anodocumento);
      dts.Open;
      if dts.FieldByName('ultimo').IsNull then
         sequencia := 1
      else
      if dts.FieldByName('ultimo').asinteger = 0 then
         sequencia := 1
      else
         sequencia := dts.FieldByName('ultimo').AsInteger+1;

      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbcontrolenotas(anodocumento, sequencia, anomesreferencia, totaldeatos, valortotal, '+
                            'fgstatus, empresa_ligada_agrupar, fgatingiuvalorlimite, idescritorio, numeronota, idplanilha) ' +
                            'VALUES('+ IntToStr(anodocumento) + ', ' + IntToStr(sequencia) + ', ''' + aomesreferencia + ''',  0, 0, 0, ' + IntToStr(empresaLigada) + ', 0, ' +
                            IntToStr(idescritorio) + ', ' + IntToStr(anodocumento) + LeftZeroStr('%05d', [sequencia]) + ', ' +
                            IntToStr(idplanilha) + ')';
      adoCmd.execute;

      result := ObtemNumeroNotaEmpresa(idescritorio, idplanilha, empresaLigada, anodocumento, aomesreferencia);
      exit;
   end
   else
      sequencia := dts.FieldByname('sequencia').AsInteger;
   //verifica se nao escaparam valores sem gravar
   adoDts.Close;
   adoDts.Parameters.Clear;
   adoDts.CommandText := 'select sum(valorcorrigido) as valortotal, count(sequencia) as totaldeatos '+
                         'from tbplanilhasatos ' +
                         'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = '+ IntToStr(idplanilha) + ' and ' +
                         'empresaligadaagrupar = ' + IntToStr(empresaLigada) + ' and ' +
                         'numeronota = ' + IntToStr(dmHonorarios.dts.FieldByName('anodocumento').AsInteger) + LeftZeroStr('%05d', [sequencia]);
   adoDts.Open;
   if (FormatFloat('0.00',adoDts.FieldByName('valortotal').AsFloat) <> FormatFloat('0.00',dts.FieldByName('valortotal').AsFloat)) or
      (adoDts.FieldByName('totaldeatos').AsInteger <> dts.FieldByName('totaldeatos').AsInteger) then
   begin
//      EsperaLiberacaoGravacao;
      try
         adoCmd.Parameters.Clear;
         adoCmd.CommandText := 'update tbcontrolenotas set valortotal=:valor, totaldeatos='+adoDts.FieldByName('totaldeatos').AsString + ' ' +
                               'where anodocumento = ' + IntToStr(anodocumento) + ' and anomesreferencia = ''' + aomesreferencia + ''' and ' +
                               'empresa_ligada_agrupar = ' + IntToStr(empresaLigada) + ' and fgatingiuvalorlimite = 0 and ' +
                               'sequencia = ' + inttostr(sequencia);
         adoCmd.Parameters.ParamByName('valor').Value := adoDts.FieldByName('valortotal').AsFloat;
         adoCmd.execute;
      finally
//         LiberaGravacao;
      end;
      result := ObtemNumeroNotaEmpresa(idescritorio, idplanilha, empresaLigada, anodocumento, aomesreferencia);
      exit;
   end;
   result := StrToInt(IntToStr(dmHonorarios.dts.FieldByName('anodocumento').AsInteger) + LeftZeroStr('%05d', [sequencia]));
end;

procedure TdmHonorarios.MarcaNotaAtingiuValor(anodocumento,
  sequencia: integer; anomesreferencia: string; empresaligada: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbcontrolenotas set fgatingiuvalorlimite=1 ' +
                            'where anodocumento = ' + IntToStr(anodocumento) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                            'empresa_ligada_agrupar = ' + IntToStr(empresaLigada) + ' and fgatingiuvalorlimite = 0 ';
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.GravaNotaNoProcesso(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; numeroDaNota: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhasatos set numeronota=' + IntToStr(numeroDaNota) + ', fgliberadodigitacao=1 '+
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.execute;
   finally
//      LiberaGravacao;
   end;

end;

function TdmHonorarios.NotaJaFoiDigitada(numerodocumento: Integer): boolean;
begin
   dts.Close;
   dts.CommandText := 'select count(sequencia) as totaldigitado from tbplanilhasatos where numeronota = '+ IntToStr(numerodocumento) + ' and '+
                      'fgdigitado = 1';
   dts.Open;
   result := (dts.FieldByName('totaldigitado').AsInteger > 0);
end;

function TdmHonorarios.ObtemValorTotaldaNota(anodocumento,
  sequencia: integer; var totalAtos: integer): double;
begin
   dts.Close;
   dts.CommandText := 'select valortotal, totaldeatos ' +
                      'from tbcontrolenotas ' +
                      'where anodocumento = '+ IntToStr(anodocumento) + ' and ' +
                      'sequencia = ' + IntToStr(sequencia);
   dts.Open;
   totalAtos := dts.fieldbyName('totaldeatos').AsInteger;
   result := dts.fieldbyName('valortotal').AsFloat;
end;

procedure TdmHonorarios.MarcaDocumentoFinalizado(anodocumento,
  sequencia: integer; anomesreferencia: string; empresaligada: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      try
         adoCmd.Parameters.Clear;
         adoCmd.CommandText := 'update tbcontrolenotas set fgstatus=9 ' +
                               'where anodocumento = ' + IntToStr(anodocumento) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                               'sequencia = ' + IntToStr(sequencia) + ' and ' +
                               'empresa_ligada_agrupar = ' + IntToStr(empresaLigada);
         adoCmd.Execute;
      except
         on merr : Exception do
         begin
            ShowMessage('Erro MarcaDocumentoFinalizado: ' + merr.message);
         end;
      end;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.SalvaConfiguracaoWintask(diretorioExe,
  diretorioRob: string);
begin
   dts.Close;
   dts.CommandText := 'select id from tbconfigurawintask where id=1';
   dts.Open;
   if dts.eof then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'insert into tbconfigurawintask(id, diretorioexe, diretoriorob) '+
                            'values(1, ''' + diretorioExe + ''', ''' + diretoriorob + ''')';
      adoCmd.execute;
      exit;
   end;

//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbconfigurawintask set diretorioexe = ''' + diretorioExe + ''', ''' + diretorioRob + ''' ' +
                            'where id = 1 ';
      adocmd.execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.ObtemConfiguracaoWintask;
begin
   dts.Close;
   dts.CommandText := 'select diretorioexe, diretoriorob from tbconfigurawintask where id=1';
   dts.Open;
end;

procedure TdmHonorarios.ObtemEscritoriosCadastrados(const soAtivos: boolean = false);
begin
   dts.Close;
   dts.CommandText := 'select idescritorio, nomeescritorio, cnpjescritorio, nomedigitar, fgativo, fgpagaribi, dataatosibi, codgcpjescritorio from tbescritorios';
   if soAtivos then
      dts.CommandText := dts.CommandText + ' where fgativo=1';
   dts.CommandText := dts.CommandText + ' order by nomeescritorio';
   dts.Open;
end;

function TdmHonorarios.JaExistePlanilhaDoMes(idescritorio: integer; anomesreferencia: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select idplanilha from tbplanilhas where fgimportada = 1 and idescritorio = ' + IntToStr(idescritorio) + ' and anomesreferencia = ''' + anomesreferencia + '''';
   dts.open;
   if dts.eof then
      result := 0
   else
      result := dts.FieldByName('idplanilha').AsInteger;
end;

function TdmHonorarios.CadastraPlanilha(idEscritorio: integer;
  anomesreferencia: string; var sequencia: integer): integer;
begin
   dts.Close;
   dts.CommandText := 'select idplanilha,sequencia, fgimportada from tbplanilhas ' +
                      'where idescritorio = ' + IntToStr(idEscritorio) + ' and anomesreferencia = ''' + anomesreferencia + '''';
   dts.Open;
   if dts.Eof then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'insert into tbplanilhas(idescritorio, anomesreferencia, sequencia, fgimportada, fgfinalizada) ' +
                            'values(' + IntToStr(idEscritorio) + ', ''' + anomesreferencia + ''', 1, 0, 0) ';
      adoCmd.Execute;

      sequencia := 1;
      result := CadastraPlanilha(idEscritorio, anomesreferencia, sequencia);
      exit;
   end;

   //localiza o ultimo
   while not dts.eof do
   begin
      if dts.FieldByName('fgimportada').AsInteger = 0 then
      begin
         sequencia := dts.FieldByName('sequencia').AsInteger;
         result := dts.FieldByName('idplanilha').AsInteger;
         exit;
      end;
      dts.Next;
   end;

   dts.Close;
   dts.CommandText := 'select max(sequencia) as ultima from tbplanilhas ' +
                      'where idescritorio = ' + IntToStr(idEscritorio) + ' and anomesreferencia = ''' + anomesreferencia + '''';
   dts.Open;

   sequencia := dts.FieldByName('ultima').AsInteger+1;


   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'insert into tbplanilhas(idescritorio, anomesreferencia, sequencia, fgimportada, fgfinalizada) ' +
                         'values(' + IntToStr(idEscritorio) + ', ''' + anomesreferencia + ''', ' + IntToStr(sequencia) + ', 0, 0) ';
   adoCmd.Execute;

   result := CadastraPlanilha(idEscritorio, anomesreferencia, sequencia);
end;

procedure TdmHonorarios.CadastraAto(idescritorio, idplanilha: integer;
  anomesreferencia: string; sequencia, linhaplanilha: integer; Cliente,
  processo, partecontraria, gcpj, Comarca, Vara, tipoandamento, UF,
  divisaodafilial, idprocessounico, Porcentagem,
  valBeneficioEconomico: string; Valor: double; datadoato: TdateTime; valorBase: double);
begin
   dts.Close;
   dts.CommandText := 'select idplanilha ' +
                      'from tbplanilhasatos ' +
                      'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                      'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                      'linhaplanilha = ' + IntToStr(linhaplanilha);
   dts.Open;
   if Not Dts.Eof then
      exit;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'insert into tbplanilhasatos(idescritorio, idplanilha, anomesreferencia, sequencia, linhaplanilha, Cliente, '+
                         'processo, partecontraria, gcpj, Comarca, Vara, tipoandamento, UF, divisaodafilial, idprocessounico, Porcentagem, ' +
                         'valBeneficioEconomico, valor, datadoato, valorbase) ' +
                         'values(' + IntToStr(idescritorio) + ', ' + IntToStr(idplanilha) + ', ''' + anomesreferencia + ''', ' + IntToStr(sequencia) + ', ' +
                         IntToStr(linhaplanilha) + ', ''' + Cliente + ''', ''' + processo + ''', "' + partecontraria + '", ' + gcpj + ', :Comarca, ' +
                         '''' + Vara + ''', ''' + tipoandamento + ''', ''' + UF + ''', ''' + divisaodafilial + ''', ''' + idprocessounico + ''', ''' + Porcentagem + ''', '+
                         '''' + valBeneficioEconomico + ''', :valor, :datadoato, :valorbase)';
   adoCmd.Parameters.ParamByName('valor').Value := Valor;
   adoCmd.Parameters.ParamByName('datadoato').Value := datadoato;
   adoCmd.Parameters.ParamByName('comarca').Value := Comarca;
   adoCmd.Parameters.ParamByName('valorbase').Value := ValorBase;
   adoCmd.Execute;
end;

procedure TdmHonorarios.MarcaPlanilhaImportada(idEscritorio,
  idplanilha: integer; anomesreferencia: string;sequencia:integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhas set fgimportada = 1, dtimportacao = :hoje '+
                            'where idEscritorio = ' + IntToStr(idEscritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia=''' + anomesreferencia + ''' and sequencia = '  + IntToStr(sequencia);
      adoCmd.Parameters.ParamByName('hoje').Value := Date;
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.ObtemPlanilhasValidar(var dtsRetorno: TAdoDataSet);
begin
   dtsRetorno.Close;
   dtsRetorno.CommandText := 'SELECT cnpjescritorio, nomeescritorio, anomesreferencia, fgimportada, fgfinalizada, dtimportacao, a.idescritorio, idplanilha, sequencia, nomedigitar, codgcpjescritorio, ' +
                             'fgpagaribi, dataatosibi  ' +
                             'from tbescritorios a, tbplanilhas b ' +
                             'where fgimportada=1 and fgvalidada=0 and a.idescritorio = b.idescritorio ' +
                             'order by dtimportacao desc';
   dtsRetorno.Open;
end;

procedure TdmHonorarios.RemoveReclamadas(idescritorio, idplanilha: integer;
  anomesreferencia: string; sequencia, linhaplanilha: integer);
begin
   try
//      EsperaLiberacaoGravacao;
      try
         adoCmd.Parameters.Clear;
         adoCmd.CommandText := 'delete from tbplanilhasreclamadas '+
                               'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + INTTOSTR(idplanilha)  + '  and '+
                               'anomesreferencia = ''' + anomesreferencia + ''' and ' +
                               'sequencia = ' + IntToStr(sequencia) + ' and linhaplanilha = ' + IntToStr(linhaplanilha) + ' and codigoreclamada >= 0 and codigopessoaexterna >= 0';
         adoCmd.Execute;
      finally
//         LiberaGravacao;
      end;
   except
   end;
end;

procedure TdmHonorarios.GravaValorCalculo(idescritorio, idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha: integer; valor: double);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhasatos set valordopedido = :valor ' +
                            'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and ' +
                            'sequencia = ' + IntToStr(sequencia) + ' and linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Parameters.ParamByName('valor').Value := valor;
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.ObtemReclamadasDoProcesso(idescritorio,
  idplanilha: integer; anomesreferencia: string; sequencia,
  linhaplanilha: integer; var dtsRetorno: TAdoDataset; tipoProcesso: string);
begin
   dtsRetorno.Close;
   dtsRetorno.CommandText := 'select codigoreclamada, nomereclamada, tiporeclamada, sequenciagcpj, codempresa, codexfuncionario, codigopessoaexterna ' +
                             'from tbplanilhasreclamadas ' +
                             'where idplanilha = ' + INTTOSTR(idplanilha)  + '  and idescritorio = ' + IntToStr(idescritorio) + ' and ' +
                             'anomesreferencia = ''' + anomesreferencia + ''' and ' +
                             'sequencia = ' + IntToStr(sequencia) + ' and linhaplanilha = ' + IntToStr(linhaplanilha);
   if (tipoProcesso <> 'TR') and (tipoProcesso <> 'TA') and (tipoProcesso <> 'TO') then
      dtsRetorno.CommandText := dtsRetorno.CommandText + ' and ' + '(codigopessoaexterna is null or codigopessoaexterna = 0) and (codexfuncionario = 0) ';
   dtsRetorno.CommandText := dtsRetorno.CommandText + ' order by sequenciagcpj';
   dtsRetorno.Open;
end;

procedure TdmHonorarios.MarcaPlanilhaValidada(idEscritorio,
  idplanilha: integer; anomesreferencia: string; sequencia: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhas set fgvalidada = 1 '+
                            'where idEscritorio = ' + IntToStr(idEscritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia=''' + anomesreferencia + ''' and sequencia = '  + IntToStr(sequencia);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.RemoveNotaProcesso(idescritorio,
  idplanilha: integer; anomesreferencia: string; sequencia,
  linhaplanilha: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'UPDATE tbplanilhasatos set numeronota = 0 ' +
                            'where idEscritorio = ' + IntToStr(idEscritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia=''' + anomesreferencia + ''' and sequencia = '  + IntToStr(sequencia);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

function TdmHonorarios.SomaTotalDigitadoDaNota(numeroNota: string): double;
begin
   dts.Close;
   dts.CommandText := 'select sum(valorcorrigido) as total ' +
                      'from tbplanilhasatos ' +
                      'where numeronota = ' + numeroNota + ' and fgdigitado = 1';
   dts.Open;

   result := dts.FieldByName('total').AsFloat;
end;

function TdmHonorarios.ObtemDataDigitar(
  anomesreferencia: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select datadigitar from tbdatadigitar where anomesreferencia = ''' + anomesreferencia + '''';
   dts.Open;
   if dts.Eof then
      result := 0
   else
      result := dts.FieldByName('datadigitar').AsDateTime;
end;

function TdmHonorarios.ObtemTotalAtosDigitados(
  numeroNota: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select count(gcpj) as total ' +
                      'from tbplanilhasatos ' +
                      'where numeronota = ' + numeroNota + ' and fgdigitado = 1';
   dts.Open;

   result := dts.FieldByName('total').AsInteger;
end;

procedure TdmHonorarios.ObtemPlanilhasDigitar(var dtsRetorno: TAdoDataset);
begin
   dtsRetorno.Close;
   dtsRetorno.CommandText := 'SELECT cnpjescritorio, nomeescritorio, anomesreferencia, fgimportada, fgfinalizada, dtimportacao, a.idescritorio, idplanilha, sequencia, nomedigitar, codgcpjescritorio, ' +
                             'fgpagaribi, dataatosibi ' +
                             'from tbescritorios a, tbplanilhas b ' +
                             'where fgimportada=1 and fgvalidada=1 and fgfinalizada=0 and a.idescritorio = b.idescritorio ' +
                             'order by dtimportacao, anomesreferencia desc, a.idescritorio';
   dtsRetorno.Open;
end;

procedure TdmHonorarios.MarcaPlanilhaFinalizada(idEscritorio,
  idplanilha: integer; anomesreferencia: string; sequencia: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhas set fgfinalizada = 1 '+
                            'where idEscritorio = ' + IntToStr(idEscritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia=''' + anomesreferencia + ''' and sequencia = '  + IntToStr(sequencia);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.CadastraDataInicial(anomesReferencia: string;
  dtInicial: Tdatetime);
begin
   if ObtemDataDigitar(anomesReferencia) = 0 then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'insert into tbdatadigitar(anomesreferencia, datadigitar) ' +
                            'values(''' + anomesReferencia + ''', :data)';
      adoCmd.Parameters.ParamByName('data').Value := StrToDate(FormatDateTime('dd/mm/yyyy',dtInicial));
      adoCmd.Execute;
   end
   else
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbdatadigitar set datadigitar = :data ' +
                            'where anomesReferencia = ''' + anomesReferencia + '''';
      adoCmd.Parameters.ParamByName('data').Value := StrToDate(FormatDateTime('dd/mm/yyyy',dtInicial));
      adoCmd.Execute;
   end;
end;

procedure TdmHonorarios.ObtemPlanilhasCadastradas(
  var dtsRetorno: TAdoDataset; const order:integer = 0);
begin
   dtsRetorno.Close;
   dtsRetorno.CommandText := 'SELECT cnpjescritorio, nomeescritorio, anomesreferencia, fgimportada, fgfinalizada, dtimportacao, a.idescritorio, idplanilha, sequencia, nomedigitar, codgcpjescritorio, ' +
                             'fgpagaribi,  dataatosibi ' +
                             'from tbescritorios a, tbplanilhas b ' +
                             'where fgimportada=1 and a.idescritorio = b.idescritorio ';
   case order of
      0 : dtsRetorno.CommandText := dtsRetorno.CommandText + 'order by anomesreferencia desc, a.idescritorio ';
      1 : dtsRetorno.CommandText := dtsRetorno.CommandText + 'order by dtimportacao desc, anomesreferencia desc, a.idescritorio ';
   end;
   dtsRetorno.Open;
end;

procedure TdmHonorarios.ObtemInconsistenciasImportacao(idescritorio,
  idplanilha: integer; anomesreferencia: string; sequencia: integer; var dtsRetorno: TAdoDataset);
begin
   dtsRetorno.Close;
(**   dtsRetorno.CommandText := 'SELECT * ' +
                             'from consInconsistenciasDeImportacao ' +
                             'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idPlanilha) + ' and ' +
                             'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' ' +
                             'order by descrtipoerro';**)
   dtsRetorno.CommandText := 'SELECT tbplanilhas.idescritorio, tbplanilhas.idplanilha, tbplanilhas.anomesreferencia, tbplanilhas.sequencia, tbescritorios.nomeescritorio, ' +
                             'tbtiposerro.descrtipoerro, tbtiposocorrencia.ocorrencia, tbescritorios.cnpjescritorio ' +
                             'FROM tbtiposocorrencia INNER JOIN (tbtiposerro INNER JOIN (tbescritorios INNER JOIN (tbplanilhas INNER JOIN tbocorrenciasprocessamento ON ' +
                             'tbplanilhas.idplanilha = tbocorrenciasprocessamento.idplanilha) ON tbescritorios.idescritorio = tbplanilhas.idescritorio) ON ' +
                             'tbtiposerro.idtipoerro = tbocorrenciasprocessamento.idtipoerro) ON tbtiposocorrencia.idtipoocorrencia = tbocorrenciasprocessamento.idtipoocorrencia ' +
                             'WHERE tbtiposerro.idtipoerro=1 AND tbplanilhas.idescritorio = ' + IntToStr(idescritorio) + ' and tbplanilhas.idplanilha = ' + IntToStr(idPlanilha) + ' and ' +
                             'tbplanilhas.anomesreferencia = ''' + anomesreferencia + ''' and tbplanilhas.sequencia = ' + IntToStr(sequencia) + ' ' +
                             'order by descrtipoerro';
   dtsRetorno.Open;
end;

procedure TdmHonorarios.ObtemInconsistenciasPosImportacao(idescritorio,
  idplanilha: integer; anomesreferencia: string; sequencia: integer;
  var dtsRetorno: TAdoDataset);
begin
   dtsRetorno.Close;
(**   dtsRetorno.CommandText := 'SELECT * ' +
                             'from consInconsistenciasPosImportacao ' +
                             'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idPlanilha) + ' and ' +
                             'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' ' +
                             'order by ocorrencia, linhaplanilha';*)
   dtsRetorno.CommandText := 'SELECT tbplanilhasatos.idescritorio, tbplanilhasatos.idplanilha, tbplanilhasatos.anomesreferencia, tbplanilhasatos.sequencia, '+
                             'tbescritorios.cnpjescritorio, tbescritorios.nomeescritorio, tbplanilhasatos.linhaplanilha, tbtiposocorrencia.ocorrencia, tbplanilhasatos.gcpj, ' +
                             'tbplanilhasatos.partecontraria, tbplanilhasatos.tipoandamento, tbplanilhasatos.datadoato, tbplanilhasatos.tipoacao, tbplanilhasatos.subtipoacao, ' +
                             'tbplanilhasatos.Vara, tbplanilhasatos.valBeneficioEconomico, tbplanilhasatos.databaixa, tbplanilhasatos.motivobaixa, tbplanilhasatos.nomeempresaligada, ' +
                             'tbplanilhasatos.Valor, tbplanilhasatos.valorcorrigido, tbplanilhasatos.valordopedido, tbplanilhasatos.valorbaixa ' +
                             'FROM tbtiposocorrencia INNER JOIN (tbtiposerro INNER JOIN ((tbplanilhasatos INNER JOIN tbocorrenciasprocessamento ON ' +
                             '(tbplanilhasatos.linhaplanilha = tbocorrenciasprocessamento.linhaplanilha) AND (tbplanilhasatos.idplanilha = tbocorrenciasprocessamento.idplanilha)) ' +
                             'INNER JOIN tbescritorios ON tbplanilhasatos.idescritorio = tbescritorios.idescritorio) ON tbtiposerro.idtipoerro = tbocorrenciasprocessamento.idtipoerro) ' +
                             'ON tbtiposocorrencia.idtipoocorrencia = tbocorrenciasprocessamento.idtipoocorrencia ' +
                             'WHERE ' + //{tbtiposocorrencia.ocorrencia<>''Valor alterado pelo sistema'' AND} +
                             ' tbplanilhasatos.fgcruzadogcpj=9 AND  tbtiposerro.idtipoerro=2 ' +
                             'and tbplanilhasatos.idescritorio = ' + IntToStr(idescritorio) + ' and tbplanilhasatos.idplanilha = ' + IntToStr(idPlanilha) + ' and ' +
                             'tbplanilhasatos.anomesreferencia = ''' + anomesreferencia + ''' and tbplanilhasatos.sequencia = ' + IntToStr(sequencia) + ' ' +
                             'ORDER BY tbplanilhasatos.linhaplanilha ' ;
   dtsRetorno.Open;
end;

procedure TdmHonorarios.ObtemValoresRecalculadosSistema(idescritorio,
  idplanilha: integer; anomesreferencia: string; sequencia: integer;
  var dtsRetorno: TAdoDataset);
begin
   dtsRetorno.Close;
   dtsRetorno.CommandText := 'SELECT * ' +
                             'from consInconsistenciasValoresDiferentes ' +
                             'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idPlanilha) + ' and ' +
                             'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' ' +
                             'order by ocorrencia, linhaplanilha';
   dtsRetorno.Open;
end;

procedure TdmHonorarios.ObtemNotasFinalizadas(idescritorio,
  idplanilha: integer; anomesreferencia: string; sequencia: integer;
  var dtsRetorno: TAdoDataset; numNota:string);
begin
   dtsRetorno.Close;
(**   dtsRetorno.CommandText := 'SELECT tbescritorios.idescritorio, tbplanilhasatos.idplanilha, tbplanilhasatos.anomesreferencia, tbplanilhasatos.sequencia, ' +
                             'tbescritorios.cnpjescritorio, tbescritorios.nomeescritorio, tbcontrolenotas.numeronota, tbplanilhasatos.linhaplanilha, tbplanilhasatos.gcpj, ' +
                             'tbplanilhasatos.partecontraria, tbplanilhasatos.tipoandamento, tbplanilhasatos.datadoato, tbplanilhasatos.tipoacao, tbplanilhasatos.subtipoacao, ' +
                             'tbplanilhasatos.Vara, tbplanilhasatos.databaixa, tbplanilhasatos.motivobaixa, tbcontrolenotas.empresa_ligada_agrupar, tbplanilhasatos.nomeempresaligada, ' +
                             'tbplanilhasatos.Valor, tbplanilhasatos.valorcorrigido, tbplanilhasatos.valordopedido, tbplanilhasatos.valorbaixa ' +
                             ' FROM (tbescritorios INNER JOIN tbplanilhasatos ON tbescritorios.idescritorio = tbplanilhasatos.idescritorio) INNER JOIN tbcontrolenotas ON ' +
                             'tbplanilhasatos.numeronota = tbcontrolenotas.numeronota ' +
                             'where tbescritorios.idescritorio = ' + IntToStr(idescritorio) + ' and tbplanilhasatos.idplanilha = ' + IntToStr(idPlanilha) + ' and ' +
                             'tbplanilhasatos.anomesreferencia = ''' + anomesreferencia + ''' and tbplanilhasatos.sequencia = ' + IntToStr(sequencia) + ' ';*)

   dtsRetorno.commandtext := 'SELECT tbescritorios.idescritorio, tbplanilhasatos.idplanilha, tbplanilhasatos.anomesreferencia, tbplanilhasatos.sequencia, tbescritorios.cnpjescritorio, ' +
                             'tbescritorios.nomeescritorio, tbcontrolenotas.numeronota, tbplanilhasatos.linhaplanilha, tbplanilhasatos.gcpj, tbplanilhasatos.partecontraria, ' +
                             'tbplanilhasatos.tipoandamento, tbplanilhasatos.datadoato, tbplanilhasatos.tipoacao, tbplanilhasatos.subtipoacao, tbplanilhasatos.Vara, ' +
                             'tbplanilhasatos.databaixa, tbplanilhasatos.motivobaixa, tbcontrolenotas.empresa_ligada_agrupar, tbplanilhasatos.nomeempresaligada, tbplanilhasatos.Valor, ' +
                             'tbplanilhasatos.valorcorrigido, tbplanilhasatos.valordopedido, tbplanilhasatos.valorbaixa ' +
                             'FROM (tbescritorios INNER JOIN tbplanilhasatos ON tbescritorios.idescritorio = tbplanilhasatos.idescritorio) INNER JOIN tbcontrolenotas ON ' +
                             '(tbplanilhasatos.idplanilha = tbcontrolenotas.idplanilha) AND (tbplanilhasatos.anomesreferencia = tbcontrolenotas.anomesreferencia) AND ' +
                             '(tbplanilhasatos.idescritorio = tbcontrolenotas.idescritorio) AND (tbplanilhasatos.numeronota = tbcontrolenotas.numeronota) ' +
                             'WHERE (((tbescritorios.idescritorio)=' + IntToStr(idescritorio) +  ') AND ((tbplanilhasatos.idplanilha)=' + IntToStr(idplanilha) + ') AND ' +
                             '((tbplanilhasatos.anomesreferencia)=''' + anomesreferencia + ''') AND ((tbplanilhasatos.sequencia)=' + IntToStr(sequencia) +')) ';
   if numNota <> '' then
      dtsRetorno.CommandText := dtsRetorno.CommandText + ' and tbplanilhasatos.numeronota = ' + numNota;
   dtsRetorno.CommandText := dtsRetorno.CommandText + ' order by tbplanilhasatos.numeronota, tbplanilhasatos.gcpj ';
   dtsRetorno.Open;
end;

procedure TdmHonorarios.ObtemProcessosIguais(idEscritorio, idplanilha,
  sequencia: integer; gcpj, tipoandamento, anomesreferencia: string; datadoato: TdateTime; valorAto: double; fgdrccontrarias: integer);
begin
   adodts.Close;
   adoDts.Parameters.Clear;
   adodts.CommandText := 'select gcpj, linhaplanilha, datadoato from tbplanilhasatos ' +
                         'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idPlanilha) + ' and ' +
                         'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and  ';

   if (fgdrccontrarias <> 1) and (tipoandamento = 'EXITO') then
      adodts.CommandText := adodts.CommandText + 'gcpj = ' + gcpj + ' and (tipoandamento = ''' + tipoandamento + ''' or tipoandamento = ''HONORARIOS FINAIS'')  and fgcruzadogcpj <> 9'
   else
   if (fgdrccontrarias <> 1) and (tipoandamento = 'HONORARIOS FINAIS') then
      adodts.CommandText := adodts.CommandText + 'gcpj = ' + gcpj + ' and (tipoandamento = ''' + tipoandamento + ''' or tipoandamento = ''EXITO'')  and fgcruzadogcpj <> 9'
   else
      adodts.CommandText := adodts.CommandText + 'gcpj = ' + gcpj + ' and tipoandamento = ''' + tipoandamento + ''' and fgcruzadogcpj <> 9';

   //alterado em 26/11/2014
   IF (tipoandamento = 'PREPOSTO') then// and (tipoandamento <> 'ACORDO') then
   begin
      adodts.CommandText := adodts.CommandText + ' and datadoato = :data ';
      adodts.Parameters.ParamByName('data').Value := StrToDate(DateToStr(datadoato));
   end;

   //alterado em 12/02/2015
   IF (tipoandamento = 'ACORDO') OR ((Pos('XITO', tipoandamento) <> 0) and (fgDrcContrarias = 1)) then
   begin
      adodts.CommandText := adodts.CommandText + ' and datadoato >= :dataDe and datadoato <= :dataAte  ';
      adodts.Parameters.ParamByName('dataDe').Value := StrToDate('01/' + IntToStr(MonthOf(datadoato)) + '/' + IntToStr(YearOf(datadoato)));
      adodts.Parameters.ParamByName('dataAte').Value := StrToDate(IntToStr(DaysInMonth(datadoAto)) +'/' + FormatDateTime('mm/yyyy', datadoato));
   end;

   //se for acompanhamento trabalhista não pode ter valor duplicado
   //regra inserida em 03/04/2014 - E-mail da Amanda de 21/03/2014
   if tipoandamento = 'ACOMPANHAMENTO' then
   begin
      adodts.CommandText := adodts.CommandText + ' and valor = :valor';
      adodts.Parameters.ParamByName('valor').Value := valorAto;
   end;

   adoDts.Open;
end;

procedure TdmHonorarios.RemovePlanilha(idPlanilha, idEscritorio: integer;
  anomesreferencia: string);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.CommandText := 'delete from tbocorrenciasprocessamento where idplanilha = ' + IntToStr(idplanilha);
      adocmd.Execute;

      adocmd.CommandText := 'delete from tbcontrolenotas where idescritorio = ' + IntToStr(idEscritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                            'idplanilha = ' + IntToStr(idplanilha);
      adoCmd.execute;

      adocmd.CommandText := 'delete from tbplanilhasreclamadas where idescritorio = ' + IntToStr(idEscritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                            'idplanilha = ' + IntToStr(idplanilha);
      adoCmd.execute;

      adocmd.CommandText := 'delete from tbplanilhasatos where idescritorio = ' + IntToStr(idEscritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                            'idplanilha = ' + IntToStr(idplanilha);
      adoCmd.execute;

      adocmd.CommandText := 'delete from tbplanilhas where idescritorio = ' + IntToStr(idEscritorio) + ' and anomesreferencia = ''' + anomesreferencia + ''' and ' +
                            'idplanilha = ' + IntToStr(idplanilha);
      adoCmd.execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.GravaDadosTipoProcesso(idescritorio,
  idplanilha: integer; anomesreferencia: string; sequencia, linhaplanilha : integer;
  tipoprocesso: string);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhasatos set fgtipoprocesso = ''' + tipoprocesso + ''' ' +
                            'where idescritorio = ' + IntToStr(idescritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia = ''' + anomesreferencia + ''' and sequencia = ' + IntToStr(sequencia) + ' and ' +
                            'linhaplanilha = ' + IntToStr(linhaplanilha);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.ObtemNotasPendentes(idplanilha: integer;
  var dtsNotasPendentes: TAdoDataset);
begin
   dtsNotasPendentes.Close;
   dtsNotasPendentes.CommandText := 'select numeronota, totaldeatos, valortotal, empresa_ligada_agrupar ' +
                                    'from tbcontrolenotas ' +
                                    'where fgstatus = 0 and idplanilha = ' + IntToStr(idplanilha) + ' ' +
                                    'order by totaldeatos';
   dtsNotasPendentes.Open;
end;

procedure TdmHonorarios.ObtemNotasDigitando(idPlanilha,
  numeronota: integer; var dtsNotasDigitando: TAdoDataSet);
begin
   dtsNotasDigitando.CommandText := 'SELECT Count(linhaplanilha) AS atosdigitados, Sum(valorcorrigido) AS valordigitado ' +
                                    'FROM tbplanilhasatos ' +
                                    'where idplanilha = ' +IntToStr(idplanilha) + ' and ' +
                                    'numeronota = ' + IntToStr(numeroNota) + ' and ' +
                                    'fgdigitado=1 ';
   dtsNotasDigitando.Open;
end;

procedure TdmHonorarios.CadastraEscritorio(cnpj, nome, codigo,
  apelido: string);
var
   ssql : string;
begin
   dts.Close;
   dts.CommandText := 'select idescritorio from tbescritorios where cnpjescritorio=''' + cnpj + '''';
   dts.Open;
   if dts.eof then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'insert into tbescritorios(nomeescRitorio, cnpjescritorio, codgcpjescritorio, nomedigitar, fgativo, fgpagaribi, dataatoibi) ' +
                            'values(''' + nome + ''', ''' + cnpj + ''', ' + codigo + ', ''' + apelido + ''', 1, 0, Null)';
      adocmd.Execute;
      exit;
   end;

   ssql := '';
   if dts.FieldByName('nomeescritorio').AsString <> nome then
      ssql := 'update tbescritorios set nomeescritorio = ''' + nome + '''';

   if dts.FieldByName('codgcpjescritorio').AsString <> codigo then
   begin
      if ssql = '' then
         ssql := 'Update tbescritorios set '
      else
         ssql := ssql + ',';
      ssql := ssql + 'codgcpjescritorio = ''' + codigo + '''';
   end;

   if dts.FieldByName('nomedigitar').AsString <> apelido then
   begin
      if ssql = '' then
         ssql := 'Update tbescritorios set '
      else
         ssql := ssql + ',';
      ssql := ssql + 'nomedigitar = ''' + apelido + '''';
   end;

   if ssql <> '' then
   begin
//      EsperaLiberacaoGravacao;
      try
         adoCmd.Parameters.Clear;
         adoCmd.CommandText := ssql + ' where cnpjescritorio = ''' + cnpj + '''';
         adoCmd.Execute;
      finally
//         LiberaGravacao;
      end;
   end;
end;

function TdmHonorarios.NotificouNotaFinalizada(
  numeronota: integer): boolean;
begin
   dts.Close;
   dts.CommandText := 'Select fgnotificadofim from tbcontrolenotas where numeronota = ' + IntToStr(numeronota);
   dts.Open;
   if dts.FieldByName('fgnotificadofim').AsInteger = 0 then
      result := false
   else
      result := true;
end;
procedure TdmHonorarios.MarcaNotaFinalizadaNotificada(numeronota: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adocmd.parameters.clear;
      adoCmd.CommandText := 'update tbcontrolenotas set fgnotificadofim=1 where numeronota=' + IntToStr(numeronota);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;
end;

procedure TdmHonorarios.MarcaNotaIniciadaNotificada(numeronota: integer);
begin
   adocmd.parameters.clear;
   adoCmd.CommandText := 'update tbcontrolenotas set fgnotificadoinicio=1 where numeronota=' + IntToStr(numeronota);
   adoCmd.Execute;
end;

function TdmHonorarios.NotificouNotaIniciada(numeronota: integer): boolean;
begin
   dts.Close;
   dts.CommandText := 'Select fgnotificadoinicio from tbcontrolenotas where numeronota = ' + IntToStr(numeronota);
   dts.Open;
   if dts.FieldByName('fgnotificadoinicio').AsInteger = 0 then
      result := false
   else
      result := true;
end;

procedure TdmHonorarios.ObtemCadastroEscritorios(nome:string);
begin
   dtsEscritorios.Close;
   dtsEscritorios.CommandText := 'select idescritorio, nomeescritorio, cnpjescritorio, nomedigitar, fgativo, fgpagaribi, dataatosibi, codgcpjescritorio from tbescritorios ';
   if nome <> '' then
      dtsEscritorios.CommandText := dtsEscritorios.CommandText + ' WHERE nomeescritorio like ''%' + nome + '%''';
   dtsEscritorios.CommandText := dtsEscritorios.CommandText + ' order by nomeescritorio';
   dtsEscritorios.Open;
end;

procedure TdmHonorarios.dtsEscritoriosNewRecord(DataSet: TDataSet);
begin
   dtsEscritorios.FieldByName('fgativo').AsInteger := 1;
   dtsEscritorios.FieldByName('fgpagaribi').AsInteger := 0;
end;

procedure TdmHonorarios.SetdirLck(const Value: string);
begin
  FdirLck := Value;
end;

procedure TdmHonorarios.EsperaLiberacaoGravacao;
var
   ret : integer;
//   tries :integer;
begin
    ret := IsFileLocked(dirLck + '\GRAVAR.LCK' , fhandle);
 //   tries := 0;
    while ret <= 0 do
    begin
      Application.ProcessMessages;
      Sleep(700);
//      inc(tries);
//      if tries >= 200 then
//         exit;

      ret := IsFileLocked(dirLck + '\GRAVAR.LCK' , fhandle);
    end;
end;

procedure TdmHonorarios.LiberaGravacao;
begin
   try
      ReleaseFile(dirLck + '\GRAVAR.LCK', fhandle);
   except
   end;
end;

procedure TdmHonorarios.Sethandle(const Value: integer);
begin
  Fhandle := Value;
end;

procedure TdmHonorarios.MarcaPlanilhaNaoFinalizada(idEscritorio,
  idplanilha: integer; anomesreferencia: string; sequencia: integer);
begin
//   EsperaLiberacaoGravacao;
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'update tbplanilhas set fgfinalizada = 0 '+
                            'where idEscritorio = ' + IntToStr(idEscritorio) + ' and idplanilha = ' + IntToStr(idplanilha) + ' and ' +
                            'anomesreferencia=''' + anomesreferencia + ''' and sequencia = '  + IntToStr(sequencia);
      adoCmd.Execute;
   finally
//      LiberaGravacao;
   end;

end;

function TdmHonorarios.ObtemOutrosPagamentosDoProcesso(idEscritorio: integer; idplanilha: integer; gcpj: string; tipoAto: string;
         escritorio: string; lstCnpjs: TStringList; dtPlanilha: TdateTime) : integer;
var
   i : integer;
begin
   result := -1;
   dts.Close;
   dts.CommandText := 'select a.nomeescritorio, datadoato as datmovto, tipoandamento as nomdes ' +
                      'from tbescritorios a, tbplanilhasatos b, tbplanilhas c ' +
                      'where tipoandamento = ''' + tipoato + ''' And c.dtImportacao = :dtImportacao And ';

   if Pos('JBM', escritorio) <> 0 then
      dts.CommandText := dts.CommandText + '(nomeescritorio like ''JBM%'' or nomeescritorio like ''RAMALHO%'' or nomeescritorio like ''ALMEIDA R%'' or ' +
                         'nomeescritorio like ''SARAIVA DE%'' or nomeescritorio like ''SILAS%'' or nomeescritorio like ''D%AVILA%'') and '
   else
   if Pos('CAMARA', escritorio) <> 0 then
      dts.CommandText := dts.CommandText + '(nomeescritorio like ''SETTE%'' or nomeescritorio like ''AZEVEDO SO%'' or nomeescritorio like ''SERRAS E CE%'' or ' +
                         'nomeescritorio like ''%BORNHAUSEN%'' or nomeescritorio like ''RUBENS SERRA%'') and '
   else
   begin
      dts.CommandText := dts.CommandText + ' codgcpjescritorio in(';
      for i := 0 to lstCnpjs.Count - 1 do
      begin
            if i <> 0 then
            dts.CommandText := dts.CommandText + ', ';
         dts.CommandText := dts.CommandText + lstCnpjs.Strings[i];
      end;
      dts.CommandText := dts.CommandText + ') and ';
   end;
   dts.CommandText := dts.CommandText + ' gcpj = ' + gcpj + ' and b.idplanilha <> ' + IntToStr(idPlanilha) + ' and ' +
                      'fgcruzadogcpj <> 9 and a.idescritorio = b.idescritorio and b.idplanilha = c.idplanilha';
   dts.Parameters.ParamByName('dtimportacao').Value := StrToDate(FormatDateTime('dd/mm/yyyy', dtPlanilha));
   dts.Open;
   result := 1;
end;

function TdmHonorarios.ObtemUltimaDataCadastrada: TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select datadigitar from tbdatadigitar where anomesreferencia in (select max(anomesreferencia) from tbdatadigitar)';
   dts.Open;

   result := StrToDate(FormatDateTime('dd/mm/yyyy', dts.FieldByName('datadigitar').AsDateTime));
end;

function TdmHonorarios.ObtemUltimoAnoMesCadastrado: string;
begin
   dts.Close;
   dts.CommandText := 'select max(anomesreferencia) as ultimoAnoMes from tbdatadigitar';
   dts.Open;

   result := dts.FieldByName('ultimoAnoMes').AsString;
end;

procedure TdmHonorarios.ObtemAtosPendentes(idplanilha: integer);
begin
   dtsAtosPendentes.Close;
   dtsAtosPendentes.CommandText := 'SELECT * FROM TBPLANILHASATOS WHERE IDPLANILHA = ' + IntToStr(idplanilha) + ' and fgcruzadogcpj not in (6,9)';
   dtsAtosPendentes.Open;
end;

procedure TdmHonorarios.CriaColunaValorBase;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE TBPLANILHASATOS ADD valorbase Float';
   adoCmd.Execute;
end;

function TdmHonorarios.ObtemDataImportacaoPlanilha(
  idPlanilha: integer): TDateTime;
begin
   dts.Close;
   dts.Parameters.Clear;
   dts.COmmandText := 'select dtImportacao from tbplanilhas where idPlanilha = ' + IntToStr(idPlanilha);
   dts.Open;
   result := dts.FieldByName('dtImportacao').AsDatetime;
end;

function TdmHonorarios.TemConsolidacaoPagaNoSistema(gcpj: string): boolean;
begin
   dts.Close;
   dts.CommandText := 'select idplanilha, anomesreferencia, gcpj, fgcruzadogcpj, tipoandamento ' +
                      'from tbplanilhasatos ' +
                  'where gcpj = ' + gcpj + ' and tipoandamento = ''CONSOLIDACAO DE PROPRIEDADE'' and ' +
                      'fgcruzadogcpj <> 9';
   dts.open;

   result := (dts.RecordCount > 0);
end;

procedure TdmHonorarios.CriaTabelaValoresNaoAtualizar;
begin
   adoCmd.Parameters.CLear;
   adoCmd.CommandText  := 'CREATE TABLE tbvaloresnaoatualizar (' +
                          '   identificador integer not null, ' +
                          '   tipoandamento varchar(40) not null, ' +
                          '   valorpagar float not null default 0, ' +
                          '   CONSTRAINT PK_tbvaloresnaoatualizar PRIMARY KEY (' +
                          '        identificador, ' +
                          '        tipoandamento ' +
                          '   )' +
                          ')';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tbvaloresnaoatualizar ADD FOREIGN KEY(identificador) REFERENCES tbtiposnaoatualizar (identificador)';
   adoCmd.Execute;

end;

procedure TdmHonorarios.CriaTabelaTiposNaoAtualizar;
begin
   adoCmd.Parameters.CLear;
   adoCmd.CommandText  := 'CREATE TABLE tbtiposnaoatualizar (' +
                          '   identificador Autoincrement(1,1) not null, ' +
                          '   datacadastro datetime not null, ' +
                          '   CONSTRAINT PK_tbtiposnaoatualizar PRIMARY KEY (' +
                          '      identificador' +
                          '   )' +
                          ')';
   adocmd.Execute;

end;

procedure TdmHonorarios.InsereTiposNaoAtualizarPadrao;
begin
   adoDts.close;
   adoDts.Parameters.Clear;
   adoDts.CommandText := 'select identificador from tbtiposnaoatualizar where identificador = 1';
   adoDts.Open;

   if adoDts.RecordCount = 0 then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbtiposnaoatualizar(datacadastro)' +
                            ' Values(:hoje)';
      adoCmd.Parameters.ParamByName('hoje').value := date;
      adoCmd.Execute;
   end;
end;

procedure TdmHonorarios.InsereValoresNaoAtualizarPadrao;
begin
   adoDTs.Close;
   adoDts.CommandText := 'SELECT 1 FROM tbvaloresnaoatualizar where identificador = 1 and tipoandamento = ''HONORARIOS INICIAIS''';
   adoDTs.Open;
   if adoDts.RecordCount = 0 then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbvaloresnaoatualizar (identificador, tipoandamento, valorpagar) ' +
                            'Values(1, ''HONORARIOS INICIAIS'', 56.00)';
      adoCmd.Execute;
   end;

   adoDTs.Close;
   adoDts.CommandText := 'SELECT 1 FROM tbvaloresnaoatualizar where identificador = 1 and tipoandamento = ''HONORARIOS FINAIS''';
   adoDTs.Open;
   if adoDts.RecordCount = 0 then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbvaloresnaoatualizar (identificador, tipoandamento, valorpagar) ' +
                            'Values(1, ''HONORARIOS FINAIS'', 56.00)';
      adoCmd.Execute;
   end;

   adoDTs.Close;
   adoDts.CommandText := 'SELECT 1 FROM tbvaloresnaoatualizar where identificador = 1 and tipoandamento = ''PREPOSTO''';
   adoDTs.Open;
   if adoDts.RecordCount = 0 then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbvaloresnaoatualizar (identificador, tipoandamento, valorpagar) ' +
                            'Values(1, ''PREPOSTO'', 64.00)';
      adoCmd.Execute;
   end;

   adoDTs.Close;
   adoDts.CommandText := 'SELECT 1 FROM tbvaloresnaoatualizar where identificador = 1 and ' +
                         '(tipoandamento = ''TUTELA'' or tipoandamento = ''RECURSO'' OR tipoandamento = ''AJUIZAMENTO'')';
   adoDTs.Open;
   if adoDts.RecordCount = 0 then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbvaloresnaoatualizar (identificador, tipoandamento, valorpagar) ' +
                            'Values(1, ''RECURSO'', 470.00)';
      adoCmd.Execute;

      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbvaloresnaoatualizar (identificador, tipoandamento, valorpagar) ' +
                            'Values(1, ''AJUIZAMENTO'', 470.00)';
      adoCmd.Execute;

      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'INSERT INTO tbvaloresnaoatualizar (identificador, tipoandamento, valorpagar) ' +
                            'Values(1, ''TUTELA'', 470.00)';
      adoCmd.Execute;
   end;
end;

procedure TdmHonorarios.dtsTiposNaoAtualizarBeforeDelete(
  DataSet: TDataSet);
begin
   if dataSet.FieldByName('identificador').AsInteger = 1 then
   begin
      ShowMessage('Não pode excluir o tipo padrão');
      Abort;
      exit;
   end;
end;

procedure TdmHonorarios.ObtemTiposNaoAtualizarCadastrados;
begin
   dtsTiposNaoAtualizar.Close;
   dtsTiposNaoAtualizar.CommandText := 'select identificador, datacadastro, datadoato from tbtiposnaoatualizar order by identificador';
   dtsTiposNaoAtualizar.Open;
end;

procedure TdmHonorarios.dsTiposNaoAtualizarDataChange(Sender: TObject;
  Field: TField);
begin
   dtsValoresNaoAtualizar.Close;

   frmCadastraValoresNaoAtualizar.BitBtn8.Enabled := (dtsTiposNaoAtualizar.RecordCount <> 0);

   if dtsTiposNaoAtualizar.RecordCount = 0 then
      exit;

   if dsTiposNaoAtualizar.State = dsInsert then
      exit;

   frmCadastraValoresNaoAtualizar.BitBtn3.Enabled := (not (dtsTiposNaoAtualizar.FieldByName('identificador').AsInteger = 1));

   dtsValoresNaoAtualizar.CommandText := 'select identificador, tipoandamento, valorpagar from tbvaloresnaoatualizar ' +
                                         'where identificador = ' + dtsTiposNaoAtualizar.FieldByName('identificador').AsString;
   dtsValoresNaoAtualizar.Open;
end;

procedure TdmHonorarios.dtsTiposNaoAtualizarNewRecord(DataSet: TDataSet);
begin
   dtsTiposNaoAtualizar.FieldByName('datacadastro').Value := date;
end;

procedure TdmHonorarios.dtsValoresNaoAtualizarNewRecord(DataSet: TDataSet);
begin
   dtsValoresNaoAtualizar.FieldByName('identificador').AsInteger := dtsTiposNaoAtualizar.FieldByName('identificador').AsInteger;
end;

procedure TdmHonorarios.dtsValoresNaoAtualizarBeforePost(
  DataSet: TDataSet);
begin
   if (dataSet.FieldByName('tipoandamento').AsString <> 'HONORARIOS INICIAIS') and
      (dataSet.FieldByName('tipoandamento').AsString <> 'HONORARIOS FINAIS') and
      (dataSet.FieldByName('tipoandamento').AsString <> 'PREPOSTO') and
      (dataSet.FieldByName('tipoandamento').AsString <> 'TUTELA') and
      (dataSet.FieldByName('tipoandamento').AsString <> 'RECURSO') and
      (dataSet.FieldByName('tipoandamento').AsString <> 'AJUIZAMENTO') then
   begin
      ShowMessage('Tipo de andamento inválido');
      dtsValoresNaoAtualizar.Tag := 0;
      abort;
      exit;
   end;

   if dataSet.FieldByName('valorpagar').AsFloat <= 0.00 then
   begin
      SHowMessage('valor inválido');
      dtsValoresNaoAtualizar.Tag := 0;
      abort;
      exit;
   end;
   dtsValoresNaoAtualizar.Tag := 1;
end;

procedure TdmHonorarios.ExcluiTiposNaoAtualizar(identificador: Integer);
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'delete from tbtiposnaoatualizar where identificador = ' + IntToStr(identificador);
   adoCmd.Execute;
end;

procedure TdmHonorarios.ExcluiValoresNaoAtualizarFK(
  identificador: Integer);
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'delete from tbvaloresnaoatualizar where identificador = ' + IntToStr(identificador);
   adoCmd.Execute;
end;

procedure TdmHonorarios.ExcluiValoresNaoAtualizarPK(identificador: Integer;
  tipoandamento: string);
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'delete from tbvaloresnaoatualizar where identificador = ' + IntToStr(identificador) + ' and ' +
                         'tipoandamento = ''' + tipoandamento + '''';
   adoCmd.Execute;
end;

procedure TdmHonorarios.CriaColunaIdentificador;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tbescritorios ADD identificador integer default 1';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'Update tbescritorios set identificador = 1 where identificador is null';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'create index idx1 on tbescritorios (cnpjescritorio, identificador)';
   adoCmd.Execute;

end;

procedure TdmHonorarios.ObtemEscritoriosAssociadosAoTipo(
  identificador: integer);
begin
   dtsEscritoriosIn.Close;
   dtsEscritoriosIn.CommandText := 'select a.idescritorio, nomeescritorio, cnpjescritorio ' +
                                   'from tbescritorios a, tblinkescritoriosvalores b ' +
                                   'where b.identificador = ' + IntToStr(identificador) + ' and a.idescritorio = b.idescritorio ' +
                                   'order by nomeescritorio';
   dtsEscritoriosIn.Open;
end;

procedure TdmHonorarios.ObtemEscritoriosNaoAssociadosAoTipo(
  identificador: integer; nomeescritorio: string);
begin
   dtsEscritoriosOut.Close;
   dtsEscritoriosOut.CommandText := 'select idescritorio, nomeescritorio, cnpjescritorio ' +
                                   'from tbescritorios ' +
                                   'where idescritorio not in (select idescritorio from tblinkescritoriosvalores where identificador = ' + IntToStr(identificador) + ') ';
   if nomeescritorio <> '' then
      dtsEscritoriosOut.CommandText := dtsEscritoriosOut.CommandText + ' and nomeescritorio like ''%' +nomeescritorio + '%''';
   dtsEscritoriosOut.CommandText := dtsEscritoriosOut.CommandText +'order by nomeescritorio';
   dtsEscritoriosOut.Open;
end;

procedure TdmHonorarios.MarcaEscritorioIn(identificador,
  idEscritorio: integer);
begin
   adoCmd.Parameters.Clear;
//   adoCmd.CommandText := 'update tbescritorios set identificador = ' + IntToStr(identificador) + ' ' +
//                         'where idescritorio = ' + IntToStr(idEscritorio);
   adoCmd.CommandText := 'INSERT INTO tblinkescritoriosvalores (identificador, idescritorio) ' +
                         'values(' + IntToStr(identificador) + ', ' + inttostr(idEscritorio) + ')';
   adoCmd.Execute;
end;

procedure TdmHonorarios.MarcaEscritorioOut(identificador, idEscritorio: integer);
begin
   adoCmd.Parameters.Clear;
//   adoCmd.CommandText := 'update tbescritorios set identificador = 1 ' +
//                         'where idescritorio = ' + IntToStr(idEscritorio);
   adoCmd.CommandText := 'DELETE FROM tblinkescritoriosvalores ' +
                         'where identificador = ' + IntToStr(identificador) + ' and idescritorio = ' + inttostr(idEscritorio);
   adoCmd.Execute;
end;

function TdmHonorarios.ObtemValorSemReajuste(cnpjEscritorio,
  tipoAndamento: string; dataDoAto: TDateTime): double;
var
   identificador : integer;
begin
   //localiza a data do ato
   adoDts.Close;
   adoDts.CommandText := 'select identificador, datadoato ' +
                         'from tbtiposnaoatualizar ' +
                         'order by datadoato desc';
   adoDts.Open;
   if adoDts.Eof then
   begin
      result := 0.00;
      exit;
   end;

   identificador := 0;

   while not adoDts.Eof do
   begin
      if FormatDateTime('yyyy/mm/dd', dataDoAto) < FormatDateTime('yyyy/mm/dd', adoDts.FieldByName('datadoato').AsDateTime) then
      begin
         adoDts.Next;
         continue;
      end;
      
      identificador := adoDts.FieldByName('identificador').AsInteger;
      Dts.Close;
      Dts.CommandText := 'select valorpagar ' +
                            'from tbescritorios a, tblinkescritoriosvalores b, tbvaloresnaoatualizar c ' +
                            'where cnpjescritorio = ''' + cnpjEscritorio + ''' and tipoandamento like ''%' + tipoAndamento + '%'' and ' +
                            'b.identificador = ' + IntToStr(identificador) + ' and ' +
                            'a.idescritorio = b.idescritorio and b.identificador = c.identificador';
      Dts.Open;

      if Dts.RecordCount = 0 then
      begin
         adoDts.next;
         continue;
      end;
      result := dts.FieldByName('valorpagar').AsFloat;
      exit;
   end;
   result := 0;
end;

procedure TdmHonorarios.CriaTabelaAdvogadosInternos;
begin
   adoCmd.Parameters.CLear;
   adoCmd.CommandText  := 'CREATE TABLE tbadvogadosinternos (' +
                          '   codigofuncional integer not null, ' +
                          '   nomeadvogado varchar(60) not null, ' +
                          '   CONSTRAINT PK_tbadvogadosinternos PRIMARY KEY (' +
                          '        codigofuncional ' +
                          '   )' +
                          ')';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'create index idx1 on tbadvogadosinternos (nomeadvogado)';
   adoCmd.Execute;
end;

procedure TdmHonorarios.ObtemAdvogadosCadastrados(codfuncional: integer);
begin
   dtsAdvogados.Close;
   dtsAdvogados.CommandText := 'select codigofuncional, nomeadvogado ' +
                               'from tbadvogadosinternos ';
   if codfuncional <> 0 then
      dtsAdvogados.CommandText := dtsAdvogados.CommandText + ' Where  codigofuncional = ' + IntToStr(codfuncional);

   dtsAdvogados.CommandText := dtsAdvogados.CommandText + ' order by nomeadvogado';
   dtsAdvogados.Open;
end;

procedure TdmHonorarios.CriaColunaDataDoAto_TiposNaoAtualizar;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tbtiposnaoatualizar ADD datadoato datetime';
   adoCmd.Execute;

   adoDts.Close;
   adoDts.CommandText := 'select identificador from tbtiposnaoatualizar where identificador=1';
   adoDTs.open;
   if not adoDts.Eof then
   begin
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'UPDATE tbtiposnaoatualizar SET datadoato = :data where identificador = 1';
      adoCmd.Parameters.ParamByName('data').Value := EncodeDate(2014, 05, 01);
      adoCmd.Execute;
   end;
end;

function TdmHonorarios.ConfiguracaoSemReajusteEstaOK: boolean;
begin
   adoDts.Close;
   adoDts.Parameters.Clear;
   adoDts.CommandText := 'SELECT count(*) as total from tbtiposnaoatualizar where datadoato is not null';
   adoDts.Open;

   result := (adoDts.FieldByName('total').AsInteger <> 0);
end;

procedure TdmHonorarios.dtsTiposNaoAtualizarBeforePost(DataSet: TDataSet);
begin
   if dataSet.FieldByName('datadoato').isnull then
   begin
      ShowMessage('Data do ato inválida');
      dtsTiposNaoAtualizar.Tag := 0;
      abort;
      exit;
   end;
   dtsTiposNaoAtualizar.Tag := 1;
end;

procedure TdmHonorarios.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index dupl_idx1 on tbplanilhasatos(tipoandamento,gcpj, fgcruzadogcpj)';
      adoCmd.Execute;
   except
   end;

   try
      adoCmd.CommandText := 'create index dupl_dx1 on tbplanilhas(dtImportacao)';
      adoCmd.Execute;
   except
   end;

   try
      adoCmd.CommandText := 'create index dupl_dx1 on tbescritorios(nomeescritorio, codgcpjescritorio)';
      adoCmd.Execute;
   except
   end;
end;

procedure TdmHonorarios.ResetFlagsDaPlanilha(idPlanilha: integer);
begin
   adoCMd.Parameters.Clear;
   adoCmd.CommandText := 'update tbplanilhas set fgvalidada = 0 where idplanilha = ' + IntToStr(idPlanilha);
   adoCmd.Execute;

   adoCMd.Parameters.Clear;
   adoCmd.CommandText := 'update tbplanilhasatos set fgcruzadogcpj = 0  where idplanilha = ' + IntToStr(idPlanilha);
   adoCmd.Execute;

   adoCMd.Parameters.Clear;
   adoCmd.CommandText := 'delete from tbcontrolenotas where idplanilha = ' + IntToStr(idPlanilha);
   adoCmd.Execute;

   adoCMd.Parameters.Clear;
   adoCmd.CommandText := 'delete from tbocorrenciasprocessamento where idplanilha = ' + IntToStr(idPlanilha);
   adoCmd.Execute;
end;

procedure TdmHonorarios.ExcluiColunaIdentificador;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'drop index idx1 on tbescritorios';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tbescritorios DROP COLUMN identificador';
   adoCmd.Execute;
end;

procedure TdmHonorarios.CriaTabelaLinkEscritoriosValores;
begin
   adoCmd.Parameters.CLear;
   adoCmd.CommandText  := 'CREATE TABLE tblinkescritoriosvalores (' +
                          '   identificador integer not null, ' +
                          '   idescritorio integer not null, ' +
                          '   CONSTRAINT PK_tblinkescritoriosvalores PRIMARY KEY (' +
                          '        identificador, ' +
                          '        idescritorio ' +
                          '   )' +
                          ')';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tblinkescritoriosvalores ADD FOREIGN KEY(identificador) REFERENCES tbtiposnaoatualizar (identificador)';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tblinkescritoriosvalores ADD FOREIGN KEY(idescritorio) REFERENCES tbescritorios (idescritorio)';
   adoCmd.Execute;

end;

procedure TdmHonorarios.ObtemPlanilhasEnviarGcpj(
  var dtsRetorno: TAdoDataset; const filtro: integer);
begin
   dtsRetorno.Close;
   dtsRetorno.CommandText := 'SELECT cnpjescritorio, nomeescritorio, anomesreferencia, fgimportada, fgfinalizada, dtimportacao, a.idescritorio, idplanilha, sequencia, nomedigitar, codgcpjescritorio, ' +
                             'fgpagaribi,  dataatosibi, fgenviadogcpj, dtenviogcpj, fgretornogcpj ' +
                             'from tbescritorios a, tbplanilhas b ' +
                             'where fgimportada=1 and fgvalidada = 1 and a.idescritorio = b.idescritorio ';
   case filtro of
      1 : dtsRetorno.CommandText := dtsRetorno.CommandText + 'and (fgenviadogcpj = 0 or fgenviadogcpj is null) ';
      2 : dtsRetorno.CommandText := dtsRetorno.CommandText + 'and fgenviadogcpj = 1  ';
   end;
   dtsRetorno.CommandText := dtsRetorno.CommandText + 'order by dtimportacao desc, anomesreferencia desc, a.idescritorio ';
   dtsRetorno.Open;
end;

procedure TdmHonorarios.CriaColunafgenviadogcpj_tbplanilhas;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tbplanilhas ADD fgenviadogcpj integer default 0';
   adoCmd.Execute;
end;

procedure TdmHonorarios.CriaColunadtenviogcpj_tbplanilhas;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tbplanilhas ADD dtenviogcpj datetime';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'Update tbplanilhas ' +
                         ' Set fgenviadogcpj=0';
   adoCmd.Execute;
end;

procedure TdmHonorarios.CriaColunafgretornogcpj_tbplanilhas;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE tbplanilhas ADD fgretornogcpj integer default 0';
   adoCmd.Execute;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'Update tbplanilhas ' +
                         ' Set fgretornogcpj=0';
   adoCmd.Execute;
end;

function TdmHonorarios.ObtemSomaJaPagaADescontar(gcpj,
  tipodeandamento: string; idplanilha: integer; const recuperacaoFinal : boolean = false): double;
begin
   dts.Close;
   dts.Parameters.Clear;
   dts.CommandText := 'Select Sum(valorcorrigido) as valorPago from tbplanilhasatos ' +
                      'where idplanilha = ' + IntToStr(idplanilha) + ' and gcpj = ' + gcpj + ' and fgcruzadogcpj <> 9 ';
   if recuperacaoFinal then
   begin
      if (tipodeandamento = 'ACORDO') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (tipodeandamento = 'ENTREGA AMIGAVEL') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if tipodeandamento = 'CONSOLIDACAO DE PROPRIEDADE' then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' +
                            '''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'', ' + //''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'',
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (tipodeandamento = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' + //''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (tipodeandamento = 'HONORARIOS DE VENDA') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''ADJUDICACAO/ARREMATACAO'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (tipodeandamento = 'ADJUDICACAO/ARREMATACAO') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (tipodeandamento = 'LEVANTAMENTO DE ALVARA') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'')'
      else
      if (tipodeandamento = 'AJUIZAMENTO') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'', ''CONSOLIDACAO DE PROPRIEDADE'', ''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ' +
                            '''ENTREGA AMIGAVEL'', ''LEVANTAMENTO DE ALVARA'', ''ADJUDICACAO/ARREMATACAO'')'
      else
      if (tipodeandamento = 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE') or (tipodeandamento = 'BUSCA E APREENSÃO/REINTEGRAÇÃO DE POSSE') then
         dts.CommandText := dts.CommandText + ' and tipoandamento in(''ACORDO'')';
   end
   else
      dts.CommandText := dts.CommandText + 'and tipoandamento = ''' + tipodeandamento + '''';

   dts.CommandText := dts.CommandText + ' group by idplanilha, gcpj, tipoandamento';
   dts.Open;

   result := dts.FieldByName('valorpago').AsFloat;
end;

procedure TdmHonorarios.RemoveOcorrenciasDaPlanilha(idplanilha: integer);
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'DELETE FROM tbocorrenciasprocessamento where idplanilha = ' + IntToStr(idplanilha);
   adoCmd.Execute;
end;

procedure TdmHonorarios.CriaColunaCodExFuncionario;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'ALTER TABLE TBPLANILHASRECLAMADAS ADD codexfuncionario integer default 0';
   adoCmd.Execute;
end;

procedure TdmHonorarios.IndexaTabelaReclamadas;
begin
   try
      adoCmd.Parameters.Clear;
      adoCmd.CommandText := 'create index idx1 on tbplanilhasreclamadas(idescritorio, idplanilha, anomesreferencia, sequencia, linhaplanilha,codigoreclamada, ' +
                            'codigopessoaexterna)';
      adoCmd.Execute;
   except
   end;

end;

procedure TdmHonorarios.CadastraDeParaEmpresaGrupo(empresaDe: integer;
  nomeEmpresaDe: string; empresaPara: integer; nomeEmpresaPara,
  tipoProcesso: string);
begin
   dts.Close;
   dts.CommandText := 'SELECT * FROM TBDEPARAEMPRESAGRUPO WHERE CODIGO_EMPRESA_GRUPO_DE = ' + IntToStr(empresaDe) + ' and ' +
                      '(tipo_processo = ''' + tipoProcesso + ''' or tipo_processo = ''Z'') and  codigo_empresa_grupo_para=' + IntToStr(empresaPara);// + ' and nome_empresa_grupo_para = ''' + nomeempresapara + '''';
   dts.Open;

   if not dts.Eof then
      exit;

   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'INSERT INTO TBDEPARAEMPRESAGRUPO(codigo_empresa_grupo_de, nome_empresa_grupo_de, codigo_empresa_grupo_para, nome_empresa_grupo_para, tipo_processo) ' +
                         'values(' + IntToStr(empresaDe) + ', ''' + nomeEmpresaDe + ''', ' + IntToStr(empresaPara) + ', ''' + nomeEmpresaPara + ''', ''' + tipoProcesso + ''')';
   adoCmd.Execute;
end;

procedure TdmHonorarios.LimpaTabelaDePara;
begin
   adoCmd.Parameters.Clear;
   adoCmd.CommandText := 'DELETE FROM TBDEPARAEMPRESAGRUPO WHERE codigo_empresa_grupo_de > 0 ';
   adoCmd.Execute;
end;

procedure TdmHonorarios.ReindexTabelaDePara(tipo: integer);
begin
   if (tipo = 0) or (tipo = 9) then
   begin
      try
         adoCmd.Parameters.Clear;
         adoCmd.CommandText := 'DROP INDEX PrimaryKey ON TBDEPARAEMPRESAGRUPO';
         adoCmd.Execute;
      except
      end;

      try
         adoCmd.Parameters.Clear;
         adoCmd.CommandText := 'DROP INDEX IDX1 ON TBDEPARAEMPRESAGRUPO';
         adoCmd.Execute;
      except
      end;
   end;

   if (tipo = 1) or (tipo = 9) then
   begin
      try
         adoCmd.Parameters.Clear;
         adoCmd.CommandText := 'create unique index idx1 on TBDEPARAEMPRESAGRUPO(codigo_empresa_grupo_de, codigo_empresa_grupo_para, tipo_processo)';
         adoCmd.Execute;
      except
      end;
  end;
end;

procedure TdmHonorarios.executabackup(users: string;tempo:integer;dataatual:Tdatetime;FormBackup:TForm);

Var
   ini : TInifile;
   dataBaseName : string;
   Access: OleVariant;
   datafinal  : TDatetime;

   i: integer;
   OutraData: TDatetime;
begin



   Application.ProcessMessages;
   for i:= 0 to (FormBackup.ComponentCount -1) do
     begin
      if FormBackup.Components[i] is TImage then
       Begin
        TImage(FormBackup.Components[i]).Picture.LoadFromFile(ExtractFilePath(Application.exename) + 'loading.gif');
//        TImage(FormBackup.Components[i]).Repaint;
//        TImage(FormBackup.Components[i]).Refresh;
//        TImage(FormBackup.Components[i]).Show;
//        Application.ProcessMessages;
       end;

     end;
     Application.ProcessMessages;





   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   dataBaseName := ExtractFilePath(ini.readstring('honorarios', 'databasename', ''));



  // faz a compactação do bak
   Application.ProcessMessages;
   dmHonorarios.adoConn.Connected := False;
   access := CreateOleObject ('DAO.DBEngine.36');

   Access.CompactDatabase(dataBaseName + 'honorarios_' + users + '.mdb', dataBaseName + 'honorarios_' + users + '_BAK_'+FormatDateTime('YYYYMMDD',date)+'.mdb');


   Application.ProcessMessages;

   for i:= 0 to (FormBackup.ComponentCount -1) do
     begin
      if FormBackup.Components[i] is TImage then
       Begin
        Application.ProcessMessages;
       end;

     end;


    //volta a conectar no BD
   dmHonorarios.adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + dataBaseName + 'honorarios_' + users + '.mdb' + ';Persist Security Info=True';


    for i:= 0 to (FormBackup.ComponentCount -1) do
     begin
      if FormBackup.Components[i] is TGauge then
       TGauge(FormBackup.Components[i]).progress := 65;
     end;


// Coleta os registros para deletar.
//   EsperaLiberacaoGravacao;
   try
      dts.Parameters.Clear;
      dts.CommandText := 'SELECT * FROM tbplanilhas WHERE tbplanilhas.dtimportacao < #' +DatetoStr(INCMONTH(dataatual,-tempo))+'#';



// EXCLUI O PAI e DEPOIS 0 FILHO   ==> TBPLANILHA e TBPLANILHAATOS
     dts.open;
     while not dts.Eof do
       begin
        Application.ProcessMessages;
        adoCmd.Parameters.Clear;
        adoCmd.CommandText := 'DELETE FROM tbplanilhasatos WHERE tbplanilhasatos.idplanilha = :ID';
        adoCmd.Parameters.ParamByName('ID').Value := dTS.FieldByName('idplanilha').AsString;
        adoCmd.Execute;
        dts.Delete;
       end;


// atualiza o gauge
    for i:= 0 to (FormBackup.ComponentCount -1) do
     begin
      if FormBackup.Components[i] is TGauge then
       TGauge(FormBackup.Components[i]).progress := 85;
     end;




   finally
//      LiberaGravacao;
      ini.Free;
   end;

end;

procedure TdmHonorarios.Timer1Timer(Sender: TObject);
begin
// Application.ProcessMessages;
// Sleep(6000);
end;

end.
