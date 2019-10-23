unit datamodulegcpj_recuperados;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, datamodulegcpj_baixados;

type
  Tdmgcpj_recuperados = class(TDataModule)
    adoConn: TADOConnection;
    dts: TADODataSet;
    adoCmd: TADOCommand;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function ObtemReclamada(gcpj: string; var codjuncao, nomejuncao: string):integer;
    procedure CreateIndex;
    procedure ObtemDadosDaBaixa(gcpj: string; var motivoBaixa: string; var dataBaixa: tdatetime; var valorbaixa: double);
    function ObtemNomeEnvolvido(gcpj: string):string;
    function ObtemCodEmpresaGrupoCivel(gcpj: string; var pdts : TAdoDataset):integer;
    function ObtemCodEmpresaGrupoTrabalhista(gcpj: string; var pdts : TAdoDataset):integer;
    function ObtemTipoSubtipo(gcpj: string; var pdts : TAdoDataset):integer;
    function ProcessoExisteGcpj(gcpj: string):integer;
    function ObtemAreaProcesso(gcpj: string):string;
    function ObtemCodOrgProcesso(gcpj: string):integer;
    function ObtemNumeroDoProcesso(gcpj: string):string;
    function ObtemDataCadastro(gcpj: string) : TDateTime;
    function ObtemDataBaixa(gcpj: string) : TDateTime;
  end;

var
  dmgcpj_recuperados: Tdmgcpj_recuperados;

implementation

{$R *.dfm}

{ Tdmgcpj_compartilhado }

function Tdmgcpj_recuperados.ObtemReclamada(gcpj: string;
  var codjuncao, nomejuncao: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select codjunCon, nomjuncon, nomjun from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := dmgcpj_baixados.ObtemReclamada(gcpj, codjuncao, nomejuncao);
      //result := 0;
      exit;
   end;

   codjuncao := dts.FieldByName('codjuncon').AsString;
   nomejuncao := dts.FieldByName('nomjuncon').AsString;
   result := 1;
end;

procedure Tdmgcpj_recuperados.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('gcpj_base_recuperados', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

procedure Tdmgcpj_recuperados.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index idx1 on _00_Dados_Processos_RC (bradesco)';
      adoCmd.Execute;
   except
   end;
end;

procedure Tdmgcpj_recuperados.ObtemDadosDaBaixa(gcpj: string;
  var motivoBaixa: string; var dataBaixa: tdatetime; var valorbaixa: double);
begin
   dts.Close;
   dts.CommandText := 'select datbax, nomBax, vlrbax from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;

   if not dts.eof then
   begin
      if Not dts.FieldByName('datbax').IsNull then
      begin
         dataBaixa := dts.FieldByName('datbax').AsDateTime;
         motivoBaixa := dts.FieldByName('nombax').AsString;
         valorbaixa := dts.FieldByName('vlrbax').AsFloat;
      end;
      exit;
   end;

   dmgcpj_baixados.ObtemDadosDaBaixa(gcpj, motivoBaixa, dataBaixa, valorbaixa);
end;

function Tdmgcpj_recuperados.ObtemNomeEnvolvido(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select envolvido from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
      result := dmgcpj_baixados.ObtemNomeEnvolvido(gcpj)
//      result := ''
   else
      result := dts.FieldByName('envolvido').AsString;
end;

function Tdmgcpj_recuperados.ObtemCodEmpresaGrupoCivel(gcpj: string;var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select codjunemp, codjun, nomjun, codjunagrup, codjuncon, nomjuncon from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      result := dmgcpj_baixados.ObtemCodEmpresaGrupoCivel(gcpj, pdts)
   else
      result := 1;
end;

function Tdmgcpj_recuperados.ObtemTipoSubtipo(gcpj: string; var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select codacao, nomacao, codsub, nomsub, processo, coddejur from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      Result := 0
   else
      result := 1;
end;

function Tdmgcpj_recuperados.ProcessoExisteGcpj(gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select bradesco from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then //não encontrou
      result := dmgcpj_baixados.ProcessoExisteGcpj(gcpj)
   else
      result := 1;
end;

function Tdmgcpj_recuperados.ObtemCodEmpresaGrupoTrabalhista(
  gcpj: string; var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select codjunconemp, codjunemp, codjun, nomjun, codjunagrup, codjuncon, nomjuncon from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      result := dmgcpj_baixados.ObtemCodEmpresaGrupoTrabalhista(gcpj, pdts)
   else
      result := 1;
end;

function Tdmgcpj_recuperados.ObtemAreaProcesso(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select area from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := dmgcpj_baixados.ObtemAreaProcesso(gcpj);
      exit;
   end;

   result := dts.FieldByName('area').AsString;;
end;

function Tdmgcpj_recuperados.ObtemCodOrgProcesso(gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select codOrg from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
      result := dmgcpj_baixados.ObtemCodOrgProcesso(gcpj)
   else
      result := dts.FieldByName('codorg').AsInteger;
end;

function Tdmgcpj_recuperados.ObtemNumeroDoProcesso(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select processo from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if (dts.Eof) or (dts.FieldByname('processo').IsNull) or (dts.FieldByName('processo').AsString = '0') then
      result := dmgcpj_baixados.ObtemNumerodoProcesso(gcpj)
//      result := ''
   else
      result := dts.FieldByName('processo').AsString;
end;

function Tdmgcpj_recuperados.ObtemDataCadastro(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select DatCad from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := dmgcpj_baixados.ObtemDataCadastro(gcpj);
//      result := 0;
      exit;
   end;
   result := dts.FieldByName('DatCad').AsDateTime;
end;

function Tdmgcpj_recuperados.ObtemDataBaixa(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select DatBax from _00_Dados_Processos_RC ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := dmgcpj_baixados.ObtemDataBaixa(gcpj);
//      result := 0;
      exit;
   end;
   result := dts.FieldByName('DatBax').AsDateTime;
end;

end.
