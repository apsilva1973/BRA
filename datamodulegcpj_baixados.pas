unit datamodulegcpj_baixados;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms;

type
  Tdmgcpj_baixados = class(TDataModule)
    adoConn: TADOConnection;
    adoCmd: TADOCommand;
    dts: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
    function ObtemReclamada(gcpj: string; var codjuncao, nomejuncao: string):integer;
    procedure ObtemDadosDaBaixa(gcpj: string;
         var motivoBaixa: string; var dataBaixa: tdatetime; var valorbaixa: double);
    function ObtemNomeEnvolvido(gcpj: string): string;
    function ObtemCodEmpresaGrupoCivel(gcpj: string; var pdts : TAdoDataset): integer;
    function ObtemTipoSubtipo(gcpj: string; var pdts: TadoDataset): integer;
    function ProcessoExisteGcpj(gcpj: string): integer;
    function ObtemCodEmpresaGrupoTrabalhista(gcpj: string; var pdts : TAdoDataset):integer;
    function ObtemAreaProcesso(gcpj: string): string;
    function ObtemCodOrgProcesso(gcpj: string): integer;
    function ObtemNumeroDoProcesso(gcpj: string): string;
    function ObtemDataCadastro(gcpj: string): TDateTime;
    function ObtemDataBaixa(gcpj: string): TDateTime;
   end;

var
  dmgcpj_baixados: Tdmgcpj_baixados;

implementation

{$R *.dfm}

procedure Tdmgcpj_baixados.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index idx1 on _Civel_Baixas(bradesco)';
      adoCmd.Execute;
   except
   end;
end;

procedure Tdmgcpj_baixados.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('gcpj_base_baixados', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

function Tdmgcpj_baixados.ObtemAreaProcesso(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select ''CI'' as area from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := '';
      exit;
   end;

   result := dts.FieldByName('area').AsString;;

end;

function Tdmgcpj_baixados.ObtemCodEmpresaGrupoCivel(gcpj: string; var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select Gestor_Empresa as codjunemp, codjun, nomjun, codjunagrup, Gestor_Codigo as codjuncon, Gestor_Descricao as nomjunco from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      result := 0
   else
      result := 1;
end;

function Tdmgcpj_baixados.ObtemCodEmpresaGrupoTrabalhista(
  gcpj: string; var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select Gestor_Codigoemp, Gestor_Empresa as codjunemp, codjun, nomjun, codjunagrup, Gestor_Codigo as codjuncon, Gestor_Descricao as nomjuncon from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      result := 0
   else
      result := 1;
end;

function Tdmgcpj_baixados.ObtemCodOrgProcesso(gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select codOrg from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
      result := 0
   else
      result := dts.FieldByName('codorg').AsInteger;
end;

procedure Tdmgcpj_baixados.ObtemDadosDaBaixa(gcpj: string;
  var motivoBaixa: string; var dataBaixa: tdatetime;
  var valorbaixa: double);
begin
   dts.Close;
   dts.CommandText := 'select datbax, nomBax, VlrBaxContabil from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;

   if not dts.eof then
   begin
      if Not dts.FieldByName('datbax').IsNull then
      begin
         dataBaixa := dts.FieldByName('datbax').AsDateTime;
         motivoBaixa := dts.FieldByName('nombax').AsString;
         valorbaixa := dts.FieldByName('VlrBaxContabil').AsFloat;
      end;
      exit;
   end;
end;

function Tdmgcpj_baixados.ObtemDataBaixa(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select DatBax from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := 0;
      exit;
   end;
   result := dts.FieldByName('DatBax').AsDateTime;
end;

function Tdmgcpj_baixados.ObtemDataCadastro(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select DatCad from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := 0;
      exit;
   end;
   result := dts.FieldByName('DatCad').AsDateTime;
end;

function Tdmgcpj_baixados.ObtemNomeEnvolvido(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select envolvido from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
      result := ''
   else
      result := dts.FieldByName('envolvido').AsString;

end;

function Tdmgcpj_baixados.ObtemNumeroDoProcesso(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select processo from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if (dts.Eof) or (dts.FieldByname('processo').IsNull) or (dts.FieldByName('processo').AsString = '0') then
      result := ''
   else
      result := dts.FieldByName('processo').AsString;

end;

function Tdmgcpj_baixados.ObtemReclamada(gcpj: string; var codjuncao,
  nomejuncao: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select Gestor_Codigo, Gestor_Descricao, nomjun from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := 0;
      exit;
   end;

   codjuncao := dts.FieldByName('Gestor_Codigo').AsString;
   nomejuncao := dts.FieldByName('Gestor_Descricao').AsString;
   result := 1;
end;

function Tdmgcpj_baixados.ObtemTipoSubtipo(gcpj: string; var pdts: TadoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select codacao, nomacao, codsub, nomsub, processo, coddejur from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      result := 0
   else
      result := 1;
end;

function Tdmgcpj_baixados.ProcessoExisteGcpj(gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select bradesco from _Civel_Baixas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then //não encontrou
      Result := 0
   else
      result := 1;

end;

end.
