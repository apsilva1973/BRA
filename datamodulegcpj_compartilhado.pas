unit datamodulegcpj_compartilhado;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, datamodulegcpj_baixados, datamodulegcpj_recuperados;

type
  Tdmgcpj_compartilhado = class(TDataModule)
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
  dmgcpj_compartilhado: Tdmgcpj_compartilhado;

implementation

{$R *.dfm}

{ Tdmgcpj_compartilhado }

function Tdmgcpj_compartilhado.ObtemReclamada(gcpj: string;
  var codjuncao, nomejuncao: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select codjunCon, nomjuncon, nomjun from _00_dados_processos ' +
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

procedure Tdmgcpj_compartilhado.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('gcpj_base_compartilhada', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

procedure Tdmgcpj_compartilhado.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index idx1 on _00_dados_processos (bradesco)';
      adoCmd.Execute;
   except
   end;
end;

procedure Tdmgcpj_compartilhado.ObtemDadosDaBaixa(gcpj: string;
  var motivoBaixa: string; var dataBaixa: tdatetime; var valorbaixa: double);
begin
   dts.Close;
   dts.CommandText := 'select datbax, nomBax, vlrbax from _00_dados_processos ' +
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

   dmgcpj_recuperados.ObtemDadosDaBaixa(gcpj, motivoBaixa, dataBaixa, valorbaixa);
end;

function Tdmgcpj_compartilhado.ObtemNomeEnvolvido(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select envolvido from _00_dados_processos ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
      result := dmgcpj_recuperados.ObtemNomeEnvolvido(gcpj)
   else
      result := dts.FieldByName('envolvido').AsString;
end;

function Tdmgcpj_compartilhado.ObtemCodEmpresaGrupoCivel(gcpj: string;var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select codjunemp, codjun, nomjun, codjunagrup, codjuncon, nomjuncon from _00_dados_processos ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      result := dmgcpj_recuperados.ObtemCodEmpresaGrupoCivel(gcpj, pdts)
   else
      result := 1;
end;

function Tdmgcpj_compartilhado.ObtemTipoSubtipo(gcpj: string; var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select codacao, nomacao, codsub, nomsub, processo, coddejur from _00_dados_processos ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
   begin
      Result := dmgcpj_baixados.ObtemTipoSubtipo(gcpj, pdts);
      if result <> 1 then
         Result := dmgcpj_recuperados.ObtemTipoSubtipo(gcpj, pdts)
      else
         result := 1;
   end
   else
      result := 1;
end;

function Tdmgcpj_compartilhado.ProcessoExisteGcpj(gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select bradesco from _00_Dados_Processos ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then //não encontrou
      Result := dmgcpj_recuperados.ProcessoExisteGcpj(gcpj)
   else
      result := 1;
end;

function Tdmgcpj_compartilhado.ObtemCodEmpresaGrupoTrabalhista(
  gcpj: string; var pdts : TAdoDataset): integer;
begin
   pdts.Close;
   pdts.Connection := dts.Connection;
   pdts.CommandText := 'select codjunconemp, codjunemp, codjun, nomjun, codjunagrup, codjuncon, nomjuncon from _00_dados_processos ' +
                      'where bradesco = ' + gcpj;
   pdts.Open;
   if pdts.Eof then
      result := dmgcpj_recuperados.ObtemCodEmpresaGrupoTrabalhista(gcpj, pdts)
   else
      result := 1;
end;

function Tdmgcpj_compartilhado.ObtemAreaProcesso(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select area from _00_dados_processos ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
   begin
      result := dmgcpj_recuperados.ObtemAreaProcesso(gcpj);
      exit;
   end;

   result := dts.FieldByName('area').AsString;;
end;

function Tdmgcpj_compartilhado.ObtemCodOrgProcesso(gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'select codOrg from _00_dados_processos ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if dts.Eof then
      result := dmgcpj_recuperados.ObtemCodOrgProcesso(gcpj)
   else
      result := dts.FieldByName('codorg').AsInteger;
end;

function Tdmgcpj_compartilhado.ObtemNumeroDoProcesso(gcpj: string): string;
begin
   dts.Close;
   dts.CommandText := 'select processo from _00_dados_processos ' +
                      'where bradesco = ' + gcpj;
   dts.Open;
   if (dts.Eof) or (dts.FieldByname('processo').IsNull) or (dts.FieldByName('processo').AsString = '0') then
      result := dmgcpj_recuperados.ObtemNumerodoProcesso(gcpj)
   else
      result := dts.FieldByName('processo').AsString;
end;

function Tdmgcpj_compartilhado.ObtemDataCadastro(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select DatCad from _00_dados_processos ' +
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

function Tdmgcpj_compartilhado.ObtemDataBaixa(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'select DatBax from _00_dados_processos ' +
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
