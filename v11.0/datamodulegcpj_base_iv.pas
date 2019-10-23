unit datamodulegcpj_base_iv;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, datamodulegcpj_base_ix;

type
  Tdmgcpj_base_iv = class(TDataModule)
    adoConn50: TADOConnection;
    dts50: TADODataSet;
    adoCmd50: TADOCommand;
    dts66: TADODataSet;
    adoConn66: TADOConnection;
    adoCmd66: TADOCommand;
    dts3T: TADODataSet;
    adoConn3T: TADOConnection;
    adoCmd3T: TADOCommand;
    dtsMesu: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function ObtemCodigoReclamada(codLink: string; var codReclamada, empresa, pessoaexterna, fgcodexfuncionario: integer; buscaAutor: boolean; GCPJ: string):integer;
    procedure CreateIndex;
    function ObtemNomeEmpresaGrupo(codEmpresaGrupo: integer):integer;
    function ObtemTipoDependencia(codreclamada, empresa: integer):string;
    function ObtemComplementoEmpresaGrupo(empresa, depdc: integer):integer;
    function ObtemUfProcesso(orgaojulgador:integer):string;
    function ObtemNomeOrgaoJulgador(orgaojulgador:integer):string;
    function ObtemCodigoReclamante(codLink: string; var codReclamada, empresa, pessoaexterna, fgcodexfuncionario: integer):integer;
    function ExistePrimeiraReclamadaDoGrupo(gcpj: string) : boolean;
  end;

var
  dmgcpj_base_iv: Tdmgcpj_base_iv;

implementation

{$R *.dfm}

{ Tdmgcpj_base_iv }

function Tdmgcpj_base_iv.ObtemCodigoReclamada(codLink: string;
  var codReclamada, empresa, pessoaexterna, fgcodexfuncionario: integer; buscaAutor: boolean; GCPJ: string): integer;
begin
   dts50.Close;
   dts50.CommandText := 'select CDEPDC, CEMPR_INC, CPSSOA_EXTER_BDSCO, CFUNC_EMPR_GRP, CFUNC_BDSCO from GCPJB050 where CENVDO_PROCS_JUDIC = ' + codLink;
   dts50.Open;

   if dts50.Eof then
   begin
      result := 0;
      exit;
   end;

   result := 1;

   codReclamada := dts50.FieldByName('CDEPDC').AsInteger;
   empresa := dts50.FieldByName('CEMPR_INC').AsInteger;
   pessoaexterna := dts50.FieldByName('CPSSOA_EXTER_BDSCO').AsInteger;

   fgcodexfuncionario := 0;
   if dts50.FieldByName('CFUNC_EMPR_GRP').AsInteger <> 0 then
   begin
      fgcodexfuncionario := dts50.FieldByName('CFUNC_EMPR_GRP').AsInteger;
      codReclamada := dts50.FieldByName('CFUNC_EMPR_GRP').AsInteger;;
   end
   else
   begin
     if dts50.FieldByName('CFUNC_BDSCO').AsInteger <> 0 then
     begin
        codReclamada := dts50.FieldByName('CFUNC_BDSCO').AsInteger;
        fgcodexfuncionario := dts50.FieldByName('CFUNC_BDSCO').AsInteger;
     end;

     if (fgcodexfuncionario = 0) and (buscaAutor) then
     begin
        //obtem o número do GCPJ
        if ExistePrimeiraReclamadaDoGrupo(gcpj) then
        begin
           fgcodExfuncionario := 999999;
            if pessoaexterna <> 0 then
               codReclamada := pessoaexterna
            else
               codReclamada := 999999;
        end;
     end;
   end;
end;

procedure Tdmgcpj_base_iv.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn50.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB050', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn66.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB066', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn3T.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB0J8_MESU9021_GCPJB0K1', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

procedure Tdmgcpj_base_iv.CreateIndex;
begin
   try
      adoCmd50.CommandText := 'create index idx1 on GCPJB050 (CENVDO_PROCS_JUDIC)';
      adoCmd50.Execute;
   except
   end;

   try
      adoCmd66.CommandText := 'create index idx1 on GCPJB066 (CORG_JULGA_PROCS)';
      adoCmd66.Execute;
   except
   end;


   try
      adoCmd3T.CommandText := 'create index idx1 on gcpjb0j8 (CAGRUP_BDSCO, CFUNC_GCPJ)';
      adoCmd3T.Execute;
   except
   end;

   try
      adoCmd50.CommandText := 'create index idx1 on gcpjb050 (CPSSOA_EXTER_BDSCO)';
      adoCmd50.Execute;
   except
   end;

   try
      adoCmd3T.CommandText := 'create index idx1 on MESU9021(CEMP, CDEPDC)';
      adoCmd3T.Execute;
   except
   end;

   try
      adoCmd3T.CommandText := 'create index idx1 on GCPJB0K1(CEMPR_INC_RUTIL, CDEPDC_RUTIL)';
      adoCmd3T.Execute;
   except
   end;

end;


function Tdmgcpj_base_iv.ObtemNomeEmpresaGrupo(codEmpresaGrupo: integer): integer;
begin
   dts3T.Close;
   dts3T.CommandText := 'select IAGRUP_BDSCO from gcpjb0j8 ' +
                      'where CAGRUP_BDSCO = ' + IntToStr(codEmpresaGrupo) + ' AND CFUNC_GCPJ = 90';
   dts3T.Open;
   if dts3T.eof then
      result := 0
   else
      result := 1;
end;

function Tdmgcpj_base_iv.ObtemTipoDependencia(
  codreclamada, empresa: integer): string;
begin
   dts3T.close;
   dts3T.CommandText := 'SELECT CTPO_DEPDC ' +
                      'FROM MESU9021 WHERE CEMP = ' + IntToStr(empresa) + ' AND CDEPDC='+ IntToStr(codreclamada);
   dts3T.Open;
   if dts3T.Eof then
      result := ''
   else
      result := dts3T.FieldByname('CTPO_DEPDC').AsString;
end;

function Tdmgcpj_base_iv.ObtemComplementoEmpresaGrupo(empresa,
  depdc: integer): integer;
begin
   dtsMesu.close;
   dtsMesu.CommandText := 'SELECT CTPO_DEPDC, ITPO_DEPDC ' +
                          'FROM MESU9021 WHERE CEMP = ' + IntToStr(empresa) + ' AND ' +
                          'CDEPDC='+ IntToStr(depdc);
   dtsMesu.Open;
   if dtsMesu.Eof then
   begin
      dtsMesu.close;
      dtsMesu.CommandText := 'SELECT CIDTFD_NATUZ_DEPDC AS CTPO_DEPDC, '''' AS ITPO_DEPDC ' +
                          'FROM GCPJB0K1 WHERE CEMPR_INC_RUTIL = ' + IntToStr(empresa) + ' AND ' +
                          'CDEPDC_RUTIL = '+ IntToStr(depdc);
      dtsMesu.Open;
   end;

   if dtsMesu.Eof then
      result := 0
   else
      result := 1;
end;

function Tdmgcpj_base_iv.ObtemUfProcesso(orgaojulgador: integer): string;
begin
   dts66.Close;
   dts66.CommandText := 'Select csgl_uf_julga from gcpjb066 where corg_julga_procs=' + IntToStr(orgaojulgador);
   dts66.Open;
   if dts66.eof then
      result := ''
   else
      result := dts66.fieldByName('csgl_uf_julga').AsString;
end;

function Tdmgcpj_base_iv.ObtemNomeOrgaoJulgador(
  orgaojulgador: integer): string;
begin
   dts66.Close;
   dts66.CommandText := 'Select RORG_JULGA_PROCS from gcpjb066 where corg_julga_procs=' + IntToStr(orgaojulgador);
   dts66.Open;
   if dts66.eof then
      result := ''
   else
      result := dts66.fieldByName('RORG_JULGA_PROCS').AsString;
end;

function Tdmgcpj_base_iv.ObtemCodigoReclamante(codLink: string;
  var codReclamada, empresa, pessoaexterna,
  fgcodexfuncionario: integer): integer;
begin
   dts50.Close;
   dts50.CommandText := 'select CDEPDC, CEMPR_INC, CPSSOA_EXTER_BDSCO, CFUNC_EMPR_GRP from GCPJB050 where CENVDO_PROCS_JUDIC = ' + codLink;
   dts50.Open;
   if dts50.Eof then
      result := 0
   else
   begin
      codReclamada := dts50.FieldByName('CDEPDC').AsInteger;
      empresa := dts50.FieldByName('CEMPR_INC').AsInteger;
      pessoaexterna := dts50.FieldByName('CPSSOA_EXTER_BDSCO').AsInteger;
      fgcodexfuncionario := dts50.FieldByName('CFUNC_EMPR_GRP').AsInteger;
      result := 1;
   end;
end;

function Tdmgcpj_base_iv.ExistePrimeiraReclamadaDoGrupo(
  gcpj: string): boolean;
var
   codReclamada : integer;
   tipoDep : string;
begin
   result := false;

   codReclamada := dmgcpj_base_IX.ObtemCodigoPrimeiraReclamada(gcpj);
   if codReclamada = 0 then
      exit;

   dts50.close;
   dts50.CommandText := 'select CDEPDC, CEMPR_INC, CPSSOA_EXTER_BDSCO, CFUNC_EMPR_GRP, CFUNC_BDSCO from GCPJB050 where CENVDO_PROCS_JUDIC = ' + IntToStr(codReclamada);
   dts50.Open;
   if dts50.Eof then
      exit;

   tipoDep := ObtemTipoDependencia(codreclamada, dts50.FieldByName('CEMPR_INC').AsInteger);
   if (tipoDep <> '') then
      result := true;

end;

end.
