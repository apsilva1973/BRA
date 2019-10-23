unit datamodulegcpj_base_i;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, dialogs;

type
  Tdmgcpcj_base_I = class(TDataModule)
    adoConn57: TADOConnection;
    dts57: TADODataSet;
    adoCmd57: TADOCommand;
    ADOConnXII: TADOConnection;
    ADOCmdXII: TADOCommand;
    dtsXII: TADODataSet;
    adoConn75: TADOConnection;
    adoCmd75: TADOCommand;
    dts75: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
//    function ObtemDadosReclamada1(gcpj, nomereclamada2: string; var codReclamada, nomereclamada: string):integer;
//    function ProcessoCadastradoBaseI(gcpj: string):boolean;
//    function ObtemEnvolvidosNoProcesso(gcpj: string; fgDrcAtivas: integer):integer;
    procedure CreateIndex;
//    function ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso: string):integer;
    function ObtemEscritorioDistribuidoPara(gcpj: string):integer;
//    function ObtemDataCadastroGCPJ(gcpj: string):TDateTime;
//    procedure ObtemEventosDoProcesso(gcpj: string);
    function ProcessoDistribuidoAdvogadoInterno_GCPJ(gcpj: string) : integer;
    function ProcessoDistribuidoAdvogadoInterno_SGCP(gcpj: string) : integer;
    function EhProcessoPreocupante(gcpj: string) : boolean;
    function ObtemDataDaDistribuicao(gcpj: string):TDateTime;
  end;

var
  dmgcpcj_base_I: Tdmgcpcj_base_I;

implementation

{$R *.dfm}

{ Tdmgcpcj_base_I }
(**
function Tdmgcpcj_base_I.ObtemDadosReclamada1(gcpj, nomereclamada2: string; var codReclamada,
  nomereclamada: string): integer;
begin
   result := 0;
   dts.Close;
   dts.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU from GCPJB077 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'IENVDO_REFRD_PROCS <> "' + nomereclamada2 + '"';
   dts.Open;
   if Not dts.Eof then
   begin
      codReclamada := dts.FieldByName('CENVDO_PROCS_REU').AsString;
      nomereclamada := dts.FieldByName('IENVDO_REFRD_PROCS').AsString;
      result := 1;
      exit;
   end;
end; **)

(*mudou pra a base IX
function Tdmgcpcj_base_I.ObtemEnvolvidosNoProcesso(gcpj: string; fgDrcAtivas: integer): integer;
begin
   dts.Close;
   if fgdrcAtivas = 1 then
      dts.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_AUTOR as CENVDO_PROCS_REU, CSEQ_PROCS_AUTOR as CSEQ_PROCS_REU ' +
                         'from GCPJB076 ' +
                         'where CNRO_PROCS_BDSCO = ' + gcpj + '  ' +
                         'order by CSEQ_PROCS_AUTOR'
   else
      dts.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU, CSEQ_PROCS_REU ' +
                         'from GCPJB077 ' +
                         'where CNRO_PROCS_BDSCO = ' + gcpj + '  ' +
                         'order by CSEQ_PROCS_REU';
   dts.Open;
   if dts.Eof then
      result := 0
   else
      result := 1;
end;
**)

(**
function Tdmgcpcj_base_I.ProcessoCadastradoBaseI(gcpj: string): boolean;
begin
   dts.Close;
   dts.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU from GCPJB077 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj;
   dts.Open;
   result := (Not dts.eof);
end;          **)

procedure Tdmgcpcj_base_I.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn57.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB057', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn75.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB075', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConnXII.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB056', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

end;

procedure Tdmgcpcj_base_I.CreateIndex;
begin
(**
   try
      adoCmd.CommandText := 'create index idx1 on GCPJB077 (CNRO_PROCS_BDSCO)';
      adoCmd.Execute;
   except
   end;

   try
      adoCmd.CommandText := 'create index idx2 on GCPJB077 (CSEQ_PROCS_REU)';
      adoCmd.Execute;
   except
   end;

   try
      adoCmd.CommandText := 'create index idx3 on GCPJB077 (CNRO_PROCS_BDSCO, IENVDO_REFRD_PROCS)';
      adoCmd.Execute;
   except
   end;
   **)
   try
      ADOCmdXII.CommandText := 'create index idx1 on GCPJB056(CNRO_PROCS_BDSCO)';
      ADOCmdXII.Execute;
   except
   end;

   try
      ADOCmdXII.CommandText := 'create index idx2 on GCPJB056(CNRO_PROCS_BDSCO, DFIM_HIST_PROCS)';
      ADOCmdXII.Execute;
   except
   end;


   try
      ADOCmdXII.CommandText := 'create index idx3 on GCPJB056(CNRO_PROCS_BDSCO, DFIM_HIST_PROCS, CSEQ_HIST_CONTT)';
      ADOCmdXII.Execute;
   except
   end;

   try
      adoCmd57.CommandText := 'create index idxp1 on GCPJB057(CNRO_PROCS_BDSCO, CNRO_INSTN_PROCS)';
      adoCmd57.Execute;
   except
   end;

   try
      adoCmd57.CommandText := 'create index idxp2 on GCPJB057(CADVOG_INTRN_BDSCO_N)';
      adoCmd57.Execute;
   except
   end;

   (**try
      adoCmd.CommandText := 'create index idx1 on GCPJB076 (CNRO_PROCS_BDSCO)';
      adoCmd.Execute;
   except
   end;
      **)

  (** try
      adoCmd.CommandText := 'create index idx2 on GCPJB076 (CNRO_PROCS_BDSCO, CSEQ_PROCS_AUTOR)';
      adoCmd.Execute;
   except
   end;
     **)

   try
      adoCmd75.CommandText := 'create index idxp1 on GCPJB075 (CNRO_PROCS_BDSCO, CIND_PROCS_PREOC)';
      adoCmd75.Execute;
   except
   end;
end;

(***
function Tdmgcpcj_base_i.ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso: string):integer;
begin
   result := -1;
   dts.Close;
   dts.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                      'from NFISCDT ' +
                      'where ';
   if Pos('TUTELA', UpperCase(tipoato )) <> 0 then
      dts.CommandText := dts.CommandText + '(nomdes = ''CONTESTACAO'' or nomdes = ''AUDIENCIA DE CONCILIACAO'' or nomdes = ''AUDIENCIA INICIAL'' or nomdes = ''ACOMPANHAMENTO'') and '
   else
   if Pos('XITO', tipoato) <> 0 then
      dts.CommandText := dts.CommandText + 'nomdes like ''EXITO%'' and '
   else
   if (tipoato = 'RECURSO NA FASE DE EXECUCAO') or (tipoato = 'RECURSO EM FASE DE EXECU«√O') then
      dts.CommandText := dts.CommandText + 'nomdes = ''RECURSO NA FASE DE EXECUCAO'' and '
   else
   if (Pos('RECURSO', tipoato) <> 0) OR
      (Pos('AGRAVO FAT', tipoato) <> 0) then
      dts.CommandText := dts.CommandText + 'nomdes = ''RECURSO'' and '
   else
   if (tipoato = 'AUDIÍNCIA AVULSA') or (tipoato = 'AUDI NCIA AVULSA') or (tipoato = 'AUDIENCIA AVULSA') or
      (tipoato = 'AVULSO TRABALHISTA') or (tipoato = 'AUDIENCIA TRABALHISTA') or (tipoato = 'AUDI NCIA TRABALHISTA') or
      (tipoato = 'INSTRUCAO')  then
   begin
      if (tipoprocesso = 'TR') OR (tipoprocesso = 'TA') OR (tipoprocesso = 'TO') then
         dts.CommandText := dts.CommandText + 'nomdes = ''INSTRUCAO'' and '
      else
         dts.CommandText := dts.CommandText + 'nomdes = ''ACOMPANHAMENTO'' and '
   end
   else
   if (Tipoato = 'AVULSO CÕVEL (AUDITORIA)') or (Tipoato = 'AVULSO') then
      dts.CommandText := dts.CommandText + 'nomdes = ''ASSESSORIA JURIDICA'' and '
   else
   begin
//      ShowMessage('Erro tipoato: ' + tipoato);
      exit;
   end;

   if Pos('JBM', escritorio) <> 0 then
      dts.CommandText := dts.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
                         'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '
   else
   if Pos('CAMARA', escritorio) <> 0 then
      dts.CommandText := dts.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                         'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and ';
   dts.CommandText := dts.CommandText + ' bradesco = ' + gcpj;
   dts.Open;
   result := 1;
end;
***)


function Tdmgcpcj_base_I.ObtemEscritorioDistribuidoPara(gcpj: string):integer;
var
   ultimo : integer;

begin
   result := 0;

   dtsXII.Close;
   dtsXII.CommandText := 'select max(CSEQ_HIST_CONTT) as ultimo ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null';
   dtsXII.Open;
   if dtsXII.FieldByName('ultimo').AsInteger = 0 then
      exit; //n„o distribuido para nenhum escritorio

   ultimo := dtsXII.FieldByName('ultimo').AsInteger;

   dtsXII.close;
   dtsXII.CommandText := 'select CPSSOA_EXTER_ADVOG, DINIC_HIST_PROCS, DFIM_HIST_PROCS ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null and ' +
                      'CSEQ_HIST_CONTT = ' + IntToStr(ultimo);
   dtsXII.Open;

   if dtsXII.recordCount > 0 then
      result := 1;
end;

(**function Tdmgcpcj_base_I.ObtemDataCadastroGCPJ(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.CommandText := 'SELECT DCAD_PROCS_JUDIC AS DT_CADASTRO FROM GCPJB081 WHERE CNRO_PROCS_BDSCO=' + gcpj;
   dts.Open;

   result := dts.FieldByName('DT_CADASTRO').AsDatetime;
end;   *)

(**
procedure Tdmgcpcj_base_I.ObtemEventosDoProcesso(gcpj: string);
begin
   dts.Close;
   dts.CommandText := 'SELECT CID_TPO_EVNTO FROM GCPJB052 WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' and CID_TPO_EVNTO = 533';
   dts.Open;
end;                    **)

function Tdmgcpcj_base_I.ProcessoDistribuidoAdvogadoInterno_GCPJ(gcpj: string): integer;
begin
   dts57.Close;
   dts57.CommandText := 'SELECT CADVOG_INTRN_BDSCO_N ' +
                      'FROM GCPJB057 ' +
                      'WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' AND CNRO_INSTN_PROCS=1 AND CSEQ_HIST_RESP IN(SELECT MAX(CSEQ_HIST_RESP) ' +
                      'FROM GCPJB057 WHERE CNRO_PROCS_BDSCO = ' + gcpj + ')';
   dts57.Open;

   if dts57.RecordCount = 0 then
      result := 0
   else
      result := dts57.FieldByName('CADVOG_INTRN_BDSCO_N').AsInteger;
end;

function Tdmgcpcj_base_I.ProcessoDistribuidoAdvogadoInterno_SGCP(
  gcpj: string): integer;
begin
   dts57.Close;
   dts57.CommandText := 'SELECT CADVOG_INTRN_BDSCO_N ' +
                      'FROM GCPJB057 ' +
                      'WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' AND CNRO_INSTN_PROCS=1 AND DFIM_RESP_ADVOG is null and CSEQ_HIST_RESP IN(SELECT MAX(CSEQ_HIST_RESP) ' +
                      'FROM GCPJB057 WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' and DFIM_RESP_ADVOG is null)';
   dts57.Open;

   if dts57.RecordCount = 0 then
      result := 0
   else
      result := dts57.FieldByName('CADVOG_INTRN_BDSCO_N').AsInteger;
end;

function Tdmgcpcj_base_I.EhProcessoPreocupante(gcpj: string): boolean;
begin
   dts75.Close;
   dts75.CommandText := 'SELECT CNRO_PROCS_BDSCO, CIND_PROCS_PREOC ' +
                      'FROM GCPJB075 ' +
                      'WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' AND CIND_PROCS_PREOC = 1';
   dts75.Open;

   if dts75.RecordCount = 0 then
      result := false
   else
      result := true;
end;

function Tdmgcpcj_base_I.ObtemDataDaDistribuicao(gcpj: string): TDateTime;
var
   ultimo : integer;
begin
   result := 0;
      //alterado em 13/12/2016 - alteraÁ„o sugerida pelo Henrique
  (** dts.Close;
   dts.CommandText := 'select max(CSEQ_HIST_CONTT) as ultimo ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null';
   dts.Open;
   if dts.FieldByName('ultimo').AsInteger = 0 then
      exit; //n„o distribuido para nenhum escritorio

   ultimo := dts.FieldByName('ultimo').AsInteger;    **)

   dtsXII.close;
   dtsXII.CommandText := 'select CPSSOA_EXTER_ADVOG, DINIC_HIST_PROCS, DFIM_HIST_PROCS ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null ';
                     // 'CSEQ_HIST_CONTT = ' + IntToStr(ultimo);
   dtsXII.Open;

   if dtsXII.recordCount > 0 then
      result := dtsXII.FieldByName('DINIC_HIST_PROCS').AsDateTime
   else
      result := 0;
end;

end.
