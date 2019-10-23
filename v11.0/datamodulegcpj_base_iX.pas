unit datamodulegcpj_base_ix;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, dialogs;

type
  Tdmgcpj_base_IX = class(TDataModule)
    adoConn76: TADOConnection;
    dts76: TADODataSet;
    adoCmd76: TADOCommand;
    adoConn77: TADOConnection;
    dts77: TADODataSet;
    adoCmd77: TADOCommand;
    dtsPesq: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function ObtemDadosReclamada(gcpj, nomereclamada2: string; var codReclamada, nomereclamada: string):integer;
    function ProcessoCadastradoBaseI(gcpj: string):boolean;
    function ObtemEnvolvidosNoProcesso(gcpj: string; fgDrcAtivas: integer; buscarAutor: boolean):integer;
    procedure CreateIndex;
//  function ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso: string):integer;
//    function ObtemEscritorioDistribuidoPara(gcpj: string):integer;
//    function ObtemDataCadastroGCPJ(gcpj: string):TDateTime;
//    procedure ObtemEventosDoProcesso(gcpj: string);
//    function ProcessoDistribuidoAdvogadoInterno_GCPJ(gcpj: string) : integer;
//    function ProcessoDistribuidoAdvogadoInterno_SGCP(gcpj: string) : integer;
//    function EhProcessoPreocupante(gcpj: string) : boolean;
//    function ObtemDataDaDistribuicao(gcpj: string):TDateTime;
    function ObtemNumeroGcpj(codLink: string)  : string;
    function ObtemCOdigoPrimeiraReclamada(gcpj: string) : integer;

  end;

var
  dmgcpj_base_IX: Tdmgcpj_base_IX;

implementation

{$R *.dfm}

{ Tdmgcpcj_base_I }

function Tdmgcpj_base_IX.ObtemDadosReclamada(gcpj, nomereclamada2: string; var codReclamada,
  nomereclamada: string): integer;
begin
   result := 0;
   dts77.Close;
   dts77.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU from GCPJB077 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'IENVDO_REFRD_PROCS <> "' + nomereclamada2 + '"';
   dts77.Open;
   if Not dts77.Eof then
   begin
      codReclamada := dts77.FieldByName('CENVDO_PROCS_REU').AsString;
      nomereclamada := dts77.FieldByName('IENVDO_REFRD_PROCS').AsString;
      result := 1;
      exit;
   end;

(**   dts.Close;
   dts.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU from GCPJB077 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj;
   dts.open;
   if Not dts.eof then
   begin
      codReclamada := dts.FieldByName('CENVDO_PROCS_REU').AsString;
      nomereclamada := dts.FieldByName('IENVDO_REFRD_PROCS').AsString;
      result := 1;
   end;**)
end;

function Tdmgcpj_base_IX.ObtemEnvolvidosNoProcesso(gcpj: string; fgDrcAtivas: integer; buscarAutor: boolean): integer;
begin
   if (fgdrcAtivas = 1) or (buscarAutor) then
   begin
      dtsPesq.Close;
      dtsPesq.Connection := dts76.Connection;
      dtsPesq.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_AUTOR as CENVDO_PROCS_REU, CSEQ_PROCS_AUTOR as CSEQ_PROCS_REU ' +
                         'from GCPJB076 ' +
                         'where CNRO_PROCS_BDSCO = ' + gcpj + '  ' +
                         'order by CSEQ_PROCS_AUTOR';
      dtsPesq.Open;
      if dtsPesq.Eof then
         result := 0
      else
         result := 1;
   end
   else
   begin
      dtsPesq.Close;
      dtsPesq.Connection := dts77.Connection;
      dtsPesq.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU, CSEQ_PROCS_REU ' +
                         'from GCPJB077 ' +
                         'where CNRO_PROCS_BDSCO = ' + gcpj + '  ' +
                         'order by CSEQ_PROCS_REU';
      dtsPesq.Open;
      if dtsPesq.Eof then
         result := 0
      else
         result := 1;
   end;
end;

function Tdmgcpj_base_IX.ProcessoCadastradoBaseI(gcpj: string): boolean;
begin
   dts77.Close;
   dts77.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU from GCPJB077 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj;
   dts77.Open;
   result := (Not dts77.eof);
end;

procedure Tdmgcpj_base_IX.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn76.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB076', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn77.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB077', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

procedure Tdmgcpj_base_IX.CreateIndex;
begin
   try
      adoCmd77.CommandText := 'create index idx1 on GCPJB077 (CNRO_PROCS_BDSCO)';
      adoCmd77.Execute;
   except
   end;

   try
      adoCmd77.CommandText := 'create index idx2 on GCPJB077 (CSEQ_PROCS_REU)';
      adoCmd77.Execute;
   except
   end;

   try
      adoCmd77.CommandText := 'create index idx3 on GCPJB077 (CNRO_PROCS_BDSCO, IENVDO_REFRD_PROCS)';
      adoCmd77.Execute;
   except
   end;

   try
      adoCmd76.CommandText := 'create index idx2 on GCPJB076 (CNRO_PROCS_BDSCO, CSEQ_PROCS_AUTOR)';
      adoCmd76.Execute;
   except
   end;

   try
      adoCmd76.CommandText := 'create index idx3 on GCPJB076 (CSEQ_PROCS_AUTOR)';
      adoCmd76.Execute;
   except
   end;

end;

(***
function Tdmgcpj_base_IX.ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso: string):integer;
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
end;***)



(***
function Tdmgcpj_base_IX.ObtemEscritorioDistribuidoPara(gcpj: string):integer;
var
   ultimo : integer;

begin
   result := 0;

   dts.Close;
   dts.CommandText := 'select max(CSEQ_HIST_CONTT) as ultimo ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null';
   dts.Open;
   if dts.FieldByName('ultimo').AsInteger = 0 then
      exit; //n„o distribuido para nenhum escritorio

   ultimo := dts.FieldByName('ultimo').AsInteger;

   dts.close;
   dts.CommandText := 'select CPSSOA_EXTER_ADVOG, DINIC_HIST_PROCS, DFIM_HIST_PROCS ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null and ' +
                      'CSEQ_HIST_CONTT = ' + IntToStr(ultimo);
   dts.Open;

   if dts.recordCount > 0 then
      result := 1;
end;
***)

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


(***
function Tdmgcpj_base_IX.ProcessoDistribuidoAdvogadoInterno_GCPJ(gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'SELECT CADVOG_INTRN_BDSCO_N ' +
                      'FROM GCPJB057 ' +
                      'WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' AND CNRO_INSTN_PROCS=1 AND CSEQ_HIST_RESP IN(SELECT MAX(CSEQ_HIST_RESP) ' +
                      'FROM GCPJB057 WHERE CNRO_PROCS_BDSCO = ' + gcpj + ')';
   dts.Open;

   if dts.RecordCount = 0 then
      result := 0
   else
      result := dts.FieldByName('CADVOG_INTRN_BDSCO_N').AsInteger;
end;
    ***)

(**
function Tdmgcpj_base_IX.ProcessoDistribuidoAdvogadoInterno_SGCP(
  gcpj: string): integer;
begin
   dts.Close;
   dts.CommandText := 'SELECT CADVOG_INTRN_BDSCO_N ' +
                      'FROM GCPJB057 ' +
                      'WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' AND CNRO_INSTN_PROCS=1 AND DFIM_RESP_ADVOG is null and CSEQ_HIST_RESP IN(SELECT MAX(CSEQ_HIST_RESP) ' +
                      'FROM GCPJB057 WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' and DFIM_RESP_ADVOG is null)';
   dts.Open;

   if dts.RecordCount = 0 then
      result := 0
   else
      result := dts.FieldByName('CADVOG_INTRN_BDSCO_N').AsInteger;
end;
***)

(***
function Tdmgcpj_base_IX.EhProcessoPreocupante(gcpj: string): boolean;
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
end;***)


(**
function Tdmgcpj_base_IX.ObtemDataDaDistribuicao(gcpj: string): TDateTime;
var
   ultimo : integer;
begin
   result := 0;

   dts.Close;
   dts.CommandText := 'select max(CSEQ_HIST_CONTT) as ultimo ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null';
   dts.Open;
   if dts.FieldByName('ultimo').AsInteger = 0 then
      exit; //n„o distribuido para nenhum escritorio

   ultimo := dts.FieldByName('ultimo').AsInteger;

   dts.close;
   dts.CommandText := 'select CPSSOA_EXTER_ADVOG, DINIC_HIST_PROCS, DFIM_HIST_PROCS ' +
                      'from GCPJB056 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and ' +
                      'DFIM_HIST_PROCS is null and ' +
                      'CSEQ_HIST_CONTT = ' + IntToStr(ultimo);
   dts.Open;

   if dts.recordCount > 0 then
      result := dts.FieldByName('DINIC_HIST_PROCS').AsDateTime
   else
      result := 0;
end;
   **)

function Tdmgcpj_base_IX.ObtemNumeroGcpj(codLink: string): string;
begin
   dts76.Close;
   dts76.CommandText := 'select CNRO_PROCS_BDSCO FROM GCPJB076 where  CSEQ_PROCS_AUTOR = ' + codLink;
   dts76.Open;
   if dts76.Eof then
      result := ''
   else
      result := dts76.FieldByName('CNRO_PROCS_BDSCO').AsString;
end;

function Tdmgcpj_base_IX.ObtemCOdigoPrimeiraReclamada(
  gcpj: string): integer;
begin
   dts77.Close;
   dts77.CommandText := 'select IENVDO_REFRD_PROCS, CENVDO_PROCS_REU from GCPJB077 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and CSEQ_PROCS_REU=1';
   dts77.Open;

   if dts77.Eof then
      result := 0
   else
      result := dts77.FieldByName('CENVDO_PROCS_REU').AsInteger;

   if result = 0 then
      exit;


end;

end.
