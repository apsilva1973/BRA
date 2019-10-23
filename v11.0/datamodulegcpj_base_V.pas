unit datamodulegcpj_base_V;

interface

uses
  SysUtils, Classes, ADODB, DB, INIFILES, FORMS, dateutils;

type
  Tdmgcpcj_base_v = class(TDataModule)
    adoConn: TADOConnection;
    dts: TADODataSet;
    adoCmd: TADOCommand;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
//    function ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso, codescritorio: string; lstCnpjs: TStringList):integer;
    function ObtemDataCadastroGCPJ(gcpj: string): TDateTime;
//    procedure ObtemPagamentosLaudoPericial(gcpj, codEscritorio: string);
//    function ObtemSomaJaPagaADescontar(gcpj: string; drcContraria: integer; var qtdePaga: integer; const somenteAto: string = '';const excluirVolumetria : boolean = false; const recuperacaoFinal: boolean = false) : double;
//    function ObtemOutrosPagamentosDoProcessoAtivas(gcpj, tipoato, escritorio, tipoprocesso, codescritorio: string; lstCnpjs: TStringList):integer;
//    function ObtemQualquerPagamentoDoProcesso(gcpj, escritorio, tipoprocesso, codescritorio: string; lstCnpjs: TStringList):integer;
//    function PagouAcordoNestaData(gcpj: string; dataAto: TdateTime) : boolean;
    function ObtemDataAjuizamento(gcpj: string) : TdateTime;
//    procedure ObtemAcordosPagos(gcpj: string);

  end;

var
  dmgcpcj_base_v: Tdmgcpcj_base_v;

implementation

{$R *.dfm}

{ TDataModule1 }

procedure Tdmgcpcj_base_v.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index idx1 on GCPJB081(CNRO_PROCS_BDSCO)';
      adoCmd.Execute;
   except
   end;

end;

procedure Tdmgcpcj_base_v.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB081', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

function Tdmgcpcj_base_v.ObtemDataCadastroGCPJ(gcpj: string): TDateTime;
begin
   dts.Close;
   dts.Parameters.Clear;
   dts.CommandText := 'SELECT DCAD_PROCS_JUDIC AS DT_CADASTRO FROM GCPJB081 WHERE CNRO_PROCS_BDSCO=' + gcpj;
   dts.Open;

   result := dts.FieldByName('DT_CADASTRO').AsDatetime;
end;





function Tdmgcpcj_base_v.ObtemDataAjuizamento(gcpj: string): TdateTime;
begin
   dts.Close;
   dts.Parameters.Clear;
   dts.CommandText := 'SELECT DRECLC_PROCS_JUDIC AS DT_CADASTRO FROM GCPJB081 WHERE CNRO_PROCS_BDSCO=' + gcpj;
   dts.Open;

   if dts.FieldByName('DT_CADASTRO').IsNull then
      result := 0
   else
      result := dts.FieldByName('DT_CADASTRO').AsDatetime;
end;

end.
