unit datamodulegcpj_base_migrados;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms;

type
  TdmGcpj_migrados = class(TDataModule)
    adoConn: TADOConnection;
    dts: TADODataSet;
    adoCmd: TADOCommand;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function ObtemDadosReclamada(gcpj: string):integer;
    procedure CreateIndex;
  end;

var
  dmGcpj_migrados: TdmGcpj_migrados;

implementation

{$R *.dfm}

{ TDataModule1 }

function TdmGcpj_migrados.ObtemDadosReclamada(gcpj: string):integer;
begin
   dts.Close;
   dts.CommandText := 'select cprocs_judic_bdsco, a.cjunc_depdc, idepdc_bdsco, creu_exeq ' +
                      'from dba_depdcia_c_procjur  a, dba_depdc_bdsco_procs b ' +
                      'where cprocs_judic_bdsco = ' + gcpj + ' and nsequen <= 2 and a.cjunc_depdc = b.cjunc_depdc ' +
                      'order by nsequen';
   dts.Open;
   if dts.Eof then
      result := 0
   else
      result := 1;
end;


procedure TdmGcpj_migrados.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('gcpj_base_migrados', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

procedure TdmGcpj_migrados.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index idx1 on dba_depdcia_c_procjur (cprocs_judic_bdsco,cjunc_depdc)';
      adoCmd.Execute;
   except
   end;

   try
      adoCmd.CommandText := 'create index idx2 on dba_depdc_bdsco_procs (cjunc_depdc)';
      adoCmd.Execute;
   except
   end;

end;

end.
