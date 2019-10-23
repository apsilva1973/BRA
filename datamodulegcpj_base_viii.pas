unit datamodulegcpj_base_viii;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms;

type
  Tdmgcpj_base_VIII = class(TDataModule)
    adoConn: TADOConnection;
    adoCmd: TADOCommand;
    dts: TADODataSet;
    adoConn_ate_2013: TADOConnection;
    adoCmd_ate_2013: TADOCommand;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
    procedure ObtemEventosDoProcesso(gcpj: string);
    
  end;

var
  dmgcpj_base_VIII: Tdmgcpj_base_VIII;

implementation

{$R *.dfm}

{ Tdmgcpj_base_VIII }

procedure Tdmgcpj_base_VIII.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index idx1 on GCPJB052(CNRO_PROCS_BDSCO, CID_TPO_EVNTO)';
      adoCmd.Execute;
   except
   end;
{
//alexandre 13/04
   try
      adoCmd_ate_2013.CommandText := 'create index idx1 on GCPJB052(CNRO_PROCS_BDSCO, CID_TPO_EVNTO)';
      adoCmd_ate_2013.Execute;
   except
   end;
//alexandre 13/04}


end;

procedure Tdmgcpj_base_VIII.ObtemEventosDoProcesso(gcpj: string);
begin
   dts.Close;
   dts.CommandText := 'SELECT CID_TPO_EVNTO FROM GCPJB052 WHERE CNRO_PROCS_BDSCO = ' + gcpj + ' and CID_TPO_EVNTO = 533';
   dts.Open;
end;

procedure Tdmgcpj_base_VIII.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.ReadString('gcpj_base_viii', 'databasename', 'C:\presenta\basediaria\basediaria_gcpj_viii.mdb') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
{
//alexandre 13/04
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn_ate_2013.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.ReadString('gcpj_base_viii', 'databasename', 'C:\presenta\basediaria\BaseDiaria_GCPJB052_ate_2013.mdb') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
//alexandre 13/04}



end;

end.
