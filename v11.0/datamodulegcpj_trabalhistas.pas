unit datamodulegcpj_trabalhistas;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, datamodulegcpj_baixados, datamodulegcpj_recuperados;

type
  Tdmgcpj_trabalhistas = class(TDataModule)
    adoConn: TADOConnection;
    dts: TADODataSet;
    adoCmd: TADOCommand;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
    function ObtemTipoTrabalhista(gcpj: string):integer;
  end;

var
  dmgcpj_trabalhistas: Tdmgcpj_trabalhistas;

implementation

{$R *.dfm}

{ Tdmgcpj_trabalhistas }

procedure Tdmgcpj_trabalhistas.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('gcpj_base_trabalhistas', 'databasename', 'c:\presenta\basediaria\Relatorios_Gerenciais_DADOS_TA_TO_TR.mdb') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

procedure Tdmgcpj_trabalhistas.CreateIndex;
begin
   try
      adoCmd.CommandText := 'create index idx1 on _Trabalhista_Entradas(bradesco)';
      adoCmd.Execute;
   except
   end;
end;

function Tdmgcpj_trabalhistas.ObtemTipoTrabalhista(gcpj: string): integer;
begin
   //
   dts.Close;
   dts.CommandText := 'select bradesco, CodJunAgrup, NomJun, TipoProcesso from _Trabalhista_Entradas ' +
                      'where bradesco = ' + gcpj;
   dts.Open;

   result := 0;

   if dts.RecordCount > 1 then
   begin
      result := 99;
      exit;
   end;

   if dts.FieldByName('TipoProcesso').AsString = 'EX-FUNCIONARIOS' then
      result := 3
   else
   begin
      if dts.FieldByName('CodJunAgrup').AsString = '531003' then //bvp
      begin
         if (dts.FieldByName('TipoProcesso').AsString = 'TERCEIRIZADO')  and (Pos('COR', Uppercase(dts.FieldByName('NomJun').AsString)) <> 0) then
            result := 2
         else
            result := 1;
      end
      else
      begin
         if (dts.FieldByName('TipoProcesso').AsString = 'TERCEIRIZADO')  then
            result := 1
         else
            result := 2;
      end;
   end;
end;

end.
