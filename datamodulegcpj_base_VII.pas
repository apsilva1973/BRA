unit datamodulegcpj_base_VII;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms;

type
  Tdmgcpj_base_vii = class(TDataModule)
    adoConn: TADOConnection;
    adoCmd: TADOCommand;
    dts: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
//    function ProcessoTemInicial(gcpj: string):boolean;
//    function ProcessoTemRequerimento(gcpj: string):boolean;
//    procedure VerificaInicialERequerimento( gcpj: string; var temInicial: boolean; var temRequerimento: boolean);
  end;

var
  dmgcpj_base_vii: Tdmgcpj_base_vii;

implementation

{$R *.dfm}

{ Tdmgcpcj_base_vii }

procedure Tdmgcpj_base_vii.CreateIndex;
begin
   try
//      adoCmd.CommandText := 'create index idx1 on GCPJB0K7(CNRO_PROCS_BDSCO,RANEXO_PROCS_JUDIC,IANEXO_PROCS_JUDIC)';
      adoCmd.CommandText := 'create index idx2 on GCPJB0K7(CNRO_PROCS_BDSCO)';
      adoCmd.Execute;
   except
   end;
end;

(**
function Tdmgcpj_base_vii.ProcessoTemInicial(gcpj: string): boolean;
begin
   result := false;

   dts.Close;
   dts.CommandText := 'select IANEXO_PROCS_JUDIC, RANEXO_PROCS_JUDIC from GCPJB0K7 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj;
   dts.Open;
   if dts.Eof then
      exit;

   while not dts.Eof do
   begin
      if (Pos('INICIAL', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('INICIAL', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) then
      begin
         result := true;
         exit;
      end;
      dts.Next;
   end;
end;
   **)
procedure Tdmgcpj_base_vii.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.ReadString('gcpj_base_vii', 'databasename', 'C:\presenta\basediaria\basediaria_gcpj_vii.mdb') + ';Persist Security Info=True';
//      adoConn.Open;
   finally
      ini.free;
   end;

end;
(**

function Tdmgcpj_base_vii.ProcessoTemRequerimento(gcpj: string): boolean;
begin
   result := false;

   dts.Close;
   dts.CommandText := 'select IANEXO_PROCS_JUDIC, RANEXO_PROCS_JUDIC, IANEXO_PROCS_JUDIC from GCPJB0K7 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj;

   dts.Open;
   if dts.eof then
      exit;

   while not dts.eof do
   begin
      if (Pos('REQUEIMENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUEIMENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMNTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMNTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENT', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENT', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMETO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMETO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMNETO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMNETO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) then
      begin
         result := true;
         exit;
      end;
      dts.Next;
   end;
end;
   ***)

   (**
procedure Tdmgcpj_base_vii.VerificaInicialERequerimento(gcpj: string;
  var temInicial, temRequerimento: boolean);
begin
   temInicial := false;
   temRequerimento := false;

   dts.Close;
   dts.CommandText := 'select IANEXO_PROCS_JUDIC, RANEXO_PROCS_JUDIC, IANEXO_PROCS_JUDIC from GCPJB0K7 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj;

   dts.Open;
   if dts.eof then
      exit;

   while not dts.eof do
   begin
      if (Pos('INICIAL', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('INICIAL', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) then
         temInicial := true
      else
      if (Pos('REQUEIMENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUEIMENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMNTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIEMNTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENT', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENT', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENTO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMENTO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMETO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMETO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMNETO', dts.FieldByName('RANEXO_PROCS_JUDIC').AsString) <> 0) or
         (Pos('REQUERIMNETO', dts.FieldByName('IANEXO_PROCS_JUDIC').AsString) <> 0) then
         temRequerimento := true;
      dts.Next;
   end;
end;  ***)

end.
