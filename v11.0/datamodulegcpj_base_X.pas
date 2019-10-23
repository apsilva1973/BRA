unit datamodulegcpj_base_X;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, mywin, mgenlib;

type
  Tdmgcpj_base_X = class(TDataModule)
    adoConn: TADOConnection;
    adoCmd: TADOCommand;
    dts: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
    procedure ObtemValorCalculoOriginal(gcpj: string);
    function ProcessoTemInicial(gcpj: string):boolean;
    //alterado em 26/11/2014
    function ProcessoTemReferencia(gcpj: string; codReferencia: string):boolean;
  end;

var
  dmgcpj_base_X: Tdmgcpj_base_X;

implementation

{$R *.dfm}

procedure Tdmgcpj_base_X.CreateIndex;
begin
(**
   try
      adoCmd.CommandText := 'create index idx1 on GCPJB0B7(CNRO_PROCS_BDSCO, CTPO_CALCD_PROCS)';
      adoCmd.Execute;
   except
   end;
   **)

   try
      adoCmd.CommandText := 'create index idx1 on GCPJB010(CNRO_PROCS_BDSCO)';//, CREFT_ANDAM_PROCS)';
      adoCmd.Execute;
   except
   end;
   (**
   try
      adoCmd.CommandText := 'create index idx1 on GCPJB088(CREFT_ANDAM_PROCS)';
      adoCmd.Execute;
   except
   end;**)
end;

procedure Tdmgcpj_base_X.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB010', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

procedure Tdmgcpj_base_X.ObtemValorCalculoOriginal(gcpj: string);
begin
   dts.Close;
   dts.CommandText := 'select VORIGN_CALC_PROCS from GCPJB0B7 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and CTPO_CALCD_PROCS = 21'; //VALOR DO PEDIDO - CI
   dts.Open;
end;

function Tdmgcpj_base_X.ProcessoTemInicial(gcpj: string): boolean;
begin
   dts.Close;
   dts.CommandText := 'select IANEXO_PROCS_JUDIC from GCPJB0K7 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and (RANEXO_PROCS_JUDIC like ''%INICIAL%'' OR IANEXO_PROCS_JUDIC LIKE ''%INICIAL%'')';
   dts.Open;
   result := (not dts.Eof);
end;

//alterado em 26/11/2014

function Tdmgcpj_base_X.ProcessoTemReferencia(gcpj: string; codReferencia: string): boolean;
var
   i : integer;
   str : TstringList;
begin
   result := false;
   dts.Close;

  //alterado em 26/11/2014
//   if codReferencia = 0 then
//      dts.CommandText := 'SELECT GCPJB010.CNRO_PROCS_BDSCO, GCPJB010.CSEQ_ANDAM_ESCRI, GCPJB088.RREFT_ANDAM_PROCS ' +
//                         'FROM GCPJB010, GCPJB088 ' +
//                         'WHERE GCPJB010.CNRO_PROCS_BDSCO=' + gcpj + ' AND GCPJB010.CREFT_ANDAM_PROCS=GCPJB088.CREFT_ANDAM_PROCS '
//   else
      dts.CommandText := 'SELECT CNRO_PROCS_BDSCO, CSEQ_ANDAM_ESCRI, CREFT_ANDAM_PROCS ' +
                         'FROM GCPJB010 ' +
                         'WHERE CNRO_PROCS_BDSCO=' + gcpj;// + ' AND CREFT_ANDAM_PROCS IN(' + codReferencia + ')';

(**      for i := 0 to codReferencia.Count - 1 do
      begin
         if i <> 0 then
            dts.CommandText := dts.CommandText + ', ';
         dts.CommandText := dts.CommandText + codReferencia.Strings[i];
      end;
      dts.CommandText := dts.CommandText + ')';
   **)
   dts.Open;
//   result := (Not dts.eof);
   if dts.eof then
   begin
      result := false;
      exit;
   end;

   str := TstringList.Create;
   try
      strToList(codReferencia, ',', str);

      while Not dts.Eof do
      begin
         for i := 0 to str.count - 1 do
         begin
            //if UpperCase(dts.FieldByName('RREFT_ANDAM_PROCS').AsString) = UpperCase(referencia) then
            if dts.FieldByName('CREFT_ANDAM_PROCS').AsString = str.strings[i] then
            begin
               result := true;
               exit;
            end;
         end;
         dts.Next;
      end;
   finally
      str.free;
   end;
end;

end.
