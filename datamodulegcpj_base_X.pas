unit datamodulegcpj_base_X;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, mywin, mgenlib;

type
  Tdmgcpj_base_X = class(TDataModule)
    adoConn: TADOConnection;
    adoCmd: TADOCommand;
    dts: TADODataSet;
    adoConn_de_2016_ate_2017: TADOConnection;
    adoCmd_de_2016_ate_2017: TADOCommand;
    dts_de_2016_ate_2017: TADODataSet;
    adoConn_ate_2015: TADOConnection;
    adoCmd_ate_2015: TADOCommand;
    dts_ate_2015: TADODataSet;
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

// alexandre 13/04
   try
      adoCmd_de_2016_ate_2017.CommandText := 'create index idx1 on GCPJB010(CNRO_PROCS_BDSCO)';//, CREFT_ANDAM_PROCS)';
      adoCmd_de_2016_ate_2017.Execute;
   except
   end;


   try
      adoCmd_ate_2015.CommandText := 'create index idx1 on GCPJB010(CNRO_PROCS_BDSCO)';//, CREFT_ANDAM_PROCS)';
      adoCmd_ate_2015.Execute;
   except
   end;

// alexandre 13/04




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

//alexandre 13/04
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn_de_2016_ate_2017.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB010_de_2016_ate_2017', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
//alexandre 13/04

//alexandre 13/04
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn_ate_2015.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB010_ate_2015', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
//alexandre 13/04


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


   dts.CommandText := 'SELECT CNRO_PROCS_BDSCO, CSEQ_ANDAM_ESCRI, CREFT_ANDAM_PROCS ' +
                      'FROM GCPJB010 ' +
                       'WHERE CNRO_PROCS_BDSCO=' + gcpj;// + ' AND CREFT_ANDAM_PROCS IN(' + codReferencia + ')';
   dts.Open;

// Edison alteração inicio 27/07/2018

    dts_de_2016_ate_2017.Close;
    dts_de_2016_ate_2017.CommandText := 'SELECT CNRO_PROCS_BDSCO, CSEQ_ANDAM_ESCRI, CREFT_ANDAM_PROCS ' +
                                          'FROM GCPJB010 ' +
                                          'WHERE CNRO_PROCS_BDSCO=' + gcpj;// + ' AND CREFT_ANDAM_PROCS IN(' + codReferencia + ')';

     dts_de_2016_ate_2017.Open;

//   Edison alteração Fim 27/07/2018

{

  if dts.eof then
   begin

// alexandre 13/04
      dts_de_2016_ate_2017.Close;
      dts_de_2016_ate_2017.CommandText := 'SELECT CNRO_PROCS_BDSCO, CSEQ_ANDAM_ESCRI, CREFT_ANDAM_PROCS ' +
                                          'FROM GCPJB010 ' +
                                          'WHERE CNRO_PROCS_BDSCO=' + gcpj;// + ' AND CREFT_ANDAM_PROCS IN(' + codReferencia + ')';

      dts_de_2016_ate_2017.Open;

      if dts_de_2016_ate_2017.Eof then
       begin
        result := false;
        exit;
       end
      else
       dts.Clone(dts_de_2016_ate_2017);
   end;
}

// alexandre 13/04}

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

// Edison alteração inicio 27/07/2018

      while Not dts_de_2016_ate_2017.Eof do
      begin
         for i := 0 to str.count - 1 do
         begin
            if dts_de_2016_ate_2017.FieldByName('CREFT_ANDAM_PROCS').AsString = str.strings[i] then
            begin
               result := true;
               exit;
            end;
         end;
         dts_de_2016_ate_2017.Next;
      end;

// Edison alteração fim 27/07/2018

   finally
      str.free;
   end;
end;

end.
