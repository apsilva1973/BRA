unit datamodulegcpj_base_ii;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, mywin, mgenlib;

type
  Tdmgcpj_base_ii = class(TDataModule)
    adoConn24: TADOConnection;
    adoCmd24: TADOCommand;
    dts24: TADODataSet;
    adoConn88: TADOConnection;
    adoCmd88: TADOCommand;
    dts88: TADODataSet;
    adoConnB7: TADOConnection;
    adoCmdb7: TADOCommand;
    dtsB7: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
    procedure ObtemValorCalculoOriginal(gcpj: string);
//    function ProcessoTemInicial(gcpj: string):boolean;
    //alterado em 26/11/2014
//    function ProcessoTemReferencia(gcpj: string; codReferencia: string):boolean;
    function ObtemPagamentosCadastradosGCPJ(gcpj, tipoato, tipoProcesso: string; tiposDeAto: TstringList):integer;

  end;

var
  dmgcpj_base_ii: Tdmgcpj_base_ii;

implementation

{$R *.dfm}

procedure Tdmgcpj_base_ii.CreateIndex;
begin
   try
      adoCmd24.CommandText := 'create index idxp1 on GCPJB024(CNRO_PROCS_BDSCO, CIDTFD_ATO_DESP)';
      adoCmd24.Execute;
   except
   end;

   try
      adoCmdb7.CommandText := 'create index idxp1 on GCPJB0B7(CNRO_PROCS_BDSCO, CTPO_CALCD_PROCS)';
      adoCmdb7.Execute;
   except
   end;

(**   try
      adoCmd.CommandText := 'create index idx1 on GCPJB010(CNRO_PROCS_BDSCO)';//, CREFT_ANDAM_PROCS)';
      adoCmd.Execute;
   except
   end;**)

   try
      adoCmd88.CommandText := 'create index idx1 on GCPJB088(CREFT_ANDAM_PROCS)';
      adoCmd88.Execute;
   except
   end;
end;

procedure Tdmgcpj_base_ii.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn24.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB024', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn88.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB088', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConnB7.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJB0B7', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

function Tdmgcpj_base_ii.ObtemPagamentosCadastradosGCPJ(gcpj, tipoato,
  tipoProcesso: string; tiposDeAto: TstringList): integer;
var
   i, p : integer;
   tipos : string;
begin
   result := -1;
   tipos := '';
   if Pos('TUTELA', UpperCase(tipoato )) <> 0 then
   begin
      for i := 0 to tiposDeAto.Count - 1 do
      begin
         p := Pos(';', tiposDeAto.strings[i]);
         if Pos('CONTESTACAO', tiposDeAto.Strings[i]) <> 0 then
         begin
            if tipos <> '' then
               tipos := tipos + ',';
            tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
         end;

         if Pos('AUDIENCIA DE CONCILIACAO', tiposDeAto.Strings[i]) <> 0 then
         begin
            if tipos <> '' then
               tipos := tipos + ',';
            tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
         end;

         if Pos('AUDIENCIA INICIAL', tiposDeAto.Strings[i]) <> 0 then
         begin
            if tipos <> '' then
               tipos := tipos + ',';
            tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
         end;

         if Pos('ACOMPANHAMENTO', tiposDeAto.Strings[i]) <> 0 then
         begin
            if tipos <> '' then
               tipos := tipos + ',';
            tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
         end;
      end;
   end
   else
   begin
      if Pos('XITO', tipoato) <> 0 then
      begin
         for i := 0 to tiposDeAto.Count - 1 do
         begin
            p := Pos(';', tiposDeAto.strings[i]);
            if Pos('EXITO', tiposDeAto.Strings[i]) <> 0 then
            begin
               if tipos <> '' then
                  tipos := tipos + ',';
               tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
            end;
         end;
      end
      else
      begin
         if (tipoato = 'RECURSO NA FASE DE EXECUCAO') or
            (tipoato = 'RECURSO NA FASE DE EXECU«√O') or
            (tipoato = 'RECURSO EM FASE DE EXECU«√O') or
            (tipoato = 'RECURSO EM FASE DE EXECUCAO') then
         begin
            for i := 0 to tiposDeAto.Count - 1 do
            begin
               p := Pos(';', tiposDeAto.strings[i]);
               if Pos('RECURSO NA FASE DE EXECUCAO', tiposDeAto.Strings[i]) <> 0 then
               begin
                  if tipos <> '' then
                     tipos := tipos + ',';
                  tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
               end;
            end;
         end
         else
         begin
            if (Pos('RECURSO', tipoato) <> 0) OR
               (Pos('AGRAVO FAT', tipoato) <> 0) then
            begin
               for i := 0 to tiposDeAto.Count - 1 do
               begin
                  p := Pos(';', tiposDeAto.strings[i]);
                  if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='RECURSO' then
                  begin
                     if tipos <> '' then
                        tipos := tipos + ',';
                     tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                  end;
               end;
            end
            else
            begin
               if (tipoato = 'AUDIÍNCIA AVULSA') or (tipoato = 'AUDI NCIA AVULSA') or (tipoato = 'AUDIENCIA AVULSA') or
                  (tipoato = 'AVULSO TRABALHISTA') or (tipoato = 'AUDIENCIA TRABALHISTA') or (tipoato = 'AUDI NCIA TRABALHISTA') or
                  (tipoato = 'INSTRUCAO')  then
               begin
                  for i := 0 to tiposDeAto.Count - 1 do
                  begin
                     p := Pos(';', tiposDeAto.strings[i]);
                     if ((tipoprocesso = 'TR') OR (tipoprocesso = 'TO') or (tipoprocesso = 'TA')) and
                        (Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) = 'INSTRUCAO') then
                     begin
                        if tipos <> '' then
                           tipos := tipos + ',';
                        tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                     end
                     else
                     begin
                        if (Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) = 'ACOMPANHAMENTO') then
                        begin
                           if tipos <> '' then
                              tipos := tipos + ',';
                           tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                        end;
                     end;
                  end;
               end
               else
               begin
                  if (Tipoato = 'AVULSO CÕVEL (AUDITORIA)') or (Tipoato = 'AVULSO') or (Tipoato = 'AVULSO CÕVEL') or
                     (Tipoato = 'AVULSO CIVEL')then
                  begin
                     for i := 0 to tiposDeAto.Count - 1 do
                     begin
                        p := Pos(';', tiposDeAto.strings[i]);
                        if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='ASSESSORIA JURIDICA' then
                        begin
                           if tipos <> '' then
                              tipos := tipos + ',';
                           tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                        end;
                     end;
                  end
                  else
                  begin
                     if (Tipoato = 'HONOR¡RIOS INICIAIS') or (Tipoato = 'HONORARIOS INICIAIS') then
                     begin
                        for i := 0 to tiposDeAto.Count - 1 do
                        begin
                           p := Pos(';', tiposDeAto.strings[i]);
                           if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='HONORARIOS INICIAIS' then
                           begin
                              if tipos <> '' then
                                 tipos := tipos + ',';
                              tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                           end;
                        end;
                     end
                     else
                     begin
                        if (Tipoato = 'HONOR¡RIOS FINAIS') or (Tipoato = 'HONORARIOS FINAIS') then
                        begin
                           for i := 0 to tiposDeAto.Count - 1 do
                           begin
                              p := Pos(';', tiposDeAto.strings[i]);
                              if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='HONORARIOS DE EXITO' then
                              begin
                                 if tipos <> '' then
                                    tipos := tipos + ',';
                                 tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                              end;
                           end;
                        end
                        else
                        begin
                           if (Tipoato = 'AJUIZAMENTO') then
                           begin
                              for i := 0 to tiposDeAto.Count - 1 do
                              begin
                                 p := Pos(';', tiposDeAto.strings[i]);
                                 if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='AJUIZAMENTO' then
                                 begin
                                    if tipos <> '' then
                                       tipos := tipos + ',';
                                    tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                                 end;
                              end;
                           end
                           else
                           begin
                              for i := 0 to tiposDeAto.Count - 1 do
                              begin
                                 p := Pos(';', tiposDeAto.strings[i]);
                                 if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) = tipoato then
                                 begin
                                    if tipos <> '' then
                                       tipos := tipos + ',';
                                    tipos := tipos + Copy(tiposDeAto.strings[i], 1, p-1);
                                 end;
                              end;
                           end;
                        end;
                     end;
                  end;
               end;
            end;
         end;
      end;
   end;

   dts24.Close;
   dts24.CommandText := 'select CIDTFD_ATO_DESP, CNRO_PROCS_BDSCO, DATO_DESP_HONRS ' +
                      'from GCPJB024 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' AND CIDTFD_ATO_DESP IN (' + tipos + ')';
   dts24.Open;
   result := 1;
end;

procedure Tdmgcpj_base_ii.ObtemValorCalculoOriginal(gcpj: string);
begin
   dtsB7.Close;
   dtsB7.CommandText := 'select VORIGN_CALC_PROCS from GCPJB0B7 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and CTPO_CALCD_PROCS = 21'; //VALOR DO PEDIDO - CI
   dtsB7.Open;
end;

(***function Tdmgcpj_base_ii.ProcessoTemInicial(gcpj: string): boolean;
begin
   dts.Close;
   dts.CommandText := 'select IANEXO_PROCS_JUDIC from GCPJB0K7 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' and (RANEXO_PROCS_JUDIC like ''%INICIAL%'' OR IANEXO_PROCS_JUDIC LIKE ''%INICIAL%'')';
   dts.Open;
   result := (not dts.Eof);
end;***)

//alterado em 26/11/2014

(**
function Tdmgcpj_base_ii.ProcessoTemReferencia(gcpj: string; codReferencia: string): boolean;
var
   i : integer;
   str : TstringList;
begin
   result := false;
   dts.Close;
   **)
  //alterado em 26/11/2014
//   if codReferencia = 0 then
//      dts.CommandText := 'SELECT GCPJB010.CNRO_PROCS_BDSCO, GCPJB010.CSEQ_ANDAM_ESCRI, GCPJB088.RREFT_ANDAM_PROCS ' +
//                         'FROM GCPJB010, GCPJB088 ' +
//                         'WHERE GCPJB010.CNRO_PROCS_BDSCO=' + gcpj + ' AND GCPJB010.CREFT_ANDAM_PROCS=GCPJB088.CREFT_ANDAM_PROCS '
//   else
    (**  dts.CommandText := 'SELECT CNRO_PROCS_BDSCO, CSEQ_ANDAM_ESCRI, CREFT_ANDAM_PROCS ' +
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
//   dts.Open;
//   result := (Not dts.eof);
//   if dts.eof then
//   begin
//      result := false;
//      exit;
//   end;

(**   str := TstringList.Create;
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
end;   **)

end.
