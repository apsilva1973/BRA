unit datamodulegcpj_base_iii;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, forms, dialogs;

type
  Tdmgcpj_base_iii = class(TDataModule)
    adoConn: TADOConnection;
    adoCmd: TADOCommand;
    dts: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
    function ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso: string):integer;
//    function ObtemPagamentosCadastradosGCPJ(gcpj, tipoato, tipoProcesso: string; tiposDeAto: TstringList):integer;
  end;

var
  dmgcpj_base_iii: Tdmgcpj_base_iii;

implementation

{$R *.dfm}

{ Tdmgcpj_base_iii }

procedure Tdmgcpj_base_iii.CreateIndex;
begin
(**   try
      adoCmd.CommandText := 'create index idx1 on GCPJB024(CNRO_PROCS_BDSCO, CIDTFD_ATO_DESP)';
      adoCmd.Execute;
   except
   end;**)
end;

procedure Tdmgcpj_base_iii.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('gcpj_base_iII', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

function Tdmgcpj_base_iii.ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso: string):integer;
begin
   result := -1;
   dts.Close;
   dts.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                      'from NFISCDT ' +
                      'where ';
   if Pos('TUTELA', UpperCase(tipoato )) <> 0 then
      dts.CommandText := dts.CommandText + '(nomdes = ''CONTESTACAO'' or nomdes = ''AUDIENCIA DE CONCILIACAO'' or nomdes = ''AUDIENCIA INICIAL'') and '
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
      (tipoato = 'AVULSO TRABALHISTA') OR (tipoato = 'INSTRUCAO')  then
   begin
      if tipoprocesso = 'T' then
         dts.CommandText := dts.CommandText + 'nomdes = ''INSTRUCAO'' and '
      else
         dts.CommandText := dts.CommandText + 'nomdes = ''ACOMPANHAMENTO'' and '
   end
   else
   begin
      ShowMessage('Erro tipoato: ' + tipoato);
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

(***function Tdmgcpj_base_iii.ObtemPagamentosCadastradosGCPJ(gcpj,
  tipoato, tipoProcesso: string; tiposDeAto: TstringList): integer;
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

   dts.Close;
   dts.CommandText := 'select CIDTFD_ATO_DESP, CNRO_PROCS_BDSCO, DATO_DESP_HONRS ' +
                      'from GCPJB024 ' +
                      'where CNRO_PROCS_BDSCO = ' + gcpj + ' AND CIDTFD_ATO_DESP IN (' + tipos + ')';
   dts.Open;
   result := 1;
end;   **)

end.
