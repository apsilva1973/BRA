unit fEnviaGcpj;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls, ExtCtrls, Grids, DBGrids, uplanjbm,
  {$ifndef COMISD}
  bradisd,
  {$ENDIF}
  datamodule_honorarios, inifiles, mywin, mgenlib;

type
  TfrmEnviaGcpj = class(TForm)
    DBGrid1: TDBGrid;
    Label1: TLabel;
    rgFiltro: TRadioGroup;
    re: TRichEdit;
    BitBtn4: TBitBtn;
    BitBtn2: TBitBtn;
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
  private
    FplanJbm: TPlanilhaJBM;
    FdirArqEnviar: string;
    procedure SetplanJbm(const Value: TPlanilhaJBM);
    procedure SetdirArqEnviar(const Value: string);
    { Private declarations }
  public
    { Public declarations }
    property planJbm : TPlanilhaJBM read FplanJbm write SetplanJbm;
    property dirArqEnviar : string read FdirArqEnviar write SetdirArqEnviar;

  end;

var
  frmEnviaGcpj: TfrmEnviaGcpj;

implementation

uses fcaddatainicial, main;

{$R *.dfm}

procedure TfrmEnviaGcpj.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfrmEnviaGcpj.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   planJbm.Finalizar;
   planJbm.Free;
   action := caFree;
end;

procedure TfrmEnviaGcpj.SetplanJbm(const Value: TPlanilhaJBM);
begin
  FplanJbm := Value;
end;

procedure TfrmEnviaGcpj.FormCreate(Sender: TObject);
begin
   planJbm := TPlanilhaJBM.Create;
   DBGrid1.DataSource := planJbm.dsPlan;

   //verifica se o campo fgenviadogcpj existe na tabela tbplanilhas;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 fgenviadogcpj from tbplanilhas';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O parâmetro fgenviadogcpj não tem valor padrão', merr.Message) <> 0) then
            dmHonorarios.CriaColunafgenviadogcpj_tbplanilhas
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se o campo dtenviogcpj existe na tabela tbplanilhas;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 dtenviogcpj from tbplanilhas';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O parâmetro dtenviogcpj não tem valor padrão', merr.Message) <> 0) then
            dmHonorarios.CriaColunadtenviogcpj_tbplanilhas
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se o campo fgretornogcpj existe na tabela tbplanilhas;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 fgretornogcpj from tbplanilhas';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O parâmetro fgretornogcpj não tem valor padrão', merr.Message) <> 0) then
            dmHonorarios.CriaColunafgretornogcpj_tbplanilhas
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   planJbm.ObtemPlanilhasEnviarGpj;
   if planJbm.dtsPlanilha.RecordCount = 0 then
      BitBtn4.Enabled := false;

   if Not DirectoryExists(dmHonorarios.dirLck + '\' + FormatDatetime('yyyymmdd', date)) then
      MkDir(dmHonorarios.dirLck + '\' + FormatDatetime('yyyymmdd', date));

   dirArqEnviar := dmHonorarios.dirLck + '\' + FormatDatetime('yyyymmdd', date) + '\';
end;

procedure TfrmEnviaGcpj.BitBtn4Click(Sender: TObject);
var
   ini : Tinifile;
   dataDigitar : string;
   arquivoSaida : TStringList;
   linhaDetalhe : string;
   andamentoDigitar : string;
   tipoNotaFiscal : integer;
   codigoAto : integer;
   ultimaNota : string;

   function ObtemCodigoAto(tipoato, tipoprocesso: string) : integer;
   var
      i,p : integer;
      tiposDeAto : TstringList;

   begin
      tiposDeAto := TstringList.Create;
      try
      result := 0;

      tiposDeAto.LoadFromFile(ExtractFilePath(Application.ExeName) + 'TIPOS_ATOS.TXT');

      if Pos('TUTELA', UpperCase(tipoato )) <> 0 then
      begin
         for i := 0 to tiposDeAto.Count - 1 do
         begin
            p := Pos(';', tiposDeAto.strings[i]);
            if Pos('CONTESTACAO', tiposDeAto.Strings[i]) <> 0 then
            begin
               result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
               exit;
            end;
            if Pos('AUDIENCIA DE CONCILIACAO', tiposDeAto.Strings[i]) <> 0 then
            begin
               result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
               exit;
            end;

            if Pos('AUDIENCIA INICIAL', tiposDeAto.Strings[i]) <> 0 then
            begin
               result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
               exit;
            end;

            if Pos('ACOMPANHAMENTO', tiposDeAto.Strings[i]) <> 0 then
            begin
               result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
               exit;
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
                  result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                  exit;
               end;
            end;
         end
         else
         begin
            if (tipoato = 'RECURSO NA FASE DE EXECUCAO') or
               (tipoato = 'RECURSO NA FASE DE EXECUÇÃO') or
               (tipoato = 'RECURSO EM FASE DE EXECUÇÃO') or
               (tipoato = 'RECURSO EM FASE DE EXECUCAO') then
            begin
               for i := 0 to tiposDeAto.Count - 1 do
               begin
                  p := Pos(';', tiposDeAto.strings[i]);
                  if Pos('RECURSO NA FASE DE EXECUCAO', tiposDeAto.Strings[i]) <> 0 then
                  begin
                     result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                     exit;
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
                        result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                        exit;
                     end;
                  end;
               end
               else
               begin
                  if (tipoato = 'AUDIêNCIA AVULSA') or (tipoato = 'AUDIÊNCIA AVULSA') or (tipoato = 'AUDIENCIA AVULSA') or
                     (tipoato = 'AVULSO TRABALHISTA') or (tipoato = 'AUDIENCIA TRABALHISTA') or (tipoato = 'AUDIÊNCIA TRABALHISTA') or
                     (tipoato = 'INSTRUCAO')  then
                  begin
                     for i := 0 to tiposDeAto.Count - 1 do
                     begin
                        p := Pos(';', tiposDeAto.strings[i]);
                        if ((tipoprocesso = 'TR') OR (tipoprocesso = 'TO') or (tipoprocesso = 'TA')) and
                           (Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) = 'INSTRUCAO') then
                        begin
                           result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                           exit;
                        end
                        else
                        begin
                           if (Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) = 'ACOMPANHAMENTO') then
                           begin
                              result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                              exit;
                           end;
                        end;
                     end;
                  end
                  else
                  begin
                     if (Tipoato = 'AVULSO CÍVEL (AUDITORIA)') or (Tipoato = 'AVULSO') or (Tipoato = 'AVULSO CÍVEL') or
                        (Tipoato = 'AVULSO CIVEL')then
                     begin
                        for i := 0 to tiposDeAto.Count - 1 do
                        begin
                           p := Pos(';', tiposDeAto.strings[i]);
                           if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='ASSESSORIA JURIDICA' then
                           begin
                              result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                              exit;
                           end;
                        end;
                     end
                     else
                     begin
                        if (Tipoato = 'HONORÁRIOS INICIAIS') or (Tipoato = 'HONORARIOS INICIAIS') then
                        begin
                           for i := 0 to tiposDeAto.Count - 1 do
                           begin
                              p := Pos(';', tiposDeAto.strings[i]);
                              if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='HONORARIOS INICIAIS' then
                              begin
                                 result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                                 exit;
                              end;
                           end;
                        end
                        else
                        begin
                           if (Tipoato = 'HONORÁRIOS FINAIS') or (Tipoato = 'HONORARIOS FINAIS') then
                           begin
                              for i := 0 to tiposDeAto.Count - 1 do
                              begin
                                 p := Pos(';', tiposDeAto.strings[i]);
                                 //19/04/2017
//                                 if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='HONORARIOS DE EXITO' then
                                 if Trim(Uppercase(Copy(tiposDeAto.Strings[i], p+1, Length(tiposDeAto.Strings[i])))) ='HONORARIOS FINAIS' then
                                 begin
                                    result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                                    exit;
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
                                       result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                                       exit;
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
                                       result :=  StrToInt(Copy(tiposDeAto.strings[i], 1, p-1));
                                       exit;
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
      finally
         tiposDeAto.free;
      end;
   end;


begin
(*   if Not FileExists(ExtractFilePath(Application.Exename) + 'ISD.INI') then
   begin
      re.Lines.Add('Não encontrado o arquivo de parâmetros ISD.INI');
      exit;
   end;   *)

   ini := TIniFile.Create(ExtractFilePath(Application.Exename) + 'ISD.INI') ;
   try
      tipoNotaFiscal := ini.ReadInteger('layout', 'TIPO_DCTO_HONORARIOS', 1);
   finally
      ini.Free;
   end;

   //se a planilha já foi enviada confirma novo envio
   if planJbm.dtsPlanilha.FieldByName('fgenviadogcpj').AsInteger = 1 then
   begin
      if MessageDlg('Planilha já foi enviada para o GCPJ. Deseja enviar novamente?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
         exit;
   end;

   re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Carregando dados das tabelas da planilha');

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').Asinteger;
   planJbm.cnpjEscritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.CarregaFiliais;

   dataDigitar := FormatDateTime('dd.mm.yyyy', Date);//planJbm.ObtemDataDigitar(2);
   if dataDigitar = '' then
   begin
      if MessageDlg('Não encontrada a data inicial para o ano/mês de referência. Deseja cadastrar?', mtConfirmation, mbYesNoCancel, 0) <> mrYes then
      begin
//         planJbm.NotificaErro('Sistema não pode digitar para o mês ' + planJbm.anomesreferencia + ' até ser cadastrada a data inicial');
         ShowMessage('Sistema não pode digitar para o mês ' + planJbm.anomesreferencia + ' até ser cadastrada a data inicial');
         exit;
      end;

      frmCadDataInicial := TfrmCadDataInicial.Create(nil);
      try
         if frmCadDataInicial.ShowModal <> mrOk then
            exit;

         planJbm.CadastraDataInicial(frmCadDataInicial.dataInicial.date);
      finally
         frmCadDataInicial.Free;
      end;
      dataDigitar := planJbm.ObtemDataDigitar(2);
   end;

   //verifica se já foi enviado arquivo para o escritório hoje
   if (FileExists(dirArqEnviar + planJbm.cnpjEscritorio + '.OK')) or (FileExists(dirArqEnviar + planJbm.cnpjEscritorio + '.txt')) then
   begin
      re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Já foi enviado arquivo para este escritório hoje. Somente um arquivo por dia pode ser enviado');
      exit;
   end;

   planJbm.ObtemProcessosDigitar('');
   if planJbm.dtsAtos.RecordCount = 0 then
   begin
      re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Não encontrada nenhuma nota para enviar ao GCPJ');
      exit;
   end;

   arquivoSaida := TStringList.Create;
   try
      //coloca o header
      re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Incluindo o HEADER do arquivo');
      arquivoSaida.Add('0' + planJbm.cnpjEscritorio + '0' + FormatDateTime('dd.mm.yyyy', date) + '00001' + Format('%-48.48s', [' ']));

      ultimaNota := '';

      while not planJbm.dtsAtos.Eof do
      begin
         if planJbm.dtsAtos.FieldByName('numeronota').AsInteger = 0 then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end;

         andamentoDigitar := planjbm.ObtemTipoAndamento(planJbm.dtsAtos.FieldByName('tipoandamento').AsString, planJbm.tipoProcesso, planJbm.dtsAtos.FieldByName('motivobaixa').AsString);
         if andamentoDigitar = '' then
         begin
            re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-GCPJ: ' + planJbm.dtsAtos.FieldByName('GCPJ').AsString);
            exit;
         end;

         re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Verificando inclusão do Processo: ' + planJbm.dtsAtos.FieldByName('numeronota').AsString + '/' + planJbm.dtsAtos.FieldByName('gcpj').ASString + '/' + andamentoDigitar);
         Application.ProcessMessages;

         planJbm.CarregaCamposDaTabela;

         if ultimaNota <> planJbm.dtsAtos.FieldByName('numeronota').AsString then
         begin
            linhaDetalhe := '0' + planJbm.cnpjEscritorio + '1' + LeftZeroStr('%10s', [planJbm.dtsAtos.FieldByName('numeronota').AsString]) +  '00001' + '1' + LeftZeroStr('%2d', [tipoNotaFiscal]);
            linhaDetalhe := linhaDetalhe + dataDigitar;

            planJbm.ObtemValorTotalDaNota;
            linhaDetalhe := linhaDetalhe +  LeftZeroStr('%15s', [RemoveNaoNumericos(FormatFloat('0.00', planJbm.valorTotalNota))]) + '00237';
            if planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger <> 237 then
               linhaDetalhe := linhaDetalhe + LeftZeroStr('%5s', [planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString])
            else
               linhaDetalhe := linhaDetalhe + '00000';
            linhaDetalhe := linhaDetalhe + Format('%-10.10s', [' ']);
            arquivoSaida.Add(linhaDetalhe);
            ultimaNota := planJbm.dtsAtos.FieldByName('numeronota').AsString;
         end;

         if Not FileExists(ExtractFilePath(Application.ExeName) + 'TIPOS_ATOS.TXT') then
         begin
            re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Não encontrado o arquivo de parâmetros TIPO_ATOS.TXT');
            exit;
         end;

         linhaDetalhe := '0' + planJbm.cnpjEscritorio + '2' + LeftZeroStr('%10s', [planJbm.dtsAtos.FieldByName('numeronota').AsString]) +  '00001';
         linhaDetalhe := linhaDetalhe + LeftZeroStr('%10s', [planJbm.dtsAtos.FieldByname('gcpj').AsString]);
         linhaDetalhe := linhaDetalhe + LeftZeroStr('%5s', [planJbm.ObtemDependenciaDigitar]);
         linhaDetalhe := linhaDetalhe + LeftZeroStr('%5s', [planJbm.ObtemDependenciaDigitar]);//[planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString]);

         if andamentoDigitar = '' then
            codigoAto := ObtemCodigoAto(planJbm.dtsAtos.FieldByName('tipoandamento').asString, planJbm.dtsAtos.FieldByName('fgtipoprocesso').asString)
         else
            codigoAto := ObtemCodigoAto(andamentoDigitar, planJbm.dtsAtos.FieldByName('fgtipoprocesso').asString);
         if codigoAto = 0 then
         begin
            re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Não encontrado o código do ato na tabela TIPOS_ATOS.TXT');
            exit;
         end;

         linhaDetalhe := linhaDetalhe + LeftZeroStr('%3d', [codigoAto]) + FormatDateTime('dd.mm.yyyy',planjbm.dtsatos.FieldByName('datadoato').AsDateTime) + LeftZeroStr('%15s', [RemoveNaoNumericos(FormatFloat('0.00', planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat))]);
         arquivoSaida.Add(linhaDetalhe);

         planJbm.dtsAtos.Next;
      end;
   finally
      re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Incluindo o TRAILLER do arquivo');
      arquivoSaida.Add('0' + planJbm.cnpjEscritorio + '9' + LeftZeroStr('%10d', [arquivoSaida.Count-1]) + Format('%-53.53s', [' ']));

      arquivoSaida.SaveToFile(dirArqEnviar + planJbm.cnpjEscritorio + '.TXT');
      arquivoSaida.free;
   end;
   re.Lines.Add('[' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ']-Encerrado o envio do arquivo para o GCPJ');
end;

procedure TfrmEnviaGcpj.SetdirArqEnviar(const Value: string);
begin
  FdirArqEnviar := Value;
end;

end.
