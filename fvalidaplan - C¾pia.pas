unit fvalidaplan;

interface

uses
  windows,SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Buttons, Grids, DBGrids, uplanjbm, inifiles,
  datamodulegcpj_base_i, Func_Wintask_Obj, mgenlib,  Messages, datamodulegcpj_base_V;

type
  TfrmValidaPlan = class(TForm)
    Label1: TLabel;
    DBGrid1: TDBGrid;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    re: TRichEdit;
    BitBtn3: TBitBtn;
    BitBtn5: TBitBtn;
    processando: TEdit;
    Label2: TLabel;
    numNota: TEdit;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    cbSemVolumetria: TCheckBox;
    BitBtn4: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
  private
    FplanJbm: TPlanilhaJBM;
    procedure SetplanJbm(const Value: TPlanilhaJBM);
    { Private declarations }
  public
    { Public declarations }
    property planJbm : TPlanilhaJBM read FplanJbm write SetplanJbm;
    procedure ObtemCodigoReclamadas;
    procedure IndexaBasesGCPJ;
    procedure CopiaBasesGCPJ;
    procedure ChecaNomeAutor;
    procedure ObtemTipoAcao;
    procedure ObtemNomeEmpresaGrupo;
    procedure ConfereValoresContestacao;
    procedure ValidaOutrosValores;
    procedure CalculaNumeroDaNota;
    procedure CorrigirNota;
    procedure VerificaDuplicidades;
    procedure RemoveEstadosDeVolumetria;
  end;


var
  frmValidaPlan: TfrmValidaPlan;
  ie8 : boolean;
  ultimaPagina, passagem : integer;
  
implementation


{$R *.dfm}

{ TfrmValidaPlan }

procedure TfrmValidaPlan.SetplanJbm(const Value: TPlanilhaJBM);
begin
  FplanJbm := Value;
end;

procedure TfrmValidaPlan.FormCreate(Sender: TObject);
begin
   planJbm := TPlanilhaJBM.Create;
   DBGrid1.DataSource := planJbm.dsPlan;
   planJbm.ObtemPlanilhasValidar;
   if planJbm.dtsPlanilha.RecordCount = 0 then
      BitBtn1.Enabled := false;

   if IsParam('/ie8') then
      ie8 := true
   else
      ie8 := false;
end;

procedure TfrmValidaPlan.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfrmValidaPlan.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   planJbm.Finalizar;
   planJbm.Free;
   action := CaFree;
end;

procedure TfrmValidaPlan.BitBtn1Click(Sender: TObject);
var
   ini : TIniFile;
begin
   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
   try
      if Not IsParam('/presenta') then
      begin
         if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'basescopiadas','') then
         begin
            CopiaBasesGCPJ;
            ini.WriteString('gcpj', 'basescopiadas',FormatDateTime('dd/mm/yyyy', date));
         end;
      end;

      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').Asinteger;

   ObtemCodigoReclamadas;
   ChecaNomeAutor;
   ObtemTipoAcao;
   ObtemNomeEmpresaGrupo;
   VerificaDuplicidades;
   ValidaOutrosValores;
   ConfereValoresContestacao;
   re.Lines.Add('Final do processo de validações');
end;

procedure TfrmValidaPlan.ObtemCodigoReclamadas;
var
   reclamada : TProcessoReclamdas;
   ret : integer;
begin
   re.lines.Clear;

   processando.Text := 'Capturando dados das reclamadas';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(0);
   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Cruzando processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString + ' com dados do GCPJ');
      Application.ProcessMessages;

      if planJbm.dtsAtos.FieldByName('gcpj').IsNull then
      begin
         planJbm.GravaOcorrencia(2, 'Sem número do GCPJ');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if Length(planJbm.dtsAtos.FieldByName('gcpj').AsString) > 11 then
      begin
         planJbm.GravaOcorrencia(2, 'Número do GCPJ inválido');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;
      if planJbm.dtsAtos.FieldByName('valor').AsFloat = 0 then
      begin
         planJbm.GravaOcorrencia(2, 'Valor do ato zerado');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      //processo existe no gcpj?
      if planJbm.ProcessoExisteNoGcpj = 0 then
      begin
         planJbm.GravaOcorrencia(2, 'Processo não cadastrado no GCPJ');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      reclamada := TProcessoReclamdas.Create(planJbm.idEscritorio,
                                             planJbm.idplanilha,
                                             planJbm.anomesreferencia,
                                             planJbm.sequencia,
                                             planJbm.linhaPlanilha,
                                             planJbm.gcpj);
      try
         reclamada.RemoveReclamadas;

         re.lines.Add('   Obtendo dados nas bases I e IV  ');
         ret := reclamada.ObtemEnvolvidosDoProcesso;
         if ret = 0 then //erro. Não existe no GCPJ
         begin
            planJbm.GravaOcorrencia(2, 'Processo não cadastrado no GCPJ');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if ret < 0 then
         begin
            planJbm.GravaOcorrencia(2, reclamada.errorMessage);
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         //verifica se tem data da baixa
         planJbm.VerificaDataDaBaixa(planJbm.dtsAtos.FieldByName('gcpj').AsString);
         planJbm.MarcaProcessoCruzadoGcpj(1);
      finally
         reclamada.Free;
      end;
      planJbm.dtsAtos.Next;
   end;
end;


procedure TfrmValidaPlan.IndexaBasesGCPJ;
begin
   processando.Text := 'Indexando as bases do GCPJ';
   Application.ProcessMessages;

   re.lines.Add('Criando índice na base de dados do GCPJ Base I. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseI;

   re.lines.Add('Criando índice na base de dados do GCPJ Base II. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseII;

   re.lines.Add('Criando índice na base de dados do GCPJ Base III. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseIII;

   re.lines.Add('Criando índice na base de dados do GCPJ Base IV. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseIV;

   re.lines.Add('Criando índice na base de dados do GCPJ Base V. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseV;

   re.lines.Add('Criando índice na base de dados do GCPJ Base Compartilhada. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseCompartilhada;

   re.lines.Add('Criando índice na base de dados do GCPJ Base Migrados. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseMigrados;
end;

procedure TfrmValidaPlan.ChecaNomeAutor;
begin
   re.lines.Clear;

   processando.Text := 'Verificando nome do autor';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(1);
   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Verificando nome do Autor do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      if planJbm.NomeEnvolvidoOk then
         planJbm.MarcaNomeEnvolvidoOk;

      planJbm.MarcaProcessoCruzadoGcpj(2);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ObtemTipoAcao;
begin
   re.lines.Clear;

   processando.Text := 'Capturando dados do Tipo/Subtipo Ação';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(2);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Verificando tipo acao do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      planJbm.GravaTipoProcesso;
      planJbm.GravaTipoSubtipo;
      planJbm.MarcaProcessoCruzadoGcpj(3);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ObtemNomeEmpresaGrupo;
var
  ret :integer;
begin
   re.lines.Clear;

   processando.Text := 'Capturando dados da Empresa Grupo';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(3);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Verificando Empresa Ligada: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      if planJbm.GravaEmpresaGrupo(0) = 1 then
         planJbm.MarcaProcessoCruzadoGcpj(4);

      //se for avulso trabalhista ou avulso civel, não pode estar distribuído para nenhum escritório
      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO TRABALHISTA') or
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIÊNCIA TRABALHISTA') or
         (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'AVULSO CÍVEL (AUDITORIA)') or
         (Pos('AVULSO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         ret := planJbm.AtoDistribuidoParaEscritorio;
         if ret <> 0 then //está distribuído para um escritório. Não pode
         begin
            planJbm.GravaOcorrencia(2, 'Ato distribuído para escritório externo não pode ser pago');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;
      end
      else
      begin
         if planJbm.tipoProcesso = 'CI' then
         begin
         //valida se o ato foi distribuido para este excritorio.
           ret := planJbm.AtoDistribuidoParaEscritorio;
           if ret <> 9 then
           begin
              if ret = 0 then
                planJbm.GravaOcorrencia(2, 'Processo não distribuído')
              else
                planJbm.GravaOcorrencia(2, 'Processo distribuído para outro escritório');

              planJbm.MarcaProcessoCruzadoGcpj(9);
              planJbm.dtsAtos.Next;
              continue;
           end;
         end;
      end;
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ConfereValoresContestacao;
var
   faixa1, faixa2,faixa3,faixa4,faixa5,faixa6,
   faixa7,faixa8,faixa9 : integer;
   total, faixa : integer;
   valorCorreto : double;
begin
   re.lines.Clear;

   processando.Text := 'Conferindo valores de Contestação';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(4);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
      exit;

   planJbm.cnpjescritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').AsInteger;

   //remove estados que não pode considerar volumetria
   if (planJbm.cnpjEscritorio = '08833747000132') or
      (planJbm.cnpjEscritorio = '05159996000104') or
      (planJbm.cnpjEscritorio = '10508423000170') then
      RemoveEstadosDeVolumetria;

   planJbm.ObtemProcessosCruzarGcpj(5);
   if planJbm.dtsAtos.RecordCount = 0 then
      exit;


   //total de atos de contestação com a nova regra de cálculo
   total := planJbm.dtsAtos.RecordCount;

(**   if (planJbm.cnpjEscritorio = '05159996000104') and (planJbm.anomesreferencia = '2012/03') then
   begin
      while not planjbm.dtsAtos.Eof do
      begin
         planJbm.CarregaCamposDaTabela;
         re.Lines.add('Verificando valores de contestação: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
         Application.ProcessMessages;

         valorCorreto := planjbm.dtsAtos.FieldByName('valor').AsFloat;
         if FormatDateTime('yyyy/mm', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) = '2011/11' then
            valorCorreto := 378.42
         else
         if FormatDateTime('yyyy/mm', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) = '2011/12' then
            valorCorreto := 384.52
         else
         if FormatDateTime('yyyy/mm', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) = '2012/01' then
            valorCorreto := 389.77;


         if planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorCorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
           planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
      end;
      exit;
   end;*)

   faixa1 := 0;
   faixa2 := 0;
   faixa3 := 0;
   faixa4 := 0;
   faixa5 := 0;
   faixa6 := 0;
   faixa7 := 0;
   faixa8 := 0;
   faixa9 := 0;
   faixa := 1;
   while total > 0 do
   begin
      case faixa of
         1 : begin
            if total >= 50 then
               faixa1 := 50
            else
               faixa1 := total;
            total := total - faixa1;
         end;
         2 : begin
            if total >= 50 then
               faixa2 := 50
            else
               faixa2 := total;
            total := total - faixa2;
         end;
         3 : begin
            if total >= 150 then
               faixa3 := 150
            else
               faixa3 := total;
            total := total - faixa3;
         end;
         4 : begin
            if total >= 250 then
               faixa4 := 250
            else
               faixa4 := total;
            total := total - faixa4;
         end;
         5 : begin
            if total >= 250 then
               faixa5 := 250
            else
               faixa5 := total;
            total := total - faixa5;
         end;
         6 : begin
            if total >= 250 then
               faixa6 := 250
            else
               faixa6 := total;
            total := total - faixa6;
         end;
         7 : begin
            if total >= 500 then
               faixa7 := 500
            else
               faixa7 := total;
            total := total - faixa7;
         end;
         8 : begin
            if total >= 500 then
               faixa8 := 500
            else
               faixa8 := total;
            total := total - faixa8;
         end;
         9 : begin
            faixa9 := total;
            total := total - faixa9;
         end;
      end;
      Inc(faixa);
   end;

   //novos valores de volumetria 04/10/2012
   //valorCorreto := ((faixa1*422) + (faixa2*410) + (faixa3*400) + (faixa4*390) + (faixa5*378) + (faixa6*367) + (faixa7*356) + (faixa8*345) + (faixa9*335));
   valorCorreto := ((faixa1*438) + (faixa2*425) + (faixa3*415) + (faixa4*405) + (faixa5*392) + (faixa6*381) + (faixa7*369) + (faixa8*358) + (faixa9*348));
   valorCorreto := valorcorreto / planJbm.dtsAtos.RecordCount;

   while not planjbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Verificando valores de contestação: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      if Not cbSemVolumetria.Checked then
      begin
         if planjbm.dtsAtos.FieldByName('valor').AsFloat <> valorCorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
      end
      else
         planJbm.GravaValorCorrigido(planjbm.dtsAtos.FieldByName('valor').AsFloat);
      planJbm.MarcaProcessoCruzadoGcpj(6);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.ValidaOutrosValores;
var
   valorcorreto: double;
begin
   re.lines.Clear;

   processando.Text := 'Validando valores';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(4);

   planJbm.cnpjescritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').AsInteger;

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Verificando valores: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      planJbm.GravaValorCalculo;

      //nao pode pagar ibi
      if (planjbm.tipoProcesso <> 'TR') and
         (planjbm.tipoProcesso <> 'TO') and
         (planjbm.tipoProcesso <> 'TA') and
         ((planjbm.dtsatos.FieldByName('empresaligadaagrupar').AsInteger = 63) OR
          (planjbm.dtsatos.FieldByName('empresaligadaagrupar').AsInteger = 7459) or
          (planjbm.dtsatos.FieldByName('empresaligadaagrupar').AsInteger = 5404)) then
      begin
         if planJbm.dtsPlanilha.FieldByName('fgpagaribi').AsInteger <> 1 then
         begin
            planJbm.GravaOcorrencia(2, 'Processo Banco IBI Civel');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if Not planJbm.dtsPlanilha.FieldByName('dataatosibi').IsNull then
         begin
            if FormatDateTime('yyyymmdd', dmgcpcj_base_V.ObtemDataCadastroGCPJ(planJbm.dtsAtos.FieldByName('gcpj').AsString)) < FormatDateTime('yyyymmdd', planJbm.dtsPlanilha.FieldByName('dataatosibi').AsDateTime) then
            begin
               planJbm.GravaOcorrencia(2, 'Processo Banco IBI Civel');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;
      end;

      //verifica duplicidade
      if Pos('PREPOSTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0 then
      begin
         if dmgcpcj_base_V.ObtemOutrosPagamentosDoProcesso(planJbm.gcpj,
                                       planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                       planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                       planjbm.tipoProcesso,
                                       IntToStr(planJbm.codGcpjEscritorio)) <> 1 then
         begin
            planJbm.GravaOcorrencia(2, 'Tipo de ato não reconhecido pelo sistema');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if dmgcpcj_base_V.dts.RecordCount > 0 then
         begin
            if (planJbm.tipoProcesso = 'TR') OR (planJbm.tipoProcesso = 'TO') OR (planJbm.tipoProcesso = 'TA') then
            begin
               if dmgcpcj_base_V.dts.RecordCount >= 2 then
               begin
                  planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end
            end
            else
            begin
               if Pos('AVULSO', planJbm.dtsAtos.FieldByName('tipoandamento').AsString) <> 0 then
               begin
                  if dmgcpcj_base_V.dts.RecordCount >= 2 then
                  begin
                     planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
                     planJbm.MarcaProcessoCruzadoGcpj(9);
                     planJbm.dtsAtos.Next;
                     continue;
                  end
               end
               else
               begin
                  planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;
            end;
         end;
      end;

      //tem recurso em juizado
      if planjbm.dtsatos.FieldByName('fgjuizado').AsInteger = 1 then
      begin
         if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIêNCIA AVULSA') or
            (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIENCIA AVULSA') then
         begin
            planJbm.GravaOcorrencia(2, 'Não pode cobrar AUDIENCIA AVULSA em processo juizado');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
            (Pos('AGRAVO FAT', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
         begin
            planJbm.GravaOcorrencia(2, 'Não pode cobrar RECURSO em processo juizado');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;
      end;

      if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORARIOS INICIAIS') or
         (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORÁRIOS INICIAIS') then
      begin
          if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01' then
             valorCorreto := 52.00
          else
             valorCorreto := 50.00;

          if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
          begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
          end
          else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
          planJbm.MarcaProcessoCruzadoGcpj(6);
          planJbm.dtsAtos.Next;
          continue;
      end;

      if (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORARIOS FINAIS') or
         (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORÁRIOS FINAIS') then
      begin
          if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01' then
             valorCorreto := 52.00
          else
             valorCorreto := 50.00;

          if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'EXTINTO COM JULGAMENTO DE MERITO') or
             (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO COM CUSTOS') or
             (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'IMPROCEDENCIA') then
          begin
            if planJbm.valorpedido <= 5000 then
               valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.15))
            else
               valorcorreto := StrToFloat(FormatFloat('0.00', (5000 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.15));
          end;

          if valorcorreto <= 0 then
          begin
            planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
          end;

          if ((FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01') and
              (valorcorreto < 52.00)) then
             valorcorreto := 52.00
          else
            if valorCorreto < 50.00 then
               valorcorreto := 50.00;

          if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
          begin
             planJbm.GravaValorCorrigido(valorcorreto);
             planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
          end
          else
          begin
             planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
             planJbm.MarcaProcessoCruzadoGcpj(6);
             planJbm.dtsAtos.Next;
             continue;
          end;
      end;

      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIêNCIA AVULSA') or
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIENCIA AVULSA') then
      begin
         if (FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01') then
            valorcorreto := 290.00
         else
            valorCorreto := 279.00;


         if (planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto) then
         begin
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.MarcaProcessoCruzadoGcpj(6);
            planJbm.dtsAtos.Next;
            continue;
         end;
      end;

      if (Pos('EXITO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
         (Pos('ÊXITO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'EXTINTO SEM JULGAMENTO DE MERITO') or
            (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO SEM CUSTOS') or
            (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'LIQUIDACAO') or
            (planJbm.dtsAtos.FieldByName('motivobaixa').AsString = 'DESISTENCIA DA ACAO') or
            (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'PROCEDENTE SEM CUSTOS') or
            (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'APENSADO') or
            (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'TRANSFERENCIA PARA DRC') or
            (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ARQUIVADO') or
            (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'DUPLICIDADE') or
            (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'SOLUCIONADO') or
            (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'PROCEDENTE SEM CUSTOS') or
            (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'EXTINCAO DO PROCESSO') or
            (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'IMPLANTACAO DE SENTENCA') or
            (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'CADASTRO INDEVIDO PELA DEPENDENCIA') or
            (planjbm.dtsAtos.FieldByName('motivobaixa').AsString = 'CADASTRO IRREGULAR PELA DEPENDENCIA') then
         begin
            planJbm.GravaOcorrencia(2, 'EXITO não pode ser pago por causa do motivo da baixa');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         //se for do DRC tem uma regra especial para tratar
         if planJbm. EhDrcContraria then
         begin
            //tem que estar ajuizado
            if planjbm.dtsatos.FieldByName('fgjuizado').AsInteger <> 1 then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) para processo não ajuizado');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //verifica se foi acordo parcelado
            if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO A VISTA') then
            begin
               if (planjbm.dtsatos.FieldByName('motivobaixa').IsNull) and
                  (planjbm.dtsatos.FieldByName('databaixa').IsNull) then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) - ACORDO A VISTA para processo não baixado');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if planJbm.valorpedido <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO (DRC) - ACORDO A VISTA com valor de cálculo zerado');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               valorcorreto := (planJbm.valorpedido * 10) / 100;
               if valorcorreto > 50000.00 then
                  valorCorreto := 50000.00;

               if valorcorreto <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
               begin
                  planJbm.GravaValorCorrigido(valorcorreto);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
            end;
         end
         else
         begin
            if (planjbm.dtsatos.FieldByName('motivobaixa').IsNull) and
               (planjbm.dtsatos.FieldByName('databaixa').IsNull) then
            begin
               planJbm.GravaOcorrencia(2, 'Não pode pagar EXITO para processo não baixado');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.dtsAtos.Next;
               continue;
            end;

            //se for acord com custo verifica se o benefício é menor do que 5000
            if planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'ACORDO COM CUSTOS' then
            begin
               if planJbm.valorpedido <= 5000 then
                  valorcorreto := StrToFloat(FormatFloat('0.00', (planJbm.valorpedido - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.15))
               else
                  valorcorreto := StrToFloat(FormatFloat('0.00', (5000 - planjbm.dtsatos.FieldByName('valorbaixa').AsFloat)*0.15));

               if valorcorreto <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
               begin
                  planJbm.GravaValorCorrigido(valorcorreto);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if (planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'IMPROCEDENCIA') or
               (planJbm.dtsAtos.FieldByName('motivobaixa').AsString = 'EXTINTO COM JULGAMENTO DE MERITO') then
            begin
               if Pos('PLANOS ECONOMICOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0 then
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2011/04/30' then //utlizar a rotina de novo cálculo
                     valorcorreto := 309.14
                  else
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
                     valorCorreto := 342
                  else
                     valorcorreto := 355;

                  if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                  begin
                     planJbm.GravaValorCorrigido(valorcorreto);
                     planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                  end
                  else
                     planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                  planJbm.MarcaProcessoCruzadoGcpj(6);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if Pos('ACAO DE REPARACAO DE DANOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0  then
               begin
                  if planJbm.dtsAtos.FieldByName('subtipoacao').AsString = 'OBRIGACAO DE FAZER' then
                  begin
                     if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
                        valorCorreto := 342
                     else
                        valorCorreto := 355;

                     if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
                     begin
                        planJbm.GravaValorCorrigido(valorcorreto);
                        planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
                     end
                     else
                        planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
                     planJbm.MarcaProcessoCruzadoGcpj(6);
                     planJbm.dtsAtos.Next;
                     continue;
                  end;
               end;

               if planjbm.valorpedido <= 5000 then
                  valorCorreto := StrToFloat(FormatFloat('0.00', planjbm.valorpedido * 0.15))
               else
                  valorCorreto := 750;

               if valorcorreto <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
               begin
                  planJbm.GravaValorCorrigido(valorcorreto);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end;

            if planjbm.dtsatos.FieldByName('motivobaixa').AsString = 'EXTINTO COM JULGAMENTO DE MERITO' then
            begin
               if (Pos('PLANOS ECONOMICOS', planjbm.dtsatos.FieldByName('tipoacao').AsString) <> 0) or
                  (planjbm.dtsatos.FieldByName('tipoacao').AsString = 'OBRIGACAO DE FAZER') then
               begin
                  if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
                     valorCorreto := 342
                  else
                     valorcorreto := 355
               end
               else
               begin
                  if planjbm.valorpedido <= 5000 then
                     valorCorreto := StrToFloat(FormatFloat('0.00', planjbm.valorpedido * 0.15))
                  else
                     valorCorreto := 750;
               end;

               if valorcorreto <= 0 then
               begin
                  planJbm.GravaOcorrencia(2, 'Não tem direito ao pagamento de honorário');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.dtsAtos.Next;
                  continue;
               end;

               if StrToFloat(FormatFloat('0.00', planjbm.dtsatos.FieldByName('valor').AsFloat)) <> StrToFloat(FormatFloat('0.00', valorcorreto)) then
               begin
                  planJbm.GravaValorCorrigido(valorcorreto);
                  planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
               end
               else
                  planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
               planJbm.MarcaProcessoCruzadoGcpj(6);
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;
      end;

      if (Pos('ACOMPAN', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) and
         (Pos('DELEGACIA', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         planJbm.GravaOcorrencia(2, 'Não pode pagar ACOMPANHAMENTO EM DELEGACIA');
         planJbm.MarcaProcessoCruzadoGcpj(9);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if Pos('PREPOSTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0 then
      begin
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 57
         else
            valorCorreto := 60;

         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if Pos('TUTELA', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0 then
      begin
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2011/11/01' then //utlizar a rotina de novo cálculo
         begin
            planJbm.MarcaProcessoCruzadoGcpj(5);
            planJbm.dtsAtos.Next;
            continue;
         end;
      end;

      if (Pos('TUTELA',planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
         (Pos('RECURSO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) OR
         (Pos('AGRAVO FAT', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) then
      begin
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2011/05/01' then //utlizar a rotina de novo cálculo
         begin
            if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
               valorCorreto := 422
            else
               valorCorreto := 438;

            if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
            begin
               planJbm.GravaValorCorrigido(valorcorreto);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
            planJbm.MarcaProcessoCruzadoGcpj(6);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2010/05/01' then //utlizar a rotina de novo cálculo
         begin
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> 382 then
            begin
               planJbm.GravaValorCorrigido(382);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
            planJbm.MarcaProcessoCruzadoGcpj(6);
            planJbm.dtsAtos.Next;
            continue;
         end;

//         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2009/05/01' then //utlizar a rotina de novo cálculo
//         begin
            if planjbm.dtsatos.FieldByName('valor').AsFloat <> 370 then
            begin
               planJbm.GravaValorCorrigido(370);
               planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
            end
            else
               planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
            planJbm.MarcaProcessoCruzadoGcpj(6);
            planJbm.dtsAtos.Next;
            continue;
//         end;
      end;

      if (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'AVULSO TRABALHISTA') or
         (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'AUDIÊNCIA TRABALHISTA') or
         (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'AUDIÊNCIA AVULSA') or
         (UpperCase(planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 'INSTRUCAO') then
      begin
         if (planJbm.tipoProcesso <> 'TR') AND (planJbm.tipoProcesso <> 'TO') AND (planJbm.tipoProcesso <> 'TA') then
         begin
            planJbm.GravaOcorrencia(2, 'Tipo de ato somente permitido para processo TRABALHISTA');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 564
         else
            valorCorreto := 585;

         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;
      if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CÍVEL (AUDITORIA)') or
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CÍVEL') or
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CIVEL') OR
         (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO CIVEL (AUDITORIA)')then
      begin
         if planJbm.tipoProcesso <> 'CI' then
         begin
            planJbm.GravaOcorrencia(2, 'Tipo de ato somente permitido para processo CIVEL');
            planJbm.MarcaProcessoCruzadoGcpj(9);
            planJbm.dtsAtos.Next;
            continue;
         end;

         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 279
         else
            valorCorreto := 290;

         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;
      planJbm.GravaOcorrencia(2, 'Não achou regra no programa para validar o pagamento deste ato');
      planJbm.MarcaProcessoCruzadoGcpj(9);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.CalculaNumeroDaNota;
var
   valorCobrar : double;
   index : integer;
begin
   re.lines.Clear;

   processando.Text := 'Calculando valor das notas';
   Application.ProcessMessages;
   planJbm.ObtemProcessosCruzarGcpj(3);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
   begin
      re.Lines.add('Não foi possível calcular as notas. Ainda existe processo não validado. Motivo: 3');
      exit;
   end;

   planJbm.ObtemProcessosCruzarGcpj(4);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
   begin
      re.Lines.add('Não foi possível calcular as notas. Ainda existe processo não validado. Motivo: 4');
      exit;
   end;

   planJbm.ObtemProcessosCruzarGcpj(5);
   if planJbm.dtsAtos.RecordCount <> 0 then //ainda não finalizado
   begin
      re.Lines.add('Não foi possível calcular as notas. Ainda existe processo não validado. Motivo: 5');
      exit;
   end;

   for index := 1 to 2 do
   begin
      planJbm.LimpaDadosNota;
      planJbm.ObtemProcessosCruzarGcpj(6);

      while not planJbm.dtsAtos.Eof do
      begin
         planJbm.CarregaCamposDaTabela;
         planJbm.GuardaTipoProcesso;

         //index=1 processos trabalhistas
         if ((planJbm.tipoProcesso = 'TR') OR (planJbm.tipoProcesso = 'TO') OR (planJbm.tipoProcesso = 'TA')) and (index = 2) then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end;

         //index=2 processos civeis
         if (planJbm.tipoProcesso = 'CI') and (index = 1) then
         begin
            planJbm.dtsAtos.Next;
            continue;
         end;

         re.Lines.add('Calculando número da nota para o processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
         if (planJbm.empresaLigada <> planjbm.dtsAtos.FieldByName('empresaligadaagrupar').ASInteger) and
            (planJbm.empresaLigada <> 0) then
         begin
            planJbm.GravaValorTotalDaNota;
            planJbm.LimpaDadosNota;
         end;

         if planJbm.empresaLigada = 0 then
         begin
            planJbm.empresaLigada := planjbm.dtsAtos.FieldByName('empresaligadaagrupar').ASInteger;
            planJbm.ObtemNumeroNotaEmpresa;
         end;

         valorCobrar := planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat;

         if ((planJbm.valorTotalNota + valorCobrar) > 130000) or
            (planjbm.totalAtos > 150) then
         begin
            //finaliza a nota por causa do valor
            planJbm.GravaValorTotalDaNota;
            planjbm.EncerraNotaPeloValor;
            planJbm.LimpaDadosNota;
            planJbm.empresaLigada := planjbm.dtsAtos.FieldByName('empresaligadaagrupar').ASInteger;
            planJbm.ObtemNumeroNotaEmpresa;
         end;

         planJbm.valorTotalNota := planJbm.valorTotalNota + valorCobrar;
         planJbm.totalAtos := planJbm.totalAtos + 1;
         planJbm.GravaNotaNoProcesso;

         planJbm.dtsAtos.Next;

         Application.ProcessMessages;
      end;

      if (planJbm.empresaLigada <> 0) then
         planJbm.GravaValorTotalDaNota;
   end;

   //já fiinalizado?
   planJbm.ObtemProcessosCruzarGcpj(9);
   if planJbm.dtsAtos.RecordCount = 0 then
      planJbm.MarcaPlanilhaValidada;

end;

procedure TfrmValidaPlan.BitBtn3Click(Sender: TObject);
var
   arquivo : TstringList;
   linha : string;
   ini : TIniFile;
   contador : integer;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   re.lines.Clear;

   planJbm.ObtemProcessosCruzarGcpj(7);

   arquivo := TStringList.Create;
   try
      arquivo.add('Ano/Mês;Escritório;Cód. Empresa Grupo;Nome Empresa Grupo;Nota;GCPJ;Parte Contrária;Tipo Ato;Tipo Ação;Sub Tipo Ação;Data Baixa;Motivo Baixa;Data do Ato;Valor do Pedido;Valor Pagar;'+
                  'Cod.Reclamada1;Nome Reclamada1;Cod.Reclamada2;Nome Reclamada2;Cod.Reclamada3;Nome Reclamada3;Cod.Reclamada4;Nome Reclamada4');
      while not planJbm.dtsAtos.Eof do
      begin
         planJbm.CarregaCamposDaTabela;
         re.Lines.add('Exportando: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
         Application.ProcessMessages;

         linha := planJbm.anomesreferencia + ';' + planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString + ';' + planJbm.dtsAtos.FieldByName('nomeempresaligada').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('numeronota').AsString + ';' + planJbm.dtsAtos.FieldByName('gcpj').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('partecontraria').AsString + ';' + planJbm.dtsAtos.FieldByName('tipoandamento').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('tipoacao').AsString + ';' + planJbm.dtsAtos.FieldByName('subtipoacao').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('databaixa').AsString + ';' + planJbm.dtsAtos.FieldByName('motivobaixa').AsString + ';' +
                  planJbm.dtsAtos.FieldByName('datadoato').AsString + ';' + FormatFloat('#,##0.00', planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat) + ';' +
                  FormatFloat('#,##0.00', planJbm.dtsAtos.FieldByName('valordopedido').AsFloat);

         planJbm.ObtemReclamadasDoProcesso;
         if planJbm.dtsReclamadas.Eof then
            linha := linha + ';;;;;;;';

         contador := 0;
         while not planJbm.dtsReclamadas.Eof do
         begin
            Inc(contador);
            if contador > 4 then
               break;
            linha := linha + ';' + planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString + ';' + planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
            planjbm.dtsReclamadas.Next;
         end;
         arquivo.Add(linha);
         planJbm.dtsAtos.Next;
      end;
      arquivo.SaveToFile(ExtractFilePath(Application.ExeName) + 'NOTAS.TXT');
   finally
      arquivo.free;
   end;
end;

procedure TfrmValidaPlan.BitBtn5Click(Sender: TObject);
var
   ini : TIniFile;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;

   re.lines.Clear;

   planJbm.ObtemProcessosCruzarGcpj(-1);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Verificando Empresa Ligada: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;
      planJbm.GravaEmpresaGrupo(1);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.BitBtn6Click(Sender: TObject);
var
   ini : TIniFile;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');

   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;

   re.lines.Clear;

   planJbm.ObtemProcessosCruzarGcpj(-1);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Revisando duplicidades: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

(*      if (pos('TRABALHISTA', planjbm.dtsatos.FieldByName('tipoandamento').AsString) <> 0) or
         (Pos('TRABALHO', planJbm.dtsatos.FieldByName('Vara').AsString) <> 0) or
         (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'INSTRUCAO') then
         planjbm.tipoprocesso := 'T'
      else
         planjbm.tipoprocesso := 'C';*)

      if Pos('PREPOSTO', planjbm.dtsatos.FieldByName('tipoandamento').AsString) = 0 then
      begin
         dmgcpcj_base_V.ObtemOutrosPagamentosDoProcesso(planJbm.gcpj,
                                                         planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                         planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString,
                                                         planjbm.tipoprocesso,
                                                         IntToStr(planJbm.codGcpjEscritorio));
         if dmgcpcj_base_V.dts.RecordCount > 0 then
         begin
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO TRABALHISTA') or
               (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AUDIÊNCIA TRABALHISTA') or
               (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'INSTRUCAO') then
            begin
               if dmgcpcj_base_V.dts.RecordCount >= 2 then
               begin
                  planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.RemoveNotaProcesso;
               end
            end
            else
            begin
               planJbm.GravaOcorrencia(2, 'Ato já foi pago anteriormente');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.RemoveNotaProcesso;
            end;
         end;
      end;
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.BitBtn7Click(Sender: TObject);
var
   arquivo : Tstringlist;
   arquivoout: TstringList;
   i, p : integer;
   estagio : integer;
   linha, valor : string;
begin
   arquivo := TStringlist.Create;
   try
      arquivo.LoadFromFile('c:\presenta\temp\atos.txt');

      arquivoOut := TstringlIst.Create;
      try
         estagio := 0;
         for i := 0 to arquivo.Count - 1 do
         begin
            case estagio of
               0 : begin
                  if Pos('Processo Ato Data do Ato  Valor Dependência', arquivo.strings[i]) = 0 then
                     continue;
                  Inc(estagio);
               end;
               1 : begin
                  if Pos('[', arquivo.strings[i]) <> 0 then
                  begin
                     estagio := 0;
                     continue;
                  end;

                  linha := Trim(arquivo.strings[i]);
                  valor := trim(copy(linha, 1, 10));

                  linha := Copy(linha, 11, length(linha));
                  p := Pos('01/11/2011', linha);
                  valor := valor + ';' + Trim(Copy(linha, 1, p-1));

                  linha := Trim(Copy(linha, p+10, length(linha)));
                  p := pos(' ', linha);
                  valor := valor + ';' + Trim(Copy(linha, 1, p-1));
                  arquivoout.add(valor);
               end;

            end;
         end;
         arquivoout.SaveToFile('c:\presenta\temp\digitados.txt');
      finally
         arquivoout.free;
      end;
   finally
      arquivo.free;
   end;

end;

procedure TfrmValidaPlan.BitBtn8Click(Sender: TObject);
var
   ini : TIniFile;
begin
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
    ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
    try
      if FormatDateTime('dd/mm/yyyy', date) <> ini.ReadString('gcpj', 'reindexacao','') then
      begin
         IndexaBasesGCPJ;
         ini.WriteString('gcpj', 'reindexacao',FormatDateTime('dd/mm/yyyy', date));
      end;
   finally
      ini.Free;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   CalculaNumeroDaNota;
   re.Lines.add('Cálculo das notas finalizado');
end;


procedure TfrmValidaPlan.CorrigirNota;
var
   wintask :  TWintask;
   fName, fname2 : string;
   valor : string;
   achou : boolean;
   i, estagio, index, p : integer;
   arquivo : TStringList;
   dataDigitar : string;
   tipoato : string;

   function TemDestinatario : boolean;
   var
      i,p : integer;
      arquivo : TStringlist;
      codigo : string;
   begin
      result := false;
      arquivo := TStringList.Create;
      try
         arquivo.LoadFromFile(fname);

         for i := 0 to arquivo.count - 1 do
         begin
            p :=  pos('despesaHonorariosProcessoVO.composicaoDespHonorariosVO.dependenciaBradescoOnLineVO.cdDestinatario" maxlength="5" size="5" value="', arquivo.strings[i]);
            if p = 0 then
               continue;
            codigo := Copy(arquivo.strings[i], p+129, length(arquivo.strings[i]));
            p := Pos('" onkeydown="', codigo);
            if p = 0 then
               exit;

            codigo := Trim(Copy(codigo, 1, p-1));
            if codigo <> '' then
               result := true;
            exit;
         end;
      finally
         arquivo.free;
      end;
   end;

   function ObtemDependenciaDigitar:string;
   var
      i,p : integer;
      arquivo : TStringlist;
      estagio : integer;
      codigo, nome : string;
      valstr : string;
      dependencia : string;
      primeira : boolean;
   begin
      result := '';
      codigo := '';
      nome:= '';
      planJbm.ObtemReclamadasDoProcesso;

      while Not planJbm.dtsReclamadas.Eof do
      begin
         if planJbm.dtsReclamadas.FieldByName('sequenciagcpj').AsInteger = 1 then //guarda a primeira por seguranca
         begin
            codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
            nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
         end;

         //é bradesco?
         if planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger = 237 then //procura uma agencia
         begin
            if planJbm.dtsReclamadas.FieldByName('codempresa').AsInteger = planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger then
            begin
               if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString = 'A' then //agencia
               begin
                  codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
            end;
         end
         else
         begin
            if (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger) or
               (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('codempresaligada').AsInteger) then
            begin
               if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A' then //não pode ser agencia
               begin
                  codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
            end;
         end;
         planJbm.dtsReclamadas.Next;
      end;

      if (codigo = '0') or (codigo = '') then
         primeira := true
      else
         primeira := false;

      arquivo := TStringList.Create;
      try
         arquivo.LoadFromFile(fname2);

         while true do
         begin
            estagio := 0;
            dependencia := '';

            for i := 0 to arquivo.count - 1 do
            begin
               case estagio of
                  0 : begin
                     if pos('<select name="selDependencia"', arquivo.strings[i]) = 0 then
                        continue;
                     Inc(estagio);
                  end;
                  1 : begin
                     p := Pos('<option value="', arquivo.strings[i]);
                     if p = 0 then
                        break;;

                     valstr := trim(Copy(arquivo.strings[i], p+15, length(arquivo.strings[i])));
                     p := Pos('">', valstr);
                     if p = 0 then
                        exit;

                     if Not primeira then
                     begin
                        if RemoveNaoNumericos(Trim(copy(valstr, 1, p-1))) <> codigo then
                           continue;
                     end;

                     valStr := Trim(Copy(valStr, p+2, length(valstr)));
                     p := pos('</option>', valStr);
                     if p = 0 then
                     begin
                        result := '';
                        exit;
                     end;

                     valStr := Copy(valStr, 1, p-1);
                     if valStr = '' then
                        ShowMessage('Erro');

                     p := Pos('&#39;', valStr);
                     if p <> 0 then
                        valstr := copy(valstr, 1, p-1) + '''' + copy(valstr, p+5, length(valstr));

                     Result := valstr;
                     exit;
                  end;
               end;
            end;
            //nao achou
            if primeira then
            begin
               Result := nome;
               exit;
            end;
            primeira := true;
         end;
      finally
         arquivo.free;
      end;
   end;

begin
   re.lines.Clear;

   if FormatDateTime('hh:nn', Now) > '22:00' then
      exit;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.cnpjescritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.nomeescritorio := planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString;
   planJbm.nomedigitar := planJbm.dtsPlanilha.FieldByName('nomedigitar').AsString;
   planJbm.codGcpjEscritorio := planJbm.dtsPlanilha.FieldByName('codgcpjescritorio').AsInteger;

   dataDigitar := planJbm.ObtemDataDigitar;
   if dataDigitar = '' then
   begin
      ShowMessage('Não encontrada a data inicial para o ano/mês de referência');
      exit;
   end;

   if numnota.text = '' then
   begin
      ShowMessage('Informe o número da nota para ser corrigida');
      numNota.SetFocus;
      exit;
   end;

   wintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                              'Q:\Publico\Backup_sistemas\Compartilhado\Honorarios_Programa\presenta\rob\',
                              '');

   wintask.FecharTodos('TaskExec.exe');
   wintask.FecharTodos('TaskLock.exe');
   wintask.FecharTodos('Tasksync.exe');
   try
      if Not planJbm.ObtemProcessosDigitar(numnota.text, true) then
      begin
         ShowMessage('Nenhum processo encontrado para digitacao');
         exit;
      end;

      planJbm.dtsAtos.tag := 0;

      if planJbm.dtsAtos.Eof then
      begin
         ShowMessage('Não encontrados atos para a nota');
         exit;
      end;

      planJbm.empresaLigada := planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger;
      planJbm.numerodanota :=  planJbm.dtsAtos.FieldByName('numeronota').AsInteger;

      //localizar a tela da nota
      fName := wintask.GetPagina_CopyText('IEXPLORE.EXE|Internet Explorer_Server|GCPJ - Windows Internet Explorer',
                                                   'Alteração');
      if fName = '' then //esta na pagina errada
      begin
         ShowMessage('Erro posicionando na tela de alteração');
         exit;
      end;

      try
         arquivo := TStringList.Create;
         try
            arquivo.LoadFromFile(fName);

            estagio := 0;

            Index := 0;


            for i := 0 to arquivo.Count - 1 do
            begin
               case estagio of
                  0 : begin
                     if Pos('Processo Ato Repasse Afeta Resultado Data do Ato Valor Dependência', arquivo.strings[i]) = 0 then
                        continue;
                     Inc(estagio);
                  end;
                  1 : begin
                     valor := Trim(arquivo.Strings[i]);
                     if valor = '' then
                        continue;

                     p := Pos(']', valor);
                     if p <> 0 then
                     begin
                        //é ultima página?
                        valor := Trim(Copy(valor, p+1, length(valor)));
                        if valor = '' then
                        begin
                           Inc(passagem);
                           if Passagem > 3 then
                           begin
                              wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                              Sleep(100);

                              wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                              Sleep(100);

                              ShowMessage('Fim');
                              exit;
                           end;

                           wintask.ClickHTMLElement('GCPJ', 'IMG[SRC= ''http://www4.net.bradesco.com.br/imagens/setaTabelaD'']');
                           Sleep(100);
                           CorrigirNota;
                           exit;
                        end;

                        passagem := 0;
                        //mudar de pagina
                        p := Pos(' ', valor);
                        if p <> 0 then
                           valor := Trim(copy(valor, 1, p-1));

                        wintask.ClickHTMLElement('GCPJ', 'A[INNERTEXT= ''' + valor + ''']');
                        ultimaPagina := StrToInt(valor);
                        sleep(100);
                        CorrigirNota;
                        exit;
                     end;

                     tipoato := Trim(Copy(valor, 11, Length(valor)));
                     p := Pos(' ', tipoato);
                     if p <> 0 then
                        tipoato := Trim(Copy(tipoato, 1, p-1));

                     valor := IntToStr(StrToInt(Copy(valor, 1, 10)));

                     (*                     if StrToInt(valor) < ultimoGcpj then
                     begin
                        wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                        Sleep(100);

                        wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''voltar'']');
                        Sleep(100);

                        ShowMessage('Fim');
                        exit;
                     end;*)

//                     ultimoGcpj := StrToInt(valor);

                     achou := false;
                     while Not planJbm.dtsAtos.Eof do
                     begin
                        if valor <> planJbm.dtsAtos.FieldByName('gcpj').AsString then
                        begin
                           planJbm.dtsAtos.Next;
                           continue;
                        end;
                        achou := true;
                        break;
                     end;

                     if not achou then
                     begin
                        ShowMessage('Erro inesperado');
                        exit;
                     end;

                     if planJbm.JaAlterado(tipoAto) then
                     begin
                        Inc(index);
                        continue;
                     end;

                     achou := false;

                     planJbm.CarregaCamposDaTabela;
                     re.lines.add('Digitando processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
                     Application.ProcessMessages;

                     planJbm.ObtemReclamadasDoProcesso;
                     //tem agencia?
                     while not planJbm.dtsReclamadas.Eof do
                     begin
                        if planJbm.dtsReclamadas.FieldByName('codempresa').AsInteger <> 237 then
                        begin
                           planJbm.dtsReclamadas.next;
                           continue;
                        end;
                        if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A' then
                        begin
                           planJbm.dtsReclamadas.next;
                           continue;
                        end;
                        achou := true;
                        break;
                     end;

                     if Not achou then //não precisa alterar
                     begin
                        planJbm.SalvaAlterado(tipoato);
                        Inc(Index);
                        continue;
                     end;

                     //se já é a mesma não precisa alterar
                     if Pos(planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString + ' - ', arquivo.Strings[i]) <> 0 then
                     begin
                        planJbm.SalvaAlterado(tipoato);
                        Inc(Index);
                        continue;
                     end;


                     //altera
                     if index = 0 then
                        wintask.ClickHTMLElement('GCPJ', 'INPUT RADIO[NAME= ''composicoesSelecs'']')
                     else
                        wintask.ClickHTMLElement('GCPJ', 'INPUT RADIO[NAME= ''composicoesSelecs'', INDEX=''' + IntToStr(index+1) + ''']');
                     Sleep(100);

                     wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''alterar'']');
                     sleep(100);

                     fname2 := wintask.EstaNaPagina_NotepadCodePage('IEXPLORE.EXE|Internet Explorer_Server|GCPJ - Windows Internet Explorer|1',
                                                       '787','297',
                                                       '856','490',
                                                       'NOTEPAD.EXE|Notepad|finHonorariosNovoAction',
                                                       'Documento do Prestador: ',
                                                       false);
                     if fname2 = '' then
                     begin
                        re.lines.Add('Erro');
                        exit;
                     end;

                     try
                        valor := ObtemDependenciaDigitar;
                        if valor = '' then
                        begin
                           re.lines.Add('Erro');
                           exit;
                        end;

                        wintask.SelectHTMLItem('GCPJ', 'SELECT[NAME= ''selDependencia'']', Trim(valor));
                        Sleep(180);

                        if wintask.Verifica_Janela_PopUp('IEXPLORE.EXE|Static|Esta dependência foi encerrada, deseja alterá-la para', valor) = 1 then
                        begin
                           wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                           Sleep(20);
                        end;


                        wintask.ClickHTMLElement('GCPJ', 'INPUT SUBMIT[VALUE= ''incluir'']');
                        Sleep(100);

                        wintask.ClickHTMLElement('GCPJ - Gestão e Controle de Processos Jurídicos', 'INPUT BUTTON[VALUE= ''confirmar inclusão'']');
                        Sleep(100);

                        planJbm.SalvaAlterado(tipoato);
                        Inc(index);
                     finally
                        DeleteFile(fname2);
                     end;
                  end;
               end;
            end;
         finally
            arquivo.Free;
         end;
      finally
         DeleteFile(fname);
      end;
   finally
      wintask.Free;
   end;
end;
procedure TfrmValidaPlan.VerificaDuplicidades;
begin
   re.lines.Clear;

   processando.Text := 'Verificando processos duplicados na planilha';
   Application.ProcessMessages;

   planJbm.ObtemProcessosCruzarGcpj(4);

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.Lines.add('Verificando duplicidade do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      if planJbm.TemProcessoDuplicado then
      begin
         planJbm.GravaOcorrencia(2, 'Ato em duplicidade na planilha');
            planJbm.MarcaProcessoCruzadoGcpj(9);
      end;
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.RemoveEstadosDeVolumetria;
var
   estado : string;
   orgaoJulgador : integer;
   semVolumetria : boolean;
   valorcorreto : double;
begin
   //rotina desabilitada a pedido da Nilce - E-mail recebido em 03/09/2012
   exit;

   planJbm.ObtemProcessosCruzarGcpj(5);
   if planJbm.dtsAtos.RecordCount = 0 then
      exit;
   re.lines.Clear;

   processando.Text := 'Excluindo estados do calculo de volumetria';
   Application.ProcessMessages;

   while not planJbm.dtsAtos.Eof do
   begin
      Application.ProcessMessages;
      re.Lines.add('Verificando comarca do processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      planjbm.CarregaCamposDaTabela;

      orgaoJulgador := planJbm.ObtemCodOrgProcesso;
      semvolumetria := false;

      if orgaoJulgador = 0 then
      begin
         planJbm.dtsAtos.Next;
         continue;
      end;

      estado := planJbm.ObtemUfProcesso(orgaoJulgador);

      if estado = '' then
      begin
         planjbm.dtsAtos.next;
         continue;
      end;

      //remove estados que não pode considerar volumetria
      if (planJbm.cnpjEscritorio = '08833747000132') or //sette camra
         (planJbm.cnpjEscritorio = '10508423000170')then //jbm
      begin
         if (estado = 'MG') or
            (estado = 'AM') OR
            (estado = 'PA') or
            (estado = 'RO') OR
            (estado = 'RR') or
            (estado = 'AP') OR
            (estado = 'AC') or
            (estado = 'AL') or
            (estado = 'BA') or
            (estado = 'CE') or
            (estado = 'MA') or
            (estado = 'PB') or
            (estado = 'PE') or
            (estado = 'PI') or
            (estado = 'RN') or
            (estado = 'SE') or
            (estado = 'TO') then
            semvolumetria := true;
      end
      else
      begin
         if (planJbm.cnpjEscritorio = '05159996000104') then //rocha marinho
         begin
            if (estado = 'AM') OR
               (estado = 'PA') or
               (estado = 'RO') OR
               (estado = 'RR') or
               (estado = 'AP') OR
               (estado = 'AC') or
               (estado = 'AL') or
               (estado = 'BA') or
               (estado = 'CE') or
               (estado = 'MA') or
               (estado = 'PB') or
               (estado = 'PE') or
               (estado = 'PI') or
               (estado = 'RN') or
               (estado = 'SE') or
               (estado = 'TO') then
               semvolumetria := true;
         end;
      end;

      if not semvolumetria then
      begin
         planjbm.dtsAtos.Next;
         continue;
      end;

      if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2011/05/01' then //utlizar a rotina de novo cálculo
      begin
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) <= '2012/04/30' then //utlizar a rotina de novo cálculo
            valorCorreto := 422
         else
            valorCorreto := 438;

         if planjbm.dtsatos.FieldByName('valor').AsFloat <> valorcorreto then
         begin
            planJbm.GravaValorCorrigido(valorcorreto);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2010/05/01' then //utlizar a rotina de novo cálculo
      begin
         if planjbm.dtsatos.FieldByName('valor').AsFloat <> 382 then
         begin
            planJbm.GravaValorCorrigido(382);
            planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
         end
         else
            planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
         planJbm.MarcaProcessoCruzadoGcpj(6);
         planJbm.dtsAtos.Next;
         continue;
      end;

      if planjbm.dtsatos.FieldByName('valor').AsFloat <> 370 then
      begin
         planJbm.GravaValorCorrigido(370);
         planJbm.GravaOcorrencia(2, 'Valor alterado pelo sistema');
      end
      else
         planJbm.GravaValorCorrigido(planjbm.dtsatos.FieldByName('valor').AsFloat);
      planJbm.MarcaProcessoCruzadoGcpj(6);
      planJbm.dtsAtos.Next;
   end;
end;

procedure TfrmValidaPlan.CopiaBasesGCPJ;
var
   ini : TiniFile;
   arquivoOrigem : string;
   arquivoDestino : string;
begin
   //copia as bases para o sistema
//   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');

   try
      re.lines.Add('Copiando a base de dados do GCPJ Base I. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := 'Q:\Publico\Backup_sistemas\_BaseDiaria\basediaria_gcpj_i.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_i', 'databasename', '');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      re.lines.Add('Copiando a base de dados do GCPJ Base II. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := 'Q:\Publico\Backup_sistemas\_BaseDiaria\basediaria_gcpj_ii.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_ii', 'databasename', '');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      re.lines.Add('Copiando a base de dados do GCPJ Base III. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := 'Q:\Publico\Backup_sistemas\_BaseDiaria\basediaria_gcpj_iii.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_iii', 'databasename', '');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      re.lines.Add('Copiando a base de dados do GCPJ Base IV. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := 'Q:\Publico\Backup_sistemas\_BaseDiaria\basediaria_gcpj_iv.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_iv', 'databasename', '');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      re.lines.Add('Copiando a base de dados do GCPJ Base V. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := 'Q:\Publico\Backup_sistemas\_BaseDiaria\basediaria_gcpj_v.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_v', 'databasename', '');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);


      re.lines.Add('Copiando a base de dados do GCPJ Base PROCESSOS. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := 'Q:\Publico\Backup_sistemas\Compartilhado\Gerador_Relatorios\_DADOS\Relatorios_Gerenciais_DADOS_PROCESSOS.mdb';
      arquivoDestino := ini.ReadString('gcpj_base_compartilhada', 'databasename', '');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);

      re.lines.Add('Copiando a base de dados do Honorarios. Aguarde.....');
      Application.ProcessMessages;

      arquivoOrigem := 'Q:\Publico\Backup_sistemas\Compartilhado\Honorarios_Programa\Honorarios_Programa_II.mdb';
      arquivoDestino := ini.ReadString('honorarios_programas', 'databasename', '');

      if FileExists(arquivoDestino) then
         deleteFile(arquivoDestino);
      CopyFile(PChar(arquivoOrigem), PChar(arquivoDestino), false);
   finally
      ini.Free;
   end;
end;

procedure TfrmValidaPlan.BitBtn4Click(Sender: TObject);
begin
   if MessageDlg('Tem certeza que deseja excluir a planilha?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      exit;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.RemovePlanilhas(planJbm.idplanilha, planJbm.idEscritorio, planJbm.anomesreferencia);
   planJbm.dtsPlanilha.Close;
   planJbm.dtsPlanilha.Open;
   ShowMessage('Planilha removida com sucesso');
end;

end.
