unit fimportaplan;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls, XlSReadWriteII5, presenta,
  datamodule_honorarios, mywin, uplanjbm, mgenlib, fcadescritorio, fcadastravaloresnaoatualizar;

type
  TfrmImpPlan = class(TForm)
    Label1: TLabel;
    planName: TEdit;
    SpeedButton1: TSpeedButton;
    od: TOpenDialog;
    Label2: TLabel;
    cbEscritorios: TComboBox;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    re: TRichEdit;
    Label3: TLabel;
    cbMes: TComboBox;
    Label4: TLabel;
    cbAno: TComboBox;
    BitBtn3: TBitBtn;
    nomePesq: TEdit;
    Label5: TLabel;
    BitBtn4: TBitBtn;
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure nomePesqChange(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmImpPlan: TfrmImpPlan;

implementation

//uses SheetData5;

{$R *.dfm}

procedure TfrmImpPlan.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfrmImpPlan.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   action := caFree;
end;

procedure TfrmImpPlan.BitBtn1Click(Sender: TObject);
var
   xls : TXLSReadWriteII5;
   linha, idplan: integer;
   planJbm : TPlanilhaJBM;
   p : integer;
   valor : string;

begin
   if planName.Text = '' then
   begin
      ShowMessage('Informe o nome da planilha');
      planName.SetFocus;
      exit;
   end;

   if Not FileExists(planName.text) then
   begin
      ShowMessage('Arquivo inv�lido ou inexistente');
      planName.SetFocus;
      exit;
   end;

   if cbEscritorios.ItemIndex = -1 then
   begin
      ShowMessage('Selecione o nome do escrit�rio');
      cbEscritorios.SetFocus;
      exit;
   end;

   if cbMes.ItemIndex = -1 then
   begin
      ShowMessage('Informe o m�s de refer�ncia');
      cbMes.SetFocus;
      exit;
   end;

   if cbAno.ItemIndex = -1 then
   begin
      ShowMessage('Informe o ano de refer�ncia');
      cbAno.SetFocus;
      exit;
   end;

   //verifica se a coluna ValorBase existe na tabela tbplanilhasatos;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 valorbase from tbplanilhasatos';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if Pos('valorbase', merr.Message) <> 0 then
            dmHonorarios.CriaColunaValorBase;
      end;
   end;

   //verifica se a coluna identificador existe na tabela tbESCRITORIOS;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 identificador from tbescritorios';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if Pos('identfificador', merr.Message) <> 0 then
            dmHonorarios.CriaColunaIdentificador;
      end;
   end;

   //verifica se a coluna CODEXfUNCIONARIO  existe na tabela tbplanilhasreclamadas;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 codexfuncionario from tbplanilhasreclamadas';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if Pos('codexfuncionario', merr.Message) <> 0 then
            dmHonorarios.CriaColunaCodExFuncionario;
      end;
   end;

   planJbm := TPlanilhaJBM.Create;
   try
      planJbm.idEscritorio := TIdObject(cbEscritorios.Items.Objects[cbEscritorios.ItemIndex]).Id;
      planJbm.anomesreferencia := cbAno.Text + '/' + LeftZeroStr('%02d', [cbMes.ItemIndex+1]);
      p := Pos('-', cbEscritorios.Text);
      planJbm.nomeescritorio := Trim(Copy(cbEscritorios.Text, 1, p-1));

      idPlan := planJbm.JaExistePlanilhaDoMes;
      if idPlan <> 0 then
      begin
         if MessageDlg('J� existe planilha do escrit�rio [' +  cbEscritorios.Text + '] para o m�s de referencia: ' + cbAno.Text + '/' + LeftZeroStr('%02d', [cbMes.ItemIndex+1]) + '. Deseja sobrepor?',
                       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
         begin
            re.Lines.Add('Removendo dados da planilha anterior...Aguarde.');
            planJbm.RemovePlanilhas(idplan, planJbm.idEscritorio, planJbm.anomesreferencia);
         end
         else
         begin
            if MessageDlg('Cria nova planilha de importa��o?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
            begin
               ShowMessage('Importa��o interrompida pelo usu�rio');
               exit;
            end;
         end;
      end;

      planJbm.CadastraPlanilha;

      planJbm.RemoveOcorrenciasDaPlanilha;

      re.lines.clear;

      xls := TXLSReadWriteII5.Create(Nil);
      try
         xls.Filename := planName.Text;
         xls.Read;

         linha := 1;

         while true do
         begin
            re.lines.Add('Processando linha: ' + InttoStr(linha+1));
           Application.ProcessMessages;

            if (xls.Sheets[0].AsString[0, linha] = '') and
               (xls.Sheets[0].AsString[1, linha] = '') then
               break;

            planJbm.Cliente := Uppercase(xls.Sheets[0].AsString[0, linha]);

            valor := UpperCase(xls.Sheets[0].AsString[1, linha]);
            RemoveEsteCaracter('''', valor);
            RemoveEsteCaracter('`', valor);
            RemoveEsteCaracter('"', valor);
            planJbm.partecontraria := valor;

            planJbm.processo := xls.Sheets[0].AsString[2, linha];
            planJbm.gcpj := xls.Sheets[0].AsString[3, linha];

            valor := UpperCase(xls.Sheets[0].AsString[4, linha]);
            RemoveEsteCaracter('''', valor);
            planJbm.Vara := valor;

            valor := Uppercase(xls.Sheets[0].AsString[5, linha]);
            RemoveEsteCaracter('''', valor);
            planJbm.Comarca := valor;

            planJbm.UF := UpperCase(xls.Sheets[0].AsString[6, linha]);
            planJbm.tipoandamento := Trim(UpperCase(xls.Sheets[0].AsString[7, linha]));

            if (Uppercase(planJbm.tipoandamento) = 'CONTESTA��O') or
               (Uppercase(Planjbm.tipoandamento) = 'CONTESTACAO') or
               (UpperCase(planJbm.tipoandamento) = 'CONTESTA�AO') then
               planJbm.tipoandamento := 'TUTELA'
            else
            if (Uppercase(planJbm.tipoandamento) = 'INSTRU��O') or
               (Uppercase(Planjbm.tipoandamento) = 'INSTRUCAO') or
               (Uppercase(Planjbm.tipoandamento) = 'INSTRU�AO') then
               planJbm.tipoandamento := 'INSTRUCAO'
            else
//            if (Uppercase(planJbm.tipoandamento) = 'ACORDO') then
//               planJbm.tipoandamento := 'EXITO'
//            else
            if (uppercase(planJbm.tipoandamento) = 'RECURSO EM FASE DE EXECUCAO') or
               (uppercase(planJbm.tipoandamento) = 'RECURSO EM FASE DE EXECU��O') or
               (uppercase(planJbm.tipoandamento) = 'RECURSO NA FASE DE EXECU��O') or
               (uppercase(planJbm.tipoandamento) = 'RECURSO NA FASE DE EXECUCAO') then
               planJbm.tipoandamento := 'RECURSO NA FASE DE EXECUCAO'
            else
            if Pos('RECURSO', uppercase(planJbm.tipoandamento)) <> 0 THEN
            //altera��o efetuada em 09/10/2013 - Ver email
            begin
                if (uppercase(planJbm.tipoandamento) <> 'RECURSO EM FASE DE EXECUCAO') and
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO NA FASE DE EXECUCAO') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO APELA��O') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO APELACAO') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO DE APELA��O') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO DE APELACAO') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO ESPECIAL') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO ESTRAORDIN�RIO') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO NOS TRIBUNAIS SUPERIORES') AND
                   (uppercase(planJbm.tipoandamento) <> 'RECURSO EXTRAORDINARIO') then
                begin
                   planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha - Tipo de ato inv�lido: ' + planJbm.tipoandamento +  ' no GCPJ: ' + planJbm.gcpj);
                   Inc(linha);
                   continue;
                end;
            end
            else
            if Pos('CONTRARRAZ', uppercase(planJbm.tipoandamento)) <> 0 then
            begin
                if (uppercase(planJbm.tipoandamento) <> 'CONTRARRAZ�ES') AND
                   (uppercase(planJbm.tipoandamento) <> 'CONTRARRAZOES') then
                begin
                   planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha - Tipo de ato inv�lido: ' + planJbm.tipoandamento +  ' no GCPJ: ' + planJbm.gcpj);
                   Inc(linha);
                   continue;
                end;
            end;

            if (uppercase(planJbm.tipoandamento) = 'RECURSO APELA��O') OR
               (uppercase(planJbm.tipoandamento) = 'RECURSO APELACAO') OR
               (uppercase(planJbm.tipoandamento) = 'RECURSO DE APELA��O') OR
               (uppercase(planJbm.tipoandamento) = 'RECURSO DE APELACAO') OR
               (uppercase(planJbm.tipoandamento) = 'RECURSO ESPECIAL') OR
               (uppercase(planJbm.tipoandamento) = 'RECURSO EXTRAORDIN�RIO') OR
               (uppercase(planJbm.tipoandamento) = 'RECURSO EXTRAORDINARIO') or
               (uppercase(planJbm.tipoandamento) = 'RECURSO NOS TRIBUNAIS SUPERIORES') OR
               (uppercase(planJbm.tipoandamento) = 'CONTRARRAZ�ES') OR
               (uppercase(planJbm.tipoandamento) = 'CONTRARRAZOES') then
                planJbm.tipoandamento := 'RECURSO';

            if uppercase(planJbm.tipoandamento) = 'BUSCA E APREENS�O/REINTEGRA��O DE POSSE' then
               planJbm.tipoandamento := 'BUSCA E APREENSAO/REINTEGRACAO DE POSSE';

            if uppercase(planJbm.tipoandamento) = 'ADJUDICA��O/ARREMATA��O' then
               planJbm.tipoandamento := 'ADJUDICACAO/ARREMATACAO';

            if uppercase(planJbm.tipoandamento) = 'ENTREGA AMIG�VEL' then
               planJbm.tipoandamento := 'ENTREGA AMIGAVEL';

            if uppercase(planJbm.tipoandamento) = 'IRRECUPERABILIDADE AT� 1 ANO' THEN
               planJbm.tipoandamento := 'IRRECUPERABILIDADE ATE 1 ANO';

            if uppercase(planJbm.tipoandamento) = 'IRRECUPERABILIDADE AP�S 1 ANO' then
               planJbm.tipoandamento := 'IRRECUPERABILIDADE APOS 1 ANO';

            if uppercase(planJbm.tipoandamento) = 'CARTA PRECAT�RIA' then
               planJbm.tipoandamento := 'CARTA PRECATORIA';

            if (uppercase(planJbm.tipoandamento) = 'CUMPRIMENTO DE CARTA PRECAT�RIA') OR
               (uppercase(planJbm.tipoandamento) = 'CUMPRIMENTO DE CARTA PRECATORIA') then
               planJbm.tipoandamento := 'CUMPRIMENTO DE CARTA PRECATORIA';

            if uppercase(planJbm.tipoandamento) = 'CONSOLIDA��O DE PROPRIEDADE' then
               planJbm.tipoandamento := 'CONSOLIDACAO DE PROPRIEDADE';

            if uppercase(planJbm.tipoandamento) = 'CONSOLIDA��O DE PROPRIEDADE (2 P/C ADIC)' then
               planJbm.tipoandamento := 'CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)';

            if uppercase(planJbm.tipoandamento) = 'HONOR�RIOS DE VENDA' then
               planJbm.tipoandamento := 'HONORARIOS DE VENDA';

            if uppercase(planJbm.tipoandamento) = 'LEVANTAMENTO DE ALVAR�' then
               planJbm.tipoandamento := 'LEVANTAMENTO DE ALVARA';

            if uppercase(planJbm.tipoandamento) = 'CONSIGNA��O BANC�RIA' then
               planJbm.tipoandamento := 'CONSIGNACAO BANCARIA';

            if uppercase(planJbm.tipoandamento) = 'PRO LABORE' then
               planJbm.tipoandamento := 'HONORARIOS INICIAIS';

            if ((uppercase(planJbm.tipoandamento) = 'INSTRU��O') or (uppercase(planJbm.tipoandamento) = 'INSTRUCAO')) then
               planJbm.tipoandamento := 'INSTRUCAO';

            if uppercase(planJbm.tipoandamento) = 'ACOMPANHAMENTO MENSAL' then
               planJbm.tipoandamento := 'ACOMPANHAMENTO';

            if uppercase(planJbm.tipoandamento) = 'ACOMPANHAMENTO' then
               planJbm.tipoandamento := 'PARCELA';

            if uppercase(planJbm.tipoandamento) = 'HONORARIOS DE EXITO (IMPROCEDENCIA)' then
               planJbm.tipoandamento := 'EXITO';

            if uppercase(planJbm.tipoandamento) = 'HONORARIOS DE EXITO (ACORDO REALIZADO ANTES DA AUDIENCIA)' then
               planJbm.tipoandamento := 'EXITO';

            if uppercase(planJbm.tipoandamento) = 'HONORARIOS DE EXITO (ACORDO REALIZADO ENTRE DA AUDIENCIA)' then
               planJbm.tipoandamento := 'EXITO';

            if uppercase(planJbm.tipoandamento) = 'SUSTENTA��O ORAL' then
               planJbm.tipoandamento := 'SUSTENTACAO ORAL';

            if uppercase(planJbm.tipoandamento) = 'DEFESA E RECURSO ADM. EM AUTO DE INFRA��O' then
               planJbm.tipoandamento := 'DEFESA ADMINISTRATIVO';

            if uppercase(planJbm.tipoandamento) = 'MANDADO DE SEGURAN�A A��O RESCIS�RIA E INQ. PARA APURA��O FALTA GRAVE' then
               planJbm.tipoandamento := 'MANDADO DE SEGURANCA';

            if uppercase(planJbm.tipoandamento) = 'A��O CONSIGNAT�RIA DECLAT�RIA E INTERDITO PROIBIT�RIO' then
               planJbm.tipoandamento := 'INTERDITO PROIBITORIO';

            planJbm.idprocessounico := Trim(xls.Sheets[0].AsString[8, linha]);
            planJbm.divisaodafilial := xls.Sheets[0].AsString[9, linha];
            try
               planJbm.datadoato := xls.Sheets[0].AsDateTime[10, linha];
            except
               try
                  planJbm.datadoato := StrToDate(Trim(xls.Sheets[0].AsString[10, linha]));
               except
                  planJbm.datadoato := 0;
               end;
            end;

            planJbm.linhaPlanilha := linha+1;

//            if planjbm.datadoato = 0 then
//            begin
//               try
//                  planJbm.datadoato := StrToDate(xls.Sheets[0].AsString[10, linha]);
//               except
//                  planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha com data inv�lida (erro de formata��o), GCPJ: ' + planJbm.gcpj);
//                  Inc(linha);
//                  continue;
//               end;
//            end;

            if planjbm.datadoato = 0 then
            begin
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha com data inv�lida (igual a zero), GCPJ: ' + planJbm.gcpj);
               Inc(linha);
               continue;
            end
            else
            begin
               try
                  DateToStr(planJbm.datadoato);
               except
                  planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha com data inv�lida(erro de formata��o) , GCPJ: ' + planJbm.gcpj);
                  Inc(linha);
                  continue;
               end;
            end;

            if FormatDateTime('yyyymmdd',  planJbm.datadoato) > FormatDateTime('yyyymmdd', date) then
            begin
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha com data inv�lida (data do ato maior que dia atual), GCPJ: ' + planJbm.gcpj);
               Inc(linha);
               continue;
            end;

            if FormatDateTime('yyyy',  planJbm.datadoato) < '2000' then
            begin
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha com data inv�lida (data do ato inferior ao ano 2000), GCPJ: ' + planJbm.gcpj);
               Inc(linha);
               continue;
            end;

            if (DayOfWeek(planJbm.datadoato)=1) or //domingo
               (DayOfWeek(planJbm.datadoato)=7) then //s�bado
            begin
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha com data inv�lida (data do ato � final de semana), GCPJ: ' + planJbm.gcpj);
               Inc(linha);
               continue;
            end;

            if EhFeriadoNacional(planJbm.datadoato) then
            begin
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(linha+1) + ' da planilha com data inv�lida (data do ato � feriado nacional), GCPJ: ' + planJbm.gcpj);
               Inc(linha);
               continue;
            end;

            planJbm.valBeneficioEconomico := xls.Sheets[0].AsString[11, linha];
            planJbm.Porcentagem := xls.Sheets[0].AsString[12, linha];
            try
               planJbm.Valor := xls.Sheets[0].AsFloat[13, linha];
            except
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(planJbm.linhaPlanilha) + ' da planilha com formata��o do campo valor inv�lida');
               Inc(linha);
               continue;
            end;

            try
               planJbm.ValorBase := xls.Sheets[0].AsFloat[14, linha];
            except
               planJbm.ValorBase := 0;
            end;

            if Pos('/', planJbm.gcpj) <> 0 then
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(planJbm.linhaPlanilha) + ' da planilha com mais de um numero GCPJ: ' + planJbm.gcpj)
            else
            if RemoveNaoNumericos(planJbm.gcpj) = '' then
               planJbm.GravaOcorrencia(1, 'Linha ' + IntToStr(planJbm.linhaPlanilha) + ' da planilha com numero de GCPJ inconsistente: ' + planJbm.gcpj)
            else
               planJbm.CadastraAto;
            Inc(linha)
         end;

         planJbm.MarcaPlanilhaImportada;
      finally
         try
            xls.Free;
         except
         end;
      end;
   finally
      planJbm.Free;
   end;
   re.Lines.add('Importa��o finalizada com sucesso');
end;

procedure TfrmImpPlan.SpeedButton1Click(Sender: TObject);
begin
   if od.Execute then
      planName.Text := od.FileName;
end;

procedure TfrmImpPlan.BitBtn3Click(Sender: TObject);
var
   Escritorio: TIdObject;
begin
   frmCadEscritorio := TfrmCadEscritorio.Create(nil);
   try
      frmCadEscritorio.ShowModal;
      (*begin
         dmHonorarios.CadastraEscritorio(frmCadEscritorio.cnpjEscritorio.Text,
                                         frmCadEscritorio.nomeescritorio.Text,
                                         frmCadEscritorio.codigogcpj.Text,
                                         Copy(frmCadEscritorio.nomeescritorio.Text, 1, 20));

   *)
      dmHonorarios.ObtemEscritoriosCadastrados(true);
      cbEscritorios.Clear;
      while  Not dmHonorarios.dts.Eof do
      begin
         Escritorio :=TIdObject.Create;
         Escritorio.id := dmHonorarios.dts.FieldByName('idescritorio').AsInteger;
         cbEscritorios.Items.AddObject(dmHonorarios.dts.FieldByName('nomeescritorio').AsString + '-' + dmHonorarios.dts.FieldByName('cnpjescritorio').AsString, escritorio);
         dmHonorarios.dts.next;
      end;
   finally
      frmCadEscritorio.Free;
   end;
end;

procedure TfrmImpPlan.nomePesqChange(Sender: TObject);
var
   i : integer;
begin
   if nomePesq.tag = 0 then
   begin
      nomePesq.Tag := nomePesq.tag + 1;
      exit;
   end;

   if nomePesq.tag < 3 then
   begin
      nomePesq.Tag := nomePesq.Tag + 1;
      exit;
   end;

   nomePesq.tag := 0;

   for i := 0 to cbEscritorios.items.Count - 1 do
   begin
      Application.ProcessMessages;
      if Pos(nomePesq.Text, cbEscritorios.Items[i]) <> 0 then
      begin
         cbEscritorios.ItemIndex := i;
         exit;
      end;
   end;
end;

procedure TfrmImpPlan.BitBtn4Click(Sender: TObject);
begin
   //verifica se a tabela tbtiposnaoatualizar existe;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 identificador from tbtiposnaoatualizar';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O mecanismo de banco de dados', merr.Message) <> 0) and
            (Pos('n�o encontrou a tabela', merr.Message) <> 0) and
            (Pos('tbtiposnaoatualizar', merr.Message) <> 0) then
            dmHonorarios.CriaTabelaTiposNaoAtualizar
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se o campo datadoato existe na tabela tbtiposnaoatualizar existe;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 datadoato from tbtiposnaoatualizar';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O par�metro datadoato n�o tem valor padr�o', merr.Message) <> 0) then
            dmHonorarios.CriaColunaDataDoAto_TiposNaoAtualizar
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   //verifica se a tabela tbvaloresnaoatualizar existe;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 tipoandamento from tbvaloresnaoatualizar';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O mecanismo de banco de dados', merr.Message) <> 0) and
            (Pos('n�o encontrou a tabela', merr.Message) <> 0) and
            (Pos('tbvaloresnaoatualizar', merr.Message) <> 0) then
            dmHonorarios.CriaTabelaValoresNaoAtualizar
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1  identificador from tblinkescritoriosvalores';
      dmHonorarios.dts.Open;
   except
      on merr : Exception do
      begin
         if (Pos('O mecanismo de banco de dados', merr.Message) <> 0) and
            (Pos('n�o encontrou a tabela', merr.Message) <> 0) and
            (Pos('tblinkescritoriosvalores', merr.Message) <> 0) then
            dmHonorarios.CriaTabelaLinkEscritoriosValores
         else
         begin
            ShowMessage(merr.Message);
            exit;
         end;
      end;
   end;

   dmHonorarios.InsereTiposNaoAtualizarPadrao;
   dmHonorarios.InsereValoresNaoAtualizarPadrao;

   //verifica se a coluna identificador existe na tabela tbESCRITORIOS;
   try
      dmHonorarios.dts.Close;
      dmHonorarios.dts.CommandText := 'select top 1 identificador from tbescritorios';
      dmHonorarios.dts.Open;
      //existe a coluna. Tem que remover
      dmHonorarios.ExcluiColunaIdentificador;
   except
  (**    on merr : Exception do
      begin
         if Pos('identificador', merr.Message) <> 0 then
            dmHonorarios.CriaColunaIdentificador;
      end;*)
   end;

   frmCadastraValoresNaoAtualizar := TfrmCadastraValoresNaoAtualizar.Create(nil);
   try
      dmHonorarios.ObtemTiposNaoAtualizarCadastrados;
      frmCadastraValoresNaoAtualizar.ShowModal;
   finally
      frmCadastraValoresNaoAtualizar.Free;
   end;
end;

end.
