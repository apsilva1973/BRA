unit frelatorios;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Buttons, Grids, DBGrids, uplanjbm, inifiles,
  datamodulegcpj_base_i, Func_Wintask_Obj, mgenlib, XlsReadWriteII5, datamodule_honorarios,mywin,datamodulegcpj_base_ix,
  Menus;

type
  TfrmRelatorios = class(TForm)
    Label1: TLabel;
    DBGrid1: TDBGrid;
    BitBtn2: TBitBtn;
    re: TRichEdit;
    BitBtn4: TBitBtn;
    Label2: TLabel;
    numNota: TEdit;
    cbxRelatorios: TComboBox;
    Label3: TLabel;
    nomePlan: TEdit;
    SpeedButton1: TSpeedButton;
    sd: TSaveDialog;
    Label4: TLabel;
    mesRef: TComboBox;
    anoRef: TComboBox;
    ppMenu: TPopupMenu;
    Retornarplanilhaparadigitao1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure cbxRelatoriosChange(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Retornarplanilhaparadigitao1Click(Sender: TObject);
  private
    FplanJbm: TPlanilhaJBM;
    procedure SetplanJbm(const Value: TPlanilhaJBM);
    { Private declarations }
  public
    { Public declarations }
    property planJbm : TPlanilhaJBM read FplanJbm write SetplanJbm;
    procedure GeraRelatorioGeralInconsistencia;
  end;


var
  frmRelatorios: TfrmRelatorios;
  ie8 : boolean;
  ultimaPagina, passagem : integer;

implementation

uses fcaddatainicial, DB;


{$R *.dfm}

{ TfrmValidaPlan }

procedure TfrmRelatorios.SetplanJbm(const Value: TPlanilhaJBM);
begin
  FplanJbm := Value;
end;

procedure TfrmRelatorios.FormCreate(Sender: TObject);
begin
   planJbm := TPlanilhaJBM.Create;
   DBGrid1.DataSource := planJbm.dsPlan;
   planJbm.ObtemPlanilhasCadastradas(1);
   if planJbm.dtsPlanilha.RecordCount = 0 then
      BitBtn4.Enabled := false;

   if IsParam('/ie8') then
      ie8 := true
   else
      ie8 := false;
end;

procedure TfrmRelatorios.BitBtn2Click(Sender: TObject);
begin
   close;
end;

procedure TfrmRelatorios.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   planJbm.Finalizar;
   planJbm.Free;
   action := CaFree;
end;


procedure TfrmRelatorios.BitBtn4Click(Sender: TObject);
var
   xls : TXlsReadWriteII5;
begin
   re.lines.Clear;

   if cbxRelatorios.ItemIndex = 4 then
   begin
      GeraRelatorioGeralInconsistencia;
      exit;
   end;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;
   planJbm.cnpjescritorio := planJbm.dtsPlanilha.FieldByName('cnpjescritorio').AsString;
   planJbm.nomeescritorio := planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString;
   planJbm.nomedigitar := planJbm.dtsPlanilha.FieldByName('nomedigitar').AsString;

   case cbxRelatorios.ItemIndex of
      0 : planJbm.ObtemInconsistenciasImportacao;
      1 : planJbm.ObtemInconsistenciasPosImportacao;
      2 : planJbm.ObtemValoresRecalculadosSistema;
      3 : planJbm.ObtemNotasFinalizadas(numNota.text);
   end;

   if planJbm.dtsAtos.Eof then
   begin
      ShowMessage('Nenhum registro encontrado');
      exit;
   end;

   xls := TXLSReadWriteII5.Create(nil);
   try
      xls.Filename := nomePlan.Text;
      planJbm.dtsAtos.tag := 1;

      case cbxRelatorios.ItemIndex of
         0 : begin
            xls.Sheets[0].AsString[0, 0] := 'Mês Referência';
            xls.Sheets[0].AsString[1, 0] := 'CNPJ Escritório';
            xls.Sheets[0].AsString[2, 0] := 'Nome Escritório';
            xls.Sheets[0].AsString[3, 0] := 'Ocorrência';
         end;
         1 : begin
            xls.Sheets[0].AsString[0, 0] := 'Mês Referência';
            xls.Sheets[0].AsString[1, 0] := 'CNPJ Escritório';
            xls.Sheets[0].AsString[2, 0] := 'Nome Escritório';
            xls.Sheets[0].AsString[3, 0] := 'Linha Planilha';
            xls.Sheets[0].AsString[4, 0] := 'Ocorrência';
            xls.Sheets[0].AsString[5, 0] := 'GCPJ';
            xls.Sheets[0].AsString[6, 0] := 'Parte Contrária';
            xls.Sheets[0].AsString[7, 0] := 'Tipo Andamento';
            xls.Sheets[0].AsString[8, 0] := 'Data do Ato';
            xls.Sheets[0].AsString[9, 0] := 'Tipo Ação';
            xls.Sheets[0].AsString[10, 0] := 'Sub Tipo';
            xls.Sheets[0].AsString[11, 0] := 'Vara';
            xls.Sheets[0].AsString[12, 0] := 'Data Baixa';
            xls.Sheets[0].AsString[13, 0] := 'Motivo Baixa';
            xls.Sheets[0].AsString[14, 0] := 'Empresa Grupo';
            xls.Sheets[0].AsString[15, 0] := 'Valor';
            xls.Sheets[0].AsString[16, 0] := 'Valor Pedido';
            xls.Sheets[0].AsString[17, 0] := 'Valor Baixa';
         end;
         2 : begin
            xls.Sheets[0].AsString[0, 0] := 'Mês Referência';
            xls.Sheets[0].AsString[1, 0] := 'CNPJ Escritório';
            xls.Sheets[0].AsString[2, 0] := 'Nome Escritório';
            xls.Sheets[0].AsString[3, 0] := 'Linha Planilha';
            xls.Sheets[0].AsString[4, 0] := 'Ocorrência';
            xls.Sheets[0].AsString[5, 0] := 'GCPJ';
            xls.Sheets[0].AsString[6, 0] := 'Parte Contrária';
            xls.Sheets[0].AsString[7, 0] := 'Tipo Andamento';
            xls.Sheets[0].AsString[8, 0] := 'Data do Ato';
            xls.Sheets[0].AsString[9, 0] := 'Tipo Ação';
            xls.Sheets[0].AsString[10, 0] := 'Sub Tipo';
            xls.Sheets[0].AsString[11, 0] := 'Vara';
            xls.Sheets[0].AsString[12, 0] := 'Data Baixa';
            xls.Sheets[0].AsString[13, 0] := 'Motivo Baixa';
            xls.Sheets[0].AsString[14, 0] := 'Empresa Grupo';
            xls.Sheets[0].AsString[15, 0] := 'Valor';
            xls.Sheets[0].AsString[16, 0] := 'Valor Corrigido';
            xls.Sheets[0].AsString[17, 0] := 'Valor Pedido';
            xls.Sheets[0].AsString[18, 0] := 'Valor Baixa';
         end;
         3 : begin
            xls.Sheets[0].AsString[0, 0] := 'Mês Referência';
            xls.Sheets[0].AsString[1, 0] := 'CNPJ Escritório';
            xls.Sheets[0].AsString[2, 0] := 'Nome Escritório';
            xls.Sheets[0].AsString[3, 0] := 'Nota';
            xls.Sheets[0].AsString[4, 0] := 'Linha Planilha';
            xls.Sheets[0].AsString[5, 0] := 'GCPJ';
            xls.Sheets[0].AsString[6, 0] := 'Parte Contrária';
            xls.Sheets[0].AsString[7, 0] := 'Tipo Andamento';
            xls.Sheets[0].AsString[8, 0] := 'Data do Ato';
            xls.Sheets[0].AsString[9, 0] := 'Tipo Ação';
            xls.Sheets[0].AsString[10, 0] := 'Sub Tipo';
            xls.Sheets[0].AsString[11, 0] := 'Vara';
            xls.Sheets[0].AsString[12, 0] := 'Data Baixa';
            xls.Sheets[0].AsString[13, 0] := 'Motivo Baixa';
            xls.Sheets[0].AsString[14, 0] := 'Empresa Grupo';
            xls.Sheets[0].AsString[15, 0] := 'Valor';
            xls.Sheets[0].AsString[16, 0] := 'Valor Corrigido';
            xls.Sheets[0].AsString[17, 0] := 'Valor Pedido';
            xls.Sheets[0].AsString[18, 0] := 'Valor Baixa';
         end;
      end;

      while not planJbm.dtsAtos.Eof do
      begin
         re.lines.add('Gerando linha ' + IntToStr(planJbm.dtsAtos.tag) + ' da planilha');
         Application.ProcessMessages;


         if cbxRelatorios.ItemIndex = 1 then
         begin
            if planJbm.dtsAtos.FieldByName('ocorrencia').AsString = 'Valor alterado pelo sistema' then
            begin
               planJbm.dtsAtos.Next;
               continue;
            end;
         end;
               
         case cbxRelatorios.ItemIndex of
            0 : begin
               xls.Sheets[0].AsString[0, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('anomesreferencia').AsString;
               xls.Sheets[0].AsString[1, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('cnpjescritorio').AsString;
               xls.Sheets[0].AsString[2, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('nomeescritorio').AsString;
               xls.Sheets[0].AsString[3, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('ocorrencia').AsString;
            end;
            1 : begin
               xls.Sheets[0].AsString[0, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('anomesreferencia').AsString;
               xls.Sheets[0].AsString[1, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('cnpjescritorio').AsString;
               xls.Sheets[0].AsString[2, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('nomeescritorio').AsString;
               xls.Sheets[0].AsInteger[3, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('linhaplanilha').AsInteger;
               xls.Sheets[0].AsString[4, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('ocorrencia').AsString;
               try
                  xls.Sheets[0].AsInteger[5, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('gcpj').AsInteger;
               except
                  xls.Sheets[0].AsString[5, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('gcpj').AsString;
               end;
               xls.Sheets[0].AsString[6, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('partecontraria').AsString;
               xls.Sheets[0].AsString[7, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('tipoandamento').AsString;
               xls.Sheets[0].AsDateTime[8, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('datadoato').AsDatetime;
               xls.Sheets[0].AsString[9, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('tipoacao').AsString;
               xls.Sheets[0].AsString[10, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('subtipoacao').AsString;
               xls.Sheets[0].AsString[11, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('vara').AsString;
               if Not planJbm.dtsAtos.FieldByName('databaixa').IsNull then
                  xls.Sheets[0].AsDateTime[12, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('databaixa').AsDateTime;
               xls.Sheets[0].AsString[13, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('motivobaixa').AsString;
               xls.Sheets[0].AsString[14, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('nomeempresaligada').AsString;
               xls.Sheets[0].AsFloat[15, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valor').AsFloat;
               xls.Sheets[0].AsFloat[16, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valordopedido').AsFloat;
               if Not planJbm.dtsAtos.FieldByName('valorbaixa').IsNUll then
                  xls.Sheets[0].AsFloat[17, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valorbaixa').AsFloat;
            end;
            2 : begin
               xls.Sheets[0].AsString[0, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('anomesreferencia').AsString;
               xls.Sheets[0].AsString[1, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('cnpjescritorio').AsString;
               xls.Sheets[0].AsString[2, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('nomeescritorio').AsString;
               xls.Sheets[0].AsInteger[3, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('linhaplanilha').AsInteger;
               xls.Sheets[0].AsString[4, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('ocorrencia').AsString;
               xls.Sheets[0].AsInteger[5, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('gcpj').AsInteger;
               xls.Sheets[0].AsString[6, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('partecontraria').AsString;
               xls.Sheets[0].AsString[7, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('tipoandamento').AsString;
               xls.Sheets[0].AsDateTime[8, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('datadoato').AsDatetime;
               xls.Sheets[0].AsString[9, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('tipoacao').AsString;
               xls.Sheets[0].AsString[10, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('subtipoacao').AsString;
               xls.Sheets[0].AsString[11, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('vara').AsString;
               if Not planJbm.dtsAtos.FieldByName('databaixa').IsNull then
                  xls.Sheets[0].AsDateTime[12, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('databaixa').AsDateTime;
               xls.Sheets[0].AsString[13, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('motivobaixa').AsString;
               xls.Sheets[0].AsString[14, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('nomeempresaligada').AsString;
               xls.Sheets[0].AsFloat[15, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valor').AsFloat;
               xls.Sheets[0].AsFloat[16, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat;
               xls.Sheets[0].AsFloat[17, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valordopedido').AsFloat;
               if Not planJbm.dtsAtos.FieldByName('valorbaixa').IsNUll then
                  xls.Sheets[0].AsFloat[18, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valorbaixa').AsFloat;
            end;
            3 : begin
               xls.Sheets[0].AsString[0, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('anomesreferencia').AsString;
               xls.Sheets[0].AsString[1, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('cnpjescritorio').AsString;
               xls.Sheets[0].AsString[2, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('nomeescritorio').AsString;
               xls.Sheets[0].AsString[3, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('numeronota').AsString;
               xls.Sheets[0].AsInteger[4, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('linhaplanilha').AsInteger;
               xls.Sheets[0].AsInteger[5, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('gcpj').AsInteger;
               xls.Sheets[0].AsString[6, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('partecontraria').AsString;

               //alteração solicitada por Gabriela em 08/03/2017
               if planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'PARCELA' then
                  xls.Sheets[0].AsString[7, planJbm.dtsAtos.tag] := 'ACOMPANHAMENTO'
               else
                  xls.Sheets[0].AsString[7, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('tipoandamento').AsString;
               xls.Sheets[0].AsDateTime[8, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('datadoato').AsDatetime;
               xls.Sheets[0].AsString[9, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('tipoacao').AsString;
               xls.Sheets[0].AsString[10, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('subtipoacao').AsString;
               xls.Sheets[0].AsString[11, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('vara').AsString;
               if Not planJbm.dtsAtos.FieldByName('databaixa').IsNull then
                  xls.Sheets[0].AsDateTime[12, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('databaixa').AsDateTime;
               xls.Sheets[0].AsString[13, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('motivobaixa').AsString;
               xls.Sheets[0].AsString[14, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('nomeempresaligada').AsString;
               xls.Sheets[0].AsFloat[15, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valor').AsFloat;
               xls.Sheets[0].AsFloat[16, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat;
               xls.Sheets[0].AsFloat[17, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valordopedido').AsFloat;
               if Not planJbm.dtsAtos.FieldByName('valorbaixa').IsNUll then
                  xls.Sheets[0].AsFloat[18, planJbm.dtsAtos.tag] := planJbm.dtsAtos.FieldByName('valorbaixa').AsFloat;
            end;
         end;
         planJbm.dtsAtos.Tag := planJbm.dtsAtos.Tag + 1;
         planJbm.dtsAtos.Next;
      end;
      xls.Write;
   finally
      xls.Free;
   end;
   ShowMessage('Planilha gerada com sucesso!!');
end;


procedure TfrmRelatorios.BitBtn9Click(Sender: TObject);
begin
   ultimaPagina := 1;
   passagem := 0;
end;

procedure TfrmRelatorios.cbxRelatoriosChange(Sender: TObject);
begin
   numNota.Enabled := (cbxRelatorios.ItemIndex = 3);
   anoRef.Enabled  := (cbxRelatorios.ItemIndex = 4);
   mesRef.Enabled  := (cbxRelatorios.ItemIndex = 4);
end;

procedure TfrmRelatorios.SpeedButton1Click(Sender: TObject);
begin
   if sd.Execute then
      nomePlan.Text := sd.FileName;
end;

procedure TfrmRelatorios.GeraRelatorioGeralInconsistencia;
var
   ret : integer;
   sRec : TSearchRec;
   xls : TXlsReadWriteII5;
   linhaPlan : integer;
   i, col1,col2,col3,col4,totalAtos, totalInc : integer;
   ultEscritorio : string;
   ini : TIniFile;
   diretorio : string;
begin
   if mesRef.ItemIndex = -1 then
   begin
      ShowMessage('Selecione o mês de referência');
      mesRef.SetFocus;
      exit;
   end;

   if anoRef.ItemIndex = -1 then
   begin
      ShowMessage('Selecione o ano de referência');
      anoRef.SetFocus;
      exit;
   end;

   ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
   try
      diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_Programa\presenta');
   finally
      ini.Free;
   end;

   ret := FindFirst(diretorio + '\mdb\HONORARIOS_*.mdb', faAnyFile, sRec);
   if ret <> 0 then
   begin
      ShowMessage('Nenhum banco de dados encontrado para verificação');
      exit;
   end;

   xls := TXLSReadWriteII5.Create(nil);
   try
      xls.Filename := nomePlan.Text;
      xls.Sheets[0].AsString[0, 0] := 'Eacritório Contratado';
      xls.Sheets[0].AsString[1, 0] := 'Mês/Ano Referência';
      xls.Sheets[0].AsString[2, 0] := 'Solicitação de Pagto. em Duplicidade';
      xls.Sheets[0].AsString[3, 0] := 'Solicitação de Pagto. de Proc. não Cadastrado';
      xls.Sheets[0].AsString[4, 0] := 'Configuração Incorreta do Campo Planilha';
      xls.Sheets[0].AsString[5, 0] := 'Denominação Incorreta do Ato';
      xls.Sheets[0].AsString[6, 0] := 'Total de Inconsistências';
      xls.Sheets[0].AsString[7, 0] := 'Total de Atos';
      xls.Sheets[0].AsString[8, 0] := 'Percentual de Erros';

      linhaPlan := 1;

      ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
      try
         diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_Programa\presenta\mdb');
      finally
         ini.Free;
      end;

      while ret = 0 do
      begin
         dmHonorarios.adoConnRpt.Close;;
         dmHonorarios.adoConnRpt.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + diretorio + '\mdb\' + sRec.name + ';Persist Security Info=True';
         dmHonorarios.adoConnRpt.Open;

         re.lines.Add('Totalizando informações do banco de dados: ' + sRec.Name);
         Application.ProcessMessages;

         dmHonorarios.dtsRpt.Close;
         dmHonorarios.dtsRpt.CommandText := 'Select b.idescritorio, nomedigitar, b.anomesreferencia,a.idtipoerro, ocorrencia, count(b.sequencia) as totalInconsistente ' +
                                            'from tbocorrenciasprocessamento a, tbplanilhasatos b, tbescritorios c, tbtiposocorrencia d ' +
                                            'where b.anomesreferencia = :ref and a.idplanilha = b.idplanilha and a.linhaplanilha=b.linhaplanilha ' +
                                            'and b.idescritorio = c.idescritorio and a.idtipoocorrencia = d.idtipoocorrencia ' +
                                            'group by b.idescritorio, nomedigitar, b.anomesreferencia, a.idtipoerro, ocorrencia';

         dmHonorarios.dtsRpt.Parameters.ParamByName('ref').Value := anoRef.Text + '/' + LeftZeroStr('%02d', [mesRef.itemindex+1]);
         dmHonorarios.dtsRpt.Open;

         ultEscritorio := '';
         col1 := 0;
         col2 := 0;
         col3 := 0;
         col4 := 0;
         totalAtos := 0;
         totalInc := 0;
         dmHonorarios.dtsRpt.Tag := 0;

         while Not dmHonorarios.dtsRpt.Eof do
         begin
            re.lines.Add('Totalizando dados do escritório: ' + dmHonorarios.dtsRpt.FieldByName('nomedigitar').AsString);
            Application.ProcessMessages;

            if (dmHonorarios.dtsRpt.FieldByName('nomedigitar').AsString <> ultEscritorio) and
               (ultEscritorio <> '') then
            begin
               //grava na planilha
               xls.Sheets[0].AsString[0, linhaPlan] := ultEscritorio;
               xls.Sheets[0].AsString[1, linhaPlan]:= mesRef.Text + '/' + anoRef.Text;
               xls.Sheets[0].AsInteger[2, linhaPlan]:=col1;
               xls.Sheets[0].AsInteger[3, linhaPlan]:=col2;
               xls.Sheets[0].AsInteger[4, linhaPlan]:=col3;
               xls.Sheets[0].AsInteger[5, linhaPlan]:=col4;

               totalInc := col1 + col2 + col3 + col4;
               xls.Sheets[0].AsInteger[6, linhaPlan]:=totalInc;

               xls.Sheets[0].AsInteger[7, linhaPlan]:=totalAtos+col3;

               xls.Sheets[0].AsFloat[8, linhaPlan]:=(totalInc*100)/totalAtos;

               col1 := 0;
               col2 := 0;
               col3 := 0;
               col4 := 0;
               totalAtos := 0;
               totalInc := 0;
               Inc(linhaPlan);
               dmHonorarios.dtsRpt.Tag := 0;
            end;
            ultEscritorio := dmHonorarios.dtsRpt.FieldByName('nomedigitar').AsString;

            if dmHonorarios.dtsRpt.Tag = 0 then
            begin
               dmHonorarios.dtsRptTotal.Close;
               dmHonorarios.dtsRptTotal.CommandText := 'Select count(sequencia) as totalAtosImportados ' +
                                                  'from tbplanilhasatos ' +
                                                  'where anomesreferencia = :ref and idescritorio = ' + dmHonorarios.dtsRpt.FieldByName('idescritorio').AsString;

               dmHonorarios.dtsRptTotal.Parameters.ParamByName('ref').Value := anoRef.Text + '/' + LeftZeroStr('%02d', [mesRef.itemindex+1]);
               dmHonorarios.dtsRptTotal.Open;

               totalAtos := dmHonorarios.dtsRptTotal.FieldByName('totalAtosImportados').AsInteger;

               dmHonorarios.dtsRpt.Tag := 1;
            end;

            if (dmHonorarios.dtsRpt.FieldByName('ocorrencia').AsString = 'Ato em duplicidade na planilha') or
               (dmHonorarios.dtsRpt.FieldByName('ocorrencia').AsString = 'Ato já foi pago anteriormente') then
               col1 := col1 + dmHonorarios.dtsRpt.FieldByName('totalInconsistente').AsInteger
            else
            if (dmHonorarios.dtsRpt.FieldByName('ocorrencia').AsString = 'Processo não cadastrado no GCPJ') then
               col2 := col2 + dmHonorarios.dtsRpt.FieldByName('totalInconsistente').AsInteger
            else
            if (dmHonorarios.dtsRpt.FieldByName('idtipoerro').AsInteger = 1) then
               col3 := col3 + dmHonorarios.dtsRpt.FieldByName('totalInconsistente').AsInteger
            else
            if (dmHonorarios.dtsRpt.FieldByName('ocorrencia').AsString = 'Tipo de ato não reconhecido pelo sistema') then
               col4 := col4 + dmHonorarios.dtsRpt.FieldByName('totalInconsistente').AsInteger;
            dmHonorarios.dtsRpt.Next;
         end;

         if ultEscritorio <> '' then
         begin
            //grava na planilha
            xls.Sheets[0].AsString[0, linhaPlan] := ultEscritorio;
            xls.Sheets[0].AsString[1, linhaPlan]:= mesRef.Text + '/' + anoRef.Text;
            xls.Sheets[0].AsInteger[2, linhaPlan]:=col1;
            xls.Sheets[0].AsInteger[3, linhaPlan]:=col2;
            xls.Sheets[0].AsInteger[4, linhaPlan]:=col3;
            xls.Sheets[0].AsInteger[5, linhaPlan]:=col4;

            totalInc := col1 + col2 + col3 + col4;
            xls.Sheets[0].AsInteger[6, linhaPlan]:=totalInc;

            xls.Sheets[0].AsInteger[7, linhaPlan]:=totalAtos+col3;

            xls.Sheets[0].AsFloat[8, linhaPlan]:=(totalInc*100)/totalAtos;
         end;

         ret := FindNext(sRec);
      end;
      xls.Write;
   finally
      xls.Free;
   end;
   ShowMessage('Relatório Gerado com Sucesso');
end;

procedure TfrmRelatorios.Retornarplanilhaparadigitao1Click(
  Sender: TObject);
begin
   if MessageDlg('Deseja realmente retornar a planilha com status de Não Finalizada?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      exit;

   planJbm.idEscritorio := planJbm.dtsPlanilha.FieldByName('idescritorio').AsInteger;
   planJbm.idplanilha := planJbm.dtsPlanilha.FieldByName('idplanilha').AsInteger;
   planJbm.sequencia :=  planJbm.dtsPlanilha.FieldByName('sequencia').AsInteger;
   planJbm.anomesreferencia := planJbm.dtsPlanilha.FieldByName('anomesreferencia').AsString;

   planJbm.MarcaPlanilhaNaoFinalizada;
end;

end.
