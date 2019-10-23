unit fdigitaplan;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Buttons, Grids, DBGrids, uplanjbm, inifiles,
  datamodulegcpj_base_i, Func_Wintask_Obj, mgenlib, ExtCtrls, fatosDigitados, datamodule_honorarios,
  fretornarnota, Menus, adodb, dateutils, fgcpj, OleCtrls, SHDocVw, mshtml, fGCPJConfirmar, datamodulegcpj_base_ix;

type

  TfrmDigitaPlan = class(TForm)
    Label1: TLabel;
    DBGrid1: TDBGrid;
    BitBtn2: TBitBtn;
    re: TRichEdit;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    processando: TEdit;
    Label2: TLabel;
    numNota_: TEdit;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    DBGrid2: TDBGrid;
    DBGrid3: TDBGrid;
    BitBtn1: TBitBtn;
    Timer1: TTimer;
    pmNotas: TPopupMenu;
    Retornarnotafinalizada1: TMenuItem;
    ConferirAtosdaNotaComoGCPJ1: TMenuItem;
    pnlWb: TPanel;
    wb: TWebBrowser;
    BitBtn8: TBitBtn;
    Marcarnotacomofinalizada1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure DBGrid2DblClick(Sender: TObject);
    procedure Retornarnotafinalizada1Click(Sender: TObject);
    procedure ConferirAtosdaNotaComoGCPJ1Click(Sender: TObject);
    procedure wbDocumentComplete(Sender: TObject; const pDisp: IDispatch;
      var URL: OleVariant);
    procedure wbNewWindow2(Sender: TObject; var ppDisp: IDispatch;
      var Cancel: WordBool);
    procedure BitBtn8Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Marcarnotacomofinalizada1Click(Sender: TObject);
  private
    FplanJbm: TPlanilhaJBM;
    FehPresenta: boolean;
    procedure SetplanJbm(const Value: TPlanilhaJBM);
    procedure SetehPresenta(const Value: boolean);
    { Private declarations }
  public
    { Public declarations }
    property planJbm : TPlanilhaJBM read FplanJbm write SetplanJbm;
    property ehPresenta : boolean read FehPresenta write SetehPresenta;

    procedure IndexaBasesGCPJ;
    function ObtemTipoAndamento(andamento: string; tipoProcesso: string; motivoBaixa: string) : string;
    function ConfereAtosDaNota(numeroNota: string; wintask: TWintask) :integer;
    function PosicionaNotaDigitando(notificar: boolean; nota: string; excluir: boolean;wintask: twintask) : integer;
    procedure MyDeleteFile(nome: string);
    procedure DigitaNotas(notaRetornar: string);
    procedure NovaRotinaDigitacaoNotas(notaRetornar: string; jaPosicionado: boolean);
    procedure SelecionaTodaLista;

    function GetInputField(fromForm: IHTMLFormElement;  const inputName: string;  const instance: integer=0): HTMLInputElement;
    function GetInputFieldById(fromForm: IHTMLFormElement;  const inputName, id: string;  const instance: integer=0): HTMLInputElement;
    function GetFormByName(const documento: IHTMLDocument2; nomeForm: string) : IHTMLFormElement;
    procedure EsperaPaginaProcessar(const janelaFechar: string='');
    function GetSelectField(fromForm: IHTMLFormElement;  const selectName: string;  const instance: integer=0): HTMLSelectElement;
    function GetScriptField(fromForm: IHTMLFormElement;  const selectName: string;  const instance: integer=0): HTMLScriptElement;
    procedure PopUpWbDocumentComplete(Sender: TObject; const pDisp: IDispatch; var URL: OleVariant);
    function DigitaNotaNova(winTask: TWintask; dataDigitar: string) : integer;
    function DigitaNotaEmAndamento(winTask: TWintask; dataDigitar: string) : integer;
    function PosicionaNotaNova(wintask: TWintask): integer;
    function DigitaAtosDaNota(wintask: TWintask) : integer;
    function NovoObtemDependenciaDigitar:integer;
    function SetSelectField(fromForm: IHTMLFormElement;  const selectName: string; const elementName: string; executar: string): integer;
    function ClicaBotaoGcpj(var fromForm: IHTMLFormElement;  const buttonName: string; const janelaFechar: string=''): integer;
    function NovoObtemTotalDigitado : double;
    function NovoPosicionaNotaDigitando:integer;
    procedure FechaJanelaGcpj(janelaFechar: string);
    function FinalizaNotaNoGcpj : integer;
  end;

const
   TIME_SLEEP_ZERO = 500;
   TIME_SLEEP_MIN = 800;
   TIME_SLEEP_MED = 1000;

var
  frmDigitaPlan: TfrmDigitaPlan;
  ie8 : boolean;
  ff : boolean;
  ultimaPagina, passagem : integer;
  leftDown, leftUp, rightDown, rightUp : string;
  winThread : TClickOkThread;
  mainPage, pageGcpj : string;
  segue : boolean;
  wbpopUp : TfrmGcpj;
  wbPopWin : integer;
  wbPopUpConfirmar : TfrmgcpjConfirmar;
  pJanelaFechar : string;

implementation

uses fcaddatainicial, datamodulegcpj_base_xi;


{$R *.dfm}

{ TfrmValidaPlan }

procedure TfrmDigitaPlan.SetplanJbm(const Value: TPlanilhaJBM);
begin
  FplanJbm := Value;
end;

procedure TfrmDigitaPlan.FormCreate(Sender: TObject);
var
   ini : TiniFile;
   configname : string;
begin
   planJbm := TPlanilhaJBM.Create;

   DBGrid1.DataSource := planJbm.dsPlan;
   DBGrid2.DataSource := planJbm.dsNotasPendentes;
   DBGrid3.DataSource := planJbm.dsNotasDigitando;

   planJbm.ObtemPlanilhasDigitar;
   if planJbm.dtsPlanilha.RecordCount = 0 then
      BitBtn4.Enabled := false;

   GetParam('/wintaskini=', configname);
   if configname = '' then
      configname := 'wintask_config.ini';

   ie8 := false;
   ff := false;

   if IsParam('/ff') then
      ff := true
   else
   if IsParam('/ie8') then
      ie8 := true
   else
   begin
      ini := TIniFile.Create(ExtractFilePath(application.ExeName) + configname);
      try
         ie8 := ini.ReadBool('ie', 'ie8', false);
         ff := ini.ReadBool('firefox', 'ff', false);
      finally
         ini.free;
      end;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + configname);
   try
      leftDown := ini.ReadString('notepad', 'leftDown', '856');
      leftup  := ini.ReadString('notepad', 'leftup', '490');
      rightDown := ini.ReadString('notepad', 'rightDown', '787');
      rightUp := ini.ReadString('notepad', 'rightUp', '297');
   finally
      ini.free;
   end;

   if ff then
   begin
      mainPage := 'PLUGIN-CONTAINER.EXE|Internet Explorer_Server|GCPJ - Mozilla Firefox|1';//'FIREFOX.EXE|MozillaWindowClass|GCPJ - Mozilla Firefox';
      pageGcpj := 'GCPJ - Gest�o e Controle de Processos Jur�dico';//'GCPJ';// - Mozilla Firefox';
   end
   else
   begin
      mainPage := 'IEXPLORE.EXE|Internet Explorer_Server|GCPJ - Gest�o e Controle de Processos Jur�dicos - Windows Internet Explorer|1';//'IEXPLORE.EXE|Internet Explorer_Server|GCPJ - Windows Internet Explorer';
      pageGcpj := 'GCPJ - Gest�o e Controle de Processos Jur�dico';//'GCPJ';
   end;

   ini := TiniFile.Create('c:\presenta\basediaria\CONFIG.INI');
   try
      ehPresenta := (ini.ReadInteger('gcpj', 'presenta', 0)=1);
   finally
      ini.free;
   end;
end;

procedure TfrmDigitaPlan.BitBtn2Click(Sender: TObject);
begin
   try
      winThread.pleaseClose:= true;
   except
   end;
   close;
end;

procedure TfrmDigitaPlan.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   planJbm.Finalizar;
   planJbm.Free;
   action := CaFree;
end;



procedure TfrmDigitaPlan.IndexaBasesGCPJ;
begin
   processando.Text := 'Indexando as bases do GCPJ';
   Application.ProcessMessages;

   re.lines.Add('Criando �ndice na base de dados do GCPJ Base I. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseI;

   re.lines.Add('Criando �ndice na base de dados do GCPJ Base II. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseII;

   re.lines.Add('Criando �ndice na base de dados do GCPJ Base III. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseIII;

   re.lines.Add('Criando �ndice na base de dados do GCPJ Base IV. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseIV;

   re.lines.Add('Criando �ndice na base de dados do GCPJ Base Compartilhada. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseCompartilhada;

   re.lines.Add('Criando �ndice na base de dados do GCPJ Base Migrados. Aguarde.....');
   Application.ProcessMessages;
   planJbm.CriaIndiceBaseMigrados;
end;


procedure TfrmDigitaPlan.BitBtn3Click(Sender: TObject);
var
   arquivo : TstringList;
   linha : string;
   ini : TIniFile;
   contador : integer;
begin
   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
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
      arquivo.add('Ano/M�s;Escrit�rio;C�d. Empresa Grupo;Nome Empresa Grupo;Nota;GCPJ;Parte Contr�ria;Tipo Ato;Tipo A��o;Sub Tipo A��o;Data Baixa;Motivo Baixa;Data do Ato;Valor do Pedido;Valor Pagar;'+
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

         planJbm.ObtemReclamadasDoProcesso('');
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

procedure TfrmDigitaPlan.BitBtn4Click(Sender: TObject);
begin
//   if IsParam('/new') then
      NovaRotinaDigitacaoNotas('', false);
//   else
//      DigitaNotas('');
end;

procedure TfrmDigitaPlan.DigitaNotas(notaRetornar: string);
var
   wintask :  TWintask;
   ultdocumento : string;
   fName : string;
   valor, valDigitar : string;
   digitando : boolean;
   i, estagio, index, p, k : integer;
   arquivo : TStringList;
   destinatarioOk : boolean;
//   nota : string;
   passagem : integer;
   totaldigitado, novototaldigitado, antigoTotalDigitado: double;
   //dataDigitar : string;
   popup : string;
   numeronota : string;
   bMark : TBookmarkList;
   fhandle :integer;
   ret : integer;
   andamentoDigitar : string;
   ini : Tinifile;
   diretorio : string;

   label
      proximo;

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
      sequencia : integer;
   begin
      result := '';
      codigo := '';
      nome:= '';
      planJbm.ObtemReclamadasDoProcesso('');
      sequencia := 1;

      while Not planJbm.dtsReclamadas.Eof do
      begin
         if planJbm.dtsReclamadas.FieldByName('sequenciagcpj').AsInteger = sequencia then //guarda a primeira por seguranca
         begin
            if planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger <> 0 then
            begin
               codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
               nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
            end
            else
            begin
               Inc(sequencia);
               planJbm.dtsReclamadas.Next;
               continue;
            end;
         end;

         //� bradesco?
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
            if (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = 4001) and (planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger = 5172) then
            begin
               codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
               nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
               break;
            end
            else
            if (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger) or
               (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('codempresaligada').AsInteger) then
            begin
               if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A' then //n�o pode ser agencia
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


      if not primeira then
      begin
         if StrToInt(codigo) = 5404 then
         begin
            //tem a depend�ncia 4027?
            planJbm.dtsReclamadas.First;

            while Not planJbm.dtsReclamadas.Eof do
            begin
               Application.ProcessMessages;
               if planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = 4027 then
               begin
                  codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
                  nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
                  break;
               end;
               planJbm.dtsReclamadas.Next;
            end;
         end;
      end;

      arquivo := TStringList.Create;
      try
         arquivo.LoadFromFile(fname);

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
                        ShowMessage('Erro valor');

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

   function ObtemTotalDigitado:double;
   var
      i, p, passagem : integer;
   begin
      result := 0;
      passagem := 1;

      Sleep(800);

      while passagem <= 3 do
      begin
         wintask.Click_Mouse(mainPage,591,234);
         Sleep(500);

         fName := wintask.GetPagina_CopyText(mainPage,
                                                'Tipo de Prestador: ');
         if fName = '' then
         begin
            Inc(passagem);
            Sleep(800);
            continue;
         end;

         //esta na pagina principal
         arquivo := TStringList.Create;
         try
            arquivo.LoadFromFile(fname);
            for i := 0 to arquivo.count - 1 do
            begin
               p := Pos('Total: ', arquivo.strings[i]);
               if p = 0 then
                  continue;

               valor := Trim(Copy(arquivo.strings[i], p+7, length(arquivo.Strings[i])));
               RemoveEsteCaracter('.', valor);
               RemoveEsteCaracter(' ', valor);

               if valor = '' then
                  result := 0
               else
                  result := StrToFloat(valor);
               exit;
            end;
         finally
            arquivo.free;
            MyDeleteFile(fName);
         end;
      end;
   end;


   function PosicionaNotaNova:integer;
   begin
      result := -1;

      wintask.Click_Mouse(mainPage, 591,234);

      //localizar a tela da nota
      if wintask.EstaNaPagina_CopyText(mainPage, 'Tipo de Prestador: ') then
      begin
         wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
         sleep(TIME_SLEEP_MIN);
      end;

      if Not wintask.EstaNaPagina_CopyText(mainPage, 'Honor�rios em Cadastramento') then
         exit;

      wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''incluir'']');
      Sleep(TIME_SLEEP_MED);

      wintask.Click_Mouse(mainPage, 591,234);

      while wintask.EstaNaPagina_CopyText(mainPage, 'Honor�rios em Cadastramento') do
      begin
         wintask.Click_Mouse(mainPage, 591,234);
         Application.ProcessMessages;
         Sleep(500);
      end;


      wintask.SelectHTMLItem(pageGcpj, 'SELECT[NAME= ''cdTipoDocumentoPesso'']''', 'CNPJ');
      Sleep(TIME_SLEEP_MED);


//      wintask.Click_Mouse(mainPage,297,184);
      wintask.Click_Mouse(mainPage,293,89);
      Sleep(TIME_SLEEP_ZERO);

      valor := planJbm.cnpjEscritorio;
      valDigitar := '<Num 0>';
      while Length(valor) > 0 do
      begin
         valDigitar := valDigitar + '<Num ' + Copy(valor, 1, 1) + '>';
         valor := Copy(valor, 2, length(valor));
      end;

      wintask.SendKey(mainPage,valDigitar);
      Sleep(TIME_SLEEP_ZERO);

      wintask.ClickHTMLElement(pageGcpj, 'INPUT BUTTON[NAME= ''btoLupa'']');
      Sleep(1800);

      if ie8 then
         popup := 'IEXPLORE.EXE|Internet Explorer_Server|Popup - Windows Internet Explorer|1'
      else
         popup:= 'IEXPLORE.EXE|Internet Explorer_Server|Popup - Windows Internet Explorer|1';

      if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
      begin
         if planJbm.nomedigitar = '' then
            wintask.ClickHTMLElement('Popup', 'A[INNERTEXT= ''' + copy(planJbm.nomeescritorio, 1, 20) + ''']')
         else
            wintask.ClickHTMLElement('Popup', 'A[INNERTEXT= ''' + planJbm.nomedigitar + ''']');
         Sleep(TIME_SLEEP_MIN);
      end
      else
      begin
        popUp := 'IEXPLORE.EXE|Static|O campo CNPJ � obrigat�rio';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if wintask.ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
           Sleep(TIME_SLEEP_MIN);
           exit;
        end;
      end;

      wintask.SelectHTMLItem(pageGcpj, 'SELECT[NAME= ''despesaHonorariosPro'']', 'NOTA FISCAL');
      Sleep(TIME_SLEEP_ZERO);

//      wintask.Click_Mouse(mainPage, 577,307);
      wintask.Click_Mouse(mainPage, 552,217);
      Sleep(TIME_SLEEP_ZERO);

      valor := planJbm.dtsAtos.FieldByName('numeronota').AsString;
      valDigitar := '';
      while Length(valor) > 0 do
      begin
         valDigitar := valDigitar + '<Num ' + Copy(valor, 1, 1) + '>';
         valor := Copy(valor, 2, length(valor));
      end;
      wintask.SendKey(mainPage,valDigitar);
      Sleep(TIME_SLEEP_ZERO);

      wintask.SendKey(mainPage,'<Tab>A<Tab>');
      Sleep(TIME_SLEEP_ZERO);

      valor := FormatDateTime('ddmmyyyy', date);
      valDigitar := '';
      while Length(valor) > 0 do
      begin
         valDigitar := valDigitar + '<Num ' + Copy(valor, 1, 1) + '>';
         valor := Copy(valor, 2, length(valor));
      end;
      wintask.SendKey(mainPage,valDigitar);
      Sleep(TIME_SLEEP_ZERO);

      wintask.Click_Mouse(mainPage,546,241);
//      wintask.Click_Mouse(mainPage,589,338);
      Sleep(TIME_SLEEP_ZERO);

      valor := RemoveNaoNumericos(FormatFloat('0.00', planJbm.valorTotalNota));
      valDigitar := '';
      while Length(valor) > 0 do
      begin
         valDigitar := valDigitar + '<Num ' + Copy(valor, 1, 1) + '>';
         valor := Copy(valor, 2, length(valor));
      end;

      wintask.SendKey(mainPage,valDigitar);
      Sleep(TIME_SLEEP_ZERO);

      wintask.WriteHTML(pageGcpj, 'INPUT TEXT[NAME= ''despesaHonorariosProcessoVO.empresaBradescoVO.c'']', '237');
      Sleep(TIME_SLEEP_ZERO);

      if planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger <> 237 then
         wintask.WriteHTML(pageGcpj, 'INPUT TEXT[NAME= ''despesaHonorariosProcessoVO.dependenciaBradesco'']', planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString);
      Sleep(TIME_SLEEP_ZERO);
      result := 1;
   end;

   function VerificaTela(comThread: boolean):integer;
   begin
      result := -1;

      if comThread then
      begin
//         winThread:= TClickOkThread.Create;
//         winThread.Resume;
      end;

      try
        popup := 'IEXPLORE.EXE|Static|CODIGO PRESTADOR INVALIDO';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
        popup := 'TASKEXEC.EXE|Static|Error at line';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|CODIGO DO TIPO DOCUMENTO HONORARIOS INVALIDO';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|Pressione a lupa para realizar a pesquisa';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|DATA DO ATO INVALIDO';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio'
         else
            popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|O campo Data do Ato � obrigat�rio'
         else
            popup := 'IEXPLORE.EXE|Static|O campo Data do Ato � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor, false) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                             Sleep(TIME_SLEEP_ZERO);
            result := 0;
            exit;
         end;

         Sleep(200);
         popUp := 'IEXPLORE.EXE|Static|O campo Data do Documento � obrigat�rio';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if wintask.ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            result := 1;
            exit;
         end;

         Sleep(200);
         popUp := 'IEXPLORE.EXE|Static|O campo N� do Processo Bradesco � obrigat�rio';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if wintask.ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            result := 1;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popUp := 'IEXPLORE.EXE|Static|O campo C�digo do Destinat�rio � obrigat�rio.'
         else
            popUp := 'IEXPLORE.EXE|Static|O campo C�digo do Destinat�rio � obrigat�rio.';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|O campo Depend�ncia � obrigat�rio'
         else
            popup := 'IEXPLORE.EXE|Static|O campo Depend�ncia � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|O campo Ato � obrigat�rio'
         else
            popup := 'IEXPLORE.EXE|Static|O campo Ato � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|Esta depend�ncia foi encerrada, deseja alter�-la para'
         else
            popup := 'IEXPLORE.EXE|Static|Esta depend�ncia foi encerrada, deseja alter�-la para';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            result := 0;
            exit;
         end;


         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|VALOR DO DOCUMENTO DEVE SER IGUAL A SOMATORIA DOS ATOS';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
            result := 0;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Internet Explorer_Server|ERRO! - Windows Internet Explorer';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            result := -2;
            exit;
         end;

         result := 1;
         exit;
      finally
         try
            if comThread then
               winThread.pleaseClose:=true;
         except
         end;
      end;


//                     planJbm.NotificaErro('Valores digitados n�o conferem');
//                     ShowMessage('Valores digitados n�o conferem');
//                     exit;
//                  end;
   end;

   procedure ExcluiDoGcpj(contador:integer);
   begin
      //exclui o ato do GCPJ
      wintask.ClickHTMLElement(pageGcpj, 'INPUT RADIO[NAME= ''composicoesSelecs'',INDEX=''' + IntToStr(contador) + ''']');
      Sleep(TIME_SLEEP_ZERO);

      wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''excluir'']');
      Sleep(TIME_SLEEP_MED);

      wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
      sleep(TIME_SLEEP_MED);
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

   (**retirado em 10/03/2015 a pedido da Nilce
   dataDigitar := planJbm.ObtemDataDigitar(0);
   if dataDigitar = '' then
   begin
      if MessageDlg('N�o encontrada a data inicial para o ano/m�s de refer�ncia. Deseja cadastrar?', mtConfirmation, mbYesNoCancel, 0) <> mrYes then
      begin
         planJbm.NotificaErro('Sistema n�o pode digitar para o m�s ' + planJbm.anomesreferencia + ' at� ser cadastrada a data inicial');
         ShowMessage('Sistema n�o pode digitar para o m�s ' + planJbm.anomesreferencia + ' at� ser cadastrada a data inicial');
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
      dataDigitar := planJbm.ObtemDataDigitar(0);
   end;
   **)

   dbGrid2.DataSource.DataSet.DisableControls;

   try
      if dbGrid2.SelectedRows.Count = 0 then //nehujm marcado. Marca todos
      begin
         //ainda tem registro?
         if DBGrid2.DataSource.DataSet.RecordCount = 0 then
         begin
            planJbm.MarcaPlanilhaFinalizada;
            planjbm.dtsPlanilha.Close;
            planjbm.dtsPlanilha.Open;
            exit;
         end;
         SelecionaTodaLista;
      end;

      bMark := dbGrid2.SelectedRows;

      ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
      try
         diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_programa\Presenta');
      finally
         ini.Free;
      end;


      for k := 0 to bMark.Count - 1 do
      begin
         wintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                                    diretorio + '\rob\',
                                    '');

         try
            wintask.ie8 := ie8;
            wintask.ff := ff;

            wintask.FecharTodos('TaskExec.exe');
            wintask.FecharTodos('TaskLock.exe');
            wintask.FecharTodos('Tasksync.exe');

            if IsParam('/testar') then
            begin
               wintask.EstaNaPagina_NotepadCodePage(mainPage,
                                                    //'813', '289',         ]
                                                    //'780', '459',
                                                    rightDown, rightUp,
                                                    leftDown, leftUp,
                                                    'NOTEPAD.EXE|Notepad|finHonorariosNovoAction',
                                                    'Documento do Prestador: ',
                                                    false);
               exit;
            end;


            if FileExists(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT') then
            begin
               RenameFile(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT', ExtractFilePath(Application.ExeName) + 'ENCERRAR.OK');
               exit;
            end;

            dbGrid2.DataSource.DataSet.BookMark := bMark.items[k];

            digitando := false;
            ultdocumento := '';
            totaldigitado := 0;
            novototaldigitado := 0;
            antigototaldigitado := 0;


            if (notaRetornar <> '') and (DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString <> notaRetornar) then
               continue;

            planJbm.empresaLigada := DBGrid2.DataSource.DataSet.FieldByName('empresa_ligada_agrupar').AsInteger;
            planJbm.numerodanota :=  DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsInteger;

            ret := IsFileLocked(planJbm.ObtemDirLck + '\' + IntToStr(planJbm.numerodanota) + '.lck', fhandle);

            if ret <= 0 then
               continue;


            if wintask.ie8 then
            begin
               ret := VerificaTela(false);
               if ret <> 1 then
               begin

                  ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                  try
                     winThread.pleaseClose:= true;
                  except
                  end;

                  if ret = -2 then
                  begin
                     ShowMessage('Erro inesperado no GCPJ');
                     exit;
                  end;

                  planJbm.dtsNotasPendentes.Close;
                  planJbm.dtsNotasPendentes.Open;;

                  //DBGrid2.SelectedRows.Clear;
                  DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                  exit;
               end;
            end;

            try
               ret := planJbm.ObtemProcessosDigitar(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);

               if ret = -1 then
                  exit;

               if ret = 0 then
               begin
                  //finaliza a nota no GCPJ
                  if PosicionaNotaDigitando(false,DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString, false, wintask)>= 0 then
                  begin
                     wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''finalizar nota'']');
                     Sleep(TIME_SLEEP_MED);

                     if ie8 then
                        popup := 'IEXPLORE.EXE|Static|Opera��o realizada com sucesso'
                     else
                        popup :='IEXPLORE.EXE|Static|Opera��o realizada com sucesso';
                     if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                     begin
                        wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
                        Sleep(TIME_SLEEP_MED);
                     end
                     else
                     begin
                        popup := 'IEXPLORE.EXE|Static|VALOR DO DOCUMENTO DEVE SER IGUAL A SOMATORIA DOS ATOS';
                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
                           Sleep(TIME_SLEEP_MED);

                           //confere todos os atos da Nota
                           if Not IsParam('/semconferencia') then
                           begin
                              ret := ConfereAtosDaNota(IntToStr(planJbm.numerodanota), wintask);
                              if ret <> 1 then
                              begin
                                 ShowMessage('Erro finalizando a nota');
                                 exit;
                              end
                              else
                              begin
                                 ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                                 try
                                    winThread.pleaseClose:= true;
                                 except
                                 end;

                                 planJbm.dtsNotasPendentes.Close;
                                 planJbm.dtsNotasPendentes.Open;;

                                 //DBGrid2.SelectedRows.Clear;
                                 DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                                 exit;
                              end;
                           end
                           else
                           begin
                              ShowMessage('Erro finalizando a nota');
                              exit;
                           end;
                        end
                        else
                        begin
                           ShowMessage('Erro finalizando a nota');
                           exit;
                        end;
                     end;
                  end;
                  planJbm.MarcaDocumentoFinalizado;
                  continue;
               end;

               planJbm.dtsAtos.tag := 0;

               while not planJbm.dtsAtos.Eof do
               begin
                  if FormatDateTime('hh:nn', Now) > '22:00' then
                     exit;

                  if FileExists(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT') then
                  begin
                     RenameFile(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT', ExtractFilePath(Application.ExeName) + 'ENCERRAR.OK');
                     exit;
                  end;

                  planJbm.CarregaCamposDaTabela;
                  re.lines.add('Digitando processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
                  Application.ProcessMessages;

                  planJbm.ObtemValorTotalDaNota;

                  if (ultdocumento <> planJbm.dtsAtos.FieldByName('numeronota').AsString) and
                     (ultdocumento <> '') then
                  begin
                     //envia e-mail de nota finalizada
                     if Not wintask.EstaNaPagina_CopyText(mainPage, 'Honor�rios em Cadastramento') then
                        wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                     planJbm.MarcaDocumentoFinalizado;
//                     planJbm.NotificaNotaFinalizada;
                     break;
                  end;

                  if (ultdocumento = '') and (not digitando) then //comeca a digitar
                  begin
//                     planJbm.NotificaNotaIniciada;

                     if planJbm.NotaJaFoiDigitada then
                     begin
                        if PosicionaNotaDigitando(true,planJbm.dtsAtos.FieldByName('numeronota').AsString, false,wintask) < 0 then
                           exit;

                        digitando := true;

                        novototaldigitado := ObtemTotalDigitado;    //tela
                        antigoTotalDigitado := planJbm.SomaTotalDigitadoDaNota; //banco e dados

                        if (FormatFloat('0.00', novototaldigitado) = FormatFloat('0.00', antigoTotalDigitado)) and
                           (planJbm.dtsAtos.FieldByName('fgdigitado').AsInteger = 1) then //ja foi digitado. Pode desprezar
                        begin
                           planJbm.dtsAtos.Next;
                           continue;
                        end;

                        if FormatFloat('0.00', novototaldigitado) <> FormatFloat('0.00', antigoTotalDigitado) then
                        begin
                           if antigoTotalDigitado < novototaldigitado then
                           begin
                              if FormatFloat('0.00',novototaldigitado) = FormatFloat('0.00', antigoTotalDigitado - planJbm.dtsAtos.fieldbyname('valorcorrigido').asFloat)then
                              begin
                                 planJbm.MarcaProcessoDigitado;
                                 planJbm.dtsAtos.Next;
                                 continue;
                              end;
                           end;

                           //bancode dados � maior
                           if antigoTotalDigitado > novototaldigitado then
                           begin
                              //remove do banco de dados
                              dmHonorarios.dtsAtosDigitados.Close;
                              dmHonorarios.dtsAtosDigitados.CommandText := 'select * from tbplanilhasatos where numeronota = ' + DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString +
                                                                           ' and fgdigitado = 1 ' +
                                                                           'order by gcpj, valorCorrigido';
                              dmHonorarios.dtsAtosDigitados.Open;

                              if not dmHonorarios.dtsAtosDigitados.Eof then
                              begin
                                 dmHonorarios.dtsAtosDigitados.Last;

                                 while not dmHonorarios.dtsAtosDigitados.Bof do
                                 begin
                                    if FormatFloat('0.00',novototaldigitado) <> FormatFloat('0.00', antigoTotalDigitado - dmHonorarios.dtsAtosDigitados.fieldbyname('valorcorrigido').asFloat)then
                                    begin
                                         dmHonorarios.dtsAtosDigitados.Prior;
                                         continue;
                                    end;

                                    //marca como n�o digitado
                                    dmHonorarios.adoCmd.CommandText := 'update tbplanilhasatos set fgdigitado=0, datahoradigitacao=null ' +
                                                                       'where idescritorio = ' + dmHonorarios.dtsAtosDigitados.FieldByName('idescritorio').AsString + ' and ' +
                                                                       'idplanilha = ' + dmHonorarios.dtsAtosDigitados.FieldByName('idplanilha').AsString + ' and ' +
                                                                       'linhaplanilha = ' + dmHonorarios.dtsAtosDigitados.FieldByName('linhaplanilha').AsString + ' and ' +
                                                                       'sequencia = ' + dmHonorarios.dtsAtosDigitados.FieldByName('sequencia').AsString;
                                    dmHonorarios.adoCmd.Execute;

                                    ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                                    try
                                       winThread.pleaseClose:= true;
                                    except
                                    end;

                                    planJbm.dtsNotasPendentes.Close;
                                    planJbm.dtsNotasPendentes.Open;;

                                    //DBGrid2.SelectedRows.Clear;
                                    DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                                    exit;
                                 end;
                                 ShowMessage('Erro inesperado');
                                 exit;
                              end;
                           end;

                           if Not IsParam('/semconferencia') then
                           begin
                              ret := ConfereAtosDaNota(IntToStr(planJbm.numerodanota), wintask);
                              if ret <> 1 then
                              begin
                                 ShowMessage('Valores digitados n�o conferem(1)');
                                 exit;
                              end
                              else
                              begin
                                 ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                                 try
                                    winThread.pleaseClose:= true;
                                 except
                                 end;

                                 planJbm.dtsNotasPendentes.Close;
                                 planJbm.dtsNotasPendentes.Open;;

                                 //DBGrid2.SelectedRows.Clear;
                                 DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                                 exit;
                              end;
                           end
                           else
                           begin
                              ShowMessage('Valores digitados n�o conferem(2)');
                              exit;
                           end;
                        end;
                     end
                     else
                     begin
                        if PosicionaNotaNova < 0 then
                           exit;
                        ultdocumento := planJbm.dtsAtos.FieldByName('numeronota').AsString;
                     end;
                  end;

                  digitando := true;

                  planjbm.tipoProcesso := planJbm.dtsAtos.FieldByName('fgtipoprocesso').asString;

                  wintask.WriteHTML(pageGcpj, 'INPUT TEXT[NAME= ''despesaHonorariosProcessoVO.co'']', planJbm.dtsAtos.FieldByName('gcpj').AsString);
                  Sleep(5);

                  wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''pesquisar'']');
                  Sleep(TIME_SLEEP_MED*2);

                  fname := wintask.EstaNaPagina_NotepadCodePage('IEXPLORE.EXE|Internet Explorer_Server|GCPJ - Gest�o e Controle de Processos Jur�dicos - Windows Internet Explorer|1',//mainPage,
                                                                rightDown, rightUp,
                                                                leftDown, leftUp,
                                                                'NOTEPAD.EXE|Notepad|finHonorariosNovoAction',
                                                                'Documento do Prestador: ',
                                                                false);
                  if fname = '' then
                  begin
                     re.lines.Add('Erro');

                     if Not wintask.EstaNaPagina_CopyText(mainPage, 'Honor�rios em Cadastramento') then
                        wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                     Sleep(TIME_SLEEP_MIN);

                     ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);

                     planJbm.dtsNotasPendentes.Close;
                     planJbm.dtsNotasPendentes.Open;;

                     //DBGrid2.SelectedRows.Clear;
                     DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                     exit;
                  end;

                  try
                     valor := ObtemDependenciaDigitar;
                     if valor = '' then
                     begin
                        Sleep(1000);
                        MyDeleteFile(fname);

                        fname := wintask.EstaNaPagina_NotepadCodePage(mainPage,
                                                                      rightDown, rightUp,
                                                                      leftDown, leftUp,
                                                                      'NOTEPAD.EXE|Notepad|finHonorariosNovoAction',
                                                                      'Documento do Prestador: ',
                                                                      false);
                        if fname = '' then
                        begin
                           if Not wintask.EstaNaPagina_CopyText(mainPage, 'Honor�rios em Cadastramento') then
                              wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                           Sleep(TIME_SLEEP_MIN);

                           ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);

                           planJbm.dtsNotasPendentes.Close;
                           planJbm.dtsNotasPendentes.Open;;

                           //DBGrid2.SelectedRows.Clear;
                           DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                           exit;
                        end;

                        valor := ObtemDependenciaDigitar;
                        if valor = '' then
                        begin
                           re.lines.Add('Erro obtendo o c�digo da depend�ncia');

                           if Not wintask.EstaNaPagina_CopyText(mainPage, 'Honor�rios em Cadastramento') then
                              wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                           Sleep(TIME_SLEEP_MIN);

                           ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
//                           winThread.PleaseClose:=true;

                           planJbm.dtsNotasPendentes.Close;
                           planJbm.dtsNotasPendentes.Open;;

                           //DBGrid2.SelectedRows.Clear;
                           DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                           exit;
                        end;
                     end;

                     (*
                     if StrToInt(valor) = 5404 then
                     begin
                        //verifica se tem o 4027
                        planJbm.ObtemReclamadasDoProcesso;
                        while not planJbm.dtsReclamadas.Eof do
                        begin
                           if planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = 4027 then
                           begin
                              valor := '4027';
                              break;
                           end;
                           planJbm.dtsReclamadas.Next;
                        end;
                     end;*)

                     if FileExists(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT') then
                     begin
                        RenameFile(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT', ExtractFilePath(Application.ExeName) + 'ENCERRAR.OK');
                        exit;
                     end;

                     wintask.SelectHTMLItem('GCPJ - Gest�o e Controle de Processos Jur�dicos', 'SELECT[NAME= ''selDependencia'']', Trim(valor));
                     Sleep(TIME_SLEEP_MIN);

                     if ie8 then
                        popup := 'IEXPLORE.EXE|Static|Esta depend�ncia foi encerrada, deseja alter�-la para'
                     else
                        popup := 'IEXPLORE.EXE|Static|Esta depend�ncia foi encerrada, deseja alter�-la para';

                     if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                     begin
                        wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                        Sleep(TIME_SLEEP_MIN);
                     end;

                     wintask.Click_Mouse(mainPage, 774,326);
                     Sleep(TIME_SLEEP_ZERO);

                     wintask.SendKey(mainPage,
                                     '<PageDown>');
                     Sleep(TIME_SLEEP_ZERO);

                     andamentoDigitar := ObtemTipoAndamento(planJbm.dtsAtos.FieldByName('tipoandamento').AsString, planJbm.tipoProcesso, planJbm.dtsAtos.FieldByName('motivobaixa').AsString);
                     if andamentoDigitar = '' then
                     begin
//                        planjbm.NotificaErro('Falta tipo ato');
                        ShowMessage('Falta tipo ato');
                        exit;
                     end;

                     wintask.SelectHTMLItem(pageGcpj, 'SELECT[NAME= ''despesaHonorariosProcessoVO.c'']', andamentoDigitar);

                     Sleep(TIME_SLEEP_med);

                     wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''incluir'']');
                     Sleep(TIME_SLEEP_MED);

                     passagem := 0;
                     while true do
                     begin
                        if ie8 then
                           popup := 'IEXPLORE.EXE|Static|O campo Data do Ato � obrigat�rio'
                        else
                           popup := 'IEXPLORE.EXE|Static|O campo Data do Ato � obrigat�rio';

                        if wintask.Verifica_Janela_PopUp(popup, valor, false) = 1 then
                        begin
                           if ie8 then
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                           else
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                           Sleep(TIME_SLEEP_MIN);

                           wintask.SendKey(mainPage,
                                           wintask.FormataNumericField(FormatDateTime('ddmmyyyy',planjbm.dtsatos.FieldByName('datadoato').AsDateTime)));
                           Sleep(TIME_SLEEP_zero);

                           wintask.SendKey(mainPage,
                                           '<Tab><Tab>');
                           Sleep(TIME_SLEEP_ZERO);

                           if (planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat = 0.00) and
                              (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORARIOS INICIAIS') then
                           begin
                               if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01' then
                                  valor := '5200'
                               else
                                  valor := '5000';
                           end
                           else
                              valor := RemoveNaoNumericos(FormatFloat('#,##0.00', planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat));

                           Sleep(time_sleep_zero);
                           wintask.SendKey(mainPage,
                                           '<Del><Del><Del><Del><Del><Del><Del><Del><Del><Del>');
                           Sleep(TIME_SLEEP_ZERO);

                           valDigitar := '';
                           while Length(valor) > 0 do
                           begin
                              valDigitar := valDigitar + '<Num ' + Copy(valor, 1, 1) + '>';
                              valor := Copy(valor, 2, length(valor));
                           end;
                           wintask.SendKey(mainPage,
                                           valDigitar);
                           Sleep(TIME_SLEEP_ZERO);

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''incluir'']');
                           Sleep(TIME_SLEEP_MIN);


                           if ie8 then
                              popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio'
                           else
                              popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio';

                           Sleep(TIME_SLEEP_ZERO);

                           if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                           begin
                              if ie8 then
                                 wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                              else
                                 wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                              Sleep(TIME_SLEEP_MIN);
                              valor := RemoveNaoNumericos(FormatFloat('0.00', planJbm.dtsAtos.FieldByName('valor').AsFloat));
                              valDigitar := '';
                              while Length(valor) > 0 do
                              begin
                                 valDigitar := valDigitar + '<Num ' + Copy(valor, 1, 1) + '>';
                                 valor := Copy(valor, 2, length(valor));
                              end;
                              wintask.SendKey(mainPage,
                                              valDigitar);
                              Sleep(TIME_SLEEP_MIN);

                              wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''incluir'']');
                              Sleep(TIME_SLEEP_MED);
                              break;
                           end;

                           break;
                        end;

                        if ie8 then
                           popUp := 'IEXPLORE.EXE|Static|O campo C�digo do Destinat�rio � obrigat�rio.'
                        else
                           popUp := 'IEXPLORE.EXE|Static|O campo C�digo do Destinat�rio � obrigat�rio.';

                        Sleep(TIME_SLEEP_ZERO);

                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           if ie8 then
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                           else
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                           Sleep(TIME_SLEEP_MIN);

                           wintask.WriteHTML(pageGcpj, 'INPUT TEXT[NAME= ''despesaHonorariosProcessoVO.composicaoDespHonorariosVO.dependenciaBradescoOnLineVO.c'']',
                                             planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString);
                           Sleep(TIME_SLEEP_MED);

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT BUTTON[NAME= ''null'',INDEX=''3'']');
                           Sleep(TIME_SLEEP_MED);

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''incluir'']');
                           Sleep(TIME_SLEEP_MED);
                           continue;
                        end;

                        if ie8 then
                           popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio'
                        else
                           popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio';

                        Sleep(TIME_SLEEP_ZERO);
                           
                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           if ie8 then
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                           else
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                           Sleep(TIME_SLEEP_MIN);
                           valor := RemoveNaoNumericos(FormatFloat('0.00', planJbm.dtsAtos.FieldByName('valor').AsFloat));
                           valDigitar := '';
                           while Length(valor) > 0 do
                           begin
                              valDigitar := valDigitar + '<Num ' + Copy(valor, 1, 1) + '>';
                              valor := Copy(valor, 2, length(valor));
                           end;
                           wintask.SendKey(mainPage,
                                           valDigitar);
                           Sleep(TIME_SLEEP_MIN);

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''incluir'']');
                           Sleep(TIME_SLEEP_MED);
                           break;
                        end;

                        if ie8 then
                           popup := 'IEXPLORE.EXE|Static|O campo Depend�ncia � obrigat�rio'
                        else
                           popup := 'IEXPLORE.EXE|Static|O campo Depend�ncia � obrigat�rio';

                        Sleep(TIME_SLEEP_ZERO);

                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           if ie8 then
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                           else
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                           Sleep(TIME_SLEEP_MIN);
//                           planJbm.NotificaErro('Erro na captura do nome da Dependencia');

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                           Sleep(TIME_SLEEP_MIN);

                           ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                           try
                              winThread.pleaseClose:= true;
                           except
                           end;

                           planJbm.dtsNotasPendentes.Close;
                           planJbm.dtsNotasPendentes.Open;;

//                           ShowMessage('Erro na captura do nome da Dependencia');
                           //DBGrid2.SelectedRows.Clear;
                           DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                           exit;
                        end;

                        if ie8 then
                           popup := 'IEXPLORE.EXE|Static|DATA DO ATO INVALIDO'
                        else
                           popup := 'IEXPLORE.EXE|Static|DATA DO ATO INVALIDO';

                        Sleep(TIME_SLEEP_ZERO);
                           
                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           if ie8 then
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                           else
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                           Sleep(TIME_SLEEP_MIN);
//                           planJbm.NotificaErro('Data digitada inv�lida');

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                           Sleep(TIME_SLEEP_MIN);

                           ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                           try
                              winThread.pleaseClose:= true;
                           except
                           end;

                           planJbm.dtsNotasPendentes.Close;
                           planJbm.dtsNotasPendentes.Open;;

//                           ShowMessage('Erro na captura do nome da Dependencia');
                           //DBGrid2.SelectedRows.Clear;
                           DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                           exit;
                        end;


                        Sleep(TIME_SLEEP_MED*20);

                        if wintask.EstaNaPagina_CopyText(mainPage, 'Atos Pagos no Processo') then
                        begin
                           wintask.ClickHTMLElement('GCPJ - Gest�o e Controle de Processos Jur�dicos', 'INPUT BUTTON[VALUE= ''confirmar inclus�o'']');
                           Sleep(TIME_SLEEP_ZERO);
                           break;
                        end;

                        if ie8 then
                           popup := 'IEXPLORE.EXE|Static|Pressione a lupa para realizar a pesquisa'
                        else
                           popup := 'IEXPLORE.EXE|Static|Pressione a lupa para realizar a pesquisa';

                        Sleep(TIME_SLEEP_ZERO);

                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           if ie8 then
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                           else
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                           Sleep(TIME_SLEEP_MIN);

                           ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                           try
                              winThread.pleaseClose:= true;
                           except
                           end;

                           planJbm.dtsNotasPendentes.Close;
                           planJbm.dtsNotasPendentes.Open;;

//                           ShowMessage('Erro na captura do nome da Dependencia');
                           //DBGrid2.SelectedRows.Clear;
                           DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                           exit;
                        end;

                        Inc(passagem);
                        if passagem > 5 then
                        begin
                           break;
//                           planJbm.NotificaErro('Erro de posicionamento de tela');
//                           ShowMessage('Erro de posicionamento de tela');
//                           exit;
                        end;
                        Sleep(500);
                     end;

                     Sleep(TIME_SLEEP_MED*20);

                     if wintask.EstaNaPagina_CopyText(mainPage, 'Atos Pagos no Processo') then
                     begin
                        wintask.ClickHTMLElement('GCPJ - Gest�o e Controle de Processos Jur�dicos', 'INPUT BUTTON[VALUE= ''confirmar inclus�o'']');
                        Sleep(TIME_SLEEP_ZERO);
                     end;
                  finally
                     MyDeleteFile(fname);
                  end;

                  Sleep(TIME_SLEEP_MIN);

                  novototaldigitado := ObtemTotalDigitado;
                  if novototaldigitado = antigoTotalDigitado then
                  begin
                     Sleep(TIME_SLEEP_MED);
                     novototaldigitado := ObtemTotalDigitado;
                  end;

                  totaldigitado := planJbm.SomaTotalDigitadoDaNota;
                  antigoTotalDigitado := novototaldigitado;

                  if FormatFloat('0.00', totaldigitado + planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat) <> FormatFloat('0.00', novototaldigitado) then
                  begin
                     if FormatFloat('0.00', totaldigitado) = FormatFloat('0.00', novototaldigitado) then   //nao entrou a ultimadigitacao
                     begin
                        ret := VerificaTela(true);
                        case ret of
                           0 : begin
                              Sleep(TIME_SLEEP_ZERO);
                              try
                                 winThread.pleaseClose:= true;
                              except
                              end;
                              continue;
                           end;

                           1 : begin
                              wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
                              Sleep(TIME_SLEEP_MIN);

                              ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                              try
                                 winThread.pleaseClose:= true;
                              except
                              end;

                              planJbm.dtsNotasPendentes.Close;
                              planJbm.dtsNotasPendentes.Open;;

                              //DBGrid2.SelectedRows.Clear;
                              DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                              exit;
                           end;
                           -1, -2 : begin
                              ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                              try
                                 winThread.pleaseClose:= true;
                              except
                              end;
                              ShowMessage('Erro inesperado');
                              exit;
                           end;
                        end;
                     end;
                  end;

      //           totaldigitado := totaldigitado + planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat;
                  planJbm.MarcaProcessoDigitado;
                  if FileExists(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT') then
                  begin
                     RenameFile(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT', ExtractFilePath(Application.ExeName) + 'ENCERRAR.OK');
                     exit;
                  end;
                  planJbm.dtsAtos.Next;
                  if planJbm.dtsAtos.eof then
                  begin
                     if ultdocumento <> '' then
                     begin
                        //se foi finalizado com os valores corretos Finaliza a nota e notifica
                        if planJbm.totalAtos = planJbm.ObtemTotalAtosDigitados then
                        begin
                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''finalizar nota'']');
                           Sleep(TIME_SLEEP_MED);

                           if ie8 then
                              popup := 'IEXPLORE.EXE|Static|Opera��o realizada com sucesso'
                           else
                              popup :='IEXPLORE.EXE|Static|Opera��o realizada com sucesso';
                           if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                           begin
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
                              Sleep(TIME_SLEEP_MED);
                           end
                           else
                           begin
                              popup := 'IEXPLORE.EXE|Static|VALOR DO DOCUMENTO DEVE SER IGUAL A SOMATORIA DOS ATOS';
                              if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                              begin
                                 wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
                                 Sleep(TIME_SLEEP_MED);

                                 //confere todos os atos da Nota
                                 ret := ConfereAtosDaNota(IntToStr(planJbm.numerodanota),wintask);
                                 if ret <> 1 then
                                 begin
                                    ShowMessage('Erro finalizando a nota');
                                    exit;
                                 end
                                 else
                                 begin
                                    ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                                    try
                                       winThread.pleaseClose:= true;
                                    except
                                    end;

                                    planJbm.dtsNotasPendentes.Close;
                                    planJbm.dtsNotasPendentes.Open;;

                                    //DBGrid2.SelectedRows.Clear;
                                    DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                                    exit;
                                 end;  
                              end
                              else
                              begin
                                 ShowMessage('Erro finalizando a nota');
                                 exit;
                              end;
                           end;
                        end
                        else
                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
//                        planJbm.NotificaNotaFinalizada;
                        planJbm.MarcaDocumentoFinalizado;
                        digitando := false;
                        break;
                     end;
                  end;
               end;

               if digitando then
               begin
                  if planJbm.totalAtos = planJbm.ObtemTotalAtosDigitados then
                  begin
                     wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''finalizar nota'']');
                     Sleep(TIME_SLEEP_MED);

                     if ie8 then
                        popup := 'IEXPLORE.EXE|Static|Opera��o realizada com sucesso'
                     else
                        popup :='IEXPLORE.EXE|Static|Opera��o realizada com sucesso';
                     if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                     begin
                        wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
                        Sleep(TIME_SLEEP_MED);
                     end
                     else
                     begin
                        popup := 'IEXPLORE.EXE|Static|VALOR DO DOCUMENTO DEVE SER IGUAL A SOMATORIA DOS ATOS';
                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
                           Sleep(TIME_SLEEP_MED);

                           //confere todos os atos da Nota
                           ret := ConfereAtosDaNota(IntToStr(planJbm.numerodanota), wintask);
                           if ret <> 1 then
                           begin
                              ShowMessage('Erro finalizando a nota');
                              exit;
                           end
                           else
                           begin
                              ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
                              try
                                 winThread.pleaseClose:= true;
                              except
                              end;

                              planJbm.dtsNotasPendentes.Close;
                              planJbm.dtsNotasPendentes.Open;;

                              //DBGrid2.SelectedRows.Clear;
                              DigitaNotas(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);
                              exit;
                           end;
                        end
                        else
                        begin
                           ShowMessage('Erro finalizando a nota');
                           exit;
                        end;
                     end;
                  end
                  else
                     wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
//                  planJbm.NotificaNotaFinalizada;
                  planJbm.MarcaDocumentoFinalizado;
               end;
            finally
               ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
            end;
         finally
            wintask.Free;
         end;
      end;
   finally
      dbGrid2.DataSource.DataSet.EnableControls;
   end;

   planJbm.dtsNotasPendentes.Close;
   planJbm.dtsNotasPendentes.Open;

   if planJbm.dtsNotasPendentes.RecordCount = 0 then
   begin
      planJbm.dtsPlanilha.Close;
      planJbm.dtsPlanilha.Open;
   end;
end;

procedure TfrmDigitaPlan.BitBtn5Click(Sender: TObject);
var
   ini : TIniFile;
begin
   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
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

procedure TfrmDigitaPlan.BitBtn6Click(Sender: TObject);
var
   ini : TIniFile;
begin
(**
   ini := TiniFile.Create(ExtractFilePath(Application.ExeName) + 'CONFIG.INI');
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

      planjbm.tipoProcesso := planJbm.dtsAtos.FieldByName('fgtipoprocesso').asString;

      if planjbm.dtsatos.FieldByName('tipoandamento').AsString <> 'PREPOSTO' then
      begin
         dmgcpcj_base_XI.ObtemOutrosPagamentosDoProcesso(planJbm.gcpj,
                                                         planjbm.dtsatos.FieldByName('tipoandamento').AsString,
                                                         UpperCase(planJbm.dtsPlanilha.FieldByName('nomeescritorio').AsString),
                                                         planjbm.tipoprocesso,
                                                         );
         if dmgcpcj_base_I.dts.RecordCount > 0 then
         begin
            if (planjbm.dtsatos.FieldByName('tipoandamento').AsString = 'AVULSO TRABALHISTA') OR
               (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'INSTRUCAO') then
            begin
               if dmgcpcj_base_I.dts.RecordCount >= 2 then
               begin
                  planJbm.GravaOcorrencia(2, 'Ato j� foi pago anteriormente');
                  planJbm.MarcaProcessoCruzadoGcpj(9);
                  planJbm.RemoveNotaProcesso;
               end
            end
            else
            begin
               planJbm.GravaOcorrencia(2, 'Ato j� foi pago anteriormente');
               planJbm.MarcaProcessoCruzadoGcpj(9);
               planJbm.RemoveNotaProcesso;
            end;
         end;
      end;
      planJbm.dtsAtos.Next;
   end;**)
end;

procedure TfrmDigitaPlan.BitBtn7Click(Sender: TObject);
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
                  if Pos('Processo Ato Data do Ato  Valor Depend�ncia', arquivo.strings[i]) = 0 then
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

procedure TfrmDigitaPlan.BitBtn1Click(Sender: TObject);
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

procedure TfrmDigitaPlan.DBGrid2DblClick(Sender: TObject);
begin
   if dbGrid2.SelectedRows.Count > 1 then
   begin
      ShowMessage('Somente uma nota pode ser selecionada para edi��o');
      exit;
   end;

   if dbGrid2.SelectedRows.Count = 0 then
   begin
      ShowMessage('Nenhuma nota selecionada para edi��o');
      exit;
   end;


   frmAdosDigitados := TfrmAdosDigitados.Create(nil);
   try
      frmAdosDigitados.lblNumeroNota.Caption := 'Atos digitados da nota ' + DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString;

      dmHonorarios.dtsAtosDigitados.Close;
      dmHonorarios.dtsAtosDigitados.CommandText := 'select * from tbplanilhasatos where numeronota = ' + DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString +
                                                   ' and fgdigitado = 1 ' +
                                                   'order by gcpj';
      dmHonorarios.dtsAtosDigitados.Open;

      if dmHonorarios.dtsAtosDigitados.RecordCount = 0 then
      begin
         ShowMessage('Nenhum ato digitado para esta nota');
         exit;
      end;

      frmAdosDigitados.valorDigitado := dbgrid3.DataSource.DataSet.FieldByName('valordigitado').AsFloat;
      frmAdosDigitados.totalDigitado.Text  := FormatFloat('#,##0.00',dbgrid3.DataSource.DataSet.FieldByName('valordigitado').AsFloat);
      frmAdosDigitados.numeronota := DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString;
      frmAdosDigitados.ShowModal;

      if frmAdosDigitados.alterado then
      begin
         planJbm.dtsNotasPendentes.Close;
         planJbm.dtsNotasPendentes.Open;

         if planJbm.dtsNotasPendentes.RecordCount = 0 then
         begin
            planJbm.dtsPlanilha.Close;
            planJbm.dtsPlanilha.Open;
         end;
      end;
   finally
      frmAdosDigitados.Free;
   end;
end;

procedure TfrmDigitaPlan.Retornarnotafinalizada1Click(Sender: TObject);
begin
   frmRetornarFinalizada := TFrmRetornarFinalizada.Create(nil);
   try
      if DBGrid2.DataSource.DataSet.RecordCount > 0 then
         frmRetornarFinalizada.numNota.Text := DBGrid2.DataSource.DataSet.FieldByName('numernota').AsString;

      frmRetornarFinalizada.numNota.tag := 0;
      if frmRetornarFinalizada.ShowModal = mrOk then
      begin
         planJbm.dtsPlanilha.Close;
         planJbm.dtsPlanilha.Open;
      end;
   finally
      frmRetornarFinalizada.Free;
   end;
end;

function TfrmDigitaPlan.ObtemTipoAndamento(andamento: string; tipoProcesso: string; motivoBaixa: string): string;
begin
   result := '';
   if (UpperCase(andamento) = 'RECURSO EM FASE DE EXECU��O') or
      (UpperCase(andamento) = 'RECURSO EM FASE DE EXECUCAO') OR
      (UpperCase(andamento) = 'RECURSO NA FASE DE EXECU��O') or
      (UpperCase(andamento) = 'RECURSO NA FASE DE EXECUCAO') then
      result := 'RECURSO NA FASE DE EXECUCAO'
   else
   if (Pos('RECURSO', UpperCase(andamento)) <> 0) OR
      (Pos('AGRAVO FAT', UpperCase(andamento)) <> 0) then
      result := 'RECURSO'
   else
   if Pos('TUTELA', UpperCase(andamento)) <> 0 then
      result := 'CONTESTACAO'
   else
   if UpperCase(andamento) = 'PREPOSTO' then
      result := 'PREPOSTO'
   else
   if (Pos('AVULSO TRAB', UpperCase(andamento)) <> 0) or
      (((tipoprocesso = 'TR') OR (tipoprocesso = 'TO') OR (tipoprocesso = 'TA')) and
      (UpperCase(andamento) = 'AUDI�NCIA TRABALHISTA')or (UpperCase(andamento) = 'AUDIENCIA TRABALHISTA') or
      (UpperCase(andamento) = 'AUDI�NCIA AVULSA') or (UpperCase(andamento) = 'AUDI�NCIA AVULSA') or
      (UpperCase(andamento) = 'INSTRUCAO') or (UpperCase(andamento) = 'AUDIENCIA AVULSA')) then
      result := 'INSTRUCAO'
   else
   if (Pos('�XITO', UpperCase(andamento)) <> 0) or (Pos('EXITO', UpperCase(andamento)) <> 0) then
      result := 'HONORARIOS DE EXITO'
   else
   if (UpperCase(andamento) = 'AVULSO') or
      (UpperCase(andamento) = 'AVULSO C�VEL (AUDITORIA)') or
      (UpperCase(andamento) = 'AVULSO CIVEL (AUDITORIA)') or
      (UpperCase(andamento) = 'AVULSO CIVEL') or
      (UpperCase(andamento) = 'AVULSO C�VEL') then
      result := 'ASSESSORIA JURIDICA'
   else
   if (UpperCase(andamento) = 'AUDI�NCIA AVULSA') or
      (UpperCase(andamento) = 'PARCELA') or
      (UpperCase(andamento) = 'AUDIENCIA AVULSA') then
      result := 'ACOMPANHAMENTO'
   else
   if (UpperCase(andamento) = 'HONORARIOS FINAIS') or
      (UpperCase(andamento) = 'HONOR�RIOS FINAIS') then
   begin
   //19/04/2017
(*      if (motivobaixa = 'ACORDO COM CUSTOS') or
         (motivobaixa = 'IMPROCEDENCIA') or
         (motivoBaixa = 'EXTINTO COM JULGAMENTO DE MERITO') then
         result := 'HONORARIOS DE EXITO'
      else*)
         result := 'HONORARIOS FINAIS';
   end
   else
   if (UpperCase(andamento) = 'HONORARIOS INICIAIS') or
      (UpperCase(andamento) = 'HONOR�RIOS INICIAIS') then
      result := 'HONORARIOS INICIAIS'
   else
   if Pos('AJUIZAMENTO', UpperCase(andamento)) <> 0 then
      result := 'AJUIZAMENTO'
   else
   if (andamento = 'CARTA PRECATORIA') or (andamento = 'CARTA PRECAT�RIA') then
      result := 'CUMPRIMENTO DE CARTA PRECATORIA'
   else
      result := UpperCase(andamento);
end;

function TfrmDigitaPlan.ConfereAtosDaNota(numeroNota: string; wintask: TWintask): integer;
var
   i, estagio,p, contador :integer;
   arquivo : TstringList;
   linha, linhaanterior, andamentoLocalizar, ultGcPj, gcpj, andamento, valorDigitado :string;
   fim : boolean;
   fName, gcpjValidado, andamentoValidado :string;
   tries :integer;
   data:TdateTime;
   ultArquivo : string;
label
   mesmaLinha;

   procedure RemoveAto(dts: TAdoDataset; tela, db: boolean);
   begin
      if tela then
      begin
         wintask.ClickHTMLElement(pageGcpj, 'INPUT RADIO[NAME= ''composicoesSelecs'',INDEX=''' + IntToStr(contador) + ''']');
         Sleep(TIME_SLEEP_ZERO);

         wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''excluir'']');
         Sleep(TIME_SLEEP_MIN);
      end;

      if db then
      begin
         //marca como n�o digitado
         dmHonorarios.adoCmd.CommandText := 'update tbplanilhasatos set fgdigitado=0, datahoradigitacao=null ' +
                                            'where idescritorio = ' + dts.FieldByName('idescritorio').AsString + ' and ' +
                                            'idplanilha = ' + dts.FieldByName('idplanilha').AsString + ' and ' +
                                            'linhaplanilha = ' + dts.FieldByName('linhaplanilha').AsString + ' and ' +
                                            'sequencia = ' + dts.FieldByName('sequencia').AsString;
         dmHonorarios.adoCmd.Execute;
      end;
   end;
begin
   result := -1;


   fname := '';

   wintask.Click_Mouse(mainPage,591,234);
   Sleep(500);

   if Not wintask.EstaNaPagina_CopyText(mainPage, 'Tipo de Prestador:') then
      exit;

   //posiciona na primeirap�gina da lista de atos
   wintask.ClickHTMLElement(pageGcpj, 'IMG[SRC= ''http://www4.net.bradesco.com.br/i'']');
   Sleep(TIME_SLEEP_MIN);

   //localiza os atos marcados como digitados no banco de dados
   dmHonorarios.dtsAtosDigitados.Close;
   dmHonorarios.dtsAtosDigitados.CommandText := 'select * from tbplanilhasatos where numeronota = ' + DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString +
                                                ' and fgdigitado = 1 ' +
                                                'order by gcpj, valorCorrigido';
   dmHonorarios.dtsAtosDigitados.Open;

   if dmHonorarios.dtsAtosDigitados.RecordCount = 0 then
   begin
      //no banco de dados nenhum ato est� marcado comodigitado.Exclui a nota doGCPJ
      wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
      Sleep(TIME_SLEEP_MIN);

      if not wintask.EstaNaPagina_CopyText(mainPage, 'Honor�rios em Cadastramento') then
         exit;

      if PosicionaNotaDigitando(false,DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString, true,wintask)<0 then
         exit;
      result := 1;
      exit;
   end;

   fName := wintask.GetPagina_CopyText(mainPage, 'Tipo de Prestador: ');
   if fName = '' then
      exit;

   //posiciona na tela do GCPJ
   arquivo :=TstringList.Create;

   i := 0;
   linhaAnterior := '';
   contador := 0;
   fim := false;
   gcpjValidado := '';
   tries := 0;
   ultArquivo := '';

   dmHonorarios.dtsAtosDigitados.First;

   try
      while true do
      begin
         arquivo.Clear;

         if ultArquivo = fName then
         begin
            MyDeleteFile(fName);
            Result := ConfereAtosDaNota(numeroNota, wintask);
            exit;
         end;

         arquivo.LoadFromFile(fname);
         estagio := 0;

         for i := 0 to arquivo.Count - 1 do
         begin
            if Trim(arquivo.strings[i]) = '' then
               continue;

            if estagio = 0 then
            begin
               if Pos('Processo Ato Data do Ato  Valor', arquivo.Strings[i]) = 0 then
                  continue;
               inc(estagio);
               contador := 0;
               if arquivo.Strings[i+1] = linhaanterior then   //acabaram as p�ginas
               begin
                  //tenta  mais uma vez
                  inc(tries);
                  if tries > 1 then
                  begin
                     fim:= true;
                     break;
                  end;

                  fName := wintask.GetPagina_CopyText(mainPage, 'Tipo de Prestador: ');
                  if fName = '' then
                     exit;
                  MyDeleteFile(fname);
                  break;
               end;
               linhaanterior := arquivo.Strings[i+1];
               continue;
            end;

         mesmaLinha:

            linha := Trim(arquivo.strings[i]);
            if (linha = '') or (Pos('/', linha)=0) then
            begin
               //clica para mover para a pr�xima janela
               wintask.ClickHTMLElement(pageGcpj, 'IMG[SRC= ''http://www4.net.bradesco.com.br/imagens/setaTabelaD'']');

               Sleep(TIME_SLEEP_MED);

               estagio := 0;
               break;
            end;

            //achou a linha. Separa os dados GCPJ / Valor / Tipo andamento
            Inc(contador);

            //se acabou o banco de dados e ainda temna tela � para excluir
            if dmHonorarios.dtsAtosDigitados.Eof then
            begin
               RemoveAto(dmHonorarios.dtsAtosDigitados, true, false);
               continue;
            end;

            p := Pos(' ', linha);
            if p = 0 then
               continue;

            gcpj := Trim(Copy(linha, 1, p-1));
            linha:= Trim(Copy(linha,p+1, length(linha)));

            p := Pos('/', linha);
            if p = 0 then
               continue;

            andamento := Trim(Copy(linha, 1, p-3));
            linha := Trim(Copy(linha, p+1, length(linha)));

            p := Pos(FormatDateTime('/yyyy', Date), linha);
            if p = 0 then
            begin
               p := Pos('/' + InttoStr(YearOf(date)-1), linha);
               if p = 0 then
                  continue;
            end;

            valorDigitado := Trim(Copy(linha, p+5, length(linha)));
            p := Pos(' ', valorDigitado);
            if p = 0 then
               continue;

            valorDigitado := Trim(copy(valordigitado, 1, p-1));

            //verifica contra o banco de dados
            andamentoLocalizar := ObtemTipoAndamento(dmHonorarios.dtsAtosDigitados.FieldByName('tipoandamento').AsString, '', '');

            if dmHonorarios.dtsAtosDigitados.FieldByName('gcpj').AsInteger = StrToInt(gcpj) then //eh omesmo gcpj
            begin
               //� o mesmo tipo de andamento?
               if andamentoLocalizar = andamento then
               begin
                  //o valor est� correto?
                  if FormatFloat('#,##0.00', dmHonorarios.dtsAtosDigitados.FieldByName('valorcorrigido').AsFloat) = valorDigitado then //est� OK
                  begin
                     //posiciona no pr�ximo registro do banco de dados
                     dmHonorarios.dtsAtosDigitados.Next;
                     continue;
                  end;
                  //o valor � diferente. Altera no GCPJ
                  RemoveAto(dmHonorarios.dtsAtosDigitados, true, true);
                  MyDeleteFile(fName);
                  result := ConfereAtosDaNota(numeroNota, wintask);
                  exit;
               end;

               //andamento diferente. Ser� que a nota tem duplicado?
               dmHonorarios.dts.Close;
               dmHonorarios.dts.CommandText := 'select * from tbplanilhasatos where numeronota = ' + DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString +
                                               ' and gcpj = ' + gcpj + ' and linhaPlanilha <> ' + dmHonorarios.dtsAtosDigitados.FieldByName('linhaplanilha').AsString;
               dmHonorarios.dts.Open;
               if dmHonorarios.dts.Eof then //n�o encontrou. Exclui do GCPJ
               begin
                  MyDeleteFile(fName);
                  RemoveAto(dmHonorarios.dtsAtosDigitados, true, false);
                  result := ConfereAtosDaNota(numeroNota, wintask);
                  exit;
               end;

               //achou.
               dmHonorarios.dts.Tag := 0;
               while Not dmHonorarios.dts.Eof do
               begin
                  Application.ProcessMessages;

                  andamentoLocalizar := ObtemTipoAndamento(dmHonorarios.dts.FieldByName('tipoandamento').AsString, '', '');
                  if andamentoLocalizar = andamento then
                  begin
                     if FormatFloat('#,##0.00', dmHonorarios.dts.FieldByName('valorcorrigido').AsFloat) = valorDigitado then //est� OK
                     begin
                        //posiciona no pr�ximo registro do banco de dados
                        //consta como digitado no banco de dados?
                        if dmHonorarios.dts.FieldByName('fgdigitado').AsInteger = 0 then
                        begin
                           dmHonorarios.adoCmd.CommandText := 'update tbplanilhasatos set fgdigitado=1, datahoradigitacao=:data ' +
                                                              'where idescritorio = ' + dmHonorarios.dts.FieldByName('idescritorio').AsString + ' and ' +
                                                              'idplanilha = ' + dmHonorarios.dts.FieldByName('idplanilha').AsString + ' and ' +
                                                              'linhaplanilha = ' + dmHonorarios.dts.FieldByName('linhaplanilha').AsString + ' and ' +
                                                              'sequencia = ' + dmHonorarios.dts.FieldByName('sequencia').AsString;
                           dmHonorarios.adoCmd.Parameters.ParamByName('data').Value := now;
                           dmHonorarios.adoCmd.Execute;
                           MyDeleteFile(fName);
                           result := ConfereAtosDaNota(numeroNota, wintask);
                           exit;
                        end
                        else
                        begin
                           gcpjValidado := gcpj;
                           andamentoValidado := andamentoLocalizar;

                           dmHonorarios.dts.tag := 1;
                           dmHonorarios.dtsAtosDigitados.Next;
                           break;
                        end;
                     end;
                     //remove
                     RemoveAto(dmHonorarios.dts, true, true);
                     MyDeleteFile(fName);
                     result := ConfereAtosDaNota(numeroNota, wintask);
                     exit;
                  end;
                  //verifica o proximo
                  dmHonorarios.dts.Next;
               end;

               if dmHonorarios.dts.Tag = 1 then
                  continue;

               MyDeleteFile(fName);
               RemoveAto(dmHonorarios.dtsAtosDigitados, true, false);
               result := ConfereAtosDaNota(numeroNota, wintask);
               exit;
            end;

            if dmHonorarios.dtsAtosDigitados.FieldByName('gcpj').AsInteger > StrToInt(gcpj) then
            begin
               if gcpj = gcpjValidado then
               begin
                  if andamentoValidado = andamento then
                     continue;
                  MyDeleteFile(fName);
                  RemoveAto(dmHonorarios.dtsAtosDigitados, true, false);
                  result := ConfereAtosDaNota(numeroNota, wintask);
                  exit;
               end
               else
               begin
                     //gcpj do banco de dados � maior que a tela. Remove do
                  MyDeleteFile(fName);
                  RemoveAto(dmHonorarios.dtsAtosDigitados, true, false);
                  result := ConfereAtosDaNota(numeroNota, wintask);
                  exit;
               end;
            end;
            //verifica dupliocados
            if gcpjValidado = dmHonorarios.dtsAtosDigitados.FieldByName('gcpj').AsString then
            begin
               if andamentoLocalizar = andamentoValidado then
               begin
                  //andamento diferente. Ser� que a nota tem duplicado?
                  dmHonorarios.dts.Close;
                  dmHonorarios.dts.CommandText := 'select * from tbplanilhasatos where numeronota = ' + DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString +
                                                  ' and gcpj = ' + gcpjValidado + ' and linhaPlanilha <> ' + dmHonorarios.dtsAtosDigitados.FieldByName('linhaplanilha').AsString;
                  dmHonorarios.dts.Open;
                  if dmHonorarios.dts.Eof then //n�o encontrou. Exclui do GCPJ
                     exit;

                  if dmHonorarios.dts.FieldByName('fgdigitado').AsInteger = 0 then
                  begin
                     dmHonorarios.dtsAtosDigitados.Next;
                     goto mesmaLinha;
                  end;
                  MyDeleteFile(fName);
                  RemoveAto(dmHonorarios.dts, false, true);
                  result := ConfereAtosDaNota(numeroNota, wintask);
                  exit;
               end
               else
               begin
                  MyDeleteFile(fName);
                  RemoveAto(dmHonorarios.dtsAtosDigitados, false, true);
                  result := ConfereAtosDaNota(numeroNota, wintask);
                  exit;
               end;
            end
            else
            begin
               MyDeleteFile(fName);
               RemoveAto(dmHonorarios.dtsAtosDigitados, false, true);
               result := ConfereAtosDaNota(numeroNota, wintask);
               exit;
            end;
         end;
         if Not Fim then
         begin
            fName := wintask.GetPagina_CopyText(mainPage, 'Tipo de Prestador: ');
            if fName = '' then
               exit;

            continue;
         end;

         //se sobrou bancode dados exclui
         if Not dmHonorarios.dtsAtosDigitados.Eof then
            result := -1
         else
            result :=1;
         exit;
      end;
   finally
      MyDeleteFile(fname);
      arquivo.free;
   end;
end;


function TfrmDigitaPlan.PosicionaNotaDigitando(notificar: boolean; nota: string; excluir: boolean;wintask: twintask):integer;
var
   i, estagio, index, p : integer;
   clicou : boolean;
   posicionarPagUm : boolean;
   primeiraNota, fname, numeronota : string;
   arquivo : TStringList;
   primeiralinha : integer;
   popup, valor : string;

begin
   result := -1;
   primeiraNota := '';
   clicou := false;

   wintask.Click_Mouse(mainPage, 591,234);

   posicionarPagUm := true;

   //localizar a tela da nota
   while true do
   begin
      fName := wintask.GetPagina_CopyText(mainPage,
                                          'Honor�rios em Cadastramento');
      if fName <> '' then //esta na pagina principal
      begin
         if posicionarPagUm then
         begin
            //ja esta na pagina 1?
            arquivo:= TStringList.Create;
            try
               arquivo.LoadFromFile(fName);
               for i := 0 to arquivo.count - 1 do
               begin
                  if Pos('[1]', arquivo.strings[i])<> 0 then
                  begin
                     posicionarPagUm := false;
                     break;
                  end;
               end;
            finally
               arquivo.Free;
            end;

            if posicionarPagUm then
            begin
               wintask.ClickHTMLElement(pageGcpj, 'IMG[SRC= ''http://www4.net.brad'']');
               Sleep(TIME_SLEEP_ZERO);
               posicionarPagUm := false;

               MyDeleteFile(fName);

               fName := wintask.GetPagina_CopyText(mainPage, 'Honor�rios em Cadastramento');

            end;
         end;

         estagio := 0;
         index := 1;

         arquivo := TStringList.Create;
         try
            arquivo.LoadFromFile(fname);
            for i := 0 to arquivo.count - 1 do
            begin
               case estagio of
                  0 : begin
                     if Pos('Empresa Prestador de Servi�o', arquivo.strings[i]) <> 0 then
                     begin
                        inc(estagio);

                        primeiralinha := i+1;
                     end;
                  end;
                  1 : begin
                     p:= Pos('NOTA FISCAL', arquivo.strings[i]);
                     if p <> 0 then
                     begin
                        numeronota := Trim(Copy(arquivo.strings[i], p+11, Length(arquivo.strings[i])));
                        p := Pos(' ', numeroNota);
                        if p <> 0 then
                           numeroNota := Trim(Copy(numeroNota, 1, p-1));
                     end;

                     if clicou then
                     begin
                        if numeronota = primeiraNota then // n�o achou
                        begin
//                           planJbm.NotificaErro('Erro localizando a nota: ' + nota);
                           if notificar then
                              ShowMessage('Erro localizando a nota: ' + nota);
                           exit;
                        end;
                        clicou:= false;
                     end;

                     if nota = numeroNota then //achou
                     begin
                        Sleep(TIME_SLEEP_Med);

                        wintask.ClickHTMLElement(pageGcpj, 'INPUT RADIO[NAME= ''selecionado'',INDEX=''' + IntToStr(index) + ''']');
                        Sleep(TIME_SLEEP_MIN);

                        if Not excluir then
                        begin
                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''alterar'']');
                           Sleep(10000);
                        end
                        else
                        begin
                           wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''excluir'']');
                           Sleep(TIME_SLEEP_MED);
                        end;

                        popup := 'IEXPLORE.EXE|Static|Selecione um dos itens da grid';

                        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
                        begin
                           if wintask.ie8 then
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
                           else
                              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

                           Sleep(TIME_SLEEP_MIN);

                           wintask.ClickHTMLElement(pageGcpj, 'INPUT RADIO[NAME= ''selecionado'',INDEX=''' + IntToStr(index) + ''']');
                           Sleep(TIME_SLEEP_MIN);

                           if Not excluir then
                           begin
                              wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''alterar'']');
                              Sleep(10000);
                           end
                           else
                           begin
                              wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''excluir'']');
                              Sleep(TIME_SLEEP_MED);
                           end;
                        end;

                        break;
                     end;
                     Inc(index)
                  end;
               end;
            end;

            if i <= (arquivo.Count - 1) then
               break;

            //guarda a primeira nota da p�gina

            p := Pos('NOTA FISCAL', arquivo.strings[primeiraLinha]);
            primeiraNota := Trim(Copy(arquivo.strings[primeiraLinha], p+11, Length(arquivo.strings[primeiraLinha])));
            p := Pos(' ', primeiraNota);
            if p <> 0 then
               primeiraNota := Trim(Copy(primeiraNota, 1, p-1));

            wintask.ClickHTMLElement(pageGcpj, 'IMG[SRC= ''http://www4.net.bradesco.com.br/imagens/setaTabelaD'']');
            Sleep(TIME_SLEEP_MED);

            clicou := true;

         finally
            arquivo.free;
            MyDeleteFile(fName);
         end;
      end
      else
      begin
         if Not  wintask.EstaNaPagina_CopyText(mainPage,
                                               'Tipo de Prestador:') then
         begin
            wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
            sleep(TIME_SLEEP_MIN);

            Result := PosicionaNotaDigitando(notificar, nota, excluir, wintask);
            exit;
         end
         else
         begin
            wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
            sleep(TIME_SLEEP_MIN);
            Result := PosicionaNotaDigitando(notificar, nota, excluir, wintask);
            exit;
         end;
      end;
   end;
   result := 1;
end;

procedure TfrmDigitaPlan.ConferirAtosdaNotaComoGCPJ1Click(Sender: TObject);
var
   wintask : TWintask;
   ret: integer;
   ini : tinifile;
   diretorio : string;
begin
   ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
   try
      diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_programa\Presenta');
   finally
      ini.Free;
   end;

   wintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                              diretorio + '\rob\',
                              '');

   try
      wintask.ie8 := ie8;
      wintask.ff := ff;

      wintask.FecharTodos('TaskExec.exe');
      wintask.FecharTodos('TaskLock.exe');
      wintask.FecharTodos('Tasksync.exe');

      if PosicionaNotaDigitando(false,DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString, false,wintask)<0 then
         exit;

      ret := ConfereAtosDaNota(DBGrid2.DataSource.DataSet.FieldByName('numeronota').asString, wintask);
      if ret = 1 then
         ShowMessage('Nota validada com sucesso')
      else
         ShowMessage('Erro na valida��o da nota');
   finally
      wintask.free;
   end;
end;

procedure TfrmDigitaPlan.MyDeleteFile(nome: string);
begin
   try
      DeleteFile(nome);
   except
   end;
end;

procedure TfrmDigitaPlan.NovaRotinaDigitacaoNotas(notaRetornar: string; jaPosicionado: boolean);
var
   wintask :  TWintask;
   ultdocumento : string;
   fName : string;
   valor, valDigitar : string;
   digitando : boolean;
   i, estagio, index, p, k : integer;
   arquivo : TStringList;
   destinatarioOk : boolean;
//   nota : string;
   passagem : integer;
   totaldigitado, novototaldigitado, antigoTotalDigitado: double;
//   dataDigitar : string;
   popup : string;
   numeronota : string;
   bMark : TBookmarkList;
   fhandle :integer;
   ret : integer;
   andamentoDigitar : string;
   busy : WordBool;
   ini : tinifile;
   diretorio : string;

   label
      proximo;

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

   function ObtemTotalDigitado:double;
   var
      i, p, passagem : integer;
   begin
      result := 0;
      passagem := 1;

      Sleep(800);

      while passagem <= 3 do
      begin
         wintask.Click_Mouse(mainPage,591,234);
         Sleep(500);

         fName := wintask.GetPagina_CopyText(mainPage,
                                                'Tipo de Prestador: ');
         if fName = '' then
         begin
            Inc(passagem);
            Sleep(800);
            continue;
         end;

         //esta na pagina principal
         arquivo := TStringList.Create;
         try
            arquivo.LoadFromFile(fname);
            for i := 0 to arquivo.count - 1 do
            begin
               p := Pos('Total: ', arquivo.strings[i]);
               if p = 0 then
                  continue;

               valor := Trim(Copy(arquivo.strings[i], p+7, length(arquivo.Strings[i])));
               RemoveEsteCaracter('.', valor);
               RemoveEsteCaracter(' ', valor);

               if valor = '' then
                  result := 0
               else
                  result := StrToFloat(valor);
               exit;
            end;
         finally
            arquivo.free;
            MyDeleteFile(fName);
         end;
      end;
   end;


   function VerificaTela(comThread: boolean):integer;
   begin
      result := -1;

      if comThread then
      begin
//         winThread:= TClickOkThread.Create;
//         winThread.Resume;
      end;

      try
        popup := 'IEXPLORE.EXE|Static|CODIGO PRESTADOR INVALIDO';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
        popup := 'TASKEXEC.EXE|Static|Error at line';

        if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
        begin
           if ie8 then
              wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
           else
              wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|CODIGO DO TIPO DOCUMENTO HONORARIOS INVALIDO';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|Pressione a lupa para realizar a pesquisa';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|DATA DO ATO INVALIDO';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            Sleep(TIME_SLEEP_ZERO);

            result :=1;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio'
         else
            popUp := 'IEXPLORE.EXE|Static|O campo Valor do Ato � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|O campo Data do Ato � obrigat�rio'
         else
            popup := 'IEXPLORE.EXE|Static|O campo Data do Ato � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor, false) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
                             Sleep(TIME_SLEEP_ZERO);
            result := 0;
            exit;
         end;

         Sleep(200);
         popUp := 'IEXPLORE.EXE|Static|O campo Data do Documento � obrigat�rio';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if wintask.ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            result := 1;
            exit;
         end;

         Sleep(200);
         popUp := 'IEXPLORE.EXE|Static|O campo N� do Processo Bradesco � obrigat�rio';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if wintask.ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            result := 1;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popUp := 'IEXPLORE.EXE|Static|O campo C�digo do Destinat�rio � obrigat�rio.'
         else
            popUp := 'IEXPLORE.EXE|Static|O campo C�digo do Destinat�rio � obrigat�rio.';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|O campo Depend�ncia � obrigat�rio'
         else
            popup := 'IEXPLORE.EXE|Static|O campo Depend�ncia � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;
         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|O campo Ato � obrigat�rio'
         else
            popup := 'IEXPLORE.EXE|Static|O campo Ato � obrigat�rio';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            if ie8 then
               wintask.Click_Button('IEXPLORE.EXE|#32770|Mensagem da p�gina da web','OK')
            else
               wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');

            result := 0;
            exit;

         end;

         Sleep(200);
         if ie8 then
            popup := 'IEXPLORE.EXE|Static|Esta depend�ncia foi encerrada, deseja alter�-la para'
         else
            popup := 'IEXPLORE.EXE|Static|Esta depend�ncia foi encerrada, deseja alter�-la para';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer', 'OK');
            result := 0;
            exit;
         end;


         Sleep(200);
         popup := 'IEXPLORE.EXE|Static|VALOR DO DOCUMENTO DEVE SER IGUAL A SOMATORIA DOS ATOS';
         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            wintask.Click_Button('IEXPLORE.EXE|#32770|Windows Internet Explorer','OK');
            result := 0;
            exit;
         end;

         Sleep(200);
         popup := 'IEXPLORE.EXE|Internet Explorer_Server|ERRO! - Windows Internet Explorer';

         if wintask.Verifica_Janela_PopUp(popup, valor) = 1 then
         begin
            result := -2;
            exit;
         end;

         result := 1;
         exit;
      finally
         try
            if comThread then
               winThread.pleaseClose:=true;
         except
         end;
      end;


//                     planJbm.NotificaErro('Valores digitados n�o conferem');
//                     ShowMessage('Valores digitados n�o conferem');
//                     exit;
//                  end;
   end;

   procedure ExcluiDoGcpj(contador:integer);
   begin
      //exclui o ato do GCPJ
      wintask.ClickHTMLElement(pageGcpj, 'INPUT RADIO[NAME= ''composicoesSelecs'',INDEX=''' + IntToStr(contador) + ''']');
      Sleep(TIME_SLEEP_ZERO);

      wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''excluir'']');
      Sleep(TIME_SLEEP_MED);

      wintask.ClickHTMLElement(pageGcpj, 'INPUT SUBMIT[VALUE= ''voltar'']');
      sleep(TIME_SLEEP_MED);
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

(***Retirado em 10/03/2015 a pedido da Nilce
   dataDigitar := planJbm.ObtemDataDigitar(1);
   if dataDigitar = '' then
   begin
      if MessageDlg('N�o encontrada a data inicial para o ano/m�s de refer�ncia. Deseja cadastrar?', mtConfirmation, mbYesNoCancel, 0) <> mrYes then
      begin
         planJbm.NotificaErro('Sistema n�o pode digitar para o m�s ' + planJbm.anomesreferencia + ' at� ser cadastrada a data inicial');
         ShowMessage('Sistema n�o pode digitar para o m�s ' + planJbm.anomesreferencia + ' at� ser cadastrada a data inicial');
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
      dataDigitar := planJbm.ObtemDataDigitar(1);
   end;
   **)

   dbGrid2.DataSource.DataSet.DisableControls;
   try
      if dbGrid2.SelectedRows.Count = 0 then //nehujm marcado. Marca todos
      begin
         //ainda tem registro?
         if DBGrid2.DataSource.DataSet.RecordCount = 0 then
         begin
            planJbm.MarcaPlanilhaFinalizada;
            planjbm.dtsPlanilha.Close;
            planjbm.dtsPlanilha.Open;
            exit;
         end;
         SelecionaTodaLista;
      end;

      bMark := dbGrid2.SelectedRows;

      ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
      try
         diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_programa\Presenta');
      finally
         ini.Free;
      end;


      for k := 0 to bMark.Count - 1 do
      begin
         wintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                                    diretorio + '\rob\',
                                    '');

         try
            wintask.ie8 := ie8;
            wintask.ff := ff;

            wintask.FecharTodos('TaskExec.exe');
            wintask.FecharTodos('TaskLock.exe');
            wintask.FecharTodos('Tasksync.exe');

            if FileExists(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT') then
            begin
               RenameFile(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT', ExtractFilePath(Application.ExeName) + 'ENCERRAR.OK');
               exit;
            end;

            dbGrid2.DataSource.DataSet.BookMark := bMark.items[k];

            digitando := false;
            ultdocumento := '';
            totaldigitado := 0;
            novototaldigitado := 0;
            antigototaldigitado := 0;

            if (notaRetornar <> '') and (DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString <> notaRetornar) then
               continue;

            planJbm.empresaLigada := DBGrid2.DataSource.DataSet.FieldByName('empresa_ligada_agrupar').AsInteger;
            planJbm.numerodanota :=  DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsInteger;

            ret := IsFileLocked(planJbm.ObtemDirLck + '\' + IntToStr(planJbm.numerodanota) + '.lck', fhandle);

            if ret <= 0 then
               continue;

            if not ehPresenta then
            begin
               if Not JaPosicionado then
               begin
                  wb.Navigate('http://www4.net.bradesco.com.br/gcpj/servlet/finHonorariosNovoAction?metodo=carregarHonorariosEmCadastramento');
                  Application.ProcessMessages;

                  while wb.busy = true do
                  begin
                     Sleep(200);
                     Application.ProcessMessages;
                  end;
               end;
            end;

            try
               if FormatDateTime('hh:nn', Now) > '22:00' then
                  exit;

               ret := planJbm.ObtemProcessosDigitar(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString);

               if ret = -1 then
                  exit;

               case ret of
                  0 : begin
                     ret := NovoPosicionaNotaDigitando;
                     if (ret = 1) then
                     begin
                        ret := FinalizaNotaNoGcpj;
                     end
                     else
                        ret := 1;
                     if ret = 1 then
                        planJbm.MarcaDocumentoFinalizado;
                  end;
               else
                  planJbm.ObtemValorTotalDaNota;
                  if planJbm.SomaTotalDigitadoDaNota > 0 then
                     ret := DigitaNotaEmAndamento(wintask, FormatDateTime('dd/mm/yyyy',planJbm.dtsAtos.FieldByName('datadoato').AsDateTime))
                  else
                     ret := DigitaNotaNova(wintask, FormatDateTime('dd/mm/yyyy',planJbm.dtsAtos.FieldByName('datadoato').AsDateTime));
               end;
               if ret < 0 then
                  exit;
            finally
               ReleaseFile(planJbm.ObtemDirLck + '\' +IntToStr(planJbm.numerodanota) + '.lck', fhandle);
            end;
         finally
            wintask.Free;
         end;
      end;
   finally
      dbGrid2.DataSource.DataSet.Close;
      dbGrid2.DataSource.DataSet.open;
      dbGrid2.DataSource.DataSet.EnableControls;
   end;
end;



procedure TfrmDigitaPlan.SelecionaTodaLista;
begin
   DBGrid2.DataSource.DataSet.DisableControls;
   try
      DBGrid2.DataSource.DataSet.first;
      while not DBGrid2.DataSource.DataSet.eof do
      begin
         dbGrid2.SelectedRows.CurrentRowSelected := true;
         DBGrid2.DataSource.DataSet.next;
      end;
   finally
      DBGrid2.DataSource.DataSet.EnableControls;
   end;
end;

procedure TfrmDigitaPlan.wbDocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
//   if pJanelaFechar <> '' then
//      FechaJanelaGcpj(pJanelaFechar);
   segue := true;
end;

procedure TfrmDigitaPlan.EsperaPaginaProcessar(const janelaFechar: string='');
begin
   Sleep(1000);

   pJanelaFechar := janelaFechar;

   while (wb.ReadyState = 0) or (wb.ReadyState = 1) or (wb.ReadyState = 3) do
      Application.ProcessMessages;

   while wb.Busy = true do
      Application.ProcessMessages;

   while Not segue do
   begin
      Sleep(1000);
      Application.ProcessMessages;
   end;
end;

function TfrmDigitaPlan.GetFormByName(const documento: IHTMLDocument2;
  nomeForm: string): IHTMLFormElement;
var
   forms : IHTMLElementCollection;
   form : IHTMLFormElement;
   idx : integer;
begin
   forms := documento.forms as  IHTMLElementCollection;
   result := Nil;
   for idx := 0 to forms.length - 1 do
   begin
      form := forms.item(idx, 0) as IHTMLFormElement;
      if UpperCase(form.name) = UpperCase(nomeForm) then
         result := form;
   end;
end;

function TfrmDigitaPlan.GetInputField(fromForm: IHTMLFormElement;
  const inputName: string; const instance: integer): HTMLInputElement;
var
  field: IHTMLElement;
begin
  field := fromForm.Item(inputName,instance) as IHTMLElement;
  if Assigned(field) then
  begin
    if UpperCase(field.tagName) = 'INPUT' then
    begin
      result := field as HTMLInputElement;
      exit;
    end;
  end;
  result := nil;
end;

function TfrmDigitaPlan.GetSelectField(fromForm: IHTMLFormElement;
  const selectName: string; const instance: integer): HTMLSelectElement;
var
  field: IHTMLElement;
begin
  field := fromForm.Item(selectName,instance) as IHTMLElement;
  if Assigned(field) then
  begin
    if UpperCase(field.tagName) = 'SELECT' then
    begin
      result := field as HTMLSelectElement;
      exit;
    end;
  end;
  result := nil;
end;


procedure TfrmDigitaPlan.wbNewWindow2(Sender: TObject;
  var ppDisp: IDispatch; var Cancel: WordBool);
begin
   case wbPopWin of
      0 : begin
         if Assigned(wbpopup) then
         begin
            wbPopup.Close;
            FreeAndNil(wbPopup);
         end;

         wbpopUp := TfrmGcpj.Create(Nil);
         wbpopUp.Caption := Self.Caption + ' (popup)';
         wbPopUp.wb.RegisterAsBrowser := true;
         wbpopUp.wb.OnDocumentComplete := PopUpWbDocumentComplete;
         ppDisp := wbpopUp.wb.Application;
      end;
      1 : begin
         if Assigned(wbPopUp) then
         begin
            wbPopUp.Close;
            FreeAndNil(wbPopUp);
         end;

         wbPopUp := TfrmGcpj.Create(Nil);
         wbPopUp.Caption := Self.Caption + ' (popup)';
         wbPopUp.wb.OnDocumentComplete := PopUpWbDocumentComplete;
         wbPopUp.wb.RegisterAsBrowser := true;
         ppDisp := wbPopUp.wb.Application;
      end;
   end;
   Cancel := false;
end;

function TfrmDigitaPlan.GetScriptField(fromForm: IHTMLFormElement;
  const selectName: string; const instance: integer): HTMLScriptElement;
var
  field: IHTMLElement;
begin
  field := fromForm.Item(selectName,instance) as IHTMLElement;
  if Assigned(field) then
  begin
    if field.tagName = 'SCRIPT' then
    begin
      result := field as HTMLScriptElement;
      exit;
    end;
  end;
  result := nil;
end;

procedure TfrmDigitaPlan.PopUpWbDocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
var
   newWb : TWebBrowser;
   form : Variant;
   elemCollection : Variant;
   idx : integer;
   newForm : IHTMLFormElement;
begin
   newWb := (Sender as TWebBrowser);
   form := newWb.OleObject.Document.Forms.Item(0);

   case wbPopWin of
      0 : begin
         elemCollection := form.getElementsByTagName('a');
         For idx := 0 to elemCollection.length - 1 do
         begin
            elemCollection.item(idx).Click;
            segue := true;
            wbPopUp.Close;
            exit;
         end;
      end;
      1 : begin
         newForm := frmDigitaPlan.GetFormByName(newWb.document as IHTMLDocument2, 'financeiroHonorariosForm');
         Sleep(1000);
         ClicaBotaoGcpj(newForm, 'btoConfirmar');
         Sleep(1000);
         wbpopUp.Close;
         segue := true;
      end;
   end;
end;

function TfrmDigitaPlan.DigitaNotaNova(winTask: TWintask; dataDigitar: string) : integer;
begin
   wbPopWin := 0;

   if not ehPresenta then
   begin
      result := PosicionaNotaNova(wintask);
      if result < 0 then
         exit;
   end;

   wbPopWin := 1;
   result := DigitaAtosDaNota(wintask);
end;

function TfrmDigitaPlan.PosicionaNotaNova(wintask: TWintask): integer;
var
   found : boolean;
   index : integer;
   inputElement : HTMLInputElement;
   newForm : IHTMLFormElement;
   doc : IHTMLDocument2;
   valor : string;
   selectElement : HTMLSelectElement;
   popup : string;
   newDoc : IHTMLDocument2;
   wbScript : HTMLScriptElement;
   ppForm : IHTMLFormElement;
   ret : integer;
begin
   result := -1;

   doc := wb.document as IHTMLDocument2;
   newForm := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
   if Not Assigned(newForm) then
   begin
      ShowMessage('Erro conectando ao GCPJ');
      exit;
   end;

   if ClicaBotaoGcpj(newForm, 'btoIncluir') < 0 then
   begin
      ShowMessage('Erro conectando ao GCPJ');
      exit;
   end;

   //digita os dados do CNPJ
   doc := wb.document as IHTMLDocument2;
   newForm := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
   if not Assigned(newForm) then
   begin
      ShowMessage('N�o est� na p�gina correta');
      exit;
   end;

      //preenche o formulario
   ret := SetSelectField(newform, 'cdTipoDocumentoPessoa', 'CNPJ', 'rotearTipoDocumento()');
   if ret < 0 then
   begin
      ShowMessage('N�o est� no formul�rio correto');
      exit;
   end;
   Sleep(5000);

   while wb.Busy = true do
   begin
      Application.ProcessMessages;
   end;

   //preenche o CNPJ
   index := 0;
   inputElement := frmDigitaPlan.GetInputField(newform, 'despesaHonorariosProcessoVO.pessoaExternaBradescoVO.cdCnpjPessoaExterna', 0);
   while Not Assigned(inputElement) do
   begin
      Sleep(1000);
      while (wb.ReadyState = 0) or (wb.ReadyState = 1) or (wb.ReadyState = 3) do
        Application.ProcessMessages;

      while wb.Busy = true do
        Application.ProcessMessages;

      Sleep(300);
      Application.ProcessMessages;
      inc(index);
      if index > 120 then
      begin
         ShowMessage('N�o est� na p�gina correta(Campo)');
         exit;
      end;

      newform := frmDigitaPlan.GetFormByName(wb.Document as IHTMLDocument2, 'financeiroHonorariosForm');
      while not Assigned(newform) do
      begin
         Sleep(1000);
         while (wb.ReadyState = 0) or (wb.ReadyState = 1) or (wb.ReadyState = 3) do
            Application.ProcessMessages;

         while wb.Busy = true do
            Application.ProcessMessages;
         newform := frmDigitaPlan.GetFormByName(wb.Document as IHTMLDocument2, 'financeiroHonorariosForm');
      end;
      inputElement := frmDigitaPlan.GetInputField(newform, 'despesaHonorariosProcessoVO.pessoaExternaBradescoVO.cdCnpjPessoaExterna', 0);
   end;

   inputElement.value := Copy(planJbm.cnpjEscritorio, 1, 8);
   inputElement := frmDigitaPlan.GetInputField(newform, 'despesaHonorariosProcessoVO.pessoaExternaBradescoVO.cdFilialCnpjPessoaExterna', 0);
   if Not Assigned(inputElement) then
   begin
      ShowMessage('N�o est� no formul�rio correto(2)');
      exit;
   end;

   inputElement.value := Copy(planJbm.cnpjEscritorio, 9, 4);
   inputElement := frmDigitaPlan.GetInputField(newform, 'despesaHonorariosProcessoVO.pessoaExternaBradescoVO.cdControleCnpjPessoaExterna', 0);
   if Not Assigned(inputElement) then
   begin
      ShowMessage('N�o est� no formul�rio correto(3)');
      exit;
   end;

   inputElement.value := Copy(planJbm.cnpjEscritorio, 13, 2);
   Sleep(500);

   inputElement := frmDigitaPlan.GetInputField(Newform, 'btoLupa', 0);
   while Not Assigned(inputElement) do
   begin
      newform := frmDigitaPlan.GetFormByName(wb.Document as IHTMLDocument2, 'financeiroHonorariosForm');
      if not Assigned(newform) then
      begin
         ShowMessage('N�o est� na p�gina correta');
         exit;
      end;
      Sleep(3000);
      Application.ProcessMessages;
      inputElement := frmDigitaPlan.GetInputField(newform, 'btoLupa', 0);
   end;


   segue := false;
   inputElement.click;
   Sleep(2000);

   EsperaPaginaProcessar;

   if Assigned(wbPopUp) then
   begin
      wbpopUp.Close;
      FreeAndNil(wbpopUp);
   end;

   if Self.WindowState <> wsMaximized then
      self.WindowState := wsMaximized;

   selectElement := GetSelectField(newform, 'despesaHonorariosProcessoVO.tipoDocumentoHonorarioVO.cdTipoDocumentoHonorarios', 0);
   if Not Assigned(selectElement) then
   begin
      ShowMessage('N�o est� no formul�rio correto');
      exit;
   end;

   segue := false;

   //seleciona NOTA FISCAL
   selectElement.selectedIndex := 6;
   selectElement.blur;
   Sleep(500);

   newForm := GetFormByName(doc, 'financeiroHonorariosForm');
   inputElement := GetInputField(newform, 'despesaHonorariosProcessoVO.cdNumeroDocumentoHonorarios', 0);
   inputElement.value := planJbm.dtsAtos.FieldByName('numeronota').AsString;

   inputElement := GetInputField(newform, 'despesaHonorariosProcessoVO.cdSerieDocumentoHonorarios', 0);
   inputElement.value := 'A';

   inputElement := GetInputField(newform, 'despesaHonorariosProcessoVO.dtDocumentoDespesaHonorarios', 0);
   inputElement.value := FormatDateTime('dd/mm/yyyy', date);

   inputElement := GetInputField(newform, 'despesaHonorariosProcessoVO.vrDocumentoDespesaHonorarios', 0);
   inputElement.value := FormatFloat('0.00', planJbm.valorTotalNota);

   inputElement := GetInputField(newform, 'despesaHonorariosProcessoVO.empresaBradescoVO.cdEmpresaIncorporada', 0);
   inputElement.value := '237';

   if planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger <> 237 then
   begin
      inputElement := GetInputField(newform, 'despesaHonorariosProcessoVO.dependenciaBradescoOnLineVO.cdDependencia', 0);
      inputElement.value := planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString;
   end;
   result := 1;
end;


function TfrmDigitaPlan.DigitaAtosDaNota(wintask: TWintask): integer;
var
   doc : IHTMLDocument2;
   form : IHTMLFormElement;
   inputElement : HTMLInputElement;
   valor : string;
   andamentoDigitar : string;
   totalDigitadoTela, totalDigitadoDb : double;
   HTMLWindow : IHTMLWindow2;
   tries : integer;

begin
   result := -1;

   while not planJbm.dtsAtos.Eof do
   begin
      planJbm.CarregaCamposDaTabela;
      re.lines.add('Digitando processo: ' + planJbm.dtsAtos.FieldByName('gcpj').AsString);
      Application.ProcessMessages;

      if not ehPresenta then
      begin
         doc := wb.document as IHTMLDocument2;
         Form := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
         if Not Assigned(Form) then
         begin
            ShowMessage('Erro conectando ao GCPJ');
            exit;
         end;
      end;

      totalDigitadoTela := NovoObtemTotalDigitado;
      TotalDigitadoDb := planJbm.SomaTotalDigitadoDaNota; //banco e dados

      if (FormatFloat('0.00', totaldigitadoTela) <> FormatFloat('0.00', TotalDigitadoDB)) then
      begin
         tries := 0;
         while (FormatFloat('0.00', totaldigitadoTela) <> FormatFloat('0.00', TotalDigitadoDB)) do
         begin
            Sleep(1000);
            Application.ProcessMessages;

            Inc(tries);
            if tries > 20 then
               break;

            doc := wb.document as IHTMLDocument2;
            Form := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
            if Not Assigned(Form) then
            begin
               ShowMessage('Erro conectando ao GCPJ');
               exit;
            end;

            totalDigitadoTela := NovoObtemTotalDigitado;
            TotalDigitadoDb := planJbm.SomaTotalDigitadoDaNota; //banco e dados
         end;

         if (FormatFloat('0.00', totaldigitadoTela) <> FormatFloat('0.00', TotalDigitadoDB)) then
         begin
            ShowMessage('Valores digitados n�o conferem(3)');
            exit;
         end;
      end;

      planjbm.tipoProcesso := planJbm.dtsAtos.FieldByName('fgtipoprocesso').asString;

      inputElement := GetInputField(form, 'despesaHonorariosProcessoVO.composicaoDespHonorariosVO.cdNumeroProcessoBradesco', 0);
      inputElement.value := planJbm.dtsAtos.FieldByName('gcpj').AsString;

      if ClicaBotaoGcpj(form, 'btoPesquisar') < 0 then
      begin
         ShowMessage('Erro pesquisando GCPJ');
         exit;
      end;

      if NovoObtemDependenciaDigitar < 0 then
      begin
         ShowMessage('Erro obtendo a depend�ncia para digitar');
         exit;
      end;

      if FileExists(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT') then
      begin
         result := 1;
         RenameFile(ExtractFilePath(Application.ExeName) + 'ENCERRAR.TXT', ExtractFilePath(Application.ExeName) + 'ENCERRAR.OK');
         exit;
      end;

      andamentoDigitar := ObtemTipoAndamento(planJbm.dtsAtos.FieldByName('tipoandamento').AsString, planJbm.tipoProcesso, planJbm.dtsAtos.FieldByName('motivobaixa').AsString);
      if andamentoDigitar = '' then
      begin
//         planjbm.NotificaErro('Falta tipo ato');
         ShowMessage('Falta tipo ato');
         exit;
      end;

      //atualiza o formul�rio
      doc := wb.document as IHTMLDocument2;
      Form := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
      if Not Assigned(Form) then
      begin
         ShowMessage('Erro conectando ao GCPJ');
         exit;
      end;

      if SetSelectField(Form, 'despesaHonorariosProcessoVO.composicaoDespHonorariosVO.cdIdentificadorAtoDespesa', andamentoDigitar, '') < 0 then
      begin
//         planjbm.NotificaErro('Tipo de ato n�o encontrado no GCPJ');
         ShowMessage('Tipo de ato n�o encontrado no GCPJ');
         exit;
      end;

      //atualiza o formul�rio
      doc := wb.document as IHTMLDocument2;
      Form := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
      if Not Assigned(Form) then
      begin
         ShowMessage('Erro conectando ao GCPJ');
         exit;
      end;

      //data do ato
      inputElement := GetInputField(form, 'despesaHonorariosProcessoVO.composicaoDespHonorariosVO.dtAtoDespesaHonorarios', 0);
      inputElement.value := FormatDateTime('dd/mm/yyyy', planJbm.dtsAtos.FieldByName('datadoato').AsDatetime);

      //valor do ato
      if (planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat = 0.00) and
         (planJbm.dtsAtos.FieldByName('tipoandamento').AsString = 'HONORARIOS INICIAIS') then
      begin
         if FormatDateTime('yyyy/mm/dd', planjbm.dtsatos.FieldByName('datadoato').AsDateTime) >= '2012/05/01' then
            valor := '52,00'
         else
            valor := '50,00';
      end
      else
         valor := FormatFloat('0.00', planJbm.dtsAtos.FieldByName('valorcorrigido').AsFloat);

      inputElement := GetInputField(form, 'despesaHonorariosProcessoVO.composicaoDespHonorariosVO.vrAtoDespesaHonorarios', 0);
      inputElement.value := valor;

      //clica no bot�o incluir
      while true do
      begin
         doc := wb.document as IHTMLDocument2;
         Form := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
         if Not Assigned(Form) then
         begin
            ShowMessage('Erro conectando ao GCPJ');
            exit;
         end;

         inputElement := GetInputField(form, 'despesaHonorariosProcessoVO.composicaoDespHonorariosVO.dependenciaBradescoOnLineVO.cdDestinatario', 0);
         if inputElement.innerText = '' then
         begin
            inputElement.Value := planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsString;
            inputElement := GetInputFieldById(form, 'null', 'btoDestinatarioCod', 0);
            inputElement.click;
         end;

         EsperaPaginaProcessar;

         Sleep(5000);

         if Assigned(wbPopUp) then
         begin
            wbpopUp.Close;
            FreeAndNil(wbpopUp);
         end;

         segue := false;
         if ClicaBotaoGcpj(Form, 'btoIncluir') < 0 then
         begin
            ShowMessage('Erro incluindo ato na nota');
            exit;
         end;

         //marca como digitado
         while Assigned(wbPopUp) do
         begin
            if segue then
               break;
            Application.ProcessMessages;
            Sleep(300);
            continue;
         end;

         if Self.WindowState <> wsMaximized then
            self.WindowState := wsMaximized;
         break;
      end;

      planJbm.MarcaProcessoDigitado;
      planJbm.dtsAtos.Next;
      segue := false;
//      EsperaPaginaProcessar;
      Sleep(500);
   end;

   Sleep(2500);

   doc := wb.document as IHTMLDocument2;
   Form := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
   if Not Assigned(Form) then
   begin
      ShowMessage('Erro conectando ao GCPJ');
      exit;
   end;

   //finaliza nota
   totalDigitadoTela := NovoObtemTotalDigitado;
   totalDigitadoDb := planJbm.SomaTotalDigitadoDaNota;

   if (FormatFloat('0.00', totaldigitadoTela) <> FormatFloat('0.00', TotalDigitadoDB)) then
   begin
      ShowMessage('Valores digitados n�o conferem(4): [' + FormatFloat('0.00', totaldigitadoTela) + '] comparando com [' + FormatFloat('0.00', TotalDigitadoDB) + ']');
      exit;
   end;

   if (FormatFloat('0.00', planJbm.valorTotalNota) = FormatFloat('0.00', TotalDigitadoDB)) then
   begin
      result := FinalizaNotaNoGcpj;
      if result = 1 then
         planJbm.MarcaDocumentoFinalizado;
   end;
   result := 1;
end;

function TfrmDigitaPlan.NovoObtemDependenciaDigitar : integer;
var
   idx,p : integer;
   arquivo : TStringlist;
   estagio : integer;
   codigo, nome : string;
   valstr : string;
   dependencia : string;
   primeira : boolean;
   sequencia : integer;
   myString : TstringList;
   selectElement : HTMLSelectElement;
   fromForm : IHTMLFormElement;
   fromField : IHTMLElement;
   options : HTMLOptionElement;

begin
   result := -1;

   codigo := '';
   nome:= '';
   planJbm.ObtemReclamadasDoProcesso('');
   sequencia := 1;

   while Not planJbm.dtsReclamadas.Eof do
   begin
      if planJbm.dtsReclamadas.FieldByName('sequenciagcpj').AsInteger = sequencia then //guarda a primeira por seguranca
      begin
         if planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger <> 0 then
         begin
            codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
            nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
         end
         else
         begin
            Inc(sequencia);
            planJbm.dtsReclamadas.Next;
            continue;
         end;
      end;

      //� bradesco?
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
         if (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = 4001) and (planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger = 5172) then
         begin
            codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
            nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
            break;
         end
         else
         if (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('empresaligadaagrupar').AsInteger) or
            (planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = planJbm.dtsAtos.FieldByName('codempresaligada').AsInteger) then
         begin
            if planJbm.dtsReclamadas.FieldByName('tiporeclamada').AsString <> 'A' then //n�o pode ser agencia
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

   if Not primeira then
   begin
      if StrToInt(codigo) = 5404 then
      begin
         //tem a depend�ncia 4027?
         planJbm.dtsReclamadas.First;

         while Not planJbm.dtsReclamadas.Eof do
         begin
            Application.ProcessMessages;
            if planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsInteger = 4027 then
            begin
               codigo := planJbm.dtsReclamadas.FieldByName('codigoreclamada').AsString;
               nome := planJbm.dtsReclamadas.FieldByName('nomereclamada').AsString;
               break;
            end;
            planJbm.dtsReclamadas.Next;
         end;
      end;
   end;

   fromForm := frmDigitaPlan.GetFormByName(wb.document as IHTMLDocument2, 'financeiroHonorariosForm');
   if Not Assigned(fromForm) then
      exit;

   if primeira then
   begin
      if SetSelectField(fromForm,'selDependencia', '', '') < 0 then
         exit;
      result := 1;
      exit;
   end;
   if SetSelectField(fromForm, 'selDependencia', nome, '') < 0 then
      SetSelectField(fromForm,'selDependencia', '', '');
   result := 1;
end;

function TfrmDigitaPlan.SetSelectField(fromForm: IHTMLFormElement;
  const selectName, elementName: string; executar: string): integer;
var
   selectElement : HTMLSelectElement;
   idx : integer;
   HTMLWindow : IHTMLWindow2;

begin
   result := -1;

   selectElement := GetSelectField(fromForm, selectName, 0);
   if Not Assigned(selectElement) then
      exit;

   segue := false;
   if elementName = '' then //sesleciona o primeiro
   begin
      selectElement.selectedIndex := 1;
      selectElement.focus;
      if executar <> '' then
      begin
         HTMLWindow := (wb.document as IHTMLDocument2).parentWindow;
         HTMLWindow.execScript(executar, '');
         Sleep(3000);
      end;
      result := 1;
      exit;
   end;

   //localiza o elemento na lista
   for idx := 0 to selectElement.length - 1 do
   begin
      if UpperCase((selectElement.Item(idx, idx) as IHTMLOptionElement).Text) = UpperCase(elementName) then
      begin
         selectElement.selectedIndex := idx;
         selectElement.focus;
         if executar <> '' then
         begin
            HTMLWindow := (wb.document as IHTMLDocument2).parentWindow;
            HTMLWindow.execScript(executar, '');
         end;
         result := idx;
         exit;
      end;
   end;
end;

function TfrmDigitaPlan.ClicaBotaoGcpj(var fromForm: IHTMLFormElement;
  const buttonName: string; const janelaFechar: string=''): integer;
var
   found : boolean;
   index : integer;
   inputElement : HTMLInputElement;
   newForm : IHTMLFormElement;
   doc : IHTMLDocument2;

begin
   result := -1;
   found := false;
   index := 0;
   while not found do
   begin
      inputElement := GetInputField(fromForm, buttonName, index);
      if Not Assigned(inputElement) then
      begin
         doc := wb.document as IHTMLDocument2;
         newForm := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
         if Not Assigned(newForm) then
            exit;
         inputElement := GetInputField(newForm, buttonName, index);
      end;
      if Not Assigned(inputElement) then
         exit;
      segue := false;
      inputElement.click;
      Sleep(1000);
      EsperaPaginaProcessar(janelaFechar);
      break;
   end;
   result := 1;
end;

function TfrmDigitaPlan.NovoObtemTotalDigitado: double;
var
   myString : TStringList;
   i, p, estagio : integer;
   valor : string;

begin
   myString := TstringList.Create;
   try
      Application.ProcessMessages;
      Sleep(500);
      myString.Text := (wb.document as IHTMLDocument2).body.outerHTML;

      estagio := 0;
      valor := '';

      for i := 0 to myString.Count - 1 do
      begin
         if estagio = 0 then
         begin
            p := Pos('<B>Total: </B>', myString.Strings[i]);
            if p = 0 then
               continue;
            Inc(estagio);
         end;

         valor := Trim(Copy(myString.Strings[i], p+14, Length(mystring.Strings[i])));
         p := Pos('</TD>', UpperCase(valor));
         valor := Trim(Copy(valor, 1, p-1));

         RemoveEsteCaracter('.', valor);
         RemoveEsteCaracter(' ', valor);

         if valor = '' then
            result := 0
         else
            result := StrToFloat(valor);
         exit;
      end;
   finally
      myString.Free;
   end;
end;

function TfrmDigitaPlan.GetInputFieldById(fromForm: IHTMLFormElement;
  const inputName, id: string; const instance: integer): HTMLInputElement;
var
  field: IHTMLElement;
  idx : integer;
begin
   idx := instance;
  field := fromForm.Item(inputName,idx) as IHTMLElement;
  if Not Assigned(field) then
  begin
     result := nil;
     exit;
  end;

  while Assigned(field) do
  begin
    if UpperCase(field.tagName) = 'INPUT' then
    begin
       if UpperCase(field.id) = UpperCase(id) then
       begin
          result:= field as HTMLInputElement;
         exit;
       end;
    end;
    inc(idx);
    field := fromForm.Item(inputName,idx) as IHTMLElement;
  end;
  result := nil;
end;

function TfrmDigitaPlan.DigitaNotaEmAndamento(winTask: TWintask;
  dataDigitar: string): integer;
begin
   result := NovoPosicionaNotaDigitando;
   if (result <> 1) then
   begin
      ShowMessage('Erro localizando a nota');
      exit;
   end;
   wbPopWin := 1;
   Result := DigitaAtosDaNota(wintask);
end;


function TfrmDigitaPlan.NovoPosicionaNotaDigitando : integer;
var
   doc : IHTMLDocument2;
   theForm : IHTMLFormElement;
   inputElement : HTMLInputElement;
   found : boolean;
   index, i, j : integer;
   valor : string;
   myString : TStringList;
   estagio : integer;
   ovTable : variant;
   pageObj : variant;
   proximaPagina : string;
   HTMLWindow : IHTMLWindow2;
   pesquisou : boolean;
   primeiraLinha : string;
begin
   result := -1;
   proximaPagina := '0';
   primeiraLinha := '';

   while true do
   begin
      myString := TStringList.Create;
      try
         ovTable := wb.OleObject.Document.All.Tags('TABLE').item(0);
         for i := 0 to (ovTable.Rows.Length - 1) do
         begin
            for j := 0 to (ovTable.Rows.Item(i).Cells.Length - 1) do
            begin
               myString.text := ovtable.Rows.Item(i).Cells.Item(j).InnerText;
            end;
         end;

         index := 0;
         estagio := 0;
         pesquisou := false;

         for i := 0 to mystring.Count - 1 do
         begin
            if Trim(myString.Strings[i]) = '' then
               continue;

            if estagio = 0 then
            begin
               if (Pos('Empresa', myString.Strings[i]) <> 0) and (Pos('Prestador', myString.Strings[i]) <> 0) then
               begin
                  if primeiraLinha = myString.Strings[i+1] then
                     exit;
                  primeiraLinha := myString.Strings[i+1];
                  Inc(estagio);
               end;
               continue;
            end;

            pesquisou := true;


            if Pos(DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString, myString.Strings[i]) = 0 then
            begin
               Inc(index);
               continue;
            end;

            doc := wb.document as IHTMLDocument2;
            theForm := GetFormByName(doc, 'financeiroHonorariosForm');
            if not Assigned(theForm) then
               exit;

            inputElement := GetInputField(theForm, 'selecionado', index);
            if Not Assigned(inputElement) then
               exit;

            inputElement.checked := true;

            if ClicaBotaoGcpj(theForm, 'btoAlterar') < 0 then
               exit;
            result := 1;
            exit;
         end;

         if not pesquisou then
         begin
            ShowMessage('Nota ' + DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString + ' n�o localizada nas telas do GCPJ');
            exit;
         end;

      finally
         myString.Free;
      end;

      //se chegou aqui � pq n�o achou ainda
      valor := (wb.document as IHTMLDocument2).body.outerHTML;
      if Pos('p�gina posterior', valor) = 0 then
      begin
         ShowMessage('Erro conectando ao GCPJ, n�o localizou a nota');
         exit;
      end;

      //clica na imagem
      found := false;
      pageObj := wb.OleObject.Document.Images;
      for i := 0 to pageObj.Length - 1 do
      begin
         if pos('setaTabelaDir1.gif',pageObj.Item(i).src) <> 0 then
         begin
            doc := wb.document as IHTMLDocument2;
            theForm := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
            if Not Assigned(theForm) then
            begin
               ShowMessage('Erro conectando ao GCPJ');
               exit;
            end;

            found := true;
            inputElement := GetInputField(theForm, 'pagina', 0);
            proximaPagina := inputElement.value;
            valor := 'javascript:document.forms[0].elements[''pagina''].value=' + IntToStr(StrToInt(inputElement.value)+1) + ';paginarFinanceiro(''paginarListaHonorariosEmCadastramento'')';

            segue := false;
            HTMLWindow := doc.parentWindow;
            HTMLWindow.execScript(valor, 'JavaScript');
            EsperaPaginaProcessar;
            break;
         end;
      end;

      if not found then
         exit;
   end;
end;

procedure TfrmDigitaPlan.BitBtn8Click(Sender: TObject);
begin
   NovaRotinaDigitacaoNotas('', true);
end;

procedure TfrmDigitaPlan.FechaJanelaGcpj(janelaFechar: string);
var
   valor : string;
   pWintask : TWintask;
   ini : tinifile;
   diretorio : string;
begin
   ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
   try
      diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_programa\Presenta');
   finally
      ini.Free;
   end;

   pwintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                               diretorio + '\rob\',
                              '');

   try
      pwintask.ie8 := ie8;
      pwintask.ff := ff;

      pwintask.FecharTodos('TaskExec.exe');
      pwintask.FecharTodos('TaskLock.exe');
      pwintask.FecharTodos('Tasksync.exe');

      if pwintask.Verifica_Janela_PopUp(janelaFechar, valor) = 1 then
      begin
         pwintask.Run_Any_File_Parametro('fc_click_bto_ok.rob', UpperCase(ExtractFileName(Application.Exename)), '', '', '', '');
         Sleep(100);
      end;
   finally
      pwintask.free;
   end;
end;

procedure TfrmDigitaPlan.Timer1Timer(Sender: TObject);
var
   valor : string;
   pWintask : TWintask;
   ini : Tinifile;
   diretorio : string;

begin
   if timer1.tag = 1 then
      exit;

   ini := TiniFile.Create('C:\presenta\basediaria\config.ini');
   try
      diretorio := ini.ReadString('honorarios', 'diretorio', '\\mz-vv-fs-083\d4040_2\Publico\BackUp_Sistemas\Honorarios_programa\Presenta');
   finally
      ini.Free;
   end;

   timer1.tag := 1;
   pwintask := TWintask.Create('c:\Arquivos de programas\WinTask\bin\TaskExec.exe',
                               diretorio + '\rob\',
                              '');

   try
      pwintask.ie8 := ie8;
      pwintask.ff := ff;

      pwintask.FecharTodos('TaskExec.exe');
      pwintask.FecharTodos('TaskLock.exe');
      pwintask.FecharTodos('Tasksync.exe');

      if pwintask.Verifica_Janela_PopUp(pjanelaFechar, valor) = 1 then
      begin
         pwintask.Run_Any_File_Parametro('fc_click_bto_ok.rob', UpperCase(ExtractFileName(Application.Exename)), '', '', '', '');
         Sleep(100);
      end;
   finally
      pwintask.free;
   end;
   timer1.Tag := 0;
end;

function TfrmDigitaPlan.FinalizaNotaNoGcpj: integer;
var
   doc : IHTMLDocument2;
   form : IHTMLFormElement;
begin
   result := -1;
   doc := wb.document as IHTMLDocument2;
   Form := frmDigitaPlan.GetFormByName(doc, 'financeiroHonorariosForm');
   if Not Assigned(Form) then
   begin
      ShowMessage('Erro conectando ao GCPJ');
      exit;
   end;

   Timer1.Enabled := true;
   try
      pJanelaFechar := UpperCase(ExtractFileName(Application.Exename)) + '|#32770|Mensagem da p�gina da web';

      if ClicaBotaoGcpj(form, 'btoSalvar', UpperCase(ExtractFileName(Application.Exename)) + '|Static|Opera��o realizada com sucesso') < 0 then
      begin
         ShowMessage('Erro finalizando a nota');
         exit;
      end;
   finally
      Timer1.Enabled := false;
   end;
   result := 1;
end;

procedure TfrmDigitaPlan.Marcarnotacomofinalizada1Click(Sender: TObject);
begin
   frmRetornarFinalizada := TfrmRetornarFinalizada.Create(nil);
   try
      if DBGrid2.DataSource.DataSet.RecordCount > 0 then
         frmRetornarFinalizada.numNota.Text := DBGrid2.DataSource.DataSet.FieldByName('numeronota').AsString;
      frmRetornarFinalizada.numNota.tag := 1;
      frmRetornarFinalizada.Caption := 'Marcar nota finalizada';
      if frmRetornarFinalizada.ShowModal = mrOk then
      begin
         planJbm.dtsPlanilha.Close;
         planJbm.dtsPlanilha.Open;
      end;
   finally
      frmRetornarFinalizada.Free;
   end;
end;

procedure TfrmDigitaPlan.SetehPresenta(const Value: boolean);
begin
  FehPresenta := Value;
end;

end.

