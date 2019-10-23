unit uReadBCCorreio;

interface

uses
  Windows, Classes, ADODB, DB, StdCtrls, Variants, SysUtils, Inifiles, Forms,
  Controls, ComCtrls,
  GLOBAL, global_autojud_2, uPresentaFunc, DataModuleCorreio, idsmcons32,
  Criptografa_bacenjud_2, uLog, uException, uIconectorLog, mgenlib, uBCCorreio,
  SHDocVw, MSHTML, StrUtils, uSmtpService_ci, ComObj,
  // Alterado 07/05/2012
  uIAcessorPesLegado, uAcessorPesLegado, uIConectorPesLegado,

  MSWBase, MSWord, uPstac10Componente, uSafraNotes, uenviaEmailNet;

type
	TEmailAddressInfo = class(TObject)
		emailAddress : string;
		userCode : string;
		notifica : integer;
	end;

	TUserName = Class(TObject)
		userName : string;
		notifica : integer;
	end;

  TReadBCCorreio = class(TObject)
    private
      FmemMsg: TRichEdit;
      FADOConn: TADOConnection;
      FHora_Logon,
      FHora_Logout: String;
      FdbDir: String;
      FrunOnWeekends: boolean;
      FIsDebug: boolean;
      FSalvarArquivo: boolean;
      FdirArqDebug: String;
      FDirExePSTAW10: String;

      sendMail: boolean;
      checkMsgInst: boolean;
      depLer: string;
      internetLogging: boolean;
      dispatchMsg: boolean;
      vim830Global: TVim830Global;
      informouDesativacao: Boolean;
      sisbacenPublico : boolean;
      mailInfo : TMailInfo;
      useSSL : boolean;
      useAuth : boolean;

      BCCorreio: TBCCorreio;

      // Alterado 26/12/2011 - Safra
      EhSafraNotes: Boolean;
      EhTamLinha74: Boolean;

      // Alterado 11/01/2012
      FProxy: String;
      FUserProxy: String;
      FSenhaProxy: String;

      // Alterado 07/05/2012
      LerWebPost : Boolean;
      ADOConnWebPost : TADOConnection;
      AcessorWebPost: IAcessorPesLegado;
      ConectorWebPost: IConectorPesLegado;
      fMainQryWP : TADOQuery;
      fSegQryWP : TADOQuery;

      procedure ExibeMensagem(Linha: String);
      procedure LoadVim830Global;
      procedure ConnectToEmail;
      procedure DisconnectEmail;
      function IsTimeToLogout: boolean;
      procedure RemoveOldMessages;
      function VaiTerQueRodar: boolean;
      function TemQueTrocarSenha: boolean;
      procedure GetUnitsInterval(cod_inst, cod_dep, transacao: String; var lastTime: TdateTime; var minutes: TDateTime; var dias: integer);
      function ThereAreNewMessages : boolean;
      procedure DoShowProxProc;
      function ExecutaEntregaDeMensagens: integer;
      function EntregaEstaMensagem(qryMsgs: TADODataSet; ext: string; ultima: boolean) : integer;
      function MessageAppliesUserFilter(qryMsgs: TADODataSet; userCode: string; fileExt: string; msgText: TStringList): boolean;
      procedure SaveMessageSent(qryMsgs: TADODataSet);
      function JaEnviou(codUsu, codMsg: string) : boolean;
      procedure MarcaEnviada(codUsu, codMsg: string);
      procedure NotifyReadMessage(codmsg: integer; msgNumber: string; codUsu: integer; assunto: string);
      function ReadMessagesFromBCCorreio: integer;
      function CheckMessageList(var trocousenha: boolean) : integer;
      function TrocaSenhaDoBCCorreio: Integer;
      function ReadMailList: integer;
      function CheckMessageQueue(msgNumber: string): integer;
      function LogonToBCCorreio: Integer;
      function VerificaFiltro(sender, subject: string): boolean;
      procedure SaveLogonInformation(tipo: integer; strUp: string);
      procedure InsertMessageHeader(posicao: integer);
      procedure AvisaQueMensagemChegou(numero, assunto, dep: string);
      function PrioridadeZero(sender, subject: string): boolean;
      function ReadMailText(tipoMsg, msgNumber, subject, transacao, sender, status_msg: string; Cod_Msg: Integer): integer;
      function UpdateMessageHeader(msgNumber, status: string; Msg: TMensagem): integer;
      procedure EhOficio(subject: string; sender: string; var EhOficio: boolean);
      function UpdateMessageReadStatus(msgNumber: string; ehOficio: boolean; tipoMsg:string): integer;

      // Alterado 07/05/2012
      procedure AtualizaDadosParaAutojud;
      function  IsThisMessageInFile(numero: string; assunto: string): boolean;
      function  SaveMessage(codmsg: integer;numero: string; assunto: string; arquivo: string): integer;
    public
      Constructor Create(Memo: TRichEdit;
                         ADOConn: TADOConnection;
                         Hora_Logon,
                         Hora_Logout,
                         dbDir: String;
                         runOnWeekends: boolean;
                         mailSystem: TMailInfo);
      Procedure Execute;
  end;

implementation

uses fMain;

{ TReadBCCorreio }


procedure TReadBCCorreio.AtualizaDadosParaAutojud;
var
  dateFrom : string;
  mName : string;
  ini : TiniFile;
  strValue : TStringList;
  str : string;
  i : integer;
  separator : string;
  nret : integer;
  assunto : string;
  entregue : boolean;
  searchRec : TSearchRec;
  buffer : array[0..255] of char;
  salvou : boolean;
  uniqueDate : string;

  oficio : Boolean;

  sSQL : String;

  x : integer;
  // rm 29/06/09 - Loop com parametro
  LoopBreak: integer;
begin
  try
    ExibeMensagem('Verificando se existe mensagens no WebPost para atualizar no Autojud');

    ini := TIniFile.Create(FdbDir + 'VIM830.INI');
    str := ini.ReadString('AUTOJUD', 'Sender', '');
    if str = '' then
    begin
      ini.free;
      exit;
    end;

    strValue := TStringList.create;
    StrToList(str, ',', strValue);

    ini := TIniFile.Create(FdbDir + 'VIM830.INI');
    uniqueDate := ini.ReadString('AUTOJUD', 'OnlyDate', '');
    dateFrom := ini.ReadString('AUTOJUD', 'DateFrom', FormatDateTime('dd/mm/yyyy', date));
    // rm 29/06/09 - Loop com parametro
    LoopBreak := ini.ReadInteger('AUTOJUD', 'LoopBreak', 20);
    ini.free;

    with fMainQryWP do
    begin
      mName := tabWebVM_MESSAGES;

      sSQL := 'SELECT * FROM ' + mName;
      if IsParam('/semof') then
        sSQL := sSQL + ' WHERE EH_OFICIO = ''S'' And ENTREGUE = ''S'' And ((GRAVOU_OFICIO IS NULL) OR (GRAVOU_OFICIO = ''N'')) AND '
      else
        sSQL := sSQL + ' WHERE LIDO = ''S'' AND ((GRAVOU_OFICIO IS NULL) OR (GRAVOU_OFICIO = ''N'')) AND ';
      if uniqueDate <> '' then
        sSQL := sSQL + ' (AMD >= ''' + Copy(uniquedate, 7, 4) + '/' + Copy(uniquedate, 4, 2) + '/' + Copy(uniquedate, 1, 2) + ' 00:00'  + ''' AND' +
                        ' AMD <= ''' + Copy(uniquedate, 7, 4) + '/' + Copy(uniquedate, 4, 2) + '/' + Copy(uniquedate, 1, 2) + ' 23:59'  + ''')'
      else
        sSQL := sSQL + ' (AMD >= ''' + Copy(dateFrom, 7, 4) + '/' + Copy(dateFrom, 4, 2) + '/' + Copy(dateFrom, 1, 2) + ' 00:00'  + ''')';
      sSQL := sSQL + ' ORDER BY AMD, NUM_SISBACEN';

      Close;
      Sql.Text := sSQL;
      Open;

      ExibeMensagem('Total de mensagens para enviar do WebPost para Autojud: ' + IntToStr(fMainQryWP.RecordCount));
      Insere_Log(INFORMATIVO, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO, 'Total de mensagens para enviar do WebPost para Autojud: ' + IntToStr(fMainQryWP.RecordCount));
    end;

    x := 0;

    fMainQryWP.tag := 0;

    while Not fMainQryWP.eof do
    begin
      // Alterado 06/03/2009
      //if x >= 10 then begin
      //if x >= 20 then begin
      // rm 29/06/09 - Loop com parametro

      if x >= LoopBreak then begin
        break;
      end;
      Inc(x);

      if pleaseClose then
        break;

      DoEvents;

      { verifica se a mensagem está no filtro selecionado }
      ExibeMensagem('Verificando mensagem: [' + fMainQryWP.FieldByName('NUM_SISBACEN').AsString + '] ' + fMainQryWP.FieldByName('Assunto').AsString);

      for i := 0 to strValue.Count - 1 do
      begin
        if Pos(UpperCase(strValue.strings[i]), Trim(UpperCase(fMainQryWP.FieldByName('cod_dep_remetente').AsString))) <> 0 then
          break;
      end;

      if (i > (strValue.count - 1)) or
         ((Copy(dateFrom, 7, 4) + '/' + Copy(dateFrom, 4,2) + '/' + Copy(dateFrom, 1, 2)) >
         (Copy(fMainQryWP.FieldByName('AMD').AsString, 1, 10))) then
      begin
        if uniqueDate = '' then begin
          { mensagem não interessa marca como entregue }
          with fSegQryWP do
          begin
            sSQL := 'UPDATE ' + mName +
                    ' SET EH_OFICIO=''N'', GRAVOU_OFICIO = ''S''' +
                    ' WHERE COD_MSG = ' + IntToStr(fMainQryWP.FieldByName('COD_MSG').AsInteger);

            Close;
            Sql.Text := sSQL;
            ExecSQL;
          end;
        end;

        fMainQryWP.tag := 0;
        fMainQryWP.next;
        continue;
      end;

      salvou := false;
      assunto := fMainQryWP.FieldByName('Assunto').AsString;

      EhOficio(fMainQryWP.FieldByName('ASSUNTO').AsString,
               fMainQryWP.FieldByName('COD_DEP_REMETENTE').AsString,
               oficio);

      if oficio then
      begin
         ExibeMensagem('É ofício. Verifica se já está na base de dados');

        entregue := IsThisMessageInFile(fMainQryWP.FieldByName('num_sisbacen').AsString, assunto);
        if Not entregue then
        begin
         ExibeMensagem('Vai salvar a mensagem no Autojud ');

          nret := SaveMessage(fMainQryWP.FieldByName('COD_MSG').AsInteger, fMainQryWP.FieldByName('NUM_SISBACEN').AsString, assunto, FdbDir + 'MSGS\TEMP\' + fMainQryWP.FieldByName('NUM_SISBACEN').AsString  + '.TXT');
          salvou := (nret > 0);
        end
        else
          salvou := true;
      end else
        salvou := true;

      if salvou then
      begin
        with fSegQryWP do
        begin
          sSQL := 'UPDATE ' + mName +
                  ' SET GRAVOU_OFICIO = ''S'', EH_OFICIO= ''S''' +
                  ' WHERE COD_MSG = ' + IntToStr(fMainQryWP.FieldByName('COD_MSG').AsInteger);
          Close;
          Sql.Text := sSQL;
          ExecSQL;
        end;
      end;

      fMainQryWP.next;
    end;

    // Alterado 19/06/2012
    fMainQryWP.Close;

    strValue.free;
  except
    on Er:EErro do begin
      Insere_Log(ERRO_ROBOT_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO,
                 'Erro na importação dos dados do WebPost - ' + Er.Mensagem);
      ExibeMensagem('Erro na importação dos dados do WebPost - ' + Er.Mensagem);
    end;
    on Er:Exception do begin
      Insere_Log(ERRO_ROBOT_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO,
                 'Erro na importação dos dados do WebPost - ' + Er.Message);
      ExibeMensagem('Erro na importação dos dados do WebPost - ' + Er.Message);
    end;
  end;
end;

function TReadBCCorreio.IsThisMessageInFile(numero: string; assunto: string): boolean;
var
  codmsg : Integer;
begin
  dmCorreio.dts.Close;
  dmCorreio.dts.CommandText := 'SELECT * FROM ' + tabVM_MESSAGES +
                               ' WHERE NUM_SISBACEN='''+fMainQryWP.FieldByName('NUM_SISBACEN').AsString+'''';
  dmCorreio.dts.Parameters.Clear;
  dmCorreio.dts.Prepared;
  dmCorreio.dts.Open;

  if dmCorreio.dts.RecordCount = 0 then
    result := false
  else
  begin
    codmsg := dmCorreio.dts.FieldByName('COD_MSG').AsInteger;

    dmCorreio.dts.Close;
    dmCorreio.dts.CommandText := 'SELECT * FROM ' + tabVM_MESSAGES +
                                 ' WHERE NUM_SISBACEN='''+fMainQryWP.FieldByName('NUM_SISBACEN').AsString+'''' +
                                 ' AND LIDO = ''S''';
    dmCorreio.dts.Parameters.Clear;
    dmCorreio.dts.Prepared;
    dmCorreio.dts.Open;
    if dmCorreio.dts.RecordCount > 0 then
      result := true
    else
    begin
      dmCorreio.ConectorRobo.Exclui_VM_Messages(codmsg);
      result := false;
    end;
  end;
end;

function TReadBCCorreio.SaveMessage(codmsg: integer;numero: string; assunto: string; arquivo: string): integer;
var
  tname : string;
  sSQL : String;
Begin
  result := -1;

  ExibeMensagem('Enviando mensagem para Autojud: [' + numero + ']-' + assunto);
  Insere_Log(INFORMATIVO, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO, 'Enviando mensagem para Autojud: [' + numero + ']');

  Ssql := 'INSERT INTO ' + tabVM_MESSAGES + '(NUM_SISBACEN, COD_INST,' +
          ' COD_DEP, ASSUNTO, COD_DEP_REMETENTE, DESCR_DEP_REMETENTE,' +
          ' USUARIO_REMETENTE, DATA_ENVIO, HORA_ENVIO, RECEBIDO_POR,' +
          ' DT_RECEBIMENTO, HR_RECEBIMENTO, LIDO, TOTAL_PAGINAS, ';
  if Not fMainQryWP.FieldByName('PAGINAS_LIDAS').IsNull then
      ssql := ssql + ' PAGINAS_LIDAS, ';
  if Not fMainQryWP.FieldByName('ENTREGUE').IsNull then
      ssql := ssql + 'ENTREGUE, ';
  ssql := ssql + 'AMD, EH_OFICIO, STATUS_SISBACEN,' +
          ' TRANSACAO, GRAVOU_OFICIO) VALUES (:0, :1, :2, :3, :4, :5, :6, :7,' +
          ' :8, :9, :10, :11, :12, :13, ';
  if Not fMainQryWP.FieldByName('PAGINAS_LIDAS').IsNull then
     Ssql := sSql + ':14, ';

  if Not fMainQryWP.FieldByName('ENTREGUE').IsNull then
     Ssql := sSql + ':15, ';

  ssql := ssql + ':16, :17, :18, :19, :20)';

  dmCorreio.cmd.Connection := dmCorreio.ADOConnection1;
  dmCorreio.cmd.CommandType := cmdText;
  dmCorreio.cmd.CommandTimeout := Timeout;

  dmCorreio.cmd.CommandText := sSQL;
  dmCorreio.cmd.Parameters.ParamByName('0').Value  := fMainQryWP.FieldByName('NUM_SISBACEN').AsString;
  dmCorreio.cmd.Parameters.ParamByName('1').Value  := fMainQryWP.FieldByName('COD_INST').AsString;
  dmCorreio.cmd.Parameters.ParamByName('2').Value  := fMainQryWP.FieldByName('COD_DEP').AsString;
  // Alterado 28/02/2012
  //dmCorreio.cmd.Parameters.ParamByName('3').Value  := fMainQryWP.FieldByName('ASSUNTO').AsString;
  dmCorreio.cmd.Parameters.ParamByName('3').Value  := Copy(Trim(fMainQryWP.FieldByName('ASSUNTO').AsString), 1, 120);
  dmCorreio.cmd.Parameters.ParamByName('4').Value  := fMainQryWP.FieldByName('COD_DEP_REMETENTE').AsString;
  dmCorreio.cmd.Parameters.ParamByName('5').Value  := fMainQryWP.FieldByName('DESCR_DEP_REMETENTE').AsString;
  dmCorreio.cmd.Parameters.ParamByName('6').Value  := fMainQryWP.FieldByName('USUARIO_REMETENTE').AsString;
  dmCorreio.cmd.Parameters.ParamByName('7').Value  := fMainQryWP.FieldByName('DATA_ENVIO').AsString;
  dmCorreio.cmd.Parameters.ParamByName('8').Value  := fMainQryWP.FieldByName('HORA_ENVIO').AsString;
  dmCorreio.cmd.Parameters.ParamByName('9').Value  := fMainQryWP.FieldByName('RECEBIDO_POR').AsString;
  // Alterado 28/02/2012
  if fMainQryWP.FieldByName('DT_RECEBIMENTO').AsString = '01/01/0001' then
  begin
    dmCorreio.cmd.Parameters.ParamByName('10').Value := FormatDateTime('dd/mm/yyyy', now);
    dmCorreio.cmd.Parameters.ParamByName('11').Value := FormatDateTime('hh:nn', now);
  end
  else
  begin
    dmCorreio.cmd.Parameters.ParamByName('10').Value := fMainQryWP.FieldByName('DT_RECEBIMENTO').AsString;
    dmCorreio.cmd.Parameters.ParamByName('11').Value := fMainQryWP.FieldByName('HR_RECEBIMENTO').AsString;
  end;
  dmCorreio.cmd.Parameters.ParamByName('12').Value := 'N';
  dmCorreio.cmd.Parameters.ParamByName('13').Value := fMainQryWP.FieldByName('TOTAL_PAGINAS').AsInteger;
  if Not fMainQryWP.FieldByName('PAGINAS_LIDAS').IsNull then
     dmCorreio.cmd.Parameters.ParamByName('14').Value := fMainQryWP.FieldByName('PAGINAS_LIDAS').AsInteger;
  if Not fMainQryWP.FieldByName('ENTREGUE').IsNull then
     dmCorreio.cmd.Parameters.ParamByName('15').Value := fMainQryWP.FieldByName('ENTREGUE').AsString;
  dmCorreio.cmd.Parameters.ParamByName('16').Value := fMainQryWP.FieldByName('AMD').AsString;
  dmCorreio.cmd.Parameters.ParamByName('17').Value := 'S';
  dmCorreio.cmd.Parameters.ParamByName('18').Value := fMainQryWP.FieldByName('STATUS_SISBACEN').AsString;
  dmCorreio.cmd.Parameters.ParamByName('19').Value := 'PMSG830';
  dmCorreio.cmd.Parameters.ParamByName('20').Value := 'N';
  dmCorreio.cmd.Prepared := True;
  dmCorreio.cmd.Execute;


  dmCorreio.dts.Close;
  dmCorreio.dts.CommandText := 'SELECT * FROM ' + tabVM_MESSAGES +
                               'WHERE ((NUM_SISBACEN='''+fMainQryWP.FieldByName('NUM_SISBACEN').AsString+'''))';
  dmCorreio.dts.Prepared := true;
  dmCorreio.dts.Open;

  tName := tabWebVM_MESSAGES_TEXTO;
  with fSegQryWP do
  begin
    sSQL := 'SELECT * FROM ' + tName + ' WHERE COD_MSG = ' + IntToStr(fMainQryWP.FieldByName('COD_MSG').AsInteger) +
            ' ORDER BY SEQUENCIA';
    Close;
    SQL.Text := sSQL;
    Open;
    while not eof do begin
      Application.ProcessMessages;

      dmCorreio.ConectorRobo.Inclui_Mensagem_Texto(dmCorreio.dts.FieldByName('COD_MSG').AsInteger,
                                                   FieldByName('SEQUENCIA').AsInteger,
                                                   FieldByName('LINHA').AsString);

      next;
    end;

    // Alterado 19/06/2012
    Close;
  end;

  dmCorreio.cmd.CommandText := 'UPDATE ' + tabVM_MESSAGES +
                               'SET LIDO = ''S'' WHERE COD_MSG = ' + IntToStr(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
  dmCorreio.cmd.Prepared;
  dmCorreio.cmd.Execute;

  result := 1;
end;

procedure TReadBCCorreio.AvisaQueMensagemChegou(numero, assunto, dep: string);
var
  address : string;
  ini : TIniFile;
  svalues : tstringlist;
  i : integer;
  mensagem : string;
  eProg : array[0..2000] of char;
begin
  try
    TraceWFile(dmCorreio.traceFName, 'Entrou na rotina AvisaQueMensagemChegou');

    if FIsDebug then
      ExibeMensagem('Entrou na rotina AvisaQueMensagemChegou');

    ini := TIniFile.Create(FdbDir + 'VIM830.INI');
    address := ini.readString('alertador', 'address', '');
    ini.free;
    if address = '' then
      exit;

    sValues := TStringList.Create;
    StrToList(address, ';', sValues);

    mensagem := 'O sistema de captura de mensagens localizou a mensagem: [' + numero +
                ' - ' + assunto + '(' + dep + ')] e vai iniciar a sua leitura no SISBACEN. Aguarde...';

    for i := 0 to sValues.count - 1 do
    begin
      StrPcopy(eProg, 'NET SEND ' + sValues.strings[i] + ' ' + mensagem);
      WinExec(PAnsiChar(AnsiString(eProg)), SW_HIDE);
    end;
  except
    // Alterado 08/05/2012
    on Er:EErro do
    begin
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na AvisaQueMensagemChegou - ' + Er.Mensagem);
      TraceWFile(dmCorreio.traceFName, 'Erro na AvisaQueMensagemChegou - ' + Er.Mensagem);
    end;

    on Er:Exception do
    begin
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na AvisaQueMensagemChegou - ' + Er.Message);
      TraceWFile(dmCorreio.traceFName, 'Erro na AvisaQueMensagemChegou - ' + Er.Message);
    end;
  end;
end;

function TReadBCCorreio.CheckMessageList(var trocousenha: boolean): integer;
var
  ret : integer;
  fh : integer;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina CheckMessageList');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina CheckMessageList');

  result := -1;
  { Hoje eh dia de trocar a senha? }
  if pleaseClose then
    exit;

  trocouSenha := false;
  if TemQueTrocarSenha then
  begin
    TraceWFile(dmCorreio.traceFName, 'Vai trocar senha');
    if IsFileLocked(FdbDir + 'SENHA.LCK', fh) <= 0 then
      exit;

    if TrocaSenhaDoBCCorreio <= 0 then
      Insere_Log(ERRO_LOGON_SISBACEN, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO,
                 'Erro na troca de senha de: ' + dmCorreio.dtsVM_Units.FieldByName('cod_inst').AsString + '/' +
                 dmCorreio.dtsVM_Units.FieldByName('cod_dep').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('Operador').AsString);
    trocouSenha := true;
    result := 1;
    ReleaseFile(FdbDir + 'senha.LCK', fh);
    exit;
  end;

  if pleaseClose then
    exit;

  ret := ReadMailList;
  result := ret;
end;

function TReadBCCorreio.CheckMessageQueue(msgNumber: string): integer;
var
  nRet : integer;
  msgNum : string;
  lckMsgHandle : integer;
  telasInfo : TTelasInfo;
  contador : integer;
  erro: string;
  localizou: boolean;
  i: integer;
label
 novamente;
begin
  result := -1;
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina CheckMessageQueue');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina CheckMessageQueue');

  contador := 1;

novamente:

  try
    { Captura as msgs não lidas }
    dmCorreio.ConectorRobo.Obtem_Msg_Nao_Lida(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                              dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,

                                              // Alterado 04/06/2012
                                              vim830Global.DataLimiteAte,

                                              dmCorreio.dtsMsg);

    if msgNumber <> '' then
    begin
      if dmCorreio.dtsMsg.Eof then
        result := 0
      else
        result := 1;
      exit;
    end;
  except
    // Alterado 08/05/2012
    on Er: EErro do
    begin
      result := -1;
      TraceWFile(dmCorreio.traceFName, 'Saiu da rotina CheckMessageQueue(erro)');
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Saiu da rotina CheckMessageQueue(erro) - ' + Er.Mensagem);
      TraceWFile(dmCorreio.traceFName, 'Saiu da rotina CheckMessageQueue(erro) - ' + Er.Mensagem);
      exit;
    end;

    on Er: Exception do
    begin
      result := -1;
      TraceWFile(dmCorreio.traceFName, 'Saiu da rotina CheckMessageQueue(erro)');
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Saiu da rotina CheckMessageQueue(erro) - ' + Er.Message);
      TraceWFile(dmCorreio.traceFName, 'Saiu da rotina CheckMessageQueue(erro) - ' + Er.Message);
      exit;
    end;
  end;

  if dmCorreio.dtsMsg.eof then
  begin
    TraceWFile(dmCorreio.traceFName, 'Saiu da rotina CheckMessageQueue');
    result := 0;
    exit;
  end;

  if BCCorreio.IsConectado then
    BCCorreio.DesconectaWS(Erro);

  { Conecta no WebSerice }
  BCCorreio.Instituicao := dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString;
  BCCorreio.Dependencia := dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString;
  BCCorreio.Operador    := dmCorreio.dtsVM_Units.FieldByName('OPERADOR').AsString;
  BCCorreio.Senha       := DescriptografaTextoNum(CHAVE_SENHA_BACEN, dmCorreio.dtsVM_Units.FieldByName('SENHA').AsString);
  BCCorreio.Data_Ini    := dmCorreio.ConectorRobo.Obtem_Menor_Data_Msg_Nao_Lida(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                                                dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString);
  // Alterado 04/06/2012
  //BCCorreio.Data_Fim    := Date;
  BCCorreio.Data_Fim    := EncodeDate(StrToInt(Copy(DateToStr(vim830Global.DataLimiteAte), 7, 4)),
                                      StrToInt(Copy(DateToStr(vim830Global.DataLimiteAte), 4, 2)),
                                      StrToInt(Copy(DateToStr(vim830Global.DataLimiteAte), 1, 2)));

  // Alterado 11/01/2012
  BCCorreio.Proxy       := FProxy;
  BCCorreio.UserProxy   := FUserProxy;
  BCCorreio.SenhaProxy  := FSenhaProxy;

  nret := LogonToBCCorreio;
  if nret <= 0 then
    exit;

  { Captura as pastas Autorizadas }
  if BCCorreio.ConsultarPastasAutorizadas(erro) <> 1 then
  begin
    Insere_Log(ERRO_CAPTURAR_PASTA_AUTORIZADA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);

    if POS('(401)', Erro) > 0 then
    begin
      // Alterado 26/03/2012
      if dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger < 2 then
        dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                              dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                              'N',
                                                              dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger + 1)
      else
      dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                            dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                            'S',
                                                            0);
    end;
    // Alterado 23/01/2012
    BCCorreio.DesconectaWS(Erro);
    exit;
  end;

  // Alterado 26/03/2012
  dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                        dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                        'N',
                                                        0);

  { Verifica se existe a pasa CaixaEntrada }
  localizou := false;
  for i:=0 to count_Pastas-1 do
  begin
    if Pastas[i].Tipo = 'CaixaEntrada' then
    begin
      localizou := true;
      break;
    end;
  end;

  if not localizou then
  begin
    Erro := 'Login [' +
             dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' +
             dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '] não tem pormissão na pasta CaixaEntrada - Não foi localizada na rotina ConsultarPastasAutorizadas';
    Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);
    exit;
  end;

  { Captura as mensagens da pasta CaixaEntrada }
  if BCCorreio.ConsultarCorreioPorPasta('CaixaEntrada', erro) <> 1 then
  begin
    Insere_Log(ERRO_CORREIO_POR_PASTA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);

    if POS('(401)', Erro) > 0 then
    begin
      // Alterado 26/03/2012
      if dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger < 2 then
        dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                              dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                              'N',
                                                              dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger + 1)
      else
      dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                            dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                            'S',
                                                            0);
    end;
    // Alterado 23/01/2012
    BCCorreio.DesconectaWS(Erro);
    exit;
  end;

  // Alterado 26/03/2012
  dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                        dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                        'N',
                                                        0);

  msgNum := '';

  while Not dmCorreio.dtsMsg.Eof do
  begin
    Application.ProcessMessages;

    { verifica encerramento do programa }
    if (pleaseClose) then
    begin
      ReleaseFile(FdbDir + 'MSGS\' + msgNum + '.LCK' , lckMsgHandle);
      result := 0;
      exit;
    end;

    if Trim(dmCorreio.dtsMsg.FieldByName('NUM_SISBACEN').AsString) = '' then
    begin
      dmCorreio.dtsMsg.Next;
      Continue;
    end;

    if Not VerificaFiltro(dmCorreio.dtsMsg.FieldByName('COD_DEP_REMETENTE').AsString, dmCorreio.dtsMsg.FieldByName('ASSUNTO').AsString) then
    begin
      dmCorreio.dtsMsg.next;
      continue;
    end;

    msgNum := dmCorreio.dtsMsg.FieldByName('NUM_SISBACEN').AsString;
    nRet := IsFileLocked(FdbDir + msgNum + '.LCK' , lckMsgHandle);
    if nRet <= 0 then
    begin
      dmCorreio.dtsMsg.Next;
      Continue;
    end;

    ExibeMensagem('Vai ler mensagem ' + dmCorreio.dtsMsg.FieldByname('NUM_SISBACEN').AsString + '-' + dmCorreio.dtsMsg.FieldByname('ASSUNTO').AsString);

    TraceWFile(dmCorreio.traceFName, 'Vai ler mensagem');

    nRet := ReadMailText(dmCorreio.dtsMsg.FieldByname('TipoMsg').AsString,
                         dmCorreio.dtsMsg.FieldByName('NUM_SISBACEN').AsString,
                         dmCorreio.dtsMsg.FieldByName('ASSUNTO').AsString,
                         dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,
                         dmCorreio.dtsMsg.FieldByName('COD_DEP_REMETENTE').AsString,
                         dmCorreio.dtsMsg.FieldByName('STATUS_SISBACEN').AsString,
                         dmCorreio.dtsMsg.FieldByName('COD_MSG').AsInteger);

    ReleaseFile(FdbDir + msgNum + '.LCK' , lckMsgHandle);

    inc(contador);
    if contador >= 20 then
      break;

    if nRet <= 0 then
    begin
      dmCorreio.ConectorRobo.Inclui_Tentativa_Leitura(dmCorreio.dtsMsg.FieldByName('COD_MSG').AsInteger);
      dmCorreio.dtsMsg.close;
      exit;
    end;

    dmCorreio.dtsMsg.close;

    { Captura as msgs não lidas }
    dmCorreio.ConectorRobo.Obtem_Msg_Nao_Lida(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                              dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,

                                              // Alterado 04/06/2012
                                              vim830Global.DataLimiteAte,

                                              dmCorreio.dtsMsg);

    result := 1;
  end;

  // Alterado 19/06/2012
  dmCorreio.dtsMsg.Close;

  BCCorreio.DesconectaWS(Erro);
end;

procedure TReadBCCorreio.ConnectToEmail;
var
  ativo : boolean;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina ConnectToemail');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina ConnectToEmail');

  if mailInfo.mailConnected then
    exit;

  if not sendMail then
  begin
    mailInfo.mailConnected := false;
    exit;
  end;

  mailinfo.mailSystem := vim830Global.MailSystem;
  if vim830Global.mailSystem = 0 then
    exit;

  if vim830Global.mailSystem = IDSM_SYS_OUTLOOK then
  begin
    try
      if FIsDebug then
        ExibeMensagem('Conectando ao sistema de Correio - Outlook');

      //Coinitialize(nil);

      mailInfo.mailSys := CreateOleObject('Outlook.Application');
      mailInfo.nameSpace := mailInfo.mailSys.GetNameSpace('MAPI');
      mailInfo.mailConnected := true;
    except
      mailInfo.mailConnected := false;
    end;
    exit;
  end;

  if vim830Global.mailSystem IN [IDSM_SYS_MAPI, IDSM_SYS_VIM, IDSM_SYS_ACTIVE_MESSAGING , IDSM_SYS_NOTES]  then //idsmail
  begin
    try
      if FIsDebug then
        ExibeMensagem('Conectando ao sistema de Correio - Idsmail');

      //OleInitialize(nil);

      mailInfo.mailSys := CreateOLEObject('IDSMailInterface32.Server');
      mailInfo.mailSys.ObjectKey := 'QRPTHBH6ARQS';
      mailInfo.mailSys.MailSystem := vim830Global.mailSystem;

      if vim830Global.mailSystem = IDSM_SYS_SMTP_POP then
      begin
        mailInfo.mailSys.SMTPServer := vim830Global.smtpInfo.SmtpServer;

        try
          mailInfo.mailSys.SMTPPort := strToInt(vim830Global.smtpInfo.SmtpPort);
        except
          mailInfo.mailSys.SMTPPort := 25;
        end;

        mailInfo.maillogged := false;

        if Trim(vim830Global.LoginName) <> '' then
        begin
          mailInfo.mailSys.LoginName := vim830Global.LoginName;
          mailInfo.mailSys.Password := vim830Global.Password;
          mailInfo.mailSys.login;
        end;
      end;

      if (vim830Global.mailSystem = IDSM_SYS_VIM) or  (vim830Global.mailSystem = IDSM_SYS_NOTES)  then
      begin
        mailInfo.mailSys.LoginName := vim830Global.loginName;
        try
          mailInfo.mailSys.Password := LowerCase(vim830Global.Password);
          mailInfo.MailSys.login;
        except
          mailInfo.mailSys.LoginDialog := true;
          mailInfo.MailSys.login;
        end;
      end;
      mailInfo.mailConnected := true;
    except
      // Alterado 08/05/2012
      on merr : EErro do
      begin
        ExibeMensagem('Erro na conexão ao e-mail: ' + merr.Mensagem);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na conexão ao e-mail: ' + merr.Mensagem);
        TraceWFile(dmCorreio.traceFName, 'Erro na conexão ao e-mail: ' + merr.Mensagem);
        mailInfo.mailConnected := false;
      end;

      on merr : Exception do
      begin
        ExibeMensagem('Erro na conexão ao e-mail: ' + merr.message);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na conexão ao e-mail: ' + merr.message);
        TraceWFile(dmCorreio.traceFName, 'Erro na conexão ao e-mail: ' + merr.message);
        mailInfo.mailConnected := false;
      end;
    end;
  end;

  if vim830Global.mailSystem = IDSM_SYS_SMTP_POP then
  begin
    try
      if FIsDebug then
        ExibeMensagem('Conectando ao sistema de Correio - SMTP');

      if InternetLogging then
        mailInfo.smtpObj := TSmtpObject.Create(30,ExibeMensagem, 0)
      else
        mailInfo.smtpObj := TSmtpObject.Create(30,Nil, 0);

      mailInfo.smtpObj.useSsl := useSsl;
      mailInfo.smtpObj.useAuth := useAuth;

      mailInfo.smtpObj.HostName := vim830Global.smtpInfo.SmtpServer;
      mailInfo.smtpObj.PortNumber := vim830Global.smtpInfo.SmtpPort;
      mailInfo.smtpObj.LoginName := vim830Global.LoginName;
      mailInfo.smtpObj.Password := vim830Global.Password;
      mailInfo.smtpObj.AddressName := vim830Global.smtpInfo.addressName;
      mailInfo.mailConnected := true;
    except
      // Alterado 08/05/2012
      on merr : EErro do
      begin
        ExibeMensagem('Erro na conexão ao e-mail: ' + merr.Mensagem);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na conexão ao e-mail: ' + merr.Mensagem);
        TraceWFile(dmCorreio.traceFName, 'Erro na conexão ao e-mail: ' + merr.Mensagem);
        mailInfo.mailConnected := false;
      end;

      on merr : Exception do
      begin
        ExibeMensagem('Erro na conexão ao e-mail: ' + merr.message);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na conexão ao e-mail: ' + merr.message);
        TraceWFile(dmCorreio.traceFName, 'Erro na conexão ao e-mail: ' + merr.message);
        mailInfo.mailConnected := false;
      end;
    end;
  end;
end;

constructor TReadBCCorreio.Create(Memo: TRichEdit; ADOConn: TADOConnection;
  Hora_Logon, Hora_Logout, dbDir: String; runOnWeekends: boolean;
  mailSystem: TMailInfo);
var
  ini: TIniFile;
  Chave: String;
begin
  FmemMsg        := Memo;
  FADOConn       := ADOConn;
  FHora_Logon    := Hora_Logon;
  FHora_Logout   := Hora_Logout;
  FdbDir         := dbDir;
  FrunOnWeekends := runOnWeekends;

   mailInfo := mailSystem;

  { Carrega os dados do Debug }
  ini            := TIniFile.Create(FdbDir + 'VIM830.INI');
  FIsDebug       := ini.ReadBool('DEBUG', 'Ativo', false);
  FSalvarArquivo := ini.ReadBool('DEBUG', 'SalvarArquivo', false);
  FdirArqDebug   := ini.ReadString('DEBUG', 'CaminhoArquivo', 'C:\');

  // Alterado 07/05/2012
  LerWebPost := ini.ReadBool(SECTION_GLOBAL, 'LerWebPost', False);

  ini.Free;

  // Alterado 07/05/2012
  if LerWebPost then
    Insere_Log(INFORMATIVO, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO, 'Ativado modo LerWebPost');

  ini            := TIniFile.Create(FdbDir + 'VIM830.INI');
  EhTamLinha74   := ini.ReadBool('QUEBRA', 'TAMLINHA74', False);
  ini.Free;

  // Alterado 11/01/2012
  ini         := TIniFile.Create(FdbDir + 'CONECTOR.INI');
  FProxy      := ini.ReadString('PROXY', 'Proxy', '');
  FUserProxy  := ini.ReadString('PROXY', 'User', '');
  FSenhaProxy := ini.ReadString('PROXY', 'SH', '');
  Chave       := ini.ReadString(SECTION_SYSTEM, ENTRY_ACESSO, '');
  if FSenhaProxy <> '' then
  begin
    // Alterado 12/01/2012 - após distribuição para clientes
    ExibeMensagem('Proxy ativado');
    try
      FSenhaProxy := DescriptografaTextoNum(Chave, FSenhaProxy);
    except
      FSenhaProxy := '';
    end;
  end
  else
  begin
    // Alterado 12/01/2012 - após distribuição para clientes
    ExibeMensagem('Proxy desativado');
  end;
  ini.Free;

  if Copy(FdirArqDebug, Length(FdirArqDebug), 1 ) <> '\' then
    FdirArqDebug := FdirArqDebug + '\';

  if not DirectoryExists(FdirArqDebug) then
    FSalvarArquivo := false;

  if FIsDebug then
    ExibeMensagem('Carregando as configurações do robô');

  { Carrega as configurações do robô }
  ini             := TInifile.Create(FdbDir + 'ROBOT.INI');
  sendMail        := ini.ReadBool('SendMail', 'Active', False);
  checkMsgInst    := ini.ReadBool(SECTION_SYSTEM, ENTRY_CHECK_INST, False);
  depLer          := ini.ReadString(SECTION_SYSTEM, ENTRY_DEP_CODE, '');
  internetLogging := ini.ReadBool(SECTION_SYSTEM, 'internetlogging', false);
  dispatchMsg 		:= ini.ReadBool(SECTION_SYSTEM, 'dispatchMsg', True);
  sisbacenPublico := ini.ReadBool(SECTION_SYSTEM, 'PublicSisbacen', false);
  useSSL          := ini.ReadBool(SECTION_SYSTEM, 'useSsl', false);
  useAuth         := ini.ReadBool(SECTION_SYSTEM, 'useAuth', false);
  ini.Free;

  if FIsDebug then
    ExibeMensagem('Carregando o diretório do PWTAW10');

  { Carrega o diretório do PSTAW10 }
  dmCorreio.Conector.Seleciona_Configuracao(dmCorreio.dts);
  FDirExePSTAW10 := dmCorreio.dts.FieldByName('DIR_EXE').AsString;
  dmCorreio.dts.Close;

  if FIsDebug then
    ExibeMensagem('Criando a classe do BCCorreio');

  // Alterado 07/05/2012
  if LerWebPost then
  begin
    try
      if not FileExists(FdbDir + 'WebPost.Ini') then
        raise Exception.Create('Não foi encontrado o arquivo WebPost.Ini');
      ADOConnWebPost  := TADOConnection.Create(nil);
      AcessorWebPost  := TAcessorPesLegado.Create(FdbDir + 'WebPost.Ini');
      ConectorwebPost := AcessorwebPost.GetConexao('WEBPOST');
      ADOConnWebPost  := ConectorWebPost.GetConexao;

      fMainQryWP := TADOQuery.Create(nil);
      fSegQryWP  := TADOQuery.Create(nil);

      fMainqryWP.Connection := ADOConnWebPost;
      fSegqryWP.Connection  := ADOConnWebPost;

      fMainQryWP.CommandTimeout := Timeout;
      fSegQryWP.CommandTimeout  := Timeout;
    except
      on Er: EErro do begin
        raise Exception.Create('Erro ao conectar com o WebPost: ' + Er.Mensagem);
			end;
      on Er: EAviso do begin
        raise Exception.Create('Erro ao conectar com o WebPost: ' + Er.Mensagem);
			end;
      on Er: EInformacao do begin
        raise Exception.Create('Erro ao conectar com o WebPost: ' + Er.Mensagem);
			end;
      on Er: EErroLog do begin
        raise Exception.Create('Erro ao conectar com o WebPost: ' + Er.Mensagem);
			end;
      on Er: Exception do begin
        raise Exception.Create('Erro ao conectar com o WebPost: ' + Er.Message);
      end;
    end;
  end
  else
  begin
    //BCCorreio := TBCCorreio.Create(FmemMsg, FIsDebug, FSalvarArquivo , FdirArqDebug, EhTamLinha74);
    if EhTamLinha74 then
      BCCorreio := TBCCorreio.Create(FmemMsg, FIsDebug, FSalvarArquivo , FdirArqDebug, 74)
    else
      BCCorreio := TBCCorreio.Create(FmemMsg, FIsDebug, FSalvarArquivo , FdirArqDebug, 80);
  end;
  vim830Global := TVim830Global.Create;

{$IFDEF MSWINDOWS}
  ExibeMensagem('Compilado com MSWINDOWS');
  TraceWFile(dmCorreio.traceFName, 'Compilado com MSWINDOWS');
{$ELSE}
  ExibeMensagem('Compilado sem MSWINDOWS');
  TraceWFile(dmCorreio.traceFName, 'Compilado sem MSWINDOWS');
{$ENDIF}
end;

procedure TReadBCCorreio.DisconnectEmail;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina DisconnectEmail');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina DisconnectEmail');

  if not mailInfo.mailConnected then
    exit;

  try
    if vim830Global.mailSystem = 0 then
      exit;

    if vim830Global.mailSystem = IDSM_SYS_OUTLOOK then
    begin
      try
        mailinfo.mailSys.quit;
      except
      end;
      exit;
    end;

    if vim830Global.mailSystem IN [IDSM_SYS_MAPI, IDSM_SYS_VIM, IDSM_SYS_ACTIVE_MESSAGING , IDSM_SYS_NOTES]  then //idsmail
    begin
      try
        mailInfo.mailSys.Logout;
      except
      end;
      exit;
    end;

    if vim830Global.mailSystem = IDSM_SYS_SMTP_POP then
    begin
      try
        FreeAndNil(mailInfo.smtpObj);
      except
      end;
      exit;
    end;
  finally
    mailInfo.mailConnected := false;
  end;
end;

procedure TReadBCCorreio.DoShowProxProc;
var
  minutes : TDateTime;
  present : TDateTime;
  dateChar : array[0..20] of char;
  msg: string;
begin
  minutes := EncodeTime(0, 1, 0, 0);
  present := Now;
  msg     := 'Aguardando próxima verificação - ' + FormatDateTime('[HH:NN:SS]', present + minutes);
  StrPCopy(dateChar, FormatDateTime('[hh:nn:ss] ', Now));
  msg := StrPas(dateChar) + msg;
  if FmemMsg <> nil then
    FmemMsg.Lines.Add(Msg);
end;

procedure TReadBCCorreio.EhOficio(subject: string; sender: string; var EhOficio: boolean);
var
  ini : TIniFile;
  sdValue : TStringList;
  str : string;
  j : integer;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina EhOficio');
  TraceWFile(dmCorreio.traceFName, 'Sender = ' + sender);

  if FIsDebug then
  begin
    ExibeMensagem('Entrou na rotina EhOficio');
    ExibeMensagem('Sender = ' + sender);
  end;

  EhOficio := false;

  ini := TIniFile.Create(FdbDir + 'VIM830.INI');
  try
     str := ini.ReadString('AUTOJUD', 'Sender', '');

     if FIsDebug then
       ExibeMensagem('[AUTOJUD] Sender = ' + str);
     TraceWFile(dmCorreio.traceFName, '[AUTOJUD] Sender = ' + str);

     if str = '' then
     begin
        ini := TIniFile.Create(FdbDir + 'VIM830.INI');
        str := ini.readString('oficios', 'subject', '');

        if FIsDebug then
          ExibeMensagem('[oficios] subject = ' + str);
        TraceWFile(dmCorreio.traceFName, '[oficios] subject = ' + str);
     end;

     if str = '' then
        exit;

     sdValue := TStringList.Create;
     try
        StrToList(str, ',', sdValue);

        TraceWFile(dmCorreio.traceFName, 'Total de Dpto: ' + IntToStr(sdValue.Count));

        if FIsDebug then
          ExibeMensagem('Total de Dpto: ' + IntToStr(sdValue.Count));

        if sdValue.count <> 0 then
        begin
          for j := 0 to sdValue.Count - 1 do
          begin
            TraceWFile(dmCorreio.traceFName, 'Verificando ' + sdValue.strings[j]);

            if FIsDebug then
              ExibeMensagem('Verificando ' + sdValue.strings[j]);

            if UPPERCASE(tRIM(sdValue.strings[j])) = Uppercase(Trim(sender)) then
            begin
              TraceWFile(dmCorreio.traceFName, 'Localizado');
              break;
            end;
          end;

          if j <= (sdValue.count-1) then
          begin
             ehOficio := true;
             TraceWFile(dmCorreio.traceFName, 'ehOficio');
          end
          else
            TraceWFile(dmCorreio.traceFName, 'Não ehOficio');
        end;
     finally
        sdValue.free;
     end;
  finally
     ini.Free;
  end;

  if FIsDebug then
    ExibeMensagem('Saiu da rotina EhOficio');
end;

function TReadBCCorreio.EntregaEstaMensagem(qryMsgs: TADODataSet; ext: string;
  ultima: boolean): integer;
var
  apply : boolean;
  emailLst, emailCC, EmailCCO : TstringList;
  tableName : string;
  userName : string;
  userCode : string;
  myHeader : TStringList;
  myText : TStringList;
  msgsSentTo : TStringList;
  tUser : TuserName;
  tUserSent : TuserName;
  i : integer;
  j : integer;
  p : integer;
  folder : array [1..5] of string;
  theRecipient : string;
  emailInfo : TEmailAddressInfo;
  achou : boolean;
  address : string;
  listEmail : TStringList;
  passo : string;
  enviarPara : string;
  naoEnviada : boolean;

  myFiles :  TstringList;
  diretorio: String;
  Arquivo_Origem: String;
  Arquivo_Destino: String;

  // Alterado 26/12/2011 - Safra
  EnviaNote: TEnviaNote;
  ret: integer;
  erro: string;

  // Alterado 14/06/2012
  Tem_Anexo: Boolean;

  EmailDotNet : TEmailNet;
  CaminhoXml: string;
  quebraInicio, quebraFim : string;

begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina EntregaEstaMensagem');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina EntregaEstaMensagem');

  // Alterado 23/01/2012
  ConnectToEmail;

  myText  := TStringList.Create;
  myFiles := TStringList.Create;
  try
    try
      passo := '';

      if FIsDebug then
        ExibeMensagem('Montando a mensagem para enviar');

      diretorio := '';
      diretorio := GetEnvironmentVariable('HOMEDRIVE');
      if diretorio <> '' then
        diretorio := diretorio + GetEnvironmentVariable('HOMEPATH') + '\'
      else
      begin
        diretorio := GetEnvironmentVariable('TEMP');
        if diretorio = '' then
          diretorio := GetEnvironmentVariable('TMP');
        diretorio := diretorio + '\';
      end;

      // Alterado 26/12/2011 - Safra
      if EhSafraNotes then
      begin
        EnviaNote := TEnviaNote.Create(FmemMsg,
                                       dmcorreio.ADOConnection1,
                                       BCCorreio.Dir_Arq_Anexo,
                                       diretorio,
                                       FdbDir,
                                       FIsDebug,
                                       FSalvarArquivo,
                                       FdirArqDebug);

        ret := EnviaNote.EnviaMensagens(qryMsgs,
                                        erro);

        if ret <> 1 then
        begin
          //
          //
          //
          //
          //
          //
          //
        end;

        result := ret;
        exit;
      end;

      { carrega o texto da mensagem }
      dmCorreio.Conector.Captura_Msg_Texto_PMSG(qryMsgs.fieldByName('COD_MSG').AsInteger,
                                                dmcorreio.dts2);

      dmcorreio.dts2.First;
      while not dmcorreio.dts2.eof do
      begin
        Application.ProcessMessages;
        myText.Add(dmcorreio.dts2.FieldByName('Linha').AsString);
        dmcorreio.dts2.next;
      end;
      dmcorreio.dts2.Close;

      // Alterado 08/05/2012
      if not LerWebPost then
      begin
        { carrega o arquivo em anexo }
        dmCorreio.ConectorRobo.Obtem_VM_MESSAGES_ANEXOS(qryMsgs.fieldByName('COD_MSG').AsInteger,
                                                        dmcorreio.dts2);


        dmcorreio.dts2.First;
        while not dmcorreio.dts2.eof do
        begin
          Application.ProcessMessages;

          // Alterado 14/06/2012
          Tem_Anexo := true;

          try
            Arquivo_Origem  := BCCorreio.Dir_Arq_Anexo +
                               qryMsgs.FieldByName('num_sisbacen').AsString + '.' + LeftZeroStr('%03d', [dmcorreio.dts2.FieldByName('SEQUENCIA').AsInteger]);
            Arquivo_Destino := diretorio +
                               dmcorreio.dts2.FieldByName('NOME_ARQUIVO').AsString;

            if CopyFile(Pchar(Arquivo_Origem),
                        Pchar(Arquivo_Destino),
                        false) then
              myFiles.Add(Arquivo_Destino);
            // Alterado 14/06/2012
            //else
            //  myFiles.Add(Arquivo_Origem);
          except
          //
          //
          //
          end;

          dmcorreio.dts2.next;
        end;
        dmcorreio.dts2.Close;
      end;

      tableName := tabVM_USERS;

      if FIsDebug then
        ExibeMensagem('Selecionando o e-mail do perfil');

      { carrega os e-mails }
      dmCorreio.ConectorRobo.Obtem_Email_Perfil(1, dmcorreio.dtsEmail);

      if dmcorreio.dtsEmail.Eof then
        dmCorreio.ConectorRobo.Obtem_Email_Perfil(2, dmcorreio.dtsEmail);

      if dmcorreio.dtsEmail.eof then
      begin
        if FIsDebug then
          ExibeMensagem('Não foi selecionado nenhum e-mail');

          SaveMessageSent(qryMsgs);
          myText.free;
          result := 1;

          // Alterado 19/06/2012
          dmcorreio.dtsEmail.Close;

          exit;
      end;

      eMailLst := TStringList.Create;
      msgsSentTo := TStringList.Create;
      listEmail := TStringList.create;
      naoEnviada := true;

      while Not dmcorreio.dtsEmail.eof do
      begin
        Application.ProcessMessages;

        if Trim(dmcorreio.dtsEmail.FieldByName('email_address').AsString) = '' then
        begin
          dmcorreio.dtsEmail.next;
          continue;
        end;

        userName := dmcorreio.dtsEmail.FieldByName('Nome').AsString;
        userCode := dmcorreio.dtsEmail.FieldByName('codigo').AsString;

        ExibeMensagem('Verificando mensagem no perfil: ' + qryMsgs.FieldByName('num_sisbacen').AsString + '-' + userName);

        { verifica se a mensagem faz parte do filtro deste usuario }
        if (Pos('DENOR', qryMsgs.FieldByName('num_sisbacen').AsString) <> 0) And
           (Pos('JUD', UpperCase(dmcorreio.dtsEmail.FieldByName('Email_Address').AsString)) = 0) then
          apply := true
        else
          apply := MessageAppliesUserFilter(qryMsgs, userCode, ext, mytext);

        if (apply) then
        begin
          ExibeMensagem('Verificando se enviou a mensagem para o perfil ' + userName);

          if Not JaEnviou(usercode, qryMsgs.FieldByName('cod_msg').AsString) then
          begin
            tUser := tUserName.Create;
            tUser.userName := userCode;
            tUser.notifica := dmcorreio.dtsEmail.FieldByName('Notifica').AsInteger;
            if dmcorreio.dtsEmail.FieldByName('Notifica').AsInteger = 0 then
              emailLst.AddObject(dmcorreio.dtsEmail.FieldByName('Email_Address').AsString, tUser);
            emailInfo := TEmailAddressInfo.Create;
            emailInfo.emailAddress := dmcorreio.dtsEmail.FieldByName('Email_Address').AsString;
            emailInfo.userCode := userCode;
            emailInfo.notifica := dmcorreio.dtsEmail.FieldByName('Notifica').AsInteger;
            msgsSentTo.AddObject(userName, emailInfo);
            naoEnviada := false;

            ExibeMensagem('Precisa enviar e-mail para:' + userName + ' - ' + dmcorreio.dtsEmail.FieldByName('Email_Address').AsString);
          end
          else
          begin
            ExibeMensagem('Já foi enviado e-mail para o perfil ' + userName);
          end;
        end;
        dmcorreio.dtsEmail.next;
      end;

      // Alterado 19/06/2012
      dmcorreio.dtsEmail.Close;

      if (naoEnviada) and (vim830Global.DirecionaExcecao <> '') then
      begin
        ExibeMensagem('Vai direcionar para grupo de exceção');

        tUser := tUserName.Create;
        tUser.userName := userCode;
        tUser.notifica := 0;
        emailLst.AddObject(vim830Global.DirecionaExcecao, tUser);
        emailInfo := TEmailAddressInfo.Create;
        emailInfo.emailAddress := vim830Global.DirecionaExcecao;
        emailInfo.userCode := '';
        emailInfo.notifica := 0;
        msgsSentTo.AddObject('', emailInfo);
        naoEnviada := false;
      end;

      if (emailLst.Count <= 0) and (naoEnviada) then
      begin
        ExibeMensagem('Mensagem não se aplicou a nenhum filtro');

        dmCorreio.ConectorRobo.Obtem_Email_Perfil(2, dmcorreio.dtsEmail);

        while Not dmcorreio.dtsEmail.Eof do
        begin
          Application.ProcessMessages;

          if Trim(dmcorreio.dtsEmail.FieldByName('email_address').AsString) = '' then
          begin
            dmcorreio.dtsEmail.next;
            continue;
          end;

          userName := dmcorreio.dtsEmail.FieldByName('Nome').AsString;
          userCode := dmcorreio.dtsEmail.FieldByName('codigo').AsString;

          tUser := tUserName.Create;
          tUser.userName := userCode;
          tUser.notifica := dmcorreio.dtsEmail.FieldByName('Notifica').AsInteger;
          if dmcorreio.dtsEmail.FieldByName('Notifica').AsInteger = 0 then
            emailLst.AddObject(dmcorreio.dtsEmail.FieldByName('Email_Address').AsString, tUser);
          emailInfo := TEmailAddressInfo.Create;
          emailInfo.emailAddress := dmcorreio.dtsEmail.FieldByName('Email_Address').AsString;
          emailInfo.userCode := userCode;
          emailInfo.notifica := dmcorreio.dtsEmail.FieldByName('Notifica').AsInteger;
          msgsSentTo.AddObject(userName, emailInfo);
          dmcorreio.dtsEmail.next;
        end;

        // Alterado 19/06/2012
        dmcorreio.dtsEmail.Close;
      end;

      if (emailLst.Count <= 0) and (naoEnviada) then
      begin
  //
  //
  //
  //
  //
  //
  //
  //
  //
  //
        SaveMessageSent(qryMsgs);
        result := 1;
        eMailLst.Free;
        msgsSentTo.Free;
        listEmail.free;
        mytext.free;
        exit;
      end;

      if Not mailInfo.mailConnected Then
      begin
        ExibeMensagem('<ERRRO> Não está conectado no e-mail');

        result := 1;
        eMailLst.Free;
        msgsSentTo.Free;
        listEmail.free;
        mytext.free;
        exit;
      end;

      if mailInfo.mailConnected Then
      begin
        if vim830Global.mailSystem = IDSM_SYS_OUTLOOK then
        begin
          ExibeMensagem('ENVIANDO E-MAIL: ' + qryMsgs.FieldByName('num_sisbacen').AsString  + '-' +
                                              qryMsgs.FieldByName('Assunto').AsString);

          { primeiro envia pasta publica }
          for i := 0 to emailLst.count - 1 do
          begin
            { verifica se eh pasta publica }
            if (Pos('PUBLIC FOLDER', UpperCase(emailLst.strings[i])) <> 0 ) Or
               (Pos('PASTAS P', UpperCase(emailLst.strings[i])) <> 0 ) then
            begin
              j := 1;
              address := emailLst.strings[i];
              p := Pos('\\', address);
              while p <> 0 DO
              begin
                folder[j] := Copy(address, 1, p-1);
                Inc(j);
                address := Copy(address, p + 2, Length(address));
                p := Pos('\\', address);
              end;

              if address <> '' then
                folder[j] := address;

              if j = 1 then
                mailInfo.mailFolder := mailInfo.nameSpace.Folders(folder[1])
              else
                if j = 2 then
                  mailInfo.mailFolder := mailInfo.nameSpace.Folders(folder[1]).Folders(folder[2])
                else
                  if j = 3 then
                    mailInfo.mailFolder := mailInfo.nameSpace.Folders(folder[1]).Folders(folder[2]).Folders(folder[3])
                  else
                    if j = 4 then
                      mailInfo.mailFolder := mailInfo.nameSpace.Folders(folder[1]).Folders(folder[2]).Folders(folder[3]).Folders(folder[4])
                    else
                      mailInfo.mailFolder := mailInfo.nameSpace.Folders(folder[1]).Folders(folder[2]).Folders(folder[3]).Folders(folder[4]).Folders(folder[5]);

              Sleep(20);
              mailInfo.mailMessage := mailInfo.mailSys.CreateItemFromTemplate(fdbDir + 'msgtemp.oft', mailInfo.mailFolder);
              Sleep(20);

              mailInfo.mailMessage.Body := myText.Text;

              mailInfo.userProp := mailInfo.mailMessage.UserProperties;
              mailInfo.userProp.Item('BacenAssunto') := qryMsgs.FieldByName('Assunto').AsString;
              mailInfo.mailMessage.subject := qryMsgs.FieldByName('Assunto').AsString;

              mailInfo.userProp.Item('BacenNumero') := qryMsgs.FieldByName('Num_Sisbacen').AsString;
              if Pos('DENOR', qryMsgs.FieldByName('Assunto').AsString) <> 0 then
              begin
                if Length(qryMsgs.FieldByName('Data_Envio').AsString) < 10 then
                  mailInfo.userProp.Item('BacenDataEnvio') := StrToDateTime(qryMsgs.FieldByName('Data_Envio').AsString + '/' +
                                                              FormatDateTime('yy', Now) + ' ' +
                                                              qryMsgs.FieldByName('Hora_Envio').AsString)
                else
                  mailInfo.userProp.Item('BacenDataEnvio') := StrToDateTime(qryMsgs.FieldByName('Data_Envio').AsString + ' ' +
                                                              qryMsgs.FieldByName('Hora_Envio').AsString);
              end
              else
              begin
                if Length(qryMsgs.FieldByName('Data_Envio').AsString) < 10 then
                  mailInfo.userProp.Item('BacenDataEnvio') := StrToDateTime(qryMsgs.FieldByName('Data_Envio').AsString + '/' +
                                                              Copy(qryMsgs.FieldByName('Amd').AsString, 3, 2) + ' ' +
                                                              qryMsgs.FieldByName('Hora_Envio').AsString)
                else
                  mailInfo.userProp.Item('BacenDataEnvio') := StrToDateTime(qryMsgs.FieldByName('Data_Envio').AsString + ' ' +
                                                              qryMsgs.FieldByName('Hora_Envio').AsString);
              end;

              mailInfo.userProp.Item('BacenRemetente') := qryMsgs.FieldByName('Cod_dep_Remetente').AsString;
              mailInfo.userProp.Item('BacenNomeRemetente') := qryMsgs.FieldByName('DESCR_DEP_REMETENTE').AsString;
              mailInfo.userProp.Item('BacenStatus') := qryMsgs.FieldByName('Status_sisbacen').AsString;
              mailInfo.userProp.Item('BacenDataRecebimento') := qryMsgs.FieldByName('Dt_Recebimento').AsString;
              mailInfo.userProp.Item('BacenHoraRecebimento') := qryMsgs.FieldByName('Hr_Recebimento').AsString;
              mailInfo.userProp.Item('BacenRecebidoPor') := qryMsgs.FieldByName('Recebido_Por').AsString;
              mailInfo.userProp.Item('BacenDependencia') := qryMsgs.FieldByName('cod_inst').AsString + ' / ' +
              qryMsgs.FieldByName('cod_dep').AsString;
              mailInfo.mailMessage.post;

              tUserSent := tUserName(emailLst.objects[i]);
              if tUserSent.username <> '' then
                MarcaEnviada(tUserSent.userName, qryMsgs.FieldByName('cod_msg').AsString);
            end;
          end;

          { agora envia uma mensagem para todos os usuarios }
          mailInfo.mailFolder := mailInfo.namespace.GetDefaultFolder(4);
          mailInfo.mailItem := mailInfo.mailsys.CreateItem(0);

          achou := false;
          for i := 0 to emailLst.count - 1 do
          begin
            { verifica se eh pasta publica }
            if (Pos('PUBLIC FOLDER', UpperCase(emailLst.strings[i])) = 0 ) And
               (Pos('PASTAS P', UpperCase(emailLst.strings[i])) = 0 ) then
            begin

              if tUserName(emailLst.objects[i]).notifica = 0 then
              begin
                achou := true;
                listEmail.Clear;
                if Pos(';', emailLst.strings[i]) <> 0 then
                begin
                  StrToList(emailLst.strings[i], ';', listEmail);
                  For p := 0 to listEmail.count - 1 do
                  begin
                    if Trim(listEmail.Strings[p]) <> '' then
                      mailinfo.mailitem.recipients.add(listEmail.Strings[p]);
                  end;
                end
                else
                  mailinfo.mailitem.recipients.add(emailLst.Strings[i]);
              end;
            end;
          end;

          if achou then
          begin
  {$ifdef BVR}
            mailInfo.mailitem.Subject := qryMsgs.FieldByName('Assunto').AsString + ' - ' +
                                         qryMsgs.FieldByName('COD_DEP_Remetente').AsString +
                                         ' ('+ qryMsgs.FieldByName('Num_sisbacen').AsString + ' ' +
                                         ')';
  {$else}
            mailInfo.mailitem.Subject := qryMsgs.FieldByName('Assunto').AsString +
                                         ' ('+ qryMsgs.FieldByName('Num_Sisbacen').AsString + ' ' +
                                         ')';
  {$endif}

            myHeader := TStringList.Create;

            // Alterado 14/06/2012
            if Tem_Anexo then
            begin
              myHeader.Add ('*********************************************************************');
              myHeader.Add ('*******          Contem arquivo em anexo na mensagem          *******');
              myHeader.Add ('*********************************************************************');
              myHeader.Add ('');
              myHeader.Add ('');
            end;

            myHeader.Add ('------------ CAPTURA AUTOMATICA DE MENSAGENS DO SISBACEN -------------');
            myHeader.Add ('Assunto: '+ qryMsgs.FieldByName('Assunto').AsString + '  Status: '+  qryMsgs.FieldByName('status_sisbacen').AsString);
            myHeader.Add ('Numero: '+  qryMsgs.FieldByName('NUM_Sisbacen').AsString + '  Data: '+  qryMsgs.FieldByName('Data_Envio').AsString + '  Hora: '+  qryMsgs.FieldByName('Hora_Envio').AsString);
            myHeader.Add ('Remetente: '+  qryMsgs.FieldByName('COD_DEP_Remetente').AsString + '  Instituicao/Dependencia: ' +  qryMsgs.FieldByName('COD_Inst').AsString + '/' + qryMsgs.FieldByName('COD_Dep').AsString);
            myHeader.Add ('Recebido em: ' + qryMsgs.FieldByName('Dt_Recebimento').AsString + ' ' + qryMsgs.FieldByName('HR_Recebimento').AsString);
            myHeader.Add ('---------------------------------------------------------------------');

            myHeader.Text := RemoveAcentuacao(myHeader.Text);

            mailInfo.mailitem.body := myHeader.Text + #13 + #10 + #13 + #10 + myText.Text;

            for i := 0 to myFiles.Count - 1 do
            begin
              mailInfo.mailItem.Attachments.Add(myFiles[i]);
            end;

            mailInfo.mailitem.send;
            myHeader.free;
          end;
        end
        else
        begin
          if (emailLst.Count <> 0) then
          begin
            ExibeMensagem(qryMsgs.FieldByName('Num_Sisbacen').AsString  + '-' +
                          qryMsgs.FieldByName('Assunto').AsString+ '-' +
                          'Vai enviar e-mail');

            TraceWFile(dmCorreio.traceFName, qryMsgs.FieldByName('Num_Sisbacen').AsString  + '-' +
                                                      qryMsgs.FieldByName('Assunto').AsString+ '-' +
                                                      'Vai enviar e-mail');

            try
              if (vim830Global.mailSystem <> IDSM_SYS_SMTP_POP) and
                 (vim830Global.mailSystem <> IDSM_SYS_VIM) and
                 (vim830Global.mailSystem <> IDSM_SYS_NOTES) then
              begin
                mailInfo.mailsys.LoginName := mailInfo.mailsys.DefaultProfile;
                try
                  mailInfo.mailSys.login;
                except
                end;
              end;

              if vim830Global.mailSystem = IDSM_SYS_MAPI then
              begin
                try
                  mailinfo.mailsys.LoginName := mailInfo.mailsys.DefaultProfile;
                except
                end;
                mailinfo.mailsys.Login;
              end;
            except
            end;

            theRecipient := '';

  {$ifdef UBS}
            for i := 0 to emailLst.count - 1 do
            begin
              if tUserName(emailLst.objects[i]).notifica <> 0 then
                continue;

              if Pos(';', emailLst.strings[i]) <> 0 then
              begin
                listEmail.clear;
                StrToList(emailLst.strings[i], ';', listEmail);
                For p := 0 to listEmail.count - 1 do
                begin
                  mailInfo.mailSys.NewMessage;

                  mailInfo.mailSys.AddRecipientTo (Trim(listEmail.strings[p]));
                  mailInfo.mailSys.RecipientDelimiter := ';';

  {$ifdef BVR}
                  mailInfo.mailsys.Subject := qryMsgs.FieldByName('Assunto').AsString + ' - ' +
                                              qryMsgs.FieldByName('COD_DEP_Remetente').AsString +
                                              ' ('+ qryMsgs.FieldByName('Num_Sisbacen').AsString + ' ' +
                                              ')';
  {$else}
                  mailInfo.mailsys.Subject := qryMsgs.FieldByName('Assunto').AsString +
                                              ' ('+ qryMsgs.FieldByName('Num_Sisbacen').AsString + ' ' +
                                              ')';
  {$endif}

                  myHeader := TStringList.Create;

                  // Alterado 14/06/2012
                  if Tem_Anexo then
                  begin
                    myHeader.Add ('*********************************************************************');
                    myHeader.Add ('*******          Contem arquivo em anexo na mensagem          *******');
                    myHeader.Add ('*********************************************************************');
                    myHeader.Add ('');
                    myHeader.Add ('');
                  end;

                  myHeader.Add ('------------ CAPTURA AUTOMATICA DE MENSAGENS DO SISBACEN -------------');
                  myHeader.Add ('Assunto: '+ qryMsgs.FieldByName('Assunto').AsString + '  Status: '+  qryMsgs.FieldByName('Status').AsString);
                  myHeader.Add ('Numero: '+  qryMsgs.FieldByName('Numero').AsString + '  Data: '+  qryMsgs.FieldByName('DataEnvio').AsString + '  Hora: '+  qryMsgs.FieldByName('HoraEnvio').AsString);
                  myHeader.Add ('Remetente: '+  qryMsgs.FieldByName('COD_DEP_Remetente').AsString + '  Instituicao/Dependencia: ' +  qryMsgs.FieldByName('COD_Inst').AsString + '/' + qryMsgs.FieldByName('COD_Dep').AsString);
                  myHeader.Add ('Recebido em: ' + qryMsgs.FieldByName('DataRecebimento').AsString + ' ' + qryMsgs.FieldByName('HoraRecebimento').AsString);
                  myHeader.Add ('---------------------------------------------------------------------');

                  myHeader.Text := RemoveAcentuacao(myHeader.Text);

                  mailInfo.mailSys.message := myHeader.Text + #13 + #10 + #13 + #10 + myText.Text;
                  if (vim830Global.mailSystem = IDSM_SYS_VIM) OR (vim830Global.mailSystem = IDSM_SYS_NOTES) then
                    mailinfo.mailSys.Signed := True;

                  for i := 0 to myFiles.Count - 1 do
                  begin
                    mailInfo.mailsys.AddAttachment(myFiles[i]);
                  end;

                  mailInfo.mailSys.send;

                  myText.free;
                  myHeader.free;
                end;
              end;
            end;
  {$endif}

  {$IFNDEF UBS}
            myHeader := TStringList.Create;

            // Alterado 14/06/2012
            if IsParam('/SalvarXml') then
            begin
               quebraInicio := '';
               quebraFim := '';
(**               myheader.Add('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">');
               myheader.Add('<html>');
               myheader.Add('<head>');
               myheader.Add('<title> New Document </title>');
               myheader.Add('<meta name="Generator" content="EditPlus">');
               myheader.Add('<meta name="Author" content="">');
               myheader.Add('<meta name="Keywords" content="">');
               myheader.Add('<meta name="Description" content="">');
               myheader.Add('</head>');
               myheader.Add('<body>');*)
               myheader.add('<pre>');
            end
            else
            begin
              quebraInicio := '';
              quebraFim := '';
            end;

            if Tem_Anexo then
            begin
              myHeader.Add (quebraInicio + '*********************************************************************'+quebraFim);
              myHeader.Add (quebraInicio +'*******          Contem arquivo em anexo na mensagem          *******'+quebraFim);
              myHeader.Add (quebraInicio +'*********************************************************************'+quebraFim);
              myHeader.Add (quebraInicio +''+quebraFim);
              myHeader.Add (quebraInicio +''+quebraFim);
            end;

            myHeader.Add (quebraInicio +'------------ CAPTURA AUTOMATICA DE MENSAGENS DO SISBACEN -------------'+quebraFim);
            myHeader.Add (quebraInicio +'Assunto: '+ qryMsgs.FieldByName('Assunto').AsString + '  Status: '+  qryMsgs.FieldByName('Status_sisbacen').AsString+quebraFim);
            myHeader.Add (quebraInicio +'Numero: '+  qryMsgs.FieldByName('Num_Sisbacen').AsString + '  Data: '+  qryMsgs.FieldByName('Data_Envio').AsString + '  Hora: '+  qryMsgs.FieldByName('Hora_Envio').AsString+quebraFim);
            myHeader.Add (quebraInicio +'Remetente: '+  qryMsgs.FieldByName('Cod_dep_Remetente').AsString + '  Instituicao/Dependencia: ' +  qryMsgs.FieldByName('cod_Inst').AsString + '/' + qryMsgs.FieldByName('cod_Dep').AsString+quebraFim);
            myHeader.Add (quebraInicio +'Recebido em: ' + qryMsgs.FieldByName('Dt_Recebimento').AsString + ' ' + qryMsgs.FieldByName('Hr_Recebimento').AsString+quebraFim);
            myHeader.Add (quebraInicio +'---------------------------------------------------------------------'+quebraFim);

            if IsParam('/SalvarXml') then
               myheader.Add('</pre>');

            myHeader.Text := RemoveAcentuacao(myHeader.Text);

            for i := 0 to emailLst.count - 1 do
            begin
              if tUserName(emailLst.objects[i]).notifica <> 0 then
                continue;

              if Pos(';', emailLst.strings[i]) <> 0 then
              begin
                StrToList(emailLst.strings[i], ';', listEmail);
                For p := 0 to listEmail.count - 1 do
                begin
                  if Trim(listEmail.strings[p]) <> '' then
                  begin
                    if enviarPara <> '' then
                      enviarPara := Concat(enviarPara, ';');
                    enviarPara := Concat(enviarPara, listEmail.strings[p]);
                  end;
                end;
              end
              else
              begin
                if enviarPara <> '' then
                  enviarPara := Concat(enviarPara, ';');
                enviarPara := Concat(enviarPara, Trim(emailLst.strings[i]));
              end;
            end;

            if (vim830Global.mailSystem <> IDSM_SYS_SMTP_POP) then
            begin
              mailInfo.mailSys.NewMessage;

              mailInfo.mailSys.RecipientDelimiter := ';';
              ExibeMensagem( 'Enviando e-mail para: ' + enviarPara);

              mailInfo.mailSys.AddRecipientTo(enviarPara);

              mailInfo.mailsys.Subject := qryMsgs.FieldByName('Assunto').AsString + ' - ' +
                                          qryMsgs.FieldByName('COD_DEP_Remetente').AsString +
                                          ' ('+ qryMsgs.FieldByName('Num_Sisbacen').AsString  + ' ' +
                                          ')';

              mailInfo.mailSys.message := myHeader.Text + #13 + #10 + #13 + #10 + myText.Text;
              if (vim830Global.mailSystem = IDSM_SYS_VIM) OR (vim830Global.mailSystem = IDSM_SYS_NOTES) then
                mailinfo.mailSys.Signed := True;

              for i := 0 to myFiles.Count - 1 do
              begin
                mailInfo.mailsys.AddAttachment(myFiles[i]);
              end;

              mailInfo.mailSys.send;
            end;

            if vim830Global.mailSystem = IDSM_SYS_SMTP_POP then
            begin
              if tUserName(emailLst.objects[0]).notifica = 0 then
              begin
                result := mailInfo.SmtpObj.Activate(vim830Global.LoginName, vim830Global.Password, vim830Global.smtpInfo.addressName);
                if result <= 0 then
                begin
                  ExibeMensagem('Erro SMTP: ' + mailInfo.smtpObj.SmtpErro);
                  Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro SMTP: ' + mailInfo.smtpObj.SmtpErro);
                  eMailLst.Free;
                  msgsSentTo.Free;
                  listEmail.free;
                  mytext.free;
                  exit;
                end;

                myHeader.Add(' ');
                myHeader.Add(' ');

                if GetParam('/SalvarXml', CaminhoXml) then
                begin
                  if CaminhoXml[Length(CaminhoXml)] <> '\' then
                     CaminhoXml := CaminhoXml + '\';

                  if FileExists(CaminhoXml +  qryMsgs.FieldByName('Num_Sisbacen').AsString + '_Emenda.html') then
                  begin
                     mytext.Clear;
                     myheader.add('<pre>');
                     myText.LoadFromFile(CaminhoXml +  qryMsgs.FieldByName('Num_Sisbacen').AsString + '_Emenda.html');
                     for i := 0 to myText.count - 1 do
                       myHeader.Add(myText.strings[i]);
                     myheader.add('</pre>');
                  end;
                  if FileExists(CaminhoXml +  qryMsgs.FieldByName('Num_Sisbacen').AsString + '_Conteudo.html') then
                  begin
                     mytext.Clear;
                     myText.LoadFromFile(CaminhoXml +  qryMsgs.FieldByName('Num_Sisbacen').AsString + '_Conteudo.html');
                     for i := 0 to myText.count - 1 do
                       myHeader.Add(myText.strings[i]);
                  end
                  else
                  begin
                    for i := 0 to myText.count - 1 do
                      myHeader.Add(myText.strings[i]);
                  end;
                  myHeader.Add('</body>');
                  myheader.Add('</html>');
                end
                else
                begin
                  for i := 0 to myText.count - 1 do
                    myHeader.Add(myText.strings[i]);
                end;

                ExibeMensagem( 'Enviando e-mail para: ' + enviarPara);

                if Not vim830Global.emailDotNet then
                  result := mailInfo.smtpObj.SendMail(qryMsgs.FieldByName('Assunto').AsString + ' - ' +
                                                      qryMsgs.FieldByName('COD_DEP_Remetente').AsString +
                                                      ' ('+ qryMsgs.FieldByName('Num_Sisbacen').AsString  + ' ' +	')',
                                                      emailLst, myFiles, myHeader)
                else
                begin
                  EmailDotNet := TEmailNet.Create(ExtractFilePath(Application.ExeName));
                  try
                     emailCC := TStringList.Create;
                     emailCCO := TStringList.Create;
                     try
                       emailCCO.Add('maria@presenta.com.br');
                       result := EmailDotNet.Enviar_Email(mailInfo.smtpObj.HostName,
                                                          mailinfo.smtpObj.PortNumber,
                                                          mailInfo.smtpObj.LoginName,
                                                          mailInfo.smtpObj.PassWord,
                                                          mailInfo.smtpObj.AddressName,
                                                          emailLst,
                                                          emailCC,
                                                          emailCCO,
                                                          'N',
                                                          myFiles,
                                                          qryMsgs.FieldByName('Assunto').AsString + ' - ' +
                                                          qryMsgs.FieldByName('COD_DEP_Remetente').AsString +
                                                          ' ('+ qryMsgs.FieldByName('Num_Sisbacen').AsString  + ' ' +	')',
                                                          myheader.text,
                                                          'N',
                                                          true,
                                                          'S',
                                                          vim830Global.totalCaixasPoremail,
                                                          Erro);
                     finally
                       emailCC.Free;
                       emailCCO.Free;
                     end;
                  finally
                      emailDotNet.Free
                  end;
                end;

                if result <> 1 then
                begin
                  ExibeMensagem( 'Erro SMTP: ' + erro);
                  Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro SMTP: ' + erro);
                  eMailLst.Free;
                  msgsSentTo.Free;
                  listEmail.free;
                  mytext.free;

                  // Alterado 22/12/2011
                  myHeader.free;
                  try
                    mailInfo.smtpObj.Disconnect;
                  except
                  end;

                  exit;
                end;

                mailInfo.smtpObj.Disconnect;
              end;
            end;

            myHeader.free;
  {$endif}
          end;
        end;
      end;

      if ultima then
      begin
        for i := 0 to msgsSentTo.Count - 1 do
        begin
          emailInfo := TEmailAddressInfo(msgsSentTo.objects[i]);
          if emailInfo.userCode <> '' then
          begin
            if emailInfo.notifica = 1 then
              NotifyReadMessage(qryMsgs.FieldByName('cod_msg').AsInteger,
                                qryMsgs.FieldByName('num_sisbacen').AsString,
                                StrToInt(emailInfo.userCode),
                                qryMsgs.FieldByName('assunto').AsString);
            MarcaEnviada(emailInfo.userCode, qryMsgs.FieldByName('cod_msg').AsString);
          end;
        end;
        passo:= 'Processo: Salvando mensagens enviadas';
        SaveMessageSent(qryMsgs);
  {$ifdef SUDAMERIS}
        passo:= 'Processo: Atualizando tabela do Sudameris';
        AtualizaIntranetSudameris(qryMsgs);
  {$endif}
      end;

      try
        for i := 0 to myFiles.Count - 1 do
        begin
          Application.ProcessMessages;

          if POS(diretorio, myfiles.Strings[i]) > 0 then
            DeleteFile(myFiles.Strings[i]);
        end;
      except
      end;

      mytext.free;
      myFiles.Free;
      msgsSentTo.free;
      emailLst.free;
      listEmail.free;
      result := 1;
    except
      // Alterado 08/05/2012
      on mErr:EErro do
      begin
        ExibeMensagem( '<ERRO> ' + merr.Mensagem);
        mytext.free;
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '<ERRO> ' + merr.Mensagem);
        TraceWFile(dmCorreio.traceFName, '<ERRO> ' + merr.Mensagem);
        result := -1;
        exit;
      end;

      on mErr:Exception do
      begin
        ExibeMensagem( '<ERRO> Exception:' + merr.message);
        mytext.free;
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '<ERRO> Exception:' + merr.message);
        TraceWFile(dmCorreio.traceFName, '<ERRO> Exception:' + merr.message);
        result := -1;
        exit;
      end;
    end;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dts2.Close;
      dmcorreio.dtsEmail.Close;
    except
    end;
  end;

  // Alterado 23/01/2012
  DisconnectEmail;
end;

function TReadBCCorreio.ExecutaEntregaDeMensagens: integer;
var
  msgsSentTo : TStringList;
  nRet : integer;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina ExecutaEntregaDeMensagens');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina ExecutaEntregaDeMensagens');

  result := 1;
  msgsSentTo := TStringList.Create;

  { insere mensagens novas no arquivo de dados }

  ExibeMensagem('Selecionando as mensagens para enviar');

  { localiza as mensagens que nao constam do arquivo do usuario}
  try
    try
      // Alterado 26/12/2011 - Safra
      if EhSafraNotes then
        dmCorreio.ConectorRobo.Obtem_Mensagens_Enviar_Units(dmCorreio.dtsMsg)
      else
        dmCorreio.ConectorRobo.Obtem_Mensagens_Enviar(dmCorreio.dtsMsg);

      while Not dmCorreio.dtsMsg.eof do
      begin
        if (pleaseClose) then
          exit;

        { a mensagem ja esta sendo enviada por outro servidor ? }
        if Not DirectoryExists(fdbDir + 'MSGS') then
        begin
          ExibeMensagem('MSGS - Diretorio de mensagens não encontrado');
          Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '[ExecutaEntregaDeMensagens] MSGS - Diretorio de mensagens não encontrado');
          exit;
        end;

        ExibeMensagem('Verificando entrega: ' + dmCorreio.dtsMsg.FieldByName('num_sisbacen').AsString + '-' + dmCorreio.dtsMsg.FieldByName('Assunto').AsString);
        nRet := EntregaEstaMensagem(dmcorreio.dtsMsg, '.TXT', true);
        if nRet <= 0 then
          break;
        dmCorreio.dtsMsg.next;
      end;
      dmCorreio.dtsMsg.Close;
      msgsSentTo.Free;
    except
      // Alterado 08/05/2012
      on mErr:EErro do
      begin
        ExibeMensagem('[ExecutaEntregaDeMensagens] ERRO: ' + merr.Mensagem);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '[ExecutaEntregaDeMensagens] ERRO: ' + merr.Mensagem);
        TraceWFile(dmCorreio.traceFName, '[ExecutaEntregaDeMensagens] ERRO: ' + merr.Mensagem);
        result := -1;
      end;

      on mErr:Exception do
      begin
        ExibeMensagem('[ExecutaEntregaDeMensagens] ERRO: ' + merr.Message);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '[ExecutaEntregaDeMensagens] ERRO: ' + merr.Message);
        TraceWFile(dmCorreio.traceFName, '[ExecutaEntregaDeMensagens] ERRO: ' + merr.Message);
        result := -1;
      end;
    end;
  // Alterado 19/06/2012
  finally
    try
      dmCorreio.dtsMsg.Close;
    except
    end;
  end;
end;

procedure TReadBCCorreio.Execute;
var
  ret: integer;
label
  exitThread,
  waitTime;

  procedure ExcluiArquivosLck;
  var
    sRec : TSearchRec;
  begin
    TraceWFile(dmCorreio.traceFName, 'entrou na rotina ExcluiArquivosLck');

    if FIsDebug then
      ExibeMensagem('Entrou na rotina ExcluiArquivoLck');

    ret := FindFirst(FdbDir + '*.LCK', faAnyFile, srec);
    while ret = 0 do
    begin
      Application.ProcessMessages;
      try
        DeleteFile(fdbDir + srec.name);
      except end;
      ret := FindNext(srec);
    end;
  end;
begin
  try
    try
      TraceWFile(dmCorreio.traceFName, 'Entrou na rotina Execute');

      if FIsDebug then
        ExibeMensagem('Entrou na rotina Execute [uReadBCCorreio]');

      if pleaseClose then
        exit;

      // Alterado 07/05/2012
      if LerWebPost then
        ExibeMensagem('Ativado modo LerWebPost');

      { Carregando as configurações regionais }
      decimalSeparator  := ',';
      thousandSeparator := '.';
      shortDateFormat   := 'dd/mm/yyyy';

      LoadVim830Global;

      // Alterado 26/12/2011 - Safra
      if EhSafraNotes then
      begin
        if not FileExists(FdbDir + 'LotusNotes.Ini') then
        begin
          ExibeMensagem('Não existe o arquivo LotusNotes.Ini');
          TraceWFile(dmCorreio.traceFName, 'Não existe o arquivo LotusNotes.Ini');
          Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Não existe o arquivo LotusNotes.Ini');
          exit;
        end;
      end;
      // Alterado 23/01/2012
      //else
      //  ConnectToEmail;

      { Verifica se o diretório DB é válido }
      if (FdbDir = '') or (Not DirectoryExists(FdbDir)) then
      begin
        TraceWFile(dmCorreio.traceFName, '<ERRO> Diretorio de dados inválido ' + FdbDir);
        ExibeMensagem('<ERRO> Diretorio de dados inválido ' + FdbDir);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '<ERRO> Diretorio de dados inválido ' + FdbDir);
        exit;
      end;

      TraceWFile(dmCorreio.traceFName, 'Verificando se está na hora de ser executado o programa');

      if FIsDebug then
        ExibeMensagem('Verificando se está na hora de ser executado o programa');

      { verifica se está na hora de ser executado o programa }
      if IsTimetoLogout then
      begin
        if Not informouDesativacao then
        begin
          TraceWFile(dmCorreio.traceFName, 'Atingido horario de desativacao');
          ExibeMensagem('Atingido horario de desativacao');
        end;
        informouDesativacao := true;
        goto waitTime;
      end;
      informouDesativacao := false;

      TraceWFile(dmCorreio.traceFName, 'Carregando as configurações do WebService');

      if FIsDebug then
        ExibeMensagem('Carregando as configurações do WebService');

      { carrega as configurações do WebService }
      // Alterado 07/05/2012
      if not LerWebPost then
      begin
        dmCorreio.ConectorRobo.Obtem_Configuracao_BCCorreio(dmcorreio.dts);
        BCCorreio.Dir_Arq_Anexo       := dmcorreio.dts.FieldByName('DIR_ARQUIVOS_ANEXOS').asString;
        BCCorreio.Url_WebService      := dmcorreio.dts.FieldByName('URL_WEBSERVICE').asString;
        BCCorreio.Url_WebService_WSDL := dmcorreio.dts.FieldByName('URL_WEBSERVICE_WSDL').asString;
        BCCorreio.Porta               := 'CorreioWSSoap12';
        BCCorreio.Servico             := 'CorreioWS';

        TraceWFile(dmCorreio.traceFName, 'Dir_Arq_Anexo = ' + dmcorreio.dts.FieldByName('DIR_ARQUIVOS_ANEXOS').asString);
        TraceWFile(dmCorreio.traceFName, 'Url_WebService = ' + dmcorreio.dts.FieldByName('URL_WEBSERVICE').asString);
        TraceWFile(dmCorreio.traceFName, 'Url_WebService_WSDL = ' + dmcorreio.dts.FieldByName('URL_WEBSERVICE_WSDL').asString);

        // Alterado 12/01/2012 - após distribuição para clientes
        ExibeMensagem('Dir_Arq_Anexo = ' + dmcorreio.dts.FieldByName('DIR_ARQUIVOS_ANEXOS').asString);
        ExibeMensagem('Url_WS = ' + dmcorreio.dts.FieldByName('URL_WEBSERVICE').asString);
        ExibeMensagem('Url_WS_WSDL = ' + dmcorreio.dts.FieldByName('URL_WEBSERVICE_WSDL').asString);

        dmcorreio.dts.Close;
      end;

      if pleaseClose then
        goto exitThread;

      ExcluiArquivosLck;

      TraceWFile(dmCorreio.traceFName, 'verifica se tem que remover as mensagens antigas');

      if FIsDebug then
        ExibeMensagem('verifica se tem que remover as mensagens antigas');

      { verifica se tem que remover as mensagens antigas }
      try
        if (vim830Global.lastDeleteDate <> FormatDateTime('dd/mm/yyyy', Now)) and
           (vim830Global.diasMantem <> 0) then
          RemoveOldMessages;
      except
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro ao verificar se tem que remover as mensagens antigas');
        TraceWFile(dmCorreio.traceFName, 'Erro ao verificar se tem que remover as mensagens antigas');
      end;

      if pleaseClose then
        goto exitThread;

      if dispatchMsg then
      begin
        try
          if ThereAreNewMessages then
          begin
            ret := ExecutaEntregaDeMensagens;
          end
          else
          begin
            ExibeMensagem('Não existe mensagem para enviar');
          end;
        except
          on Er: EErro do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on Er: EAviso do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on Er: EInformacao do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on Er: EErroLog do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on merr : exception do
          begin
            ExibeMensagem('Erro rotina de entrega: ' + merr.message);
            Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro rotina de entrega: ' + merr.message);
            TraceWFile(dmCorreio.traceFName, 'Erro rotina de entrega: ' + merr.message);
          end;
        end;

        if pleaseClose then
          goto exitThread;
      end;

      try
        TraceWFile(dmCorreio.traceFName, 'Vai iniciar pesquisas no Sisbacen');

        if FIsDebug then
          ExibeMensagem('Vai iniciar pesquisas no Sisbacen');

        ret := ReadMessagesFromBCCorreio;
      except
        on Er: EErro do
          ExibeMensagem('Erro rotina de captura: ' + er.mensagem);
        on Er: EAviso do
          ExibeMensagem('Erro rotina de captura: ' + er.mensagem);
        on Er: EInformacao do
          ExibeMensagem('Erro rotina de captura: ' + er.mensagem);
        on Er: EErroLog do
          ExibeMensagem('Erro rotina de captura: ' + er.mensagem);
        on merr : exception do
        begin
          ExibeMensagem('Erro rotina de captura: ' + merr.message);
          Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro rotina de captura: ' + merr.message);
          TraceWFile(dmCorreio.traceFName, 'Erro rotina de captura: ' + merr.message);
        end;
      end;

      if pleaseClose then
        goto exitThread;

      if dispatchMsg then
      begin
        try
          if ThereAreNewMessages then
          begin
            ExibeMensagem('Existe mensagem para enviar');
            ret := ExecutaEntregaDeMensagens;
          end;
        except
          on Er: EErro do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on Er: EAviso do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on Er: EInformacao do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on Er: EErroLog do
            ExibeMensagem('Erro rotina de entrega: ' + er.mensagem);
          on merr : exception do
          begin
            ExibeMensagem('Erro rotina de entrega: ' + merr.message);
            Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro rotina de entrega: ' + merr.message);
            TraceWFile(dmCorreio.traceFName, 'Erro rotina de entrega: ' + merr.message);
          end;
        end;

        if pleaseClose then
          goto exitThread;
      end;

exitThread:

      TraceWFile(dmCorreio.traceFName, 'Vai desconectar emulador e banco de dados');
      DoShowProxProc;

waitTime:

    except
      // Alterado 08/05/2012
      on Er: EErro do
      begin
        ExibeMensagem('Erro: ' + Er.Mensagem);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Er.Mensagem);
        TraceWFile(dmCorreio.traceFName, '[Execute] ' + er.Mensagem);
      end;

      on Er: Exception do
      begin
        ExibeMensagem('Erro: ' + Er.Message);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Er.Message);
        TraceWFile(dmCorreio.traceFName, '[Execute] ' + er.message);
      end;
    end;
  finally
    // Alterado 19/06/2012
    try
      dmCorreio.dts.Close;
      dmCorreio.dts2.Close;
      dmCorreio.dtsVM_Units.Close;
      dmCorreio.dtsMsg.Close;
      dmCorreio.dtsEmail.Close;
    except
    end;
  end;
end;

procedure TReadBCCorreio.ExibeMensagem(Linha: String);
var
  F: TextFile;
  Texto: String;
begin
  Texto :=  '[' + FormatDateTime('hh:mm:ss', now) + '] ' + Linha;

  if FIsDebug and FSalvarArquivo then
  begin
    AssignFile(F, FdirArqDebug + 'Log_Robo_Correio_' + FormatDateTime('yyyymmdd', date) + '.log');

    if FileExists(FdirArqDebug + 'Log_Robo_Correio_' + FormatDateTime('yyyymmdd', date) + '.log') then
      Append(F)
    else
      ReWrite(F);

    Writeln(F, Texto);
    CloseFile(F);
  end;

  if FmemMsg = nil then
    exit;

  While FmemMsg.Lines.Count > 20 do
  begin
    Application.ProcessMessages;
    FmemMsg.Lines.Delete(0);
  end;

  FmemMsg.Lines.Add(Texto);
end;

procedure TReadBCCorreio.GetUnitsInterval(cod_inst, cod_dep, transacao: String;
  var lastTime, minutes: TDateTime; var dias: integer);
var
  resto : integer;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina GetUnitsInterval');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina GetUnitsInterval');

  if dmCorreio.dtsVM_Units.FieldByName('LAST_TIME_CHECKED').AsString = '' then
    lastTime := 0
  else
  begin
    try
      lastTime := StrToDateTime(dmCorreio.dtsVM_Units.FieldByName('LAST_TIME_CHECKED').AsString);
    except
      lastTime := 0;
    end;
  end;

  if (Trim(dmCorreio.dtsVM_Units.FieldByName('HORA_LOGON').AsString) <> '') and
     (Trim(dmCorreio.dtsVM_Units.FieldByName('HORA_LOGOFF').AsString) <> '')then
  begin
    if (FormatDateTime('hh:nn', Now) < dmCorreio.dtsVM_Units.FieldByName('HORA_LOGON').AsString) or
       (FormatDateTime('hh:nn', Now) > dmCorreio.dtsVM_Units.FieldByName('HORA_LOGOFF').AsString) then begin
      lastTime := Now;
      minutes := -1;
      exit;
    end;
  end;

  if dmCorreio.dtsVM_Units.FieldByName('Minutos').AsInteger >= 60 then
  begin
    if dmCorreio.dtsVM_Units.FieldByName('Minutos').AsInteger = 60 then
      minutes := EncodeTime(1, 0, 0, 0)
    else
    begin
      resto := dmCorreio.dtsVM_Units.FieldByName('Minutos').AsInteger div 60;
      if resto >= 24 then
        dias := resto div 24
      else
        minutes := EncodeTime(resto, dmCorreio.dtsVM_Units.FieldByName('Minutos').AsInteger mod 60, 0,0);
    end;
  end
  else
    minutes := EncodeTime(0, dmCorreio.dtsVM_Units.FieldByName('Minutos').AsInteger, 0, 0);

  TraceWFile(dmCorreio.traceFName, 'Saiu da rotina GetUnitsInterval');
end;

procedure TReadBCCorreio.InsertMessageHeader(posicao: integer);
Var
	lida : boolean;
  codigo: integer;
begin
  try
    try
      TraceWFile(dmCorreio.traceFName, 'Vai verificar se precisa inserir dados na tabela VM_MESSAGES - ' + IntToStr(Mensagens[posicao].NumeroCorreio));

      if FIsDebug then
        ExibeMensagem('Vai verificar se precisa inserir dados na tabela VM_MESSAGES - ' + IntToStr(Mensagens[posicao].NumeroCorreio));

      { Verificar se a mensagem já foi inserida no banco de dados e foi cancelada }
      lida := false;

      { Verifica se ja esta lida }
      dmCorreio.ConectorRobo.Obtem_VM_Messages(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                               dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                               IntToStr(Mensagens[posicao].NumeroCorreio),
                                               checkMsgInst,
                                               dmCorreio.dts);

      if Not dmCorreio.dts.Eof then
      begin
        // Alterado 19/06/2012
//?????Maria        dmCorreio.dts.Close;
        if Not checkMsgInst then
        begin
          if (dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString <> dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString) Or
             (dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString <> dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString) then
          begin
            // Alterado 26/12/2011 - Safra
            codigo := dmCorreio.dts.FieldByName('COD_MSG').AsInteger;

            dmCorreio.ConectorRobo.Obtem_VM_MESSAGES_UNITS(codigo,
                                                           dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                           dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                           dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,
                                                           dmCorreio.dts2);

            if dmCorreio.dts2.RecordCount = 0 then
              dmCorreio.ConectorRobo.Insert_VM_MESSAGES_UNITS(codigo,
                                                              dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                              dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                              dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString);

            exit;
          end;
        end;
        lida := true;
      end;

      if lida then
      begin
        if (POS('CANCEL', Mensagens[posicao].Status) <> 0) and
           (POS('CANCEL', dmCorreio.dts.FieldByName('STATUS_SISBACEN').AsString) = 0) And
           (POS('DENOR', dmCorreio.dts.FieldByName('STATUS_SISBACEN').AsString) = 0) then
        begin
          dmCorreio.ConectorRobo.Atualiza_Status_Mensagem(dmCorreio.dts.FieldByName('NUM_SISBACEN').AsString,
                                                          Mensagens[posicao].Status);
        end;

        // Alterado 26/12/2011 - Safra
        codigo := dmCorreio.dts.FieldByName('COD_MSG').AsInteger;

        // Alterado 19/06/2012
        dmCorreio.dts.Close;

        dmCorreio.ConectorRobo.Obtem_VM_MESSAGES_UNITS(codigo,
                                                       dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                       dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                       dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,
                                                       dmCorreio.dts);

        if dmCorreio.dts.RecordCount = 0 then
          dmCorreio.ConectorRobo.Insert_VM_MESSAGES_UNITS(codigo,
                                                          dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                          dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                          dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString);

        exit;
      end;

      if not dmCorreio.dtsVM_Units.FieldByName('PRIORIDADE').isnull then
      begin
        if dmCorreio.dtsVM_Units.FieldByName('PRIORIDADE').AsString = '0' then
          AvisaQueMensagemChegou(IntToStr(Mensagens[posicao].NumeroCorreio), Mensagens[posicao].Assunto, Mensagens[posicao].UnidadeRemetente);
      end
      else
      begin
        if PrioridadeZero(Mensagens[0].UnidadeRemetente, Mensagens[0].Assunto) then
        begin
          AvisaQueMensagemChegou(IntToStr(Mensagens[posicao].NumeroCorreio), Mensagens[posicao].Assunto, Mensagens[posicao].UnidadeRemetente);
        end;
      end;

      { Inserir novo registro para a mensagem }
      dmcorreio.ConectorRobo.Insere_Mensagem(Mensagens[posicao].Assunto,                                  // ASSUNTO
                                             Mensagens[posicao].UnidadeRemetente,                         // COD_DEP_REMETENTE,
                                             IntToStr(Mensagens[posicao].NumeroCorreio),                  // NUM_SISBACEN,
                                             FormatDateTime('dd/mm/yyyy', Mensagens[posicao].Data),       // DATA_ENVIO,
                                             FormatDateTime('hh:nn', Mensagens[posicao].Hora),            // HORA_ENVIO,
                                             Mensagens[posicao].Status,                                   // STATUS_SISBACEN,
                                             'N',                                                         // LIDO,
                                             dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,      // COD_INST,
                                             dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,       // COD_DEP,
                                             FormatDateTime('yyyy/mm/dd hh:nn', Mensagens[posicao].Data + Mensagens[posicao].Hora), // AMD,
                                             dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,     // TRANSACAO,
                                             FormatDateTime('dd/mm/yyyy', Date),                          // DT_RECEBIMENTO,
                                             FormatDateTime('hh:nn', Time));                              // HR_RECEBIMENTO

      // Alterado 19/06/2012
      dmCorreio.dts.Close;

      // Alterado 26/12/2011 - Safra
      dmCorreio.ConectorRobo.Obtem_VM_Messages(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                               dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                               IntToStr(Mensagens[posicao].NumeroCorreio),
                                               true,
                                               dmCorreio.dts);

      codigo := dmCorreio.dts.FieldByName('COD_MSG').AsInteger;

      // Alterado 19/06/2012
      dmCorreio.dts.Close;

      dmCorreio.ConectorRobo.Obtem_VM_MESSAGES_UNITS(codigo,
                                                     dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                     dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                     dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,
                                                     dmCorreio.dts);

      if dmCorreio.dts.RecordCount = 0 then
        dmCorreio.ConectorRobo.Insert_VM_MESSAGES_UNITS(codigo,
                                                        dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                        dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                        dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString);
    except
      on Er: EErro do
        ExibeMensagem('<InsertMessageHeader> Erro: ' + er.Mensagem);
      on Er: EAviso do
        ExibeMensagem('<InsertMessageHeader> Erro: ' + er.Mensagem);
      on Er: EInformacao do
        ExibeMensagem('<InsertMessageHeader> Erro: ' + er.Mensagem);
      on Er: EErroLog do
        ExibeMensagem('<InsertMessageHeader> Erro: ' + er.Mensagem);
      on mErr: Exception do
      begin
        ExibeMensagem('<InsertMessageHeader> Erro: ' + merr.Message);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '<InsertMessageHeader> Erro: ' + merr.Message);
        TraceWFile(dmCorreio.traceFName, '<InsertMessageHeader> Erro: ' + merr.Message);
      end;
    end;
  // Alterado 19/06/2012
  finally
    try
      dmCorreio.dts.Close;
    except
    end;
  end;
end;

function TReadBCCorreio.IsTimeToLogout: boolean;
var
  agora: TdateTime;
  dayWeek: integer;
  timeNow: string;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina IsTimeToLogout');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina IsTimeToLogout');

  result := true;

  agora := Now;
  dayWeek := DayOfWeek(agora);

  if Not FrunOnWeekends then
  begin
    if (dayWeek = 1) or (dayWeek = 7) then
      exit;
  end;

  timeNow := FormatDateTime('hh:nn', Now);

  if (timeNow >= FHora_Logon) and (timeNow <= FHora_Logout) then
  begin
    result := false;
  end;
end;

function TReadBCCorreio.JaEnviou(codUsu, codMsg: string): boolean;
begin
  result := dmCorreio.ConectorRobo.Mensagem_Foi_Enviada(StrToInt(codUsu), StrToInt(codMsg));
end;

procedure TReadBCCorreio.LoadVim830Global;
var
	ini : TiniFile;
  str: string;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina configurações');

  if FIsDebug then
    ExibeMensagem('entrou na rotina LoadVim830Global');

  vim830Global.smtpInfo    := TSmtpStruct.Create;
  vim830Global.outlookInfo := TOutlookStruct.Create;

  ini := TIniFile.Create(FdbDir + 'VIM830.INI');
  vim830Global.sisbacenTimer   := ini.ReadInteger(SECTION_GLOBAL, ENTRY_INTERVALO, 30);
  vim830Global.buscaAutomatica := ini.ReadBool(SECTION_GLOBAL, ENTRY_BUSCA, True);
  vim830Global.diasMantem      := ini.ReadInteger(SECTION_GLOBAL, ENTRY_DIASMSGS, 0);
  vim830Global.emailDotNet     := ini.ReadBool(SECTION_GLOBAL, ENTRY_EMAIL_DOT_NET, False);
  vim830Global.tratarOficio    := ini.ReadBool(SECTION_GLOBAL, ENTRY_TRATAR_OFICIO, True);
  vim830Global.totalCaixasPoremail := ini.ReadInteger(SECTION_GLOBAL, ENTRY_TOTAL_CXS_EMAIL, 999);

  if vim830Global.diasMantem = 0 then
    vim830Global.dataLimite := StrToDateTime('01/01/1980')
  else
    vim830Global.dataLimite := DecrementData(date, vim830Global.diasMantem, true);

  vim830Global.DataLimiteAte := StrToDateTime(ini.ReadString('System', 'DateTo', FormatDateTime('dd/mm/yyyy', Date)));

  if ini.ReadString('SYSTEM', 'DateFrom', '') <> '' then
    vim830Global.dateInstall := StrToDateTime(ini.ReadString('SYSTEM', 'DateFrom', ''))
  else
  begin
    //vim830Global.dateInstall  := StrToDateTime(ini.ReadString('SYSTEM', 'DateInstall', FormatDateTime('dd/mm/yyyy', DecrementData(Date, 3, true))));
    str := ini.ReadString('SYSTEM', 'DateInstall', FormatDateTime('dd/mm/yyyy', DecrementData(Date, 3, true)));
    if str = '' then
      str := FormatDateTime('dd/mm/yyyy', DecrementData(Date, 3, true));
    vim830Global.dateInstall  := StrToDateTime(str);
  end;
  vim830Global.readDenor      := ini.ReadBool(SECTION_GLOBAL, ENTRY_PPAC300, False);
  vim830Global.readPtax       := ini.ReadBool(SECTION_GLOBAL, ENTRY_PTAX, False);
  vim830Global.ExcluiPedido   := ini.ReadBool(SECTION_GLOBAL, ENTRY_EXCLUIPEDIDO, true);
  vim830Global.filtroTextoAnd := ini.ReadBool(SECTION_GLOBAL, 'FiltroTextoAnd', false);

  vim830Global.notifica         := ini.ReadBool(SECTION_GLOBAL, 'NotificaUsuario', False);
  vim830Global.trocaSenha       := ini.ReadBool(SECTION_GLOBAL, ENTRY_AUTOSENHA, false);
  vim830Global.pwdDir           := ini.ReadString(SECTION_GLOBAL, ENTRY_AUTOSENHA_DIR, '');
  vim830Global.lastDeleteDate   := ini.ReadString(SECTION_GLOBAL, ENTRY_LAST_DELETE, '');
  vim830Global.DirecionaExcecao := ini.ReadString(SECTION_GLOBAL, 'DirecionarExcecoesPara', '');

  // Alterado 26/12/2011 - Safra
  EhSafraNotes                  := ini.ReadBool('SYSTEM', 'EH_SAFRA', False);

  dmCorreio.Conector.Obtem_Configuracao_Email(dmCorreio.dts);

  if dmCorreio.dts.RecordCount = 0 then
    vim830Global.mailSystem := IDSM_SYS_UNKNOWN
  else
    vim830Global.mailSystem := dmCorreio.dts.FieldByName('SISTEMA').AsInteger;

  if vim830Global.mailSystem = IDSM_SYS_SMTP_POP then
  begin
    vim830Global.smtpInfo.SmtpServer  := dmCorreio.dts.FieldByName('servidor').AsString;
    vim830Global.LoginName            := dmCorreio.dts.FieldByName('usuario').AsString;
    vim830Global.Password             := DescriptoGrafaTextoNum('',dmCorreio.dts.FieldByName('senha').AsString);
    vim830Global.smtpInfo.SmtpPort    := dmCorreio.dts.FieldByName('porta').AsString;
    vim830Global.smtpInfo.addressName := dmCorreio.dts.FieldByName('nome_postal').AsString;
  end;

  if (vim830Global.mailSystem = IDSM_SYS_VIM) or (vim830Global.mailSystem = IDSM_SYS_NOTES)then
  begin
    vim830Global.loginName := dmCorreio.dts.FieldByName('usuario').AsString;
    vim830Global.Password  := DescriptoGrafaTextoNum('',dmCorreio.dts.FieldByName('senha').AsString);
  end;

  if vim830Global.mailSystem = IDSM_SYS_OUTLOOK then
  begin
    vim830Global.outlookInfo.publicFolderName := ini.ReadString (SECTION_OUTLOOK_97, ENTRY_OUTLOOK_PUBLIC_FOLDER, 'Pastas Compartilhadas do Microsoft Mail');
    vim830Global.outlookInfo.targetFolderName := ini.ReadString (SECTION_OUTLOOK_97, ENTRY_OUTLOOK_FOLDER_NAME, 'Sisbacen');
    vim830Global.outlookInfo.templateName     := ini.ReadString (SECTION_OUTLOOK_97, ENTRY_OUTLOOK_FORM_NAME, 'C:\Arquivos de Programas\Microsoft Office\Modelos\Outlook\sisbacenform.oft');
  end;
  ini.free;

  if IsParam('/semlimpeza') then
    vim830Global.lastDeleteDate := FormatDateTime('dd/mm/yyyy', Now);
end;

function TReadBCCorreio.LogonToBCCorreio: Integer;
var
  bloqueado : integer;
  erro: string;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina LogonToBCCorreio');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina LogonToBCCorreio');

  result := -1;

  // verifica se esta bloqueado
  bloqueado := dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger;
  // Alterado 26/03/2012
  //if bloqueado >= 3 then
  if bloqueado >= 2 then
  begin
    Erro := '<ERRO DE LOGON>' + dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' +
                                dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '/' +
                                dmcorreio.dtsVM_Units.FieldByName('OPERADOR').AsString + ' - ' +
            'O sistema não vai tentar logar nesta instituição até ser trocada a senha.';

    Insere_Log(ERRO_LOGON_SISBACEN, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);

    dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                          dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                          'S',
                                                          0);

    exit;
  end;

  // Alterado 26/03/2012
  dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                        dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                        'N',
                                                        0);

  // Alterado 11/01/2012
  BCCorreio.Proxy       := FProxy;
  BCCorreio.UserProxy   := FUserProxy;
  BCCorreio.SenhaProxy  := FSenhaProxy;

  if BCCorreio.ConectaWS(Erro) <> 1 then
  begin
    Insere_Log(ERRO_LOGON_SISBACEN, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);

    if POS('(401)', Erro) > 0 then
    begin
      // Alterado 26/03/2012
      if dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger < 2 then
        dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                              dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                              'N',
                                                              dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger + 1)
      else
      dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                            dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                            'S',
                                                            0);
    end;

    exit;
  end;

  // Alterado 26/03/2012
  // Alterado 13/04/2012
  //dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
  //                                                      dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
  //                                                      'N',
  //                                                      0);

  result := 1;
end;

procedure TReadBCCorreio.MarcaEnviada(codUsu, codMsg: string);
var
  num, assunto, perfil: string;
begin
  if Not JaEnviou(codUsu, codMsg) then
    dmCorreio.ConectorRobo.Insere_Msg_Entregue(StrToInt(codUsu), StrToInt(codMsg));

  try
    dmCorreio.ConectorRobo.Obtem_VM_Messages(StrToInt(codMsg), dmCorreio.dts2);
    num := dmCorreio.dts2.FieldByName('NUM_SISBACEN').AsString;
    Assunto := dmCorreio.dts2.FieldByName('ASSUNTO').AsString;
    dmCorreio.dts2.Close;

    dmCorreio.ConectorRobo.Obtem_VM_Users(StrToInt(CodUsu), dmCorreio.dts2);
    perfil := dmCorreio.dts2.FieldByName('NOME_USUARIO').AsString;
    dmCorreio.dts2.Close;

    Insere_Log(ENVIAR_EMAIL, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO,
               ' Mensagem: ' + num + '-' + Assunto + ' - para o Perfil: ' + perfil);
  except
    // Alterado 08/05/2012
    on Er:EErro do
    begin
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na MarcaEnviada - ' + Er.Mensagem);
      TraceWFile(dmCorreio.traceFName, 'Erro na MarcaEnviada - ' + Er.Mensagem);
    end;

    on Er:Exception do
    begin
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na MarcaEnviada - ' + Er.Message);
      TraceWFile(dmCorreio.traceFName, 'Erro na MarcaEnviada - ' + Er.Message);
    end;
  end;
end;

function TReadBCCorreio.MessageAppliesUserFilter(qryMsgs: TADODataSet; userCode,
  fileExt: string; msgText: TStringList): boolean;
var
  tableName : string;
  checaCaixas : boolean;
  j : integer;
  achou : boolean;
label
  checkBoxes;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina MessageAppliesUserFilter');

  try
    if FIsDebug then
      ExibeMensagem('Entrou na rotina MessageAppliesUserFilter');

    result := true;
    tableName := tabVM_USERS;
    checaCaixas := false;

    dmCorreio.ConectorRobo.Obtem_Filtro_Distrib(StrToInt(userCode),
                                                ENTRY_CAIXAS,
                                                1,
                                                dmcorreio.dts2);

    if Not dmcorreio.dts2.eof then
      checaCaixas := true;

    { Filtro para Contém Texto }
    achou := false;

    // Alterado 19/06/2012
    dmcorreio.dts2.Close;

    dmCorreio.ConectorRobo.Obtem_Filtro_Distrib(StrToInt(userCode),
                                                ENTRY_MESSAGE_TEXTO,
                                                2,
                                                dmcorreio.dts2);

    if Not dmcorreio.dts2.Eof then
    begin
      while Not dmcorreio.dts2.Eof do
      begin
        for j:=0 to msgText.count - 1 do
        begin
          if Pos (UpperCase(dmcorreio.dts2.FieldByName('filtro').AsString), UpperCase(msgText.strings[j])) <> 0 then
            break;
        end;

        if j <= (msgText.count-1) then
        begin
          achou := true;
          break;
        end;
        dmcorreio.dts2.Next;
      end;

      if achou then
      begin
        result := true;
        if checaCaixas then
          goto checkBoxes;

        if Not vim830Global.filtroTextoAnd then
          exit;
      end
      else
      begin
        if (vim830Global.filtroTextoAnd) and (Not achou) then
        begin
          result := false;
          exit;
        end;
      end;
    end;

    // Alterado 19/06/2012
    dmcorreio.dts2.Close;

    { filtro para departamentos }
    achou := false;
    dmCorreio.ConectorRobo.Obtem_Filtro_Distrib(StrToInt(userCode),
                                                ENTRY_DEPTOS,
                                                2,
                                                dmcorreio.dts2);
    if Not dmcorreio.dts2.Eof then
    begin
      while Not dmcorreio.dts2.Eof do
      begin
        if Pos (UpperCase(Trim(dmcorreio.dts2.FieldByName('filtro').AsString)), UpperCase(qryMsgs.FieldByName('Cod_dep_Remetente').AsString)) <> 0 then
        begin
          achou := true;
          break;
        end;
        dmcorreio.dts2.Next;
      end;

      if (Not achou) then
      begin
        result := false;
        exit;
      end;
    end;

    // Alterado 19/06/2012
    dmcorreio.dts2.Close;

    { filtro para Comtém Assunto }
    achou := false;
    dmCorreio.ConectorRobo.Obtem_Filtro_Distrib(StrToInt(userCode),
                                                ENTRY_GOOD_WORDS,
                                                2,
                                                dmcorreio.dts2);
    if Not dmcorreio.dts2.Eof then
    begin
      while Not dmcorreio.dts2.Eof do
      begin
        if Pos (UpperCase(dmcorreio.dts2.FieldByName('filtro').AsString), UpperCase(qryMsgs.FieldByName('ASSUNTO').AsString)) <> 0 then
        begin
          achou := true;
          break;
        end;
        dmcorreio.dts2.Next;
      end;

      if (Not achou) then
      begin
        result := false;
        exit;
      end;
    end;

    // Alterado 19/06/2012
    dmcorreio.dts2.Close;

    { filtro para não Contém Assunto }
    achou := false;
    dmCorreio.ConectorRobo.Obtem_Filtro_Distrib(StrToInt(userCode),
                                                ENTRY_BAD_WORDS,
                                                2,
                                                dmcorreio.dts2);
    if Not dmcorreio.dts2.Eof then
    begin
      while Not dmcorreio.dts2.Eof do
      begin
        if Pos (UpperCase(dmcorreio.dts2.FieldByName('filtro').AsString), UpperCase(qryMsgs.FieldByName('ASSUNTO').AsString)) <> 0 then
        begin
          achou := true;
          break;
        end;
        dmcorreio.dts2.Next;
      end;

      if (achou) then
      begin
        result := false;
        exit;
      end;
    end;

    // Alterado 19/06/2012
    dmcorreio.dts2.Close;

  checkBoxes:

    if checaCaixas then
    begin
      if result = false then
        exit;

      dmCorreio.ConectorRobo.Obtem_Filtro_Distrib(StrToInt(userCode),
                                                  ENTRY_CAIXAS,
                                                  1,
                                                  dmcorreio.dts2);
      if Not dmcorreio.dts2.eof then
      begin
        while not dmcorreio.dts2.Eof do
        begin
          if UpperCase(dmcorreio.dts2.FieldByName('filtro').AsString) = qryMsgs.FieldByName('cod_inst').AsString + qryMsgs.FieldByName('cod_dep').AsString then
          begin
            result := true;
            exit;
          end;
          dmcorreio.dts2.next;
        end;

        // Alterado 19/06/2012
        dmcorreio.dts2.Close;

        result := false;
        exit;
      end;

      // Alterado 19/06/2012
      dmcorreio.dts2.Close;
    end;
    result := true;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dts2.Close;
    except
    end;
  end;
end;

procedure TReadBCCorreio.NotifyReadMessage(codmsg: integer; msgNumber: string;
  codUsu: integer; assunto: string);
var
	i : integer;
	address : string;
	ini : TIniFile;
	pagina : string;
	msgTexto, msgAssunto  : string;
	mText : TStringList;
	sValues : TStringList;
	eProg : PAnsiChar;
	modelo : string;
	ret : integer;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina NotifyReadMessage');

  try
    if FIsDebug then
      ExibeMensagem('Entrou na rotina NotifyReadMessage');

    // Alterado 23/01/2012
    ConnectToEmail;

    try
      { localiza os usuários que tem email setado }
      dmcorreio.ConectorRobo.Obtem_Usu_Email_Setado(codUsu, dmCorreio.dts2);

      while not dmCorreio.dts2.eof do
      begin
        Application.ProcessMessages;
        { verifica se a mensagem se aplica ao filtro do usuário }
        if Trim(dmCorreio.dts2.FieldByName('EMAIL_ADDRESS').AsString) = '' then
        begin
          dmCorreio.dts2.next;
          continue;
        end;

        ExibeMensagem('Notificando: ' + dmCorreio.dts2.FieldByName('EMAIL_ADDRESS').AsString);

        address := dmCorreio.dts2.FieldByName('EMAIL_ADDRESS').AsString;
        msgAssunto := 'Mensagem nova do Sisbacen';
        msgTexto := 'Acaba de ser lida no Sisbacen a mensagem: ';
        msgTexto := msgTexto + #13 + #10 + #13 + #10;
        msgTexto := msgTexto + msgNumber + '   ' +
                    dmCorreio.dts2.FieldByName('COD_INST').Value + '/' +
                    dmCorreio.dts2.FieldByName('COD_DEP').Value + '   ' +
                    assunto;

        if vim830Global.mailSystem = IDSM_SYS_OUTLOOK then
        begin
          try
            mailInfo.mailFolder := mailInfo.nameSpace.GetDefaultFolder(4);
            mailInfo.mailItem := mailInfo.mailsys.CreateItem(0);
            WaitSeconds(1);
            mailInfo.mailitem.subject := msgAssunto;
              i := Pos (';', address);
            while i <> 0 do
            begin
              Application.ProcessMessages;
              mailInfo.mailitem.recipients.add(Trim(Copy(address, 1, i-1)));
              address := Copy(address, i + 1, Length(address));
              i := Pos (';', address);
            end;
            if address <> '' then
              mailInfo.mailitem.recipients.add(Trim(address));
            mailInfo.mailItem.body := msgTexto;
            mailInfo.mailItem.send;
          except
          end;
          dmCorreio.dts2.next;
          continue;
        end;

        if vim830Global.mailSystem = 0 then
        begin
          i := Pos (';', address);
          while i <> 0 do
          begin
            Application.ProcessMessages;
            StrPcopy(eProg, 'NET SEND ' + Trim(Copy(address, 1, i-1)) + ' ' + '[Robot_Correio]' + msgTexto);
            WinExec(eProg, SW_HIDE);
            address := Copy(address, i + 1, Length(address));
            i := Pos (';', address);
          end;

          if Trim(address) <> '' then
          begin
            StrPcopy(eProg, 'NET SEND ' + Trim(address) + ' ' + '[Robot_Correio]' + msgTexto);
            WinExec(eProg, SW_HIDE);
          end;
          dmCorreio.dts2.next;
          continue;
        end;

        if (vim830Global.mailSystem <> IDSM_SYS_SMTP_POP) then
        begin
          try
            if not mailinfo.maillogged then
            begin
              if Trim(vim830Global.LoginName) <> '' then
              begin
                mailInfo.mailSys.LoginName := vim830Global.LoginName;
                mailInfo.mailSys.Password := vim830Global.Password;
                mailInfo.mailSys.login;
              end;

              if vim830Global.mailSystem = IDSM_SYS_MAPI then
              begin
                modelo := mailInfo.mailsys.DefaultProfile;
                mailinfo.mailsys.LoginName := modelo;
                mailinfo.mailsys.Login;
              end;
            end;

            mailInfo.mailSys.NewMessage;

            if vim830Global.mailSystem = IDSM_SYS_SMTP_POP then
              mailInfo.mailSys.From := vim830global.smtpInfo.addressName;

            i := Pos (';', address);
            while i <> 0 do
            begin
              Application.ProcessMessages;
              mailInfo.mailSys.AddRecipientTo(Trim(Copy(address, 1, i-1)));
              address := Copy(address, i + 1, Length(address));
              i := Pos (';', address);
            end;

            if address <> '' then
              mailInfo.mailSys.AddRecipientTo(Trim(address));
            mailInfo.mailSys.RecipientDelimiter := ';';
            mailInfo.mailSys.subject := msgAssunto;
            mailInfo.mailSys.message := msgTexto;
            if (vim830Global.mailSystem = IDSM_SYS_VIM) OR (vim830Global.mailSystem = IDSM_SYS_NOTES) then
              mailinfo.mailSys.Signed := True;
            mailInfo.mailSys.send;
          except
          end;
          dmCorreio.dts2.next;
          continue;
        end;

        if vim830Global.mailSystem = IDSM_SYS_SMTP_POP then
        begin
          try
            ret := mailInfo.SmtpObj.Activate(vim830Global.LoginName, vim830Global.Password, vim830Global.smtpInfo.addressName);
            if ret <= 0 then
              exit;
              mText := TStringList.Create;
            mText.Add(msgTexto);
            sValues := TStringList.Create;
            StrToList(address, ';', svalues);
            ret := mailInfo.smtpObj.SendMail(msgAssunto, sValues, Nil, mText);
            mailInfo.smtpObj.Disconnect;
            FreeAndNil(mtext);
            FreeAndNil(sValues);
          except
          end;
          dmCorreio.dts2.next;
          continue;
        end;
      end;

      // Alterado 19/06/2012
      dmCorreio.dts2.Close;
    except
    end;

    // Alterado 23/01/2012
    DisconnectEmail;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dts2.Close;
    except
    end;
  end;
end;

function TReadBCCorreio.PrioridadeZero(sender, subject: string): boolean;
var
  i: integer;
  result1, result2, result3: boolean;
  sValue : TStringList;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina PrioridadeZero');

  try
    if FIsDebug then
      ExibeMensagem('Entrou na rotina PrioridadeZero');

    result := true;
    result1 := false;
    result2 := false;
    result3 := true;

    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               10,
                                               false,
                                               dmCorreio.dts);

    if dmCorreio.dts.Eof then
    begin
      result := false;
      exit;
    end;

    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               ENTRY_DEPTOS_ZERO,
                                               true,
                                               dmCorreio.dts);

    sValue := TstringList.Create;
    while not dmCorreio.dts.eof do
    begin
      Application.ProcessMessages;
      sValue.Add(dmCorreio.dts.FieldByName('FILTRO').AsString);
      dmCorreio.dts.Next;
    end;

    // Alterado 19/06/2012
    dmCorreio.dts.Close;

    for i := 0 to sValue.Count - 1 do
      if sender = sValue.Strings[i] then
      begin
        result1 := true;
        break;
      end;
    sValue.Free;

    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               ENTRY_GOOD_WORDS_ZERO,
                                               true,
                                               dmCorreio.dts);

    sValue := TstringList.Create;
    while not dmCorreio.dts.eof do
    begin
      Application.ProcessMessages;
      sValue.Add(dmCorreio.dts.FieldByName('FILTRO').AsString);
      dmCorreio.dts.Next;
    end;

    // Alterado 19/06/2012
    dmCorreio.dts.Close;

    for i := 0 to sValue.Count - 1 do
      if Pos(sValue.Strings[i], subject) > 0 then
      begin
        result2 := true;
        break;
      end;
    sValue.Free;

    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               ENTRY_BAD_WORDS_ZERO,
                                               true,
                                               dmCorreio.dts);

    sValue := TstringList.Create;
    while not dmCorreio.dts.eof do
    begin
      Application.ProcessMessages;
      sValue.Add(dmCorreio.dts.FieldByName('FILTRO').AsString);
      dmCorreio.dts.Next;
    end;

    // Alterado 19/06/2012
    dmCorreio.dts.Close;

    if sValue.Count <= 0 then
      result3 := false
    else
    begin
      for i := 0 to sValue.Count - 1 do
      begin
        if Pos(sValue.Strings[i], subject) > 0 then
        begin
          result3 := false;
          break;
        end;
      end;
    end;
    sValue.Free;

    if (Not result1) and (Not result2) and (Not result3) then
      result := false;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dts.Close;
    except
    end;
  end;
end;

function TReadBCCorreio.ReadMailList: integer;
Var
  nRet : integer;
  dma : string;
  ini : TiniFile;
  Erro: String;
  i: integer;
  localizou: boolean;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina ReadMailList');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina ReadMailList');

  result := -1;

  if pleaseClose then
    exit;

  ExibeMensagem(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' +
                dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '/' +
                dmCorreio.dtsVM_Units.FieldByName('Operador').AsString + ' - Checando lista de Mensagens');

  { Captura a DateFrom }
  ini := TIniFile.Create(FdbDir + 'VIM830.INI');
  dma := ini.ReadString('System', 'DateFrom', '');
  ini.free;
  if dma = '' then
  begin
    // Alterado 02/04/2012
    //dma := FormatDateTime('dd/mm/yyyy', Date - 10);
    // Alterado 04/06/2012
    //dma := FormatDateTime('dd/mm/yyyy', Date - 2);
    Case DayOfWeek(Date) of
      2: dma := FormatDateTime('dd/mm/yyyy', Date - 4); // segunda-feira
      3: dma := FormatDateTime('dd/mm/yyyy', Date - 4); // terça-feira
      else
         dma := FormatDateTime('dd/mm/yyyy', Date - 2);
    end;
  end;

  // Alterado 08/05/2012
  // Alterado 04/06/2012
  //Insere_Log(INFORMATIVO, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO, 'Checando lista de Mensagens [' + dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '] - Data inicial: ' + dma);
  Insere_Log(INFORMATIVO, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO, 'Checando lista de Mensagens [' + dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '] - Data inicial: ' + dma + '   Data final: ' + FormatDateTime('dd/mm/yyyy', vim830Global.DataLimiteAte));

  if BCCorreio.IsConectado then
    BCCorreio.DesconectaWS(Erro);

  { Conecta no WebSerice }
  BCCorreio.Instituicao := dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString;
  BCCorreio.Dependencia := dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString;
  BCCorreio.Operador    := dmCorreio.dtsVM_Units.FieldByName('OPERADOR').AsString;
  BCCorreio.Senha       := DescriptografaTextoNum(CHAVE_SENHA_BACEN, dmCorreio.dtsVM_Units.FieldByName('SENHA').AsString);
  BCCorreio.Data_Ini    := EncodeDate(StrToInt(Copy(dma, 7, 4)),
                                      StrToInt(Copy(dma, 4, 2)),
                                      StrToInt(Copy(dma, 1, 2)));

  // Alterado 04/06/2012
  //BCCorreio.Data_Fim    := Date;
  BCCorreio.Data_Fim    := EncodeDate(StrToInt(Copy(DateToStr(vim830Global.DataLimiteAte), 7, 4)),
                                      StrToInt(Copy(DateToStr(vim830Global.DataLimiteAte), 4, 2)),
                                      StrToInt(Copy(DateToStr(vim830Global.DataLimiteAte), 1, 2)));

  nret := LogonToBCCorreio;
  if nret <= 0 then
    exit;

  { Captura as pastas Autorizadas }
  if BCCorreio.ConsultarPastasAutorizadas(erro) <> 1 then
  begin
    Insere_Log(ERRO_CAPTURAR_PASTA_AUTORIZADA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);

    if POS('(401)', Erro) > 0 then
    begin
      // Alterado 26/03/2012
      if dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger < 2 then
        dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                              dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                              'N',
                                                              dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger + 1)
      else
      dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                            dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                            'S',
                                                            0);
    end;
    // Alterado 23/01/2012
    BCCorreio.DesconectaWS(Erro);
    exit;
  end;

  // Alterado 26/03/2012
  dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                        dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                        'N',
                                                        0);

  { Verifica se existe a pasa CaixaEntrada }
  localizou := false;
  for i:=0 to count_Pastas-1 do
  begin
    if Pastas[i].Tipo = 'CaixaEntrada' then
    begin
      localizou := true;
      break;
    end;
  end;

  if not localizou then
  begin
    Erro := 'Login [' +
             dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' +
             dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '] não tem pormissão na pasta CaixaEntrada - Não foi localizada na rotina ConsultarPastasAutorizadas';
    Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);
    exit;
  end;

  if pleaseClose then
    exit;

  { Captura as mensagens da pasta CaixaEntrada }
  if BCCorreio.ConsultarCorreioPorPasta('CaixaEntrada', erro) <> 1 then
  begin
    Insere_Log(ERRO_CORREIO_POR_PASTA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
    ExibeMensagem(Erro);

    if POS('(401)', Erro) > 0 then
    begin
      // Alterado 26/03/2012
      if dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger < 2 then
        dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                              dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                              'N',
                                                              dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger + 1)
      else
      dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                            dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                            'S',
                                                            0);
    end;
    // Alterado 23/01/2012
    BCCorreio.DesconectaWS(Erro);
    exit;
  end;

  // Alterado 26/03/2012
  dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                        dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                        'N',
                                                        0);

  if pleaseClose then
    exit;

  // Alterado 08/05/2012
  Insere_Log(INFORMATIVO, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO, 'Retornado ' + IntToStr(count_Mensagens) + ' mensagem(ns)');

  { Ler as informacoes de cada mensagem na lista }
  for i:=0 to count_Mensagens-1 do
  begin
    Application.ProcessMessages;

    ExibeMensagem('Verifica se mensagem tem que ser inserida - ' + IntToStr(Mensagens[i].NumeroCorreio));

    if (vim830Global.diasMantem <> 0) then
    begin
      if (Mensagens[i].Data < vim830global.dataLimite) or ((Mensagens[i].Data > vim830Global.dataLimiteAte) and (vim830Global.dataLimiteAte <> 0)) then
      begin
        TraceWFile(dmCorreio.traceFName, 'Saiu na verificacao da datalimite');

        if FIsDebug then
          ExibeMensagem('Saiu na verificação da datalimite');

        continue;
      end;
    end;

{$ifdef HOMOLOGA}
    if Mensagens[i].Data > EncodeDate(2012, 06, 10) then
    begin
      TraceWFile(FdbDir + 'TRobot_Correio.TXT', 'Demonstração expirada');
      ExibeMensagem('Demonstração expirada. Entre em contato com a Presenta Sistemas');
      continue;
    end;
{$endif}

    if Mensagens[i].Data < vim830Global.dateInstall then
    begin
      TraceWFile(dmCorreio.traceFName, 'Saiu na verificacao da datainstall');

      if FIsDebug then
        ExibeMensagem('Sai na verificação da datainstall');

      continue;
    end;

    { Verifica autorizacao para ler esta mensagem }
    //if Not VerificaFiltro(Mensagens[i].DependenciaRemetente, Mensagens[i].Assunto) then
    if Not VerificaFiltro(Mensagens[i].UnidadeRemetente, Mensagens[i].Assunto) then
    begin
      if FIsDebug then
        ExibeMensagem('Sai na verificação do filtro - ' + Mensagens[i].UnidadeRemetente + '/' +  Mensagens[i].Assunto);

      continue;
    end;

    { Inserir uma entrada para a msg no banco de dados }
    TraceWFile(dmCorreio.traceFName, 'Verifica se mensagem tem que ser inserida');
    InsertMessageHeader(i);

    { verifica encerramento do programa }
    if (pleaseClose) then
      exit;
  end;

  SaveLogonInformation(1, FormatDateTime('dd/mm/yyyyy hh:nn', Now));

  if FIsDebug then
    ExibeMensagem('Vai desconectar do WS');

  TraceWFile(dmCorreio.traceFName, 'Vai desconectar do WS');

  BCCorreio.DesconectaWS(Erro);

  { Voltar para o menu principal }
  TraceWFile(dmCorreio.traceFName, 'Saindo da lista');
  result := 1;
end;

function TReadBCCorreio.ReadMailText(tipoMsg, msgNumber, subject, transacao,
  sender, status_msg: string; Cod_Msg: Integer): integer;
var
  nRet : integer;
  bufferNum : string;
  p : integer;
  denorNum : string;
  j : integer;
  inseriuDenor : boolean;
  oficio : boolean;
  Msg: TMensagem;
  Erro: String;
  i: integer;
  cl: integer;
  linha: string;
  status: string;
label
  1,
  tryAgain,
  Finaliza;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina ReadMailText');
  Result := -1;
  inseriuDenor := false;

  if FIsDebug then
    ExibeMensagem('Entrou na rotina ReadMailText');

  try
    if vim830Global.tratarOficio then
      nret := BCCorreio.LerCorreio(StrToInt(msgNumber), Msg, erro, EhOficio)
    else
      nret := BCCorreio.LerCorreio(StrToInt(msgNumber), Msg, erro, Nil);
    if nret <> 1 then
    begin
      Insere_Log(ERRO_LER_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, Erro);
      ExibeMensagem(Erro);

      if Erro = 'Não é possível ler mensagem já cancelada' then
      begin
        EhOficio(msg.Assunto, msg.De, oficio);

        UpdateMessageReadStatus(msgNumber, oficio, tipoMsg);
        result := 1;
      end;

      if POS('(401)', Erro) > 0 then
      begin
        // Alterado 26/03/2012
        if dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger < 2 then
          dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                                dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                                'N',
                                                                dmcorreio.dtsVM_Units.FieldByName('prioridade').AsInteger + 1)
        else
        dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                              dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                              'S',
                                                              0);
      end;
      // Alterado 23/01/2012
      BCCorreio.DesconectaWS(Erro);
      exit;
    end;

    // Alterado 26/03/2012
    dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                          dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                          'N',
                                                          0);

    if vim830Global.tratarOficio then
    begin
      EhOficio(msg.Assunto, msg.De, oficio);

      if FIsDebug then
        if oficio then
          ExibeMensagem('Mensagem é um Ofício')
        else
          ExibeMensagem('Mensagem não é um Ofício');
    end;

    // Alterado 03/04/2012
    if msg.RecebidaPor = '' then
    begin
      msg.RecebidaPor := dmcorreio.dtsVM_Units.FieldByName('OPERADOR').AsString;
    end;

    { Atualizar o registro da mensagem }
    nRet := UpdateMessageHeader(msgNumber, status_msg, Msg);

    if pleaseClose then
      exit;

    {limpa a tabela de texto de mensagens }
    dmCorreio.ConectorRobo.Exclui_VM_Messages_Texto(Cod_Msg);

    {limpa a tabela de anexo de mensagens }
    dmCorreio.ConectorRobo.Exclui_VM_Messages_Anexo(Cod_Msg);

    TraceWFile(dmCorreio.traceFName, 'Gravando dados do correio no Banco de Dados');
    ExibeMensagem('Gravando dados do correio no Banco de Dados');

    try
      for cl:=0 to Msg.FormattedText.Count-1 do
      begin
        Application.ProcessMessages;

        status := 'DENOR';

        { Ler uma linha do texto }
        linha := msg.FormattedText.Strings[cl];

        if Pos('HISTORICO DE RETRANSMISSOES', linha) > 0 then
          goto 1;

        if (vim830Global.readDenor) and (UpperCase(Trim(subject)) = 'ORIENTACAO DENOR') then
        begin
          { denor esta sendo cancelada? }
          p := Pos('CANCELAD', UpperCase(linha));
          if p <> 0 then
            status := 'CANCELADA';

          p := Pos('EXCLUIDA', UpperCase(linha));
          if p <> 0 then
            status := 'CANCELADA';

          p := Pos('ALTERADA', UpperCase(linha));
          if p <> 0 then
            status := 'ALTERADA';

          if (Not inseriuDenor) and (status <> 'CANCELADA') then
          begin
            j := 3;
            p := Pos('NO.', UpperCase(linha));
            if p <= 0 then
            begin
              p := Pos('NOS.', UpperCase(linha));
              if p > 0 then
                j := 4;
            end;

            if p > 0 then
            begin
              denorNum := 'DENOR';

              if p + j > 69 then
              begin
                bufferNum := msg.FormattedText.Strings[cl+1];
                p := 0;
                for i := p to Length(bufferNum) do
                begin
                  if  (bufferNum[i] <> '1') And (bufferNum[i] <> '2') And
                      (bufferNum[i] <> '3') And (bufferNum[i] <> '4') And
                      (bufferNum[i] <> '5') And (bufferNum[i] <> '6') And
                      (bufferNum[i] <> '7') And (bufferNum[i] <> '8') And
                      (bufferNum[i] <> '9') And (bufferNum[i] <> '0') And
                      (denorNum <> 'DENOR') then
                    break;

                  if bufferNum[i] <> ' ' then
                    denorNum := denorNum + bufferNum[i];
                end;
              end
              else
              begin
                for i := (p + j) to Length(linha) do
                begin
                  if  (linha[i] <> '1') And (linha[i] <> '2') And
                      (linha[i] <> '3') And (linha[i] <> '4') And
                      (linha[i] <> '5') And (linha[i] <> '6') And
                      (linha[i] <> '7') And (linha[i] <> '8') And
                      (linha[i] <> '9') And (linha[i] <> '0') And
                      (denorNum <> 'DENOR') then
                    break;

                  if linha[i] <> ' ' then
                  denorNum := denorNum + linha[i];
                end;
              end;

              if (denorNum <> 'DENOR') then
              begin
  //              mHeader.number := denorNum;
  //              mHeader.subject := subject;
  //              mHeader.sender := 'DENOR';
  //              mheader.date := FormatDateTime('dd/mm/yyyy', Now);
  //              mHeader.time := FormatDateTime('hh:nn', Now);
  //              InsertMessageHeader(mHeader);
  //              inseriuDenor := true;
              end;
            end;
          end;
        end;

        { Gravar a linha no arquivo }
        dmCorreio.ConectorRobo.Inclui_Mensagem_Texto(COD_MSG, cl+1, linha);
      end;

      { Salvar anexos das msgs }
      for i:=0 to msg.Anexos_Count-1 do
      begin
        Application.ProcessMessages;
        dmcorreio.ConectorRobo.Insere_VM_MESSAGES_ANEXOS(COD_MSG, Msg.Anexos[i].Sequencia, Msg.Anexos[i].NomeAnexo);
      end;
1:
      { Marcar a mensagem como lida }
      UpdateMessageReadStatus(msgNumber, oficio, tipoMsg);

      Result := 1;
    finally

    end;
  except
    on Er: EErro do
      ExibeMensagem('[ReadMailText] ERRO: ' + er.Mensagem);
    on Er: EAviso do
      ExibeMensagem('[ReadMailText] ERRO: ' + er.Mensagem);
    on Er: EInformacao do
      ExibeMensagem('[ReadMailText] ERRO: ' + er.Mensagem);
    on Er: EErroLog do
      ExibeMensagem('[ReadMailText] ERRO: ' + er.Mensagem);
    on mErr:Exception do
    begin
      ExibeMensagem('[ReadMailText] ERRO: ' + merr.Message);
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '[ReadMailText] ERRO: ' + merr.Message);
      TraceWFile(dmCorreio.traceFName, '[ReadMailText] ERRO: ' + merr.Message);
    end;
  end;
end;

function TReadBCCorreio.ReadMessagesFromBCCorreio: integer;
var
  present : TdateTime;
  lastTime : TdateTime;
  minutes : TdateTime;
  dias : integer;
  checkList : boolean;
  fHandle2 : integer;
  trocouSenha : boolean;
  varName : String;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina ReadMessagesFromBCCorreio');

  try
    // Alterado 07/05/2012
    if LerWebPost then
    begin
      AtualizaDadosParaAutojud;
      exit;
    end;

    if FIsDebug then
      ExibeMensagem('Entrou na rotina ReadMessagesFromBCCorreio');

    result := -1;
    try
      if FIsDebug then
        ExibeMensagem('Capturando as Inst/Dep');

      dmCorreio.ConectorRobo.Obtem_VM_Units_Rodar(dmCorreio.dtsVM_Units);

      if dmCorreio.dtsVM_Units.eof then
      begin
        TraceWFile(dmCorreio.traceFName, 'Não tem nenhuma Instituição cadastrada no Admin ou a senha pode estar bloqueada');
        ExibeMensagem('Não tem nenhuma Instituição cadastrada no Admin ou a senha pode estar bloqueada');
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Não tem nenhuma Instituição cadastrada no Admin ou a senha pode estar bloqueada');
      end;

      while Not dmCorreio.dtsVM_Units.eof do
      begin
        Application.ProcessMessages;

        TraceWFile(dmCorreio.traceFName, 'Verificando ' + dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString);

        // Alterado 13/04/2012
        if dmCorreio.dtsVM_Units.FieldByName('PRIORIDADE').AsInteger >= 3 then
        begin
          dmCorreio.ConectorRobo.Altera_Bloqueio_Senha_Sisbacen(dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                                dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                                'N',
                                                                0);
        end;

        if FIsDebug then
          ExibeMensagem('Verificando ' + dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString);

        if pleaseClose then
          break;

        checkList := false;
        if (dmCorreio.dtsVM_Units.FieldByName('transacao').AsString = 'SENHA') then
        begin
          if TemQueTrocarSenha then
          begin
            if TrocaSenhaDoBCCorreio <= 0 then
              Insere_Log(ERRO_LOGON_SISBACEN, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO,
                         'Erro na troca de senha de: ' + dmCorreio.dtsVM_Units.FieldByName('cod_inst').AsString + '/' +
                         dmCorreio.dtsVM_Units.FieldByName('cod_dep').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('Operador').AsString);
            exit;
          end;
          dmCorreio.dtsVM_Units.next;
          continue;
        end;

        { Verifica se existe um departamento para ler }
        if depLer <> '' then
        begin
          if depLer <> dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString then
          begin
            dmCorreio.dtsVM_Units.next;
            continue;
          end;
        end;

        { Verifica se a dependência está bloqueada }
        varName := dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + dmCorreio.dtsVM_Units.FieldByName('NOME_DEP').AsString+'.LCK';
        if IsFileLocked(FdbDir + varName , fHandle2) <= 0 then
        begin
          TraceWFile(dmCorreio.traceFName, dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString +
                     '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '-Depedencia bloqueada. Le outra');
          dmCorreio.dtsVM_Units.next;
          continue;
        end;

        { está na hora de ler esta dependencia? }
        GetUnitsInterval(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                         dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString ,
                         dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,
                         lastTime, minutes, dias);
        present := Now;
        if minutes > -1 then
        begin
          if lastTime = 0 then
            checkList := true
          else
            if (((present - lastTime) >= minutes)) or
               ((present - lastTime) <= 0) then
              checkList := true;
        end;

        BCCorreio.Instituicao := dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString;
        BCCorreio.Dependencia := dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString;
        BCCorreio.Operador    := dmCorreio.dtsVM_Units.FieldByName('OPERADOR').AsString;
        BCCorreio.Senha       := DescriptografaTextoNum(CHAVE_SENHA_BACEN, dmCorreio.dtsVM_Units.FieldByName('SENHA').AsString);

        if (checkList) then
        begin
          TraceWFile(dmCorreio.traceFName, dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString +
                     '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '-Vai protocolar sisbacen');

          { Captura as mensagens }
          CheckMessageList(trocouSenha);
        end;

        if trocouSenha then
        begin
          ReleaseFile(FdbDir + varName, fHandle2);
          exit;
        end;

        if pleaseClose then
        begin
          ReleaseFile(FdbDir + varName, fHandle2);
          break;
        end;

        { verifica se tem mensagens na fila }
        { Se a instituicao estiver bloqueada tenta outra }
        TraceWFile(dmCorreio.traceFName, dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString +
                            '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '-Vai verificar se tem mensagem nao lida');

        { Captura o conteúdo da mensagem }
        CheckMessageQueue('');
        ReleaseFile(FdbDir + varName, fHandle2);
        dmCorreio.dtsVM_Units.next;
      end;
      dmcorreio.dtsVM_Units.Close;
      result := 1;

      if pleaseClose then
        exit;

    except
      on Er: EErro do
        ExibeMensagem('[ReadMessagesFromBCCorreio] ERRO: ' + er.Mensagem);
      on Er: EAviso do
        ExibeMensagem('[ReadMessagesFromBCCorreio] ERRO: ' + er.Mensagem);
      on Er: EInformacao do
        ExibeMensagem('[ReadMessagesFromBCCorreio] ERRO: ' + er.Mensagem);
      on Er: EErroLog do
        ExibeMensagem('[ReadMessagesFromBCCorreio] ERRO: ' + er.Mensagem);
      on mErr:Exception do
      begin
        ExibeMensagem('[ReadMessagesFromBCCorreio] ERRO: ' + merr.Message);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '[ReadMessagesFromBCCorreio] ERRO: ' + merr.Message);
        TraceWFile(dmCorreio.traceFName, '[ReadMessagesFromBCCorreio] ERRO: ' + merr.Message);
        try
          ReleaseFile(FdbDir + varName, fHandle2);
        except
        end;
      end;
    end;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dtsVM_Units.Close;
    except
    end;
  end;
end;

procedure TReadBCCorreio.RemoveOldMessages;
var
  data1: TDateTime;
  fName: array[0..255] of char;
  contador: integer;
  ini: tInifile;
  Arquivo: String;
label
  getNextMsg;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina RemoveOldMessages');

  try
    if FIsDebug then
      ExibeMensagem('Entrou na rotina RemoveOldMessages');

    contador := 0;

    try
      dmCorreio.ConectorRobo.Obtem_VM_Messages(vim830global.dataLimite, dmCorreio.dts);

      while not dmCorreio.dts.eof do
      begin
        Application.ProcessMessages;
        if pleaseClose then
          exit;

        Inc(contador);
        if contador > 10 then
        begin
          sleep(5);
          contador := 0;
        end;

        if dmCorreio.dts.FieldByName('NUM_SISBACEN').AsString = '' then
        begin
          // Alterado 08/05/2012
          if not LerWebPost then
          begin
            //
            // excluir anexos
            //
            try
              dmCorreio.ConectorRobo.Obtem_VM_MESSAGES_ANEXOS(dmCorreio.dts.FieldByName('COD_MSG').AsInteger,
                                                              dmCorreio.dts2);

              if dmCorreio.dts2.RecordCount > 0 then
              begin
                dmCorreio.dts2.First;
                while not dmCorreio.dts2.Eof do
                begin
                  Arquivo := BCCorreio.Dir_Arq_Anexo +
                             dmCorreio.dts.FieldByName('NUM_SISBACEN').AsString + '.' + LeftZeroStr('%03d', [dmcorreio.dts2.FieldByName('SEQUENCIA').AsInteger]);
                  DeleteFile(Arquivo);
                  dmCorreio.dts2.Next;
                end;
                dmCorreio.dts2.Close;
                dmCorreio.ConectorRobo.Exclui_VM_Messages_Anexo(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
              end;
            except
            end;
          end;

          dmCorreio.ConectorRobo.Exclui_VM_Messages(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
          goto getNextMsg;
        end;

        if dmCorreio.dts.FieldByName('DATA_ENVIO').Value = '' then
        begin
          // Alterado 08/05/2012
          if not LerWebPost then
          begin
            //
            // excluir anexos
            //
            try
              dmCorreio.ConectorRobo.Obtem_VM_MESSAGES_ANEXOS(dmCorreio.dts.FieldByName('COD_MSG').AsInteger,
                                                              dmCorreio.dts2);

              if dmCorreio.dts2.RecordCount > 0 then
              begin
                dmCorreio.dts2.First;
                while not dmCorreio.dts2.Eof do
                begin
                  Arquivo := BCCorreio.Dir_Arq_Anexo +
                             dmCorreio.dts.FieldByName('NUM_SISBACEN').AsString + '.' + LeftZeroStr('%03d', [dmcorreio.dts2.FieldByName('SEQUENCIA').AsInteger]);
                  DeleteFile(Arquivo);
                  dmCorreio.dts2.Next;
                end;
                dmCorreio.dts2.Close;
                dmCorreio.ConectorRobo.Exclui_VM_Messages_Anexo(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
              end;
            except
            end;
          end;

          dmCorreio.ConectorRobo.Exclui_VM_Messages(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
          goto getNextMsg;
        end;

        try
          if length(dmCorreio.dts.FieldByname('DATA_ENVIO').Value) < 10 then
            data1 := strToDate(dmCorreio.dts.FieldByname('DATA_ENVIO').Value + '/' +
                               copy(dmCorreio.dts.FieldByname('AMD').Value, 1, 4))
          else
            data1 := StrToDate(dmCorreio.dts.FieldByname('DATA_ENVIO').Value);
        except
          data1 := vim830global.dataLimite - 10;
        end;

        if (data1 < vim830Global.dataLimite) then
        begin
          ExibeMensagem('Excluindo mensagem: ' + dmCorreio.dts.FieldByname('NUM_SISBACEN').Value + '-' +
                                                 dmCorreio.dts.FieldByName('ASSUNTO').Value + '-' +
                                                 dmCorreio.dts.FieldByName('DATA_ENVIO').Value);

          StrPCopy(fName, FdbDir+'MSGS\'+ dmCorreio.dts.FieldByname('NUM_SISBACEN').Value+'.TXT');
          DeleteFile(fName);

          // Alterado 08/05/2012
          if not LerWebPost then
          begin
            //
            // excluir anexos
            //
            try
              dmCorreio.ConectorRobo.Obtem_VM_MESSAGES_ANEXOS(dmCorreio.dts.FieldByName('COD_MSG').AsInteger,
                                                              dmCorreio.dts2);

              if dmCorreio.dts2.RecordCount > 0 then
              begin
                dmCorreio.dts2.First;
                while not dmCorreio.dts2.Eof do
                begin
                  Arquivo := BCCorreio.Dir_Arq_Anexo +
                             dmCorreio.dts.FieldByName('NUM_SISBACEN').AsString + '.' + LeftZeroStr('%03d', [dmcorreio.dts2.FieldByName('SEQUENCIA').AsInteger]);
                  DeleteFile(Arquivo);
                  dmCorreio.dts2.Next;
                end;
                dmCorreio.dts2.Close;
                dmCorreio.ConectorRobo.Exclui_VM_Messages_Anexo(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
              end;
            except
            end;
          end;

          dmCorreio.ConectorRobo.Exclui_VM_Messages_Texto(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
          dmCorreio.ConectorRobo.Exclui_VM_Messages(dmCorreio.dts.FieldByName('COD_MSG').AsInteger);
        end;

  getNextMsg:
        dmCorreio.dts.next;
      end;
      dmCorreio.dts.Close;

      ini := TiniFile.Create(FdbDir + 'VIM830.INI');
      ini.WriteString(SECTION_GLOBAL, ENTRY_LAST_DELETE, FormatDateTime('dd/mm/yyyy', Now));
      ini.free;
    except
      // Alterado 08/05/2012
      on mErr:EErro do
      begin
        ExibeMensagem('<ERRO> ' + merr.Mensagem);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '(RemoveOldMessages) - <ERRO> ' + merr.Mensagem);
        TraceWFile(dmCorreio.traceFName, '(RemoveOldMessages) - <ERRO> ' + merr.Mensagem);
      end;

      on mErr:Exception do
      begin
        ExibeMensagem('<ERRO> ' + merr.message);
        Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '(RemoveOldMessages) - <ERRO> ' + merr.message);
        TraceWFile(dmCorreio.traceFName, '(RemoveOldMessages) - <ERRO> ' + merr.message);
      end;
    end;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dts.Close;
      dmcorreio.dts2.Close;
    except
    end;
  end;
end;

procedure TReadBCCorreio.SaveLogonInformation(tipo: integer; strUp: string);
begin
  TraceWFile(dmCorreio.traceFName, 'Salvando dados na tabela UNITS');

  if FIsDebug then
    ExibeMensagem('Salvando dados na tabela UNITS');

  dmcorreio.ConectorRobo.Atualiza_VM_Units(Tipo,
                                           dmcorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                           dmcorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                           dmcorreio.dtsVM_Units.FieldByName('OPERADOR').AsString,
                                           dmcorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,
                                           dmcorreio.dtsVM_Units.FieldByName('NOME_DEP').AsString,
                                           strUp);
end;

procedure TReadBCCorreio.SaveMessageSent(qryMsgs: TADODataSet);
begin
  try
    dmCorreio.ConectorRobo.Marcar_Mensagem_Como_Entregue(qryMsgs.FieldByName('COD_MSG').AsInteger);
  except
    // Alterado 08/05/2012
    on mErr:EErro do
    begin
      ExibeMensagem('<ERRO> ' + merr.Mensagem);
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '[ExecutaEntregaDeMensagens] ERRO:' + merr.Mensagem);
      TraceWFile(dmCorreio.traceFName, '[ExecutaEntregaDeMensagens] ERRO:' + merr.Mensagem);
    end;

    on mErr:Exception do
    begin
      ExibeMensagem('<ERRO> Exception:' + merr.message);
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '[ExecutaEntregaDeMensagens] ERRO:' + merr.message);
      TraceWFile(dmCorreio.traceFName, '[ExecutaEntregaDeMensagens] ERRO:' + merr.message);
    end;
  end;
end;

function TReadBCCorreio.TemQueTrocarSenha: boolean;
var
  ini : TIniFile;
  daysToChange : integer;
  ultimaTroca : TDateTime;
begin
  result := false;

  if FIsDebug then
    ExibeMensagem('Entrou na rotina TrocaSenha');

  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina TrocaSenha');
  if Not vim830Global.trocaSenha then
    exit;

  if dmCorreio.dtsVM_Units.FieldByName('troca_senha_ativa').AsInteger = 0 then
    exit;

  ini := TIniFile.Create(vim830global.pwdDir + '\PWD.INI');
  daysToChange := ini.ReadInteger(SECTION_SYSTEM, ENTRY_DAYS_TO_CHANGE, 0);
  if daysToChange = 0 then
  begin
    TraceWFile(dmCorreio.traceFName, 'Saiu da rotina TrocaSenha');
    ini.free;
    exit;
  end;

  if dmCorreio.dtsVM_Units.FieldByName('ultima_troca').IsNull then
    ultimaTroca := EncodeDate(1980, 1, 1)
  else
    ultimaTroca := dmCorreio.dtsVM_Units.FieldByName('ultima_troca').AsDateTime;
  if FormatDateTime('dd/mm/yyyy', ultimaTroca) = FormatDateTime('dd/mm/yyyy', Date) then
  begin
    TraceWFile(dmCorreio.traceFName, 'Saiu da rotina TrocaSenha');
    ini.free;
    exit;
  end;
  ini.free;

  ultimaTroca := IncrementData(ultimaTroca, daysToChange, false);
  if ultimaTroca <= Date then
    result := true;
  TraceWFile(dmCorreio.traceFName, 'Saiu da rotina TrocaSenha');
end;

function TReadBCCorreio.ThereAreNewMessages: boolean;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina ThereAreNewMessages');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina ThereAreNewMessages');

  result := false;

  try
    result := dmCorreio.ConectorRobo.Tem_Mensagem_Enviar;

  except
    // Alterado 08/05/2012
    on Er:EErro do
    begin
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na ThereAreNewMessages - ' + Er.Mensagem);
      TraceWFile(dmCorreio.traceFName, 'Erro na ThereAreNewMessages - ' + Er.Mensagem);
    end;

    on Er:Exception do
    begin
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, 'Erro na ThereAreNewMessages - ' + Er.Message);
      TraceWFile(dmCorreio.traceFName, 'Erro na ThereAreNewMessages - ' + Er.Message);
    end;
  end;
end;

function TReadBCCorreio.TrocaSenhaDoBCCorreio: Integer;
var
  Pstac10Componente: TPstac10Componente;
  ocorrencia: WideString;
  iniFile: TIniFile;
  iniStr: String;
  pwdPrefix: String;
  lastSufix: integer;
  pwPas: string;
  ultimaTroca : TdateTime;
  present : TDateTime;
  newSufix: Integer;
  PNewPwd: WideString;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina TrocaSenhaDoBCCorreio');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina TrocaSenhaDoBCCorreio');

  result := -1;
  Pstac10Componente := TPstac10Componente.Create;
  try
    try
      if pleaseClose then
        exit;

      ExibeMensagem('Trocando senha da ' + dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' +
                                           dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString + '/' +
                                           dmCorreio.dtsVM_Units.FieldByName('OPERADOR').AsString);

      iniFile   := TIniFile.Create(vim830Global.pwdDir + '\PWD.INI');
      iniStr    := iniFile.ReadString(SECTION_SYSTEM, ENTRY_PWD_PREFIX, '');
      pwdPrefix := DescriptografaTextoNum(CHAVE_SENHA_BACEN, iniStr);
      lastSufix := dmCorreio.dtsVM_Units.FieldByName('LAST_SUFIX').AsInteger;
      pwPas := DescriptografaTextoNum(CHAVE_SENHA_BACEN, dmCorreio.dtsVM_Units.FieldByName('Senha').AsString);
      if pwPas = '' then
      begin
        iniFile.Free;
        exit;
      end;

      if Not dmCorreio.dtsVM_Units.FieldByName('ULTIMA_TROCA').IsNull then
      begin
        ultimaTroca :=  dmCorreio.dtsVM_Units.FieldByName('ultima_troca').AsDateTime;
        present := Date;
        if ultimaTroca = present then
        begin
          iniFile.free;
          exit;
        end;
      end;
      iniFile.Free;

      if Pstac10Componente.CalculaSenha(pwdPrefix,
                                        lastSufix,
                                        PNewPwd,
                                        ocorrencia,
                                        newSufix) < 0 then
      begin
        ExibeMensagem('Erro no calculo da nova senha: ' + ocorrencia);
        Insere_Log(ERRO_TROCA_SENHA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '(BCCorreio - Calculo da nova senha) - ' + ocorrencia);
        exit;
      end;

      Pstac10Componente.Instituicao  := dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString;
      Pstac10Componente.Dependencia  := dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString;
      Pstac10Componente.Operador     := dmCorreio.dtsVM_Units.FieldByName('OPERADOR').AsString;
      Pstac10Componente.SenhaAtual   := pwPas;
      Pstac10Componente.NovaSenha    := PNewPwd;
      Pstac10Componente.DiretorioExe := FDirExePSTAW10;
      Pstac10Componente.DiretorioLog := BCCorreio.Dir_Arq_Anexo;

      if Pstac10Componente.TrocaSenha(ocorrencia) < 0 then
      begin
        ExibeMensagem('Erro na troca de senha: ' + ocorrencia);
        Insere_Log(ERRO_TROCA_SENHA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '(BCCorreio) - ' + ocorrencia);
        exit;
      end;

      dmCorreio.ConectorRobo.Atualiza_Senha_Correio(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                                                    dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString,
                                                    pwdPrefix,
                                                    newSufix);

      result := 1;
    except
      // Alterado 08/05/2012
      on Er: EErro do
      begin
        ExibeMensagem('Erro na troca de senha: ' + Er.Mensagem);
        Insere_Log(ERRO_TROCA_SENHA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '(BCCorreio) - ' + Er.Mensagem);
        TraceWFile(dmCorreio.traceFName, '(BCCorreio) - ' + Er.Mensagem);
      end;

      on Er: Exception do
      begin
        ExibeMensagem('Erro na troca de senha: ' + Er.Message);
        Insere_Log(ERRO_TROCA_SENHA, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '(BCCorreio) - ' + Er.Message);
        TraceWFile(dmCorreio.traceFName, '(BCCorreio) - ' + Er.Message);
      end;
    end;
  finally
    FreeAndNil(Pstac10Componente);
  end;
end;

function TReadBCCorreio.UpdateMessageHeader(msgNumber, status: string;
  Msg: TMensagem): integer;
begin
  result := -1;

  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina UpdateMessageHeader');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina UpdateMessageHeader');

  try
    if FormatDateTime('dd/mm/yyyy', msg.RecebidaEm) = '01/01/0001' then
      msg.RecebidaEm := now;

    { Atualizar o registro da mensagem }
    dmcorreio.ConectorRobo.Atualiza_Mensagem(msg.De,                                       // DESCR_DEP_REMETENTE,
                                             msg.De,                                       // USUARIO_REMETENTE,
                                             msg.RecebidaPor,                              // RECEBIDO_POR,
                                             FormatDateTime('dd/mm/yyyy', msg.RecebidaEm), // DT_RECEBIMENTO,
                                             FormatDateTime('hh:nn', msg.RecebidaEm),      // HR_RECEBIMENTO: String;
                                             0,                                            // TOTAL_PAGINAS: Integer;
                                             status,                                       // STATUS_SISBACEN,
                                             msgNumber);                                   // NUM_SISBACEN: String;
  except
    // Alterado 08/05/2012
    on mErr:EErro do
    begin
      ExibeMensagem('<UpdateMessageHeader> ' + mErr.Mensagem);
      TraceWFile(dmCorreio.traceFName, '<UpdateMessageHeader> ' + mErr.Mensagem);
      result := -1;
      exit;
    end;

    on mErr:Exception do
    begin
      ExibeMensagem('<UpdateMessageHeader> ' + mErr.Message);
      TraceWFile(dmCorreio.traceFName, '<UpdateMessageHeader> ' + mErr.Message);
      result := -1;
      exit;
    end;
  end;
  Result := 1;
end;

function TReadBCCorreio.UpdateMessageReadStatus(msgNumber: string;
  ehOficio: boolean; tipoMsg: string): integer;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina UpdateMessageReadStatus [msgNumber = ' + msgNumber+ ']');

  if FIsDebug then
    ExibeMensagem('Entrou na rotina UpdateMessageReadStatus [msgNumber = ' + msgNumber+ ']');

  if ehOficio then
    TraceWFile(dmCorreio.traceFName, 'ehOficio = true')
  else
    TraceWFile(dmCorreio.traceFName, 'ehOficio = false');

  { Atualizar o registro da mensagem }
  try
    dmCorreio.ConectorRobo.Marca_Msg_Lida(msgNumber, ehOficio);
    Insere_Log(IMPORTOU_MENSAGEM, ID_Usuario_Log, LOG_INFO, LOG_CRIT_MEDIA, LOG_FECHADO, msgNumber);
    result := 1;
  except
    on Er: EErro do
      ExibeMensagem('<UpdateMessageReadStatus> Erro: ' + er.Mensagem);
    on Er: EAviso do
      ExibeMensagem('<UpdateMessageReadStatus> Erro: ' + er.Mensagem);
    on Er: EInformacao do
      ExibeMensagem('<UpdateMessageReadStatus> Erro: ' + er.Mensagem);
    on Er: EErroLog do
      ExibeMensagem('<UpdateMessageReadStatus> Erro: ' + er.Mensagem);
    on mErr:Exception do
    begin
      ExibeMensagem('<UpdateMessageReadStatus> Erro: ' + mErr.Message);
      Insere_Log(ERRO_BC_CORREIO, ID_Usuario_Log, LOG_ERRO, LOG_CRIT_ALTA, LOG_ABERTO, '<UpdateMessageReadStatus> Erro: ' + mErr.Message);
      TraceWFile(dmCorreio.traceFName, '<UpdateMessageReadStatus> Erro: ' + mErr.Message);
      result := -1;
    end;
  end;
end;

function TReadBCCorreio.VaiTerQueRodar: boolean;
var
  present, lasttime, minutes : TDateTime;
  dias : integer;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina VaiTerQueRodar');

  try
    if FIsDebug then
      ExibeMensagem('Entrou na rotina VaiTerQueRodar');

    result := true;

    try
      dmCorreio.ConectorRobo.Obtem_VM_Units_Rodar(dmCorreio.dtsVM_Units);

      while not dmCorreio.dtsVM_Units.Eof do
      begin
        Application.ProcessMessages;

        if pleaseClose then
        begin
          result := false;
          exit;
        end;

        if sisbacenPublico and
           (dmCorreio.dtsVM_Units.FieldByName('transacao').AsString <> 'PMSG835') and
           (dmCorreio.dtsVM_Units.FieldByName('transacao').AsString <> 'SENHA')then begin
          dmCorreio.dtsVM_Units.next;
          continue;
        end;

        if dmCorreio.dtsVM_Units.FieldByName('transacao').AsString = 'SENHA' then
        begin
          if TemQueTrocarSenha then
            exit;
          dmCorreio.dtsVM_Units.next;
          continue;
        end;

        if depLer <> '' then
        begin
          if depLer <> dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString + '/' + dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString then
          begin
            dmCorreio.dtsVM_Units.Next;
          end;
        end;

        GetUnitsInterval(dmCorreio.dtsVM_Units.FieldByName('COD_INST').AsString,
                         dmCorreio.dtsVM_Units.FieldByName('COD_DEP').AsString ,
                         dmCorreio.dtsVM_Units.FieldByName('TRANSACAO').AsString,
                         lastTime, minutes, dias);
                         present := Now;
        if minutes > -1 then
        begin
          if lastTime = 0 then
          begin
            result := true;
            exit;
          end
          else
          begin
            if (((present - lastTime) >= minutes)) or
               ((present - lastTime) <= 0) then
            begin
              result := true;
              exit;
            end;
          end;
        end;
        dmCorreio.dtsVM_Units.next;
      end;
      dmCorreio.dtsVM_Units.Close;
      result := false;
    except
      result := false;
    end;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dtsVM_Units.Close;
    except
    end;
  end;
end;

function TReadBCCorreio.VerificaFiltro(sender, subject: string): boolean;
var
  i: integer;
  result1, result2, result3: boolean;
  sValue : TStringList;
begin
  TraceWFile(dmCorreio.traceFName, 'Entrou na rotina VerificaFiltro');

  try
    if FIsDebug then
      ExibeMensagem('Verificar sender ' + sender);

    result  := true;

    //DEPENDENCIA TEM FILTRO DE ALGUM TIPO?
    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               -1,
                                               true,
                                               dmCorreio.dts);

    dmCorreio.dts.Open;
    if dmCorreio.dts.Eof then
    begin
      // Alterado 19/06/2012
      dmCorreio.dts.Close;

      exit;
    end;

    // Alterado 19/06/2012
    dmCorreio.dts.Close;

    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               ENTRY_DEPTOS,
                                               true,
                                               dmCorreio.dts);

    sValue := TstringList.Create;
    while not dmCorreio.dts.eof do
    begin
      Application.ProcessMessages;

      if FIsDebug then
        ExibeMensagem('filtro cadastrado: ' + dmCorreio.dts.FieldByName('FILTRO').AsString);

      sValue.Add(dmCorreio.dts.FieldByName('FILTRO').AsString);
      dmCorreio.dts.Next;
    end;

    // Alterado 19/06/2012
    dmCorreio.dts.Close;

    result1 := sValue.Count = 0;
    for i := 0 to sValue.Count - 1 do
      if Trim(sender) = sValue.Strings[i] then
      begin
        result1 := true;
        break;
      end;
    sValue.Free;

    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               ENTRY_GOOD_WORDS,
                                               true,
                                               dmCorreio.dts);


    sValue := TstringList.Create;
    while not dmCorreio.dts.eof do
    begin
      Application.ProcessMessages;
      sValue.Add(dmCorreio.dts.FieldByName('FILTRO').AsString);
      dmCorreio.dts.Next;
    end;


    // Alterado 19/06/2012
    dmCorreio.dts.Close;

    result2 := sValue.Count = 0;
    for i := 0 to sValue.Count - 1 do
      if Pos(sValue.Strings[i], subject) > 0 then
      begin
        result2 := true;
        break;
      end;
    sValue.Free;

    dmcorreio.ConectorRobo.Obtem_VM_Filtro_Dep(dmcorreio.dtsVM_Units.FieldByname('COD_INST').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('COD_DEP').AsString,
                                               dmcorreio.dtsVM_Units.FieldByname('TRANSACAO').AsString,
                                               ENTRY_BAD_WORDS,
                                               true,
                                               dmCorreio.dts);

    sValue := TstringList.Create;
    while not dmCorreio.dts.eof do
    begin
      Application.ProcessMessages;
      sValue.Add(dmCorreio.dts.FieldByName('FILTRO').AsString);
      dmCorreio.dts.Next;
    end;

    // Alterado 19/06/2012
    dmCorreio.dts.Close;

    result3 := true;
    for i := 0 to sValue.Count - 1 do
      if Pos(sValue.Strings[i], subject) > 0 then
      begin
        result3 := false;
        break;
      end;
    sValue.Free;

    if (Not result1) or (Not result2) or (Not result3) then
      result := false;
  // Alterado 19/06/2012
  finally
    try
      dmcorreio.dts.Close;
    except
    end;
  end;
end;

end.
