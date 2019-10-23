unit datamodulegcpj_base_xi;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, dialogs, forms;

type
  Tdmgcpcj_base_XI = class(TDataModule)
    adoConnNfiscal: TADOConnection;
    adoConnNfiscdt: TADOConnection;
    adoCmdNfiscdt: TADOCommand;
    dtsNfiscdt: TADODataSet;
    dtsNfiscal: TADODataSet;
    adoCmdNfiscal: TADOCommand;
    adoConnNfiscdt_2013: TADOConnection;
    adoCmdNfiscdt_2013: TADOCommand;
    dtsNfiscdt_2013: TADODataSet;
    adoConnNfiscdt_2013_2: TADOConnection;
    adoCmdNfiscdt_2013_2: TADOCommand;
    dtsNfiscdt_2013_2: TADODataSet;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure CreateIndex;
    function ObtemOutrosPagamentosDoProcesso(gcpj, tipoato, escritorio, tipoprocesso, codescritorio: string; lstCnpjs: TStringList):integer;
    procedure ObtemPagamentosLaudoPericial(gcpj, codEscritorio: string);
    function ObtemSomaJaPagaADescontar(gcpj: string; drcContraria: integer; var qtdePaga: integer; codEscritorio: integer; const somenteAto: string = '';const excluirVolumetria : boolean = false; const recuperacaoFinal: boolean = false) : double;
    function ObtemOutrosPagamentosDoProcessoAtivas(gcpj, tipoato, escritorio, tipoprocesso, codescritorio: string; lstCnpjs: TStringList):integer;
    function ObtemQualquerPagamentoDoProcesso(gcpj, escritorio, tipoprocesso, codescritorio: string; lstCnpjs: TStringList):integer;
    function PagouAcordoNestaData(gcpj: string; dataAto: TdateTime) : boolean;
    function PagouTrabalhistaNestaData(andamento, gcpj: string; dataAto: TDateTime) : boolean;
    procedure ObtemAcordosPagos(gcpj: string);
  end;

var
  dmgcpcj_base_XI: Tdmgcpcj_base_XI;

implementation

{$R *.dfm}

{ Tdmgcpcj_base_XI }


{ Tdmgcpcj_base_XI }

procedure Tdmgcpcj_base_XI.CreateIndex;
begin
   try
      adoCmdNfiscal.CommandText := 'create index idx1 on NFISCAL(CodEscritorio, NroNota)';
      adoCmdNfiscal.Execute;
   except
   end;

   //try
//      adoCmd.CommandText := 'create index idx1 on NFISCDT(CodEscritorio, NroNota)';
//      adoCmd.Execute;
   //except
   //end;

   //try
   //   adoCmd.CommandText := 'create index idx2 on NFISCDT(nomdes, nomescritorio)';
   //   adoCmd.Execute;
   //except
   //end;

   try
      adoCmdNfiscdt.CommandText := 'create index idx3 on NFISCDT(bradesco)';
      adoCmdNfiscdt.Execute;
   except
   end;

   try
      adoCmdNfiscdt_2013.CommandText := 'create index idx3 on NFISCDT(bradesco)';
      adoCmdNfiscdt_2013.Execute;
   except
   end;



   //try
      //adoCmd.CommandText := 'create index idx4 on NFISCDT(nomdes, CodEscritorio)';
      //adoCmd.Execute;
   //except
   //end;

   //try
      //adoCmd.CommandText := 'create index idx5 on NFISCDT(nomdes, bradesco, datOcorr)';
      //adoCmd.Execute;
   //except
   //end;

   //try
      //adoCmd.CommandText := 'create index idx6 on NFISCDT(bradesco, nomdes)';
      //adoCmd.Execute;
   //except
   //end;
end;

procedure Tdmgcpcj_base_XI.ObtemAcordosPagos(gcpj: string);
begin
   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select datOcorr, nomdes ' +
                      'from NFISCDT ' +
                      'where bradesco = ' + gcpj + ' and nomdes = ''ACORDO''';
   dtsNfiscdt.Open;

   if dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select datOcorr, nomdes ' +
                        'from NFISCDT ' +
                        'where bradesco = ' + gcpj + ' and nomdes = ''ACORDO''';
     dtsNfiscdt_2013.Open;

     if dtsNfiscdt_2013.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013);

   end;


   if dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select datOcorr, nomdes ' +
                        'from NFISCDT ' +
                        'where bradesco = ' + gcpj + ' and nomdes = ''ACORDO''';
     dtsNfiscdt_2013_2.Open;

     if dtsNfiscdt_2013_2.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013_2);
  end;


end;

function Tdmgcpcj_base_XI.ObtemOutrosPagamentosDoProcesso(gcpj, tipoato,
  escritorio, tipoprocesso, codescritorio: string;
  lstCnpjs: TStringList): integer;
var
   i : integer;
begin
   result := -1;

   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select nomescritorio, datmovto, nomdes, datOcorr ' +
                      'from NFISCDT ' +
                      'where ';
   if Pos('TUTELA', UpperCase(tipoato )) <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomdes = ''CONTESTACAO'' or nomdes = ''AUDIENCIA DE CONCILIACAO'' or nomdes = ''AUDIENCIA INICIAL'' or nomdes = ''ACOMPANHAMENTO'') and '
   else
   if Pos('XITO', tipoato) <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + 'nomdes like ''%EXITO%'' and '
   else
   if (tipoato = 'RECURSO NA FASE DE EXECUCAO') or
      (tipoato = 'RECURSO NA FASE DE EXECU«√O') or
      (tipoato = 'RECURSO EM FASE DE EXECU«√O') or
      (tipoato = 'RECURSO EM FASE DE EXECUCAO') then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + 'nomdes = ''RECURSO NA FASE DE EXECUCAO'' and '
   else
   if (Pos('RECURSO', tipoato) <> 0) OR
      (Pos('AGRAVO FAT', tipoato) <> 0) then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + 'nomdes = ''RECURSO'' and '
   else
   if (tipoato = 'AUDIÍNCIA AVULSA') or (tipoato = 'AUDI NCIA AVULSA') or (tipoato = 'AUDIENCIA AVULSA') or
      (tipoato = 'AVULSO TRABALHISTA') or (tipoato = 'AUDIENCIA TRABALHISTA') or (tipoato = 'AUDI NCIA TRABALHISTA') or
      (tipoato = 'INSTRUCAO')  then
   begin
      if (tipoprocesso = 'TR') OR (tipoprocesso = 'TO') or (tipoprocesso = 'TA') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + 'nomdes = ''INSTRUCAO'' and '
      else
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + 'nomdes = ''ACOMPANHAMENTO'' and '
   end
   else
   if (Tipoato = 'AVULSO CÕVEL (AUDITORIA)') or (Tipoato = 'AVULSO') or (Tipoato = 'AVULSO CÕVEL') or
      (Tipoato = 'AVULSO CIVEL')then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + 'nomdes = ''ASSESSORIA JURIDICA'' and '
   else
   begin
      if (Tipoato = 'HONOR¡RIOS INICIAIS') or (Tipoato = 'HONORARIOS INICIAIS') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + 'nomdes = ''HONORARIOS INICIAIS'' and '
      else
      if (Tipoato = 'HONOR¡RIOS FINAIS') or (Tipoato = 'HONORARIOS FINAIS') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomdes = ''HONORARIOS FINAIS'' or nomdes = ''HONORARIOS DE EXITO'') and '
      else
      if (Tipoato = 'AJUIZAMENTO') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomdes = ''AJUIZAMENTO'') and '
      else
      if (Tipoato = 'CARTA CONVITE') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomdes = ''CARTA CONVITE'') and '
      else
      begin
//       ShowMessage('Erro tipoato: ' + tipoato);
//         exit;
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomdes = ''' + tipoato + ''') and ';
      END;
   end;
//   if Pos('JBM', escritorio) <> 0 then
//      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
//                         'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '

{
   else
   if Pos('CAMARA', escritorio) <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                         'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
   else}
   begin
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' CodEscritorio in(';
      for i := 0 to lstCnpjs.Count - 1 do
      begin
         if i <> 0 then
            dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ', ';
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + lstCnpjs.Strings[i];
      end;
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ') and ';
   end;
   dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' bradesco = ' + gcpj;

   dtsNfiscdt.Open;

   If dtsNfiscdt.RecordCount = 0 then
   begin

     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select nomescritorio, datmovto, nomdes, datOcorr ' +
                        'from NFISCDT ' +
                        'where ';
     if Pos('TUTELA', UpperCase(tipoato )) <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomdes = ''CONTESTACAO'' or nomdes = ''AUDIENCIA DE CONCILIACAO'' or nomdes = ''AUDIENCIA INICIAL'' or nomdes = ''ACOMPANHAMENTO'') and '
     else
     if Pos('XITO', tipoato) <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + 'nomdes like ''%EXITO%'' and '
  {   else
     if (tipoato = 'RECURSO NA FASE DE EXECUCAO') or
        (tipoato = 'RECURSO NA FASE DE EXECU«√O') or
        (tipoato = 'RECURSO EM FASE DE EXECU«√O') or
        (tipoato = 'RECURSO EM FASE DE EXECUCAO') then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + 'nomdes = ''RECURSO NA FASE DE EXECUCAO'' and '}
     else
     if (Pos('RECURSO', tipoato) <> 0) OR
        (Pos('AGRAVO FAT', tipoato) <> 0) then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + 'nomdes = ''RECURSO'' and '
     else
     if (tipoato = 'AUDIÍNCIA AVULSA') or (tipoato = 'AUDI NCIA AVULSA') or (tipoato = 'AUDIENCIA AVULSA') or
        (tipoato = 'AVULSO TRABALHISTA') or (tipoato = 'AUDIENCIA TRABALHISTA') or (tipoato = 'AUDI NCIA TRABALHISTA') or
        (tipoato = 'INSTRUCAO')  then
     begin
        if (tipoprocesso = 'TR') OR (tipoprocesso = 'TO') or (tipoprocesso = 'TA') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + 'nomdes = ''INSTRUCAO'' and '
        else
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + 'nomdes = ''ACOMPANHAMENTO'' and '
     end
     else
     if (Tipoato = 'AVULSO CÕVEL (AUDITORIA)') or (Tipoato = 'AVULSO') or (Tipoato = 'AVULSO CÕVEL') or
        (Tipoato = 'AVULSO CIVEL')then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + 'nomdes = ''ASSESSORIA JURIDICA'' and '
     else
     begin
        if (Tipoato = 'HONOR¡RIOS INICIAIS') or (Tipoato = 'HONORARIOS INICIAIS') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + 'nomdes = ''HONORARIOS INICIAIS'' and '
        else
        if (Tipoato = 'HONOR¡RIOS FINAIS') or (Tipoato = 'HONORARIOS FINAIS') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomdes = ''HONORARIOS FINAIS'' or nomdes = ''HONORARIOS DE EXITO'') and '
        else
        if (Tipoato = 'AJUIZAMENTO') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomdes = ''AJUIZAMENTO'') and '
        else
        if (Tipoato = 'CARTA CONVITE') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomdes = ''CARTA CONVITE'') and '
        else
        begin
  //       ShowMessage('Erro tipoato: ' + tipoato);
  //         exit;
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomdes = ''' + tipoato + ''') and ';
        END;
     end;
  //   if Pos('JBM', escritorio) <> 0 then
  //      dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
  //                         'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '

  {
     else
     if Pos('CAMARA', escritorio) <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                           'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
     else}
     begin
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' CodEscritorio in(';
        for i := 0 to lstCnpjs.Count - 1 do
        begin
           if i <> 0 then
              dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ', ';
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + lstCnpjs.Strings[i];
        end;
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ') and ';
     end;
     dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' bradesco = ' + gcpj;

     dtsNfiscdt_2013.Open;

     if dtsNfiscdt_2013.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013);
   end;



   If dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select nomescritorio, datmovto, nomdes, datOcorr ' +
                        'from NFISCDT ' +
                        'where ';
     if Pos('TUTELA', UpperCase(tipoato )) <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomdes = ''CONTESTACAO'' or nomdes = ''AUDIENCIA DE CONCILIACAO'' or nomdes = ''AUDIENCIA INICIAL'' or nomdes = ''ACOMPANHAMENTO'') and '
     else
     if Pos('XITO', tipoato) <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + 'nomdes like ''%EXITO%'' and '
  {   else
     if (tipoato = 'RECURSO NA FASE DE EXECUCAO') or
        (tipoato = 'RECURSO NA FASE DE EXECU«√O') or
        (tipoato = 'RECURSO EM FASE DE EXECU«√O') or
        (tipoato = 'RECURSO EM FASE DE EXECUCAO') then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + 'nomdes = ''RECURSO NA FASE DE EXECUCAO'' and '}
     else
     if (Pos('RECURSO', tipoato) <> 0) OR
        (Pos('AGRAVO FAT', tipoato) <> 0) then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + 'nomdes = ''RECURSO'' and '
     else
     if (tipoato = 'AUDIÍNCIA AVULSA') or (tipoato = 'AUDI NCIA AVULSA') or (tipoato = 'AUDIENCIA AVULSA') or
        (tipoato = 'AVULSO TRABALHISTA') or (tipoato = 'AUDIENCIA TRABALHISTA') or (tipoato = 'AUDI NCIA TRABALHISTA') or
        (tipoato = 'INSTRUCAO')  then
     begin
        if (tipoprocesso = 'TR') OR (tipoprocesso = 'TO') or (tipoprocesso = 'TA') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + 'nomdes = ''INSTRUCAO'' and '
        else
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + 'nomdes = ''ACOMPANHAMENTO'' and '
     end
     else
     if (Tipoato = 'AVULSO CÕVEL (AUDITORIA)') or (Tipoato = 'AVULSO') or (Tipoato = 'AVULSO CÕVEL') or
        (Tipoato = 'AVULSO CIVEL')then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + 'nomdes = ''ASSESSORIA JURIDICA'' and '
     else
     begin
        if (Tipoato = 'HONOR¡RIOS INICIAIS') or (Tipoato = 'HONORARIOS INICIAIS') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + 'nomdes = ''HONORARIOS INICIAIS'' and '
        else
        if (Tipoato = 'HONOR¡RIOS FINAIS') or (Tipoato = 'HONORARIOS FINAIS') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomdes = ''HONORARIOS FINAIS'' or nomdes = ''HONORARIOS DE EXITO'') and '
        else
        if (Tipoato = 'AJUIZAMENTO') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomdes = ''AJUIZAMENTO'') and '
        else
        if (Tipoato = 'CARTA CONVITE') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomdes = ''CARTA CONVITE'') and '
        else
        begin
  //       ShowMessage('Erro tipoato: ' + tipoato);
  //         exit;
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomdes = ''' + tipoato + ''') and ';
        END;
     end;
  //   if Pos('JBM', escritorio) <> 0 then
  //      dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
  //                         'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '

  {
     else
     if Pos('CAMARA', escritorio) <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                           'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
     else}
     begin
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' CodEscritorio in(';
        for i := 0 to lstCnpjs.Count - 1 do
        begin
           if i <> 0 then
              dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ', ';
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + lstCnpjs.Strings[i];
        end;
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ') and ';
     end;
     dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' bradesco = ' + gcpj;

     dtsNfiscdt_2013_2.Open;

     if dtsNfiscdt_2013_2.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013_2);
   end;

   result := 1;
end;

function Tdmgcpcj_base_XI.ObtemOutrosPagamentosDoProcessoAtivas(gcpj,
  tipoato, escritorio, tipoprocesso, codescritorio: string;
  lstCnpjs: TStringList): integer;
var
   i : integer;
begin
   result := -1;

   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                      'from NFISCDT ' +
                      'where nomdes = ''' + tipoato + ''' and ';

   if Pos('JBM', escritorio) <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
                         'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '
   else
   if Pos('CAMARA', escritorio) <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                         'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
   else
   begin
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' CodEscritorio in(';
      for i := 0 to lstCnpjs.Count - 1 do
      begin
         if i <> 0 then
            dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ', ';
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + lstCnpjs.Strings[i];
      end;
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ') and ';
   end;
   dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' bradesco = ' + gcpj;
   dtsNfiscdt.Open;


   If dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                        'from NFISCDT ' +
                        'where nomdes = ''' + tipoato + ''' and ';

     if Pos('JBM', escritorio) <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
                           'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '
     else
     if Pos('CAMARA', escritorio) <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                           'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
     else
     begin
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' CodEscritorio in(';
        for i := 0 to lstCnpjs.Count - 1 do
        begin
           if i <> 0 then
              dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ', ';
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + lstCnpjs.Strings[i];
        end;
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ') and ';
     end;
     dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' bradesco = ' + gcpj;
     dtsNfiscdt_2013.Open;

     if dtsNfiscdt_2013.RecordCount <> 0 then
       dtsNfiscdt.Clone(dtsNfiscdt_2013);
   end;

   If dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                        'from NFISCDT ' +
                        'where nomdes = ''' + tipoato + ''' and ';

     if Pos('JBM', escritorio) <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
                           'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '
     else
     if Pos('CAMARA', escritorio) <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                           'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
     else
     begin
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' CodEscritorio in(';
        for i := 0 to lstCnpjs.Count - 1 do
        begin
           if i <> 0 then
              dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ', ';
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + lstCnpjs.Strings[i];
        end;
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ') and ';
     end;
     dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' bradesco = ' + gcpj;
     dtsNfiscdt_2013_2.Open;

     if dtsNfiscdt_2013_2.RecordCount <> 0 then
       dtsNfiscdt.Clone(dtsNfiscdt_2013_2);
   end;



   result := 1;
end;

procedure Tdmgcpcj_base_XI.ObtemPagamentosLaudoPericial(gcpj,
  codEscritorio: string);
begin
   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select CodEscritorio, NomEscritorio, Bradesco, DatMovto, NomDes, VlrDespesa ' +
                      'from NFISCDT ' +
                      'where Bradesco = ' + gcpj + ' AND NomDes = ''LAUDO PERICIAL''';
//                      ' and  CodEscritorio = ' + codEscritorio + '
   dtsNfiscdt.Open;

   If dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select CodEscritorio, NomEscritorio, Bradesco, DatMovto, NomDes, VlrDespesa ' +
                        'from NFISCDT ' +
                        'where Bradesco = ' + gcpj + ' AND NomDes = ''LAUDO PERICIAL''';
  //                      ' and  CodEscritorio = ' + codEscritorio + '
     dtsNfiscdt_2013.Open;

     if dtsNfiscdt_2013.RecordCount <> 0 then
       dtsNfiscdt.Clone(dtsNfiscdt_2013);
   end;

   If dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select CodEscritorio, NomEscritorio, Bradesco, DatMovto, NomDes, VlrDespesa ' +
                        'from NFISCDT ' +
                        'where Bradesco = ' + gcpj + ' AND NomDes = ''LAUDO PERICIAL''';
  //                      ' and  CodEscritorio = ' + codEscritorio + '
     dtsNfiscdt_2013_2.Open;

     if dtsNfiscdt_2013_2.RecordCount <> 0 then
       dtsNfiscdt.Clone(dtsNfiscdt_2013_2);
   end;

end;

function Tdmgcpcj_base_XI.ObtemQualquerPagamentoDoProcesso(gcpj,
  escritorio, tipoprocesso, codescritorio: string;
  lstCnpjs: TStringList): integer;
var
   i : integer;
begin
   result := -1;
   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                      'from NFISCDT ' +
                      'where ';
   if Pos('JBM', escritorio) <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
                         'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '
   else
   if Pos('CAMARA', escritorio) <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                         'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
   else
   begin
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' CodEscritorio in(';
      for i := 0 to lstCnpjs.Count - 1 do
      begin
         if i <> 0 then
            dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ', ';
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + lstCnpjs.Strings[i];
      end;
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ') and ';
   end;
   dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' bradesco = ' + gcpj;
   dtsNfiscdt.Open;

   If dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                        'from NFISCDT ' +
                        'where ';
     if Pos('JBM', escritorio) <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
                           'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '
     else
     if Pos('CAMARA', escritorio) <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                           'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
     else
     begin
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' CodEscritorio in(';
        for i := 0 to lstCnpjs.Count - 1 do
        begin
           if i <> 0 then
              dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ', ';
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + lstCnpjs.Strings[i];
        end;
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ') and ';
     end;
     dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' bradesco = ' + gcpj;
     dtsNfiscdt_2013.Open;

     if dtsNfiscdt_2013.RecordCount <> 0 then
       dtsNfiscdt.Clone(dtsNfiscdt_2013);

   end;

   If dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select nomescritorio, datmovto, nomdes ' +
                        'from NFISCDT ' +
                        'where ';
     if Pos('JBM', escritorio) <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomescritorio like ''JBM%'' or nomescritorio like ''RAMALHO%'' or nomescritorio like ''ALMEIDA R%'' or ' +
                           'nomescritorio like ''SARAIVA DE%'' or nomescritorio like ''SILAS%'' or nomescritorio like ''D%AVILA%'') and '
     else
     if Pos('CAMARA', escritorio) <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + '(nomescritorio like ''SETTE%'' or nomescritorio like ''AZEVEDO SO%'' or nomescritorio like ''SERRAS E CE%'' or ' +
                           'nomescritorio like ''%BORNHAUSEN%'' or nomescritorio like ''RUBENS SERRA%'') and '
     else
     begin
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' CodEscritorio in(';
        for i := 0 to lstCnpjs.Count - 1 do
        begin
           if i <> 0 then
              dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ', ';
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + lstCnpjs.Strings[i];
        end;
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ') and ';
     end;
     dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' bradesco = ' + gcpj;
     dtsNfiscdt_2013_2.Open;

     if dtsNfiscdt_2013_2.RecordCount <> 0 then
       dtsNfiscdt.Clone(dtsNfiscdt_2013_2);

   end;

   result := 1;
end;

function Tdmgcpcj_base_XI.ObtemSomaJaPagaADescontar(gcpj: string;
  drcContraria: integer; var qtdePaga: integer; codEscritorio: integer; const somenteAto: string = '';
  const excluirVolumetria: boolean = false; const recuperacaoFinal: boolean = false): double;
begin

   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select Sum(VlrDespesa) As TotalPago, count(*) as qtdePaga ' +
                      'from NFISCDT '+
                      'where Bradesco = ' +  gcpj;

   if codEscritorio <> 0 then
      dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and codescritorio = ' + IntToStr(codEscritorio);

   if recuperacaoFinal then
   begin
      if (somenteAto = 'ACORDO') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + 'nomdes in(''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                            '''ADJUDICACAO/ARREMATACAO'', ' + '''LEVANTAMENTO DE ALVARA'')'
      else
      if (somenteAto = 'ENTREGA AMIGAVEL') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                            '''ADJUDICACAO/ARREMATACAO'', ' + '''LEVANTAMENTO DE ALVARA'')'
      else
      if somenteAto = 'CONSOLIDACAO DE PROPRIEDADE' then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' +
                            '''HONORARIOS DE VENDA'', ' + //''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'',
                            '''ADJUDICACAO/ARREMATACAO'', ''LEVANTAMENTO DE ALVARA'')'
      else
      if (somenteAto = 'CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)') or (somenteAto = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' + //''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (somenteAto = 'HONORARIOS DE VENDA') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''ADJUDICACAO/ARREMATACAO'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (somenteAto = 'ADJUDICACAO/ARREMATACAO') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                            '''LEVANTAMENTO DE ALVARA'')'
      else
      if (somenteAto = 'LEVANTAMENTO DE ALVARA') then
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                            '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                            '''ADJUDICACAO/ARREMATACAO'')';
   end
   else
   begin
      if somenteAto = '' then
      begin
         if drcContraria <> 1 then
         begin
            if excluirVolumetria then
               dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + // DatMovto <= :umAno And ' +
                               'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''HONORARIOS DE VENDA'', ' +
                               '''BUSCA E APREENS√O/REINTEGRA«√O DE POSSE'',  ''ADJUDICACAO/ARREMATACAO'', ' +
                               '''PENHORA'', ''BUSCA E APREENSAO/REINTEGRACAO DE POSSE'')'

                               //retirado conforme e-mail Samara - 17/01/2017
                               //''IRRECUPERABILIDADE ATE 1 ANO'', ''IRRECUPERABILIDADE APOS 1 ANO'')'
            else
               dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and ' + // DatMovto <= :umAno And ' +
                               'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''HONORARIOS DE VENDA'', ''AJUIZAMENTO'', ' +
                               '''BUSCA E APREENS√O/REINTEGRA«√O DE POSSE'', ''ADJUDICACAO/ARREMATACAO'', ' +
                               '''PENHORA'', ''BUSCA E APREENSAO/REINTEGRACAO DE POSSE'')';

                               //retirado conforme e-mail Samara - 17/01/2017
                               //''IRRECUPERABILIDADE ATE 1 ANO'', ''IRRECUPERABILIDADE APOS 1 ANO'')';
         end
         else
            dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and nomdes <> ''LAUDO PERICIAL''';
      end
      else
         dtsNfiscdt.CommandText := dtsNfiscdt.CommandText + ' and  ' +
                            'nomdes = ''' + somenteAto + '''';
   end;

   dtsNfiscdt.Open;


   If dtsNfiscdt.FieldByName('qtdePaga').AsInteger = 0 then
//   If dtsNfiscdt.RecordCount = 0 then
   begin

     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select Sum(VlrDespesa) As TotalPago, count(*) as qtdePaga ' +
                        'from NFISCDT '+
                        'where Bradesco = ' +  gcpj;

     if codEscritorio <> 0 then
        dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and codescritorio = ' + IntToStr(codEscritorio);

     if recuperacaoFinal then
     begin
        if (somenteAto = 'ACORDO') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + 'nomdes in(''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''ADJUDICACAO/ARREMATACAO'', ' + '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'ENTREGA AMIGAVEL') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''ADJUDICACAO/ARREMATACAO'', ' + '''LEVANTAMENTO DE ALVARA'')'
        else
        if somenteAto = 'CONSOLIDACAO DE PROPRIEDADE' then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' +
                              '''HONORARIOS DE VENDA'', ' + //''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'',
                              '''ADJUDICACAO/ARREMATACAO'', ''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)') or (somenteAto = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' + //''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'', ' +
                              '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'HONORARIOS DE VENDA') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''ADJUDICACAO/ARREMATACAO'', ' +
                              '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'ADJUDICACAO/ARREMATACAO') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'LEVANTAMENTO DE ALVARA') then
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''ADJUDICACAO/ARREMATACAO'')';
     end
     else
     begin
        if somenteAto = '' then
        begin
           if drcContraria <> 1 then
           begin
              if excluirVolumetria then
                 dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + // DatMovto <= :umAno And ' +
                                 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''HONORARIOS DE VENDA'', ' +
                                 '''BUSCA E APREENS√O/REINTEGRA«√O DE POSSE'',  ''ADJUDICACAO/ARREMATACAO'', ' +
                                 '''PENHORA'', ''BUSCA E APREENSAO/REINTEGRACAO DE POSSE'')'

                                 //retirado conforme e-mail Samara - 17/01/2017
                                 //''IRRECUPERABILIDADE ATE 1 ANO'', ''IRRECUPERABILIDADE APOS 1 ANO'')'
              else
                 dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and ' + // DatMovto <= :umAno And ' +
                                 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''HONORARIOS DE VENDA'', ''AJUIZAMENTO'', ' +
                                 '''BUSCA E APREENS√O/REINTEGRA«√O DE POSSE'', ''ADJUDICACAO/ARREMATACAO'', ' +
                                 '''PENHORA'', ''BUSCA E APREENSAO/REINTEGRACAO DE POSSE'')';

                                 //retirado conforme e-mail Samara - 17/01/2017
                                 //''IRRECUPERABILIDADE ATE 1 ANO'', ''IRRECUPERABILIDADE APOS 1 ANO'')';
           end
           else
              dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and nomdes <> ''LAUDO PERICIAL''';
        end
        else
           dtsNfiscdt_2013.CommandText := dtsNfiscdt_2013.CommandText + ' and  ' +
                              'nomdes = ''' + somenteAto + '''';
     end;

     dtsNfiscdt_2013.Open;

     If dtsNfiscdt_2013.RecordCount <> 0 then
     dtsNfiscdt.Clone(dtsNfiscdt_2013);

   end;

   If dtsNfiscdt.FieldByName('qtdePaga').AsInteger = 0 then
//   If dtsNfiscdt.RecordCount = 0 then
   begin

     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select Sum(VlrDespesa) As TotalPago, count(*) as qtdePaga ' +
                        'from NFISCDT '+
                        'where Bradesco = ' +  gcpj;

     if codEscritorio <> 0 then
        dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and codescritorio = ' + IntToStr(codEscritorio);

     if recuperacaoFinal then
     begin
        if (somenteAto = 'ACORDO') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + 'nomdes in(''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''ADJUDICACAO/ARREMATACAO'', ' + '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'ENTREGA AMIGAVEL') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''ADJUDICACAO/ARREMATACAO'', ' + '''LEVANTAMENTO DE ALVARA'')'
        else
        if somenteAto = 'CONSOLIDACAO DE PROPRIEDADE' then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' +
                              '''HONORARIOS DE VENDA'', ' + //''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'',
                              '''ADJUDICACAO/ARREMATACAO'', ''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)') or (somenteAto = 'CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ' + //''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''HONORARIOS DE VENDA'', ''ADJUDICACAO/ARREMATACAO'', ' +
                              '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'HONORARIOS DE VENDA') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''ADJUDICACAO/ARREMATACAO'', ' +
                              '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'ADJUDICACAO/ARREMATACAO') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''LEVANTAMENTO DE ALVARA'')'
        else
        if (somenteAto = 'LEVANTAMENTO DE ALVARA') then
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''CONSOLIDACAO DE PROPRIEDADE'', ' +
                              '''CONSOLIDACAO DE PROPRIEDADE (2 P/C ADIC)'', ''CONSOLIDACAO PROPRIEDADE (2 P/C ADIC)'', ''HONORARIOS DE VENDA'', ' +
                              '''ADJUDICACAO/ARREMATACAO'')';
     end
     else
     begin
        if somenteAto = '' then
        begin
           if drcContraria <> 1 then
           begin
              if excluirVolumetria then
                 dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + // DatMovto <= :umAno And ' +
                                 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''HONORARIOS DE VENDA'', ' +
                                 '''BUSCA E APREENS√O/REINTEGRA«√O DE POSSE'',  ''ADJUDICACAO/ARREMATACAO'', ' +
                                 '''PENHORA'', ''BUSCA E APREENSAO/REINTEGRACAO DE POSSE'')'

                                 //retirado conforme e-mail Samara - 17/01/2017
                                 //''IRRECUPERABILIDADE ATE 1 ANO'', ''IRRECUPERABILIDADE APOS 1 ANO'')'
              else
                 dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and ' + // DatMovto <= :umAno And ' +
                                 'nomdes in(''ACORDO'', ''ENTREGA AMIGAVEL'', ''HONORARIOS DE VENDA'', ''AJUIZAMENTO'', ' +
                                 '''BUSCA E APREENS√O/REINTEGRA«√O DE POSSE'', ''ADJUDICACAO/ARREMATACAO'', ' +
                                 '''PENHORA'', ''BUSCA E APREENSAO/REINTEGRACAO DE POSSE'')';

                                 //retirado conforme e-mail Samara - 17/01/2017
                                 //''IRRECUPERABILIDADE ATE 1 ANO'', ''IRRECUPERABILIDADE APOS 1 ANO'')';
           end
           else
              dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and nomdes <> ''LAUDO PERICIAL''';
        end
        else
           dtsNfiscdt_2013_2.CommandText := dtsNfiscdt_2013_2.CommandText + ' and  ' +
                              'nomdes = ''' + somenteAto + '''';
     end;

     dtsNfiscdt_2013_2.Open;

     If dtsNfiscdt_2013_2.RecordCount <> 0 then
     dtsNfiscdt.Clone(dtsNfiscdt_2013_2);

   end;

   qtdePaga := dtsNfiscdt.FieldByName('qtdePaga').AsInteger;
   result := dtsNfiscdt.FieldByName('TotalPago').AsFloat;

end;

function Tdmgcpcj_base_XI.PagouAcordoNestaData(gcpj: string;
  dataAto: TdateTime): boolean;
begin
   result := false;

   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select datOcorr, nomdes ' +
                      'from NFISCDT ' +
                      'where nomdes = ''ACORDO'' and bradesco = ' + gcpj + ' and datOcorr = ' +DatetoStr(dataAto);

//           'where nomdes = ''ACORDO'' and bradesco = ' + gcpj + ' and datOcorr = :dataAto';
//           dtsNfiscdt.Parameters.ParamByName('dataAto').Value := dataAto;


   dtsNfiscdt.Open;
   if dtsNfiscdt.RecordCount > 0 then //achou pagamento da mesma data
      result := true;


   if dtsNfiscdt.RecordCount = 0 then //achou pagamento da mesma data
   begin
     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select datOcorr, nomdes ' +
                       'from NFISCDT ' +
                       'where nomdes = ''ACORDO'' and bradesco = ' + gcpj + ' and datOcorr = ' +DatetoStr(dataAto);

//                    'where nomdes = ''ACORDO'' and bradesco = ' + gcpj + ' and datOcorr = :dataAto';
//                     dtsNfiscdt_2013.Parameters.ParamByName('dataAto').Value := dataAto;

     dtsNfiscdt_2013.Open;
     if dtsNfiscdt_2013.RecordCount > 0 then //achou pagamento da mesma data
        result := true;


     If dtsNfiscdt_2013.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013);
   end;

   if dtsNfiscdt.RecordCount = 0 then //achou pagamento da mesma data
   begin
     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select datOcorr, nomdes ' +
                       'from NFISCDT ' +
                       'where nomdes = ''ACORDO'' and bradesco = ' + gcpj + ' and datOcorr = ' +DatetoStr(dataAto);

//                    'where nomdes = ''ACORDO'' and bradesco = ' + gcpj + ' and datOcorr = :dataAto';
//                     dtsNfiscdt_2013_2.Parameters.ParamByName('dataAto').Value := dataAto;

     dtsNfiscdt_2013_2.Open;
     if dtsNfiscdt_2013_2.RecordCount > 0 then //achou pagamento da mesma data
        result := true;


     If dtsNfiscdt_2013_2.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013_2);
   end;


end;

procedure Tdmgcpcj_base_XI.DataModuleCreate(Sender: TObject);
var
   ini : TiniFile;
begin
   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConnNfiscal.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJ_NFISCAL', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConnNfiscdt.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJ_NFISCDT', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConnNfiscdt_2013.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJ_NFISCDT_Ate_072013', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;

   ini := TIniFile.Create(ExtractFilePath(Application.exename) + 'config.ini');
   try
      adoConnNfiscdt_2013_2.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=' + ini.readstring('BaseDiaria_GCPJ_NFISCDT_De_2013_08_Ate_2015_11', 'databasename', '') + ';Persist Security Info=True';
   finally
      ini.free;
   end;
end;

function Tdmgcpcj_base_XI.PagouTrabalhistaNestaData(andamento,gcpj: string;
  dataAto: TDateTime): boolean;
var
   tipo : string;
begin
   result := false;

   if andamento = 'PARCELA' then
      tipo := 'ACOMPANHAMENTO'
   else
   if andamento = 'TUTELA' then
      tipo := 'CONTESTACAO'
   else
      tipo := andamento;

   dtsNfiscdt.Close;
   dtsNfiscdt.Parameters.Clear;
   dtsNfiscdt.CommandText := 'select datOcorr, nomdes ' +
                      'from NFISCDT ' +
                      'where nomdes = ''' + tipo + ''' and bradesco = ' + gcpj + ' and datOcorr = :dataAto';
   dtsNfiscdt.Parameters.ParamByName('dataAto').Value := dataAto;
   dtsNfiscdt.Open;
   if dtsNfiscdt.RecordCount > 0 then //achou pagamento da mesma data
      result := true;

   if dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013.CommandText := '';
     if andamento = 'PARCELA' then
        tipo := 'ACOMPANHAMENTO'
     else
     if andamento = 'TUTELA' then
        tipo := 'CONTESTACAO'
     else
        tipo := andamento;

     dtsNfiscdt_2013.Close;
     dtsNfiscdt_2013.Parameters.Clear;
     dtsNfiscdt_2013.CommandText := 'select datOcorr, nomdes ' +
                        'from NFISCDT ' +
                        'where nomdes = ''' + tipo + ''' and bradesco = ' + gcpj + ' and datOcorr = :dataAto';
     dtsNfiscdt_2013.Parameters.ParamByName('dataAto').Value := dataAto;
     dtsNfiscdt_2013.Open;

     If dtsNfiscdt_2013.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013);
   end;

   if dtsNfiscdt.RecordCount = 0 then
   begin
     dtsNfiscdt_2013_2.CommandText := '';
     if andamento = 'PARCELA' then
        tipo := 'ACOMPANHAMENTO'
     else
     if andamento = 'TUTELA' then
        tipo := 'CONTESTACAO'
     else
        tipo := andamento;

     dtsNfiscdt_2013_2.Close;
     dtsNfiscdt_2013_2.Parameters.Clear;
     dtsNfiscdt_2013_2.CommandText := 'select datOcorr, nomdes ' +
                        'from NFISCDT ' +
                        'where nomdes = ''' + tipo + ''' and bradesco = ' + gcpj + ' and datOcorr = :dataAto';
     dtsNfiscdt_2013_2.Parameters.ParamByName('dataAto').Value := dataAto;
     dtsNfiscdt_2013_2.Open;

     If dtsNfiscdt_2013_2.RecordCount <> 0 then
      dtsNfiscdt.Clone(dtsNfiscdt_2013_2);
   end;


end;

end.
