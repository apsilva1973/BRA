{ Este exemplo foi baixado no site www.andrecelestino.com
  Passe por lá a qualquer momento para conferir os novos artigos! :)
  contato@andrecelestino.com }

unit untThread;

interface

uses
  classes, SysUtils, Windows, ComCtrls,Forms;

type
  TMinhaThread = class(TThread)
  private
    FTempo: integer;
    FProgressBar: TProgressBar;
    FForm       : TForm;
  public
    constructor Create; overload;

    procedure Progress;
    procedure FormAtualiza;
    procedure Execute; override;

    property Tempo: Integer read FTempo write FTempo;
    property ProgressBar: TProgressBar read FProgressBar write FProgressBar;
    property Form       : TForm        read FForm        write FForm;
  end;

implementation

constructor TMinhaThread.Create;
begin
  inherited Create(True);
end;

procedure TMinhaThread.Execute;
begin
  // loop infinito
  while True do
  begin
//    Application.ProcessMessages;
    Sleep(FTempo);
    Synchronize(Progress);
    Synchronize(FormAtualiza);
//    Application.ProcessMessages;
  end;
end;

procedure TMinhaThread.FormAtualiza;
begin
  // atualiza o Form a cada iteração do método 'Execute'
  Form.Refresh;
  Application.ProcessMessages;
end;

procedure TMinhaThread.Progress;
begin
  // atualiza a ProgressBar a cada iteração do método 'Execute'
  FProgressBar.StepIt;
end;


end.
