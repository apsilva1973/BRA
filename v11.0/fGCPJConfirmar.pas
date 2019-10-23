unit fGCPJConfirmar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, SHDocVw, Func_Wintask_Obj, uplanjbm, mshtml, ExtCtrls;

type
  TfrmGcpjConfirmar = class(TForm)
    wb: TWebBrowser;
    procedure FormCreate(Sender: TObject);
  private
  public
    { Public declarations }
  end;

implementation

{$R *.dfm}


procedure TfrmGcpjConfirmar.FormCreate(Sender: TObject);
begin
   Show;
end;

end.
