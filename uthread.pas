unit uthread;

interface

uses
  Classes,fbackup;

type
  Piscando = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure Piscando.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ Piscando }

procedure Piscando.Execute;
begin
//   frmBackup.Timer1.Enabled := True;
  { Place thread code here }
end;

end.
 