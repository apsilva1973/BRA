unit fcadastravaloresnaoatualizar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, datamodule_honorarios, StdCtrls, Mask, DBCtrls, Buttons, Grids,
  DBGrids, fassociaescritorios;

type
  TfrmCadastraValoresNaoAtualizar = class(TForm)
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    Label1: TLabel;
    BitBtn1: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    DBEdit1: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    DBEdit2: TDBEdit;
    Label4: TLabel;
    DBEdit3: TDBEdit;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn10: TBitBtn;
    BitBtn11: TBitBtn;
    Label6: TLabel;
    DBEdit5: TDBEdit;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure BitBtn10Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCadastraValoresNaoAtualizar: TfrmCadastraValoresNaoAtualizar;

implementation

{$R *.dfm}

procedure TfrmCadastraValoresNaoAtualizar.BitBtn2Click(Sender: TObject);
begin
   bitbtn4.enabled := false;
   bitbtn5.Enabled := false;
   bitbtn2.Enabled := false;

   BitBtn6.Enabled := true;
   BitBtn7.Enabled := true;

   DBEdit2.Enabled := true;
   DBEdit3.Enabled := true;

   dmHonorarios.dtsValoresNaoAtualizar.Insert;

   dbedit2.SetFocus;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn6Click(Sender: TObject);
begin
   dmHonorarios.dtsValoresNaoAtualizar.Cancel;

   bitbtn2.Enabled := true;
   BitBtn4.Enabled := (dmHonorarios.dtsValoresNaoAtualizar.RecordCount <> 0);
   BitBtn5.Enabled := (dmHonorarios.dtsValoresNaoAtualizar.RecordCount <> 0);

   BitBtn6.enabled := false;
   BitBtn7.enabled := false;

   DBEdit1.Enabled := false;
   DBEdit2.Enabled := false;
   DBEdit3.Enabled := false;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn7Click(Sender: TObject);
begin
   dmHonorarios.dtsValoresNaoAtualizar.Tag := 0;

   dmHonorarios.dtsValoresNaoAtualizar.Post;

   if dmHonorarios.dtsValoresNaoAtualizar.Tag = 1 then
   begin
      bitbtn2.Enabled := true;
      BitBtn4.Enabled := (dmHonorarios.dtsValoresNaoAtualizar.RecordCount <> 0);
      BitBtn5.Enabled := (dmHonorarios.dtsValoresNaoAtualizar.RecordCount <> 0);

      BitBtn6.enabled := false;
      BitBtn7.enabled := false;

      DBEdit1.Enabled := false;
      DBEdit2.Enabled := false;
      DBEdit3.Enabled := false;
   end;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn5Click(Sender: TObject);
begin
   bitbtn4.enabled := false;
   bitbtn5.Enabled := false;
   bitbtn2.Enabled := false;

   BitBtn6.Enabled := true;
   BitBtn7.Enabled := true;

   DBEdit2.Enabled := true;
   DBEdit3.Enabled := true;

   dmHonorarios.dtsValoresNaoAtualizar.Edit;

   dbedit2.SetFocus;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn1Click(Sender: TObject);
begin
   BitBtn10.Enabled := true;
   BitBtn11.Enabled := true;

   DBEdit5.Enabled := true;

   dmHonorarios.dtsTiposNaoAtualizar.Insert;

   DBEdit5.setfocus;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn3Click(Sender: TObject);
begin
   if MessageDlg('Deseja realmente excluir o controle de valores que não devem ser atualizados?', mtConfirmation, [mbyes, mbno], 0) <> mrYes then
      exit;

   dmHonorarios.ExcluiValoresNaoAtualizarFK(dmHonorarios.dtsTiposNaoAtualizar.FieldByName('identificador').AsInteger);
   dmHonorarios.ExcluiTiposNaoAtualizar(dmHonorarios.dtsTiposNaoAtualizar.FieldByName('identificador').AsInteger);

   dmHonorarios.ObtemTiposNaoAtualizarCadastrados;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn4Click(Sender: TObject);
begin
   if MessageDlg('Deseja realmente excluir o controle de valores que não devem ser atualizados?', mtConfirmation, [mbyes, mbno], 0) <> mrYes then
      exit;

   dmHonorarios.ExcluiValoresNaoAtualizarPK(dmHonorarios.dtsValoresNaoAtualizar.FieldByName('identificador').AsInteger,
                                            dmHonorarios.dtsValoresNaoAtualizar.FieldByName('tipoandamento').AsString);
   dmHonorarios.dtsValoresNaoAtualizar.Close;
   dmHonorarios.dtsValoresNaoAtualizar.Open;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn8Click(Sender: TObject);
begin
   frmAssociaEscritorios := TfrmAssociaEscritorios.Create(nil);
   try
      frmAssociaEscritorios.tpNaoAtualizar.Text := dmHonorarios.dtsValoresNaoAtualizar.FieldByName('identificador').AsString;
      dmHonorarios.ObtemEscritoriosAssociadosAoTipo(dmHonorarios.dtsValoresNaoAtualizar.FieldByName('identificador').AsInteger);
      dmHonorarios.ObtemEscritoriosNaoAssociadosAoTipo(dmHonorarios.dtsValoresNaoAtualizar.FieldByName('identificador').AsInteger, '');
      frmAssociaEscritorios.ShowModal;
   finally
      frmAssociaEscritorios.Free;
   end;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn10Click(Sender: TObject);
begin
   dmHonorarios.dtsTiposNaoAtualizar.Cancel;

   bitbtn1.Enabled := true;
   BitBtn3.Enabled := (dmHonorarios.dtsTiposNaoAtualizar.RecordCount <> 0);
   BitBtn9.Enabled := (dmHonorarios.dtsTiposNaoAtualizar.RecordCount <> 0);

   BitBtn10.enabled := false;
   BitBtn11.enabled := false;

   DBEdit5.Enabled := false;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn11Click(Sender: TObject);
begin
   dmHonorarios.dtsTiposNaoAtualizar.Tag := 0;

   dmHonorarios.dtsTiposNaoAtualizar.Post;

   if dmHonorarios.dtsTiposNaoAtualizar.Tag = 1 then
   begin
      bitbtn1.Enabled := true;
      BitBtn3.Enabled := (dmHonorarios.dtsTiposNaoAtualizar.RecordCount <> 0);
      BitBtn9.Enabled := (dmHonorarios.dtsTiposNaoAtualizar.RecordCount <> 0);

      BitBtn10.enabled := false;
      BitBtn11.enabled := false;

      DBEdit5.Enabled := false;
   end;
end;

procedure TfrmCadastraValoresNaoAtualizar.BitBtn9Click(Sender: TObject);
begin
   bitbtn3.enabled := false;
   bitbtn9.Enabled := false;
   bitbtn1.Enabled := false;

   BitBtn10.Enabled := true;
   BitBtn11.Enabled := true;

   DBEdit5.Enabled := true;

   dmHonorarios.dsTiposNaoAtualizar.Edit;

   dbedit5.SetFocus;

end;

end.
