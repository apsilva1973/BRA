program digita_honorarios;



uses
  Forms,
  MAIN in 'MAIN.PAS' {MainForm},
  fConfigWintask in 'fConfigWintask.pas' {fConfiguraWintask},
  fconfigDb in 'fconfigDb.pas' {frmConfigDb},
  datamodule_honorarios in 'datamodule_honorarios.pas' {dmHonorarios: TDataModule},
  fimportaplan in 'fimportaplan.pas' {frmImpPlan},
  maindir in '..\..\..\sisba32b\maindir.pas' {frmDir},
  Mywin in '..\..\..\sisba32b\mywin.pas',
  Presenta in '..\..\..\sisba32b\Presenta.pas',
  uPresentaFunc in '..\..\..\sisba32b\uPresentaFunc.pas',
  uPlanJBM in 'uPlanJBM.pas',
  mgenlib in '..\..\..\sisba32b\mgenlib.pas',
  frelatorios in 'frelatorios.pas' {frmRelatorios},
  datamodulegcpj_base_i in 'datamodulegcpj_base_i.pas' {dmgcpcj_base_I: TDataModule},
  datamodulegcpj_base_iv in 'datamodulegcpj_base_iv.pas' {dmgcpj_base_iv: TDataModule},
  datamodulegcpj_base_migrados in 'datamodulegcpj_base_migrados.pas' {dmGcpj_migrados: TDataModule},
  datamodulegcpj_compartilhado in 'datamodulegcpj_compartilhado.pas' {dmgcpj_compartilhado: TDataModule},
  datamodulegcpj_base_iii in 'datamodulegcpj_base_iii.pas' {dmgcpj_base_iii: TDataModule},
  datamodulegcpj_base_ii in 'datamodulegcpj_base_ii.pas' {dmgcpj_base_ii: TDataModule},
  Func_Wintask_Obj in '..\..\..\wintask\Func_Wintask_Obj.pas',
  fvalidaplan in 'fvalidaplan.pas' {frmValidaPlan},
  fcaddatainicial in 'fcaddatainicial.pas' {frmCadDataInicial},
  fdigitaplan in 'fdigitaplan.pas' {frmDigitaPlan},
  calccpfcgc in '..\..\..\sisba32b\calccpfcgc.pas',
  fcadescritorio in 'fcadescritorio.pas' {frmCadEscritorio},
  datamodulegcpj_base_V in 'datamodulegcpj_base_V.pas' {dmgcpcj_base_v: TDataModule},
  fGCPJConfirmar in 'fGCPJConfirmar.pas' {frmGcpjConfirmar},
  datamodulegcpj_base_ix in 'datamodulegcpj_base_iX.pas' {dmgcpj_base_IX: TDataModule},
  datamodulegcpj_trabalhistas in 'datamodulegcpj_trabalhistas.pas' {dmgcpj_trabalhistas: TDataModule},
  datamodulegcpj_base_X in 'datamodulegcpj_base_X.pas' {dmgcpj_base_X: TDataModule},
  datamodulegcpj_base_xi in 'datamodulegcpj_base_xi.pas' {dmgcpcj_base_XI: TDataModule},
  datamodulegcpj_baixados in 'datamodulegcpj_baixados.pas' {dmgcpj_baixados: TDataModule},
  datamodulegcpj_recuperados in 'datamodulegcpj_recuperados.pas' {dmgcpj_recuperados: TDataModule},
  datamodulegcpj_base_VII in 'datamodulegcpj_base_VII.pas' {dmgcpj_base_vii: TDataModule},
  datamodulegcpj_base_viii in 'datamodulegcpj_base_viii.pas' {dmgcpj_base_VIII: TDataModule},
  fGCPJ in 'fGCPJ.pas' {frmGcpj};
//  about in 'about.pas' {AboutBox},
//  datamodulegcpj_base_volumetria in 'datamodulegcpj_base_volumetria.pas' {dmgcpj_base_volumetria: TDataModule},
//  uCadVolumetria in 'uCadVolumetria.pas' {frmCadVolumetria};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Digita Notas no GCPJ';
  Application.CreateForm(TdmHonorarios, dmHonorarios);
  Application.CreateForm(Tdmgcpcj_base_I, dmgcpcj_base_I);
  Application.CreateForm(Tdmgcpj_base_iv, dmgcpj_base_iv);
  Application.CreateForm(TdmGcpj_migrados, dmGcpj_migrados);
  Application.CreateForm(Tdmgcpj_compartilhado, dmgcpj_compartilhado);
  Application.CreateForm(Tdmgcpj_base_iii, dmgcpj_base_iii);
  Application.CreateForm(Tdmgcpj_base_ii, dmgcpj_base_ii);
  Application.CreateForm(Tdmgcpcj_base_v, dmgcpcj_base_v);
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(Tdmgcpj_base_IX, dmgcpj_base_IX);
  Application.CreateForm(Tdmgcpj_trabalhistas, dmgcpj_trabalhistas);
  Application.CreateForm(Tdmgcpj_base_X, dmgcpj_base_X);
  Application.CreateForm(Tdmgcpcj_base_XI, dmgcpcj_base_XI);
  Application.CreateForm(Tdmgcpj_baixados, dmgcpj_baixados);
  Application.CreateForm(Tdmgcpj_recuperados, dmgcpj_recuperados);
  Application.CreateForm(Tdmgcpj_base_vii, dmgcpj_base_vii);
  Application.CreateForm(Tdmgcpj_base_VIII, dmgcpj_base_VIII);
  Application.CreateForm(TfrmDir, frmDir);
//  Application.CreateForm(TAboutBox, AboutBox);
//  Application.CreateForm(Tdmgcpj_base_volumetria, dmgcpj_base_volumetria);
  Application.Run;
end.
