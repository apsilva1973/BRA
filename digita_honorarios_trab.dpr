program digita_honorarios_trab;



uses
  Forms,
  MAIN in 'MAIN.PAS' {MainForm},
  about in 'about.pas' {AboutBox},
  fConfigWintask in 'fConfigWintask.pas' {fConfiguraWintask},
  fconfigDb in 'fconfigDb.pas' {frmConfigDb},
  maindir in '..\..\..\..\sisba32b\maindir.pas' {frmDir},
  datamodule_honorarios in 'datamodule_honorarios.pas' {dmHonorarios: TDataModule},
  Mywin in '..\..\..\..\sisba32b\mywin.pas',
  fimportaplan in 'fimportaplan.pas' {frmImpPlan},
  Presenta in '..\..\..\..\sisba32b\Presenta.pas',
  uPlanJBM in 'uPlanJBM.pas',
  mgenlib in '..\..\..\..\sisba32b\mgenlib.pas',
  frelatorios in 'frelatorios.pas' {frmRelatorios},
  datamodulegcpj_base_i in 'datamodulegcpj_base_i.pas' {dmgcpcj_base_I: TDataModule},
  datamodulegcpj_base_iv in 'datamodulegcpj_base_iv.pas' {dmgcpj_base_iv: TDataModule},
  datamodulegcpj_base_migrados in 'datamodulegcpj_base_migrados.pas' {dmGcpj_migrados: TDataModule},
  datamodulegcpj_trabalhistas in 'datamodulegcpj_trabalhistas.pas' {dmgcpj_trabalhistas: TDataModule},
  datamodulegcpj_base_iii in 'datamodulegcpj_base_iii.pas' {dmgcpj_base_iii: TDataModule},
  datamodulegcpj_base_ii in 'datamodulegcpj_base_ii.pas' {dmgcpj_base_ii: TDataModule},
  Func_Wintask_Obj in '..\..\..\..\wintask\Func_Wintask_Obj.pas',
  fvalidaplan in 'fvalidaplan.pas' {frmValidaPlan},
  fcaddatainicial in 'fcaddatainicial.pas' {frmCadDataInicial},
  fdigitaplan in 'fdigitaplan.pas' {frmDigitaPlan},
  calccpfcgc in '..\..\..\..\sisba32b\calccpfcgc.pas',
  fcadescritorio in 'fcadescritorio.pas' {frmCadEscritorio},
  datamodulegcpj_base_V in 'datamodulegcpj_base_V.pas' {dmgcpcj_base_v: TDataModule},
  fGCPJ in 'fGCPJ.pas' {frmGcpj},
  fGCPJConfirmar in 'fGCPJConfirmar.pas' {frmGcpjConfirmar},
  datamodulegcpj_base_VII in 'datamodulegcpj_base_VII.pas' {dmgcpj_base_vii: TDataModule},
  fAtosPendentes in 'fAtosPendentes.pas' {frmAtosPendentes},
  fcadastravaloresnaoatualizar in 'fcadastravaloresnaoatualizar.pas' {frmCadastraValoresNaoAtualizar},
  fassociaescritorios in 'fassociaescritorios.pas' {frmAssociaEscritorios},
  fcadadvogadointerno in 'fcadadvogadointerno.pas' {frmCadAdvogadoInterno},
  datamodulegcpj_base_viii in 'datamodulegcpj_base_viii.pas' {dmgcpj_base_VIII: TDataModule},
  fEnviaGcpj in 'fEnviaGcpj.pas' {frmEnviaGcpj},
  BradISD in '..\..\..\..\sisba32b\BradISD.pas',
  isdcdret in '..\..\..\..\sisba32b\isdcdret.pas',
  ISDTX32_NOVO in '..\..\..\..\sisba32b\ISDTX32_NOVO.pas',
  datamodulegcpj_baixados in 'datamodulegcpj_baixados.pas' {dmgcpj_baixados: TDataModule},
  uPresentaFunc in '..\..\..\..\sisba32b\uPresentaFunc.pas',
  datamodulegcpj_base_ix in 'datamodulegcpj_base_iX.pas' {dmgcpj_base_IX: TDataModule},
  datamodulegcpj_base_X in 'datamodulegcpj_base_X.pas' {dmgcpj_base_X: TDataModule},
  datamodulegcpj_recuperados in 'datamodulegcpj_recuperados.pas' {dmgcpj_recuperados: TDataModule},
  datamodulegcpj_base_xi in 'datamodulegcpj_base_xi.pas' {dmgcpcj_base_XI: TDataModule},
  datamodulegcpj_compartilhado in 'datamodulegcpj_compartilhado.pas' {dmgcpj_compartilhado: TDataModule},
  fbackup in 'fbackup.pas' {frmBackup},
  uthread in 'uthread.pas',
  untThread in 'untThread.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Digita Notas no GCPJ';
  Application.CreateForm(TdmHonorarios, dmHonorarios);
  Application.CreateForm(Tdmgcpcj_base_I, dmgcpcj_base_I);
  Application.CreateForm(Tdmgcpj_base_iv, dmgcpj_base_iv);
  Application.CreateForm(Tdmgcpj_baixados, dmgcpj_baixados);
  Application.CreateForm(TdmGcpj_migrados, dmGcpj_migrados);
  Application.CreateForm(Tdmgcpj_trabalhistas, dmgcpj_trabalhistas);
  Application.CreateForm(Tdmgcpj_base_iii, dmgcpj_base_iii);
  Application.CreateForm(Tdmgcpj_base_ii, dmgcpj_base_ii);
  Application.CreateForm(Tdmgcpcj_base_v, dmgcpcj_base_v);
  Application.CreateForm(Tdmgcpj_base_vii, dmgcpj_base_vii);
  Application.CreateForm(Tdmgcpj_base_VIII, dmgcpj_base_VIII);
  Application.CreateForm(Tdmgcpj_base_IX, dmgcpj_base_IX);
  Application.CreateForm(Tdmgcpj_base_X, dmgcpj_base_X);
  Application.CreateForm(Tdmgcpj_recuperados, dmgcpj_recuperados);
  Application.CreateForm(Tdmgcpcj_base_XI, dmgcpcj_base_XI);
  Application.CreateForm(Tdmgcpj_compartilhado, dmgcpj_compartilhado);
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
end.
