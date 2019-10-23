program digita_honorarios_III;



uses
  Forms,
  MAIN in 'MAIN.PAS' {MainForm},
  about in 'about.pas' {AboutBox},
  fConfigWintask in 'fConfigWintask.pas' {fConfiguraWintask},
  fconfigDb in 'fconfigDb.pas' {frmConfigDb},
  maindir in '..\..\..\sisba32b\maindir.pas' {frmDir},
  datamodule_honorarios in 'datamodule_honorarios.pas' {dmHonorarios: TDataModule},
  Mywin in '..\..\..\sisba32b\mywin.pas',
  fimportaplan in 'fimportaplan.pas' {frmImpPlan},
  Presenta in '..\..\..\sisba32b\Presenta.pas',
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
  datamodulegcpj_base_V in '\\D4040S006\D4040_01\Publico\Backup_sistemas\Compartilhado\Honorarios_Programa\presenta\Projeto\utils\bradesco - honorarios\digitador\datamodulegcpj_base_V.pas' {dmgcpcj_base_v: TDataModule};

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
  Application.Run;
end.
