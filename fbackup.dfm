object frmBackup: TfrmBackup
  Left = 985
  Top = 238
  Width = 311
  Height = 360
  ActiveControl = SpinEdit1
  Caption = 'Backup'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poScreenCenter
  PrintScale = poNone
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 56
    Width = 88
    Height = 13
    Caption = 'Backup da Base.: '
  end
  object Label2: TLabel
    Left = 16
    Top = 83
    Width = 183
    Height = 13
    Caption = 'Manter no banco de dados  os '#250'ltimos '
  end
  object Label3: TLabel
    Left = 246
    Top = 84
    Width = 30
    Height = 13
    Caption = 'meses'
  end
  object Gauge1: TGauge
    Left = 10
    Top = 57
    Width = 271
    Height = 38
    ForeColor = clNavy
    Progress = 0
    Visible = False
  end
  object Image1: TImage
    Left = 48
    Top = -77
    Width = 185
    Height = 126
  end
  object Label4: TLabel
    Left = 20
    Top = 19
    Width = 113
    Height = 29
    Cursor = crDrag
    Caption = 'Aguarde..'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -24
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Button1: TButton
    Left = 25
    Top = 123
    Width = 97
    Height = 25
    Caption = 'Iniciar backup'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 192
    Top = 123
    Width = 75
    Height = 25
    Caption = 'Sair'
    TabOrder = 1
    OnClick = Button2Click
  end
  object SpinEdit1: TSpinEdit
    Left = 200
    Top = 76
    Width = 41
    Height = 22
    MaxValue = 0
    MinValue = 0
    TabOrder = 2
    Value = 0
  end
  object ProgressBar1: TProgressBar
    Left = 24
    Top = 184
    Width = 217
    Height = 17
    Min = 0
    Max = 100
    TabOrder = 3
  end
end