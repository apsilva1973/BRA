object frmCadEscritorio: TfrmCadEscritorio
  Left = 143
  Top = 115
  Width = 1085
  Height = 439
  Caption = 'Cadastra Escrit'#243'rios'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 464
    Top = 80
    Width = 593
    Height = 273
    Shape = bsFrame
  end
  object Label1: TLabel
    Left = 643
    Top = 117
    Width = 95
    Height = 13
    Caption = 'Nome do Escritt'#243'rio:'
  end
  object Label2: TLabel
    Left = 472
    Top = 117
    Width = 76
    Height = 13
    Caption = 'CNPJ Escrit'#243'rio:'
  end
  object Label3: TLabel
    Left = 472
    Top = 168
    Width = 66
    Height = 13
    Caption = 'C'#243'digo GCPJ:'
  end
  object Label4: TLabel
    Left = 8
    Top = 16
    Width = 113
    Height = 13
    Caption = 'Escrit'#243'rios Cadastrados:'
  end
  object Label5: TLabel
    Left = 643
    Top = 168
    Width = 91
    Height = 13
    Caption = 'Nome Digitar GCPJ'
  end
  object Label6: TLabel
    Left = 643
    Top = 216
    Width = 128
    Height = 13
    Caption = 'Data pagar IBI (a partir de):'
  end
  object DBGrid1: TDBGrid
    Left = 8
    Top = 40
    Width = 449
    Height = 289
    DataSource = dmHonorarios.dsEscritorios
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'cnpjescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'CNPJ'
        Width = 130
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'nomeescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'Nome'
        Width = 290
        Visible = True
      end>
  end
  object DBNavigator1: TDBNavigator
    Left = 464
    Top = 40
    Width = 594
    Height = 25
    DataSource = dmHonorarios.dsEscritorios
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete, nbEdit, nbPost, nbCancel]
    TabOrder = 1
    OnClick = DBNavigator1Click
  end
  object DBEdit1: TDBEdit
    Left = 472
    Top = 136
    Width = 169
    Height = 21
    DataField = 'cnpjescritorio'
    DataSource = dmHonorarios.dsEscritorios
    TabOrder = 2
  end
  object DBEdit2: TDBEdit
    Left = 643
    Top = 136
    Width = 406
    Height = 21
    CharCase = ecUpperCase
    DataField = 'nomeescritorio'
    DataSource = dmHonorarios.dsEscritorios
    TabOrder = 3
  end
  object DBEdit3: TDBEdit
    Left = 472
    Top = 184
    Width = 121
    Height = 21
    DataField = 'codgcpjescritorio'
    DataSource = dmHonorarios.dsEscritorios
    TabOrder = 4
  end
  object DBEdit4: TDBEdit
    Left = 643
    Top = 184
    Width = 385
    Height = 21
    DataField = 'nomedigitar'
    DataSource = dmHonorarios.dsEscritorios
    TabOrder = 5
  end
  object DBCheckBox1: TDBCheckBox
    Left = 472
    Top = 88
    Width = 97
    Height = 17
    Caption = 'Escrit'#243'rio Ativo'
    DataField = 'fgativo'
    DataSource = dmHonorarios.dsEscritorios
    TabOrder = 6
    ValueChecked = '1'
    ValueUnchecked = '0'
  end
  object DBCheckBox2: TDBCheckBox
    Left = 472
    Top = 240
    Width = 97
    Height = 17
    Caption = 'Pagar IBI'
    DataField = 'fgpagaribi'
    DataSource = dmHonorarios.dsEscritorios
    TabOrder = 7
    ValueChecked = '1'
    ValueUnchecked = '0'
  end
  object DBEdit5: TDBEdit
    Left = 643
    Top = 240
    Width = 121
    Height = 21
    DataField = 'dataatosibi'
    DataSource = dmHonorarios.dsEscritorios
    TabOrder = 8
  end
  object BitBtn1: TBitBtn
    Left = 8
    Top = 368
    Width = 75
    Height = 25
    Caption = 'Sair'
    TabOrder = 9
    OnClick = BitBtn1Click
  end
  object nomePesq: TEdit
    Left = 8
    Top = 336
    Width = 449
    Height = 21
    TabOrder = 10
    OnChange = nomePesqChange
  end
end
