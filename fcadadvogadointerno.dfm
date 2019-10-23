object frmCadAdvogadoInterno: TfrmCadAdvogadoInterno
  Left = 325
  Top = 195
  Width = 1088
  Height = 506
  Caption = 'Cadastra Advogado Interno (p/AJUIZAMENTOS)'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 16
    Width = 119
    Height = 13
    Caption = 'Advogados Cadastrados:'
  end
  object Bevel1: TBevel
    Left = 464
    Top = 96
    Width = 569
    Height = 313
    Shape = bsFrame
  end
  object Label2: TLabel
    Left = 488
    Top = 128
    Width = 97
    Height = 13
    Caption = 'Nome do advogado:'
  end
  object Label3: TLabel
    Left = 488
    Top = 232
    Width = 82
    Height = 13
    Caption = 'C'#243'digo funcional:'
  end
  object DBGrid1: TDBGrid
    Left = 16
    Top = 40
    Width = 433
    Height = 369
    DataSource = dmHonorarios.dsAdvogados
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'nomeadvogado'
        Title.Alignment = taCenter
        Title.Caption = 'Nome do Advogado'
        Width = 300
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'codigofuncional'
        Title.Alignment = taCenter
        Title.Caption = 'C'#243'd.Funcional'
        Width = 100
        Visible = True
      end>
  end
  object DBNavigator1: TDBNavigator
    Left = 464
    Top = 40
    Width = 570
    Height = 41
    DataSource = dmHonorarios.dsAdvogados
    TabOrder = 1
  end
  object DBEdit1: TDBEdit
    Left = 488
    Top = 152
    Width = 513
    Height = 21
    DataField = 'nomeadvogado'
    DataSource = dmHonorarios.dsAdvogados
    TabOrder = 2
  end
  object DBEdit2: TDBEdit
    Left = 488
    Top = 256
    Width = 153
    Height = 21
    DataField = 'codigofuncional'
    DataSource = dmHonorarios.dsAdvogados
    TabOrder = 3
  end
end
