object frmAssociaEscritorios: TfrmAssociaEscritorios
  Left = 381
  Top = 116
  Width = 1070
  Height = 571
  Caption = 'Associa escrit'#243'rios'
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
    Left = 24
    Top = 32
    Width = 201
    Height = 13
    Caption = 'Tipo de Controle de Valores n'#227'o Atualizar: '
  end
  object Label2: TLabel
    Left = 24
    Top = 64
    Width = 142
    Height = 13
    Caption = 'Escrit'#243'rios associados ao tipo:'
  end
  object Label3: TLabel
    Left = 552
    Top = 64
    Width = 128
    Height = 13
    Caption = 'Escrit'#243'rios associar ao tipo:'
  end
  object SpeedButton1: TSpeedButton
    Left = 504
    Top = 160
    Width = 23
    Height = 22
    Caption = '<'
    OnClick = SpeedButton1Click
  end
  object SpeedButton2: TSpeedButton
    Left = 504
    Top = 210
    Width = 23
    Height = 22
    Caption = '<<'
    OnClick = SpeedButton2Click
  end
  object SpeedButton3: TSpeedButton
    Left = 504
    Top = 261
    Width = 23
    Height = 22
    Caption = '>'
    OnClick = SpeedButton3Click
  end
  object SpeedButton4: TSpeedButton
    Left = 504
    Top = 320
    Width = 23
    Height = 25
    Caption = '>>'
    OnClick = SpeedButton4Click
  end
  object tpNaoAtualizar: TEdit
    Left = 232
    Top = 28
    Width = 57
    Height = 21
    ReadOnly = True
    TabOrder = 0
  end
  object DBGrid1: TDBGrid
    Left = 24
    Top = 88
    Width = 457
    Height = 377
    DataSource = dmHonorarios.dsEscritoriosIn
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgMultiSelect]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDblClick = DBGrid1DblClick
    Columns = <
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'cnpjescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'CNPJ'
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'nomeescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'Nome'
        Width = 320
        Visible = True
      end>
  end
  object DBGrid2: TDBGrid
    Left = 552
    Top = 88
    Width = 457
    Height = 377
    DataSource = dmHonorarios.dsEscritoriosOut
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgMultiSelect]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnDblClick = DBGrid2DblClick
    Columns = <
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'cnpjescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'CNPJ'
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'nomeescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'Nome'
        Width = 320
        Visible = True
      end>
  end
  object BitBtn1: TBitBtn
    Left = 24
    Top = 480
    Width = 75
    Height = 25
    Caption = 'Sair'
    ModalResult = 2
    TabOrder = 3
  end
  object pesqEscritorio: TEdit
    Left = 552
    Top = 472
    Width = 457
    Height = 21
    TabOrder = 4
    OnChange = pesqEscritorioChange
  end
end
