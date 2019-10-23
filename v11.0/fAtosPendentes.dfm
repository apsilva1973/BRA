object frmAtosPendentes: TfrmAtosPendentes
  Left = 13
  Top = 125
  Width = 1552
  Height = 659
  Caption = 'Atos Pendentes de Valida'#231#227'o'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 141
    Height = 13
    Caption = 'Atos pendentes de valida'#231#227'o:'
  end
  object Label2: TLabel
    Left = 16
    Top = 528
    Width = 71
    Height = 13
    Caption = 'Inconsist'#234'ncia:'
  end
  object DBGrid1: TDBGrid
    Left = 16
    Top = 24
    Width = 1497
    Height = 489
    DataSource = dmHonorarios.dsAtosPendentes
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'GCPJ'
        Title.Alignment = taCenter
        Width = 120
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'TIPOANDAMENTO'
        Title.Alignment = taCenter
        Title.Caption = 'Tipo de Andamento'
        Width = 250
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'VALOR'
        Title.Alignment = taCenter
        Title.Caption = 'Valor do Ato'
        Width = 150
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'CODTIPOACAO'
        Title.Alignment = taCenter
        Title.Caption = 'C'#243'd. A'#231#227'o'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'TIPOACAO'
        Title.Alignment = taCenter
        Title.Caption = 'Tipo A'#231#227'o'
        Width = 220
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'MOTIVOBAIXA'
        Title.Alignment = taCenter
        Title.Caption = 'Motivo da Baixa'
        Width = 250
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NUMERONOTA'
        Title.Alignment = taCenter
        Title.Caption = 'N'#250'mero da Nota'
        Width = 90
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'FGCRUZADOGCPJ'
        Title.Alignment = taCenter
        Title.Caption = 'Status'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'fgDrcAtivas'
        Title.Alignment = taCenter
        Title.Caption = 'DRC Ativa'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'fgDrcContrarias'
        Title.Alignment = taRightJustify
        Title.Caption = 'DRC Contr'#225'ria'
        Visible = True
      end>
  end
  object BitBtn1: TBitBtn
    Left = 16
    Top = 576
    Width = 217
    Height = 25
    Caption = 'Excluir o ato da planilha'
    TabOrder = 2
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 232
    Top = 576
    Width = 217
    Height = 25
    Caption = 'Marcar Ato como Inconsistente'
    TabOrder = 3
    OnClick = BitBtn2Click
  end
  object BitBtn3: TBitBtn
    Left = 448
    Top = 576
    Width = 217
    Height = 25
    Caption = 'Sair'
    ModalResult = 1
    TabOrder = 4
  end
  object Edit1: TEdit
    Left = 16
    Top = 544
    Width = 1209
    Height = 21
    TabOrder = 1
  end
end
