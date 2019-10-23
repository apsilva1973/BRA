object frmRelatorios: TfrmRelatorios
  Left = 11
  Top = 80
  Width = 1083
  Height = 615
  Caption = 'Consiste Dados da Planilha'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poDefault
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 104
    Height = 13
    Caption = 'Planilhas para Validar:'
  end
  object Label2: TLabel
    Left = 184
    Top = 272
    Width = 26
    Height = 13
    Caption = 'Nota:'
  end
  object Label3: TLabel
    Left = 16
    Top = 208
    Width = 85
    Height = 13
    Caption = 'Nome da planilha:'
  end
  object SpeedButton1: TSpeedButton
    Left = 952
    Top = 200
    Width = 23
    Height = 22
    Caption = '...'
    OnClick = SpeedButton1Click
  end
  object Label4: TLabel
    Left = 16
    Top = 232
    Width = 117
    Height = 13
    Caption = 'M'#234's/Ano de Refer'#234'ncia:'
  end
  object BitBtn4: TBitBtn
    Left = 16
    Top = 264
    Width = 80
    Height = 25
    Caption = 'Gerar'
    TabOrder = 6
    OnClick = BitBtn4Click
    Glyph.Data = {
      F6000000424DF600000000000000760000002800000010000000100000000100
      0400000000008000000000000000000000001000000010000000000000000000
      BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
      DDDDDDDDDDDDDDDDDDDDDDDDDDDDDD00000DD00000000006660DD08888880E00
      000DD000000000EEE080DD07778E0EEE0080DDD078E0EEE07700DDDD0E0EEE00
      0000DDD0E0EEE080DDDDDD0E0EEE07080DDDD0E0EEE0777080DD0E0EEE0D0777
      080D00EEE0DDD077700D00000DDDDD00000DDDDDDDDDDDDDDDDD}
  end
  object DBGrid1: TDBGrid
    Left = 16
    Top = 24
    Width = 961
    Height = 145
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    PopupMenu = ppMenu
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
        Title.Caption = 'CNPJ Escrit'#243'rio'
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'nomeescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'Escrit'#243'rio'
        Width = 300
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'anomesreferencia'
        Title.Alignment = taCenter
        Title.Caption = 'Ano/M'#234's Ref.'
        Width = 90
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'sequencia'
        Title.Alignment = taCenter
        Title.Caption = 'Sequencial'
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'dtimportacao'
        Title.Alignment = taCenter
        Title.Caption = 'Data Importa'#231#227'o'
        Width = 100
        Visible = True
      end>
  end
  object BitBtn2: TBitBtn
    Left = 96
    Top = 264
    Width = 80
    Height = 25
    Caption = 'Cancelar'
    TabOrder = 5
    OnClick = BitBtn2Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00330000000000
      03333377777777777F333301BBBBBBBB033333773F3333337F3333011BBBBBBB
      0333337F73F333337F33330111BBBBBB0333337F373F33337F333301110BBBBB
      0333337F337F33337F333301110BBBBB0333337F337F33337F333301110BBBBB
      0333337F337F33337F333301110BBBBB0333337F337F33337F333301110BBBBB
      0333337F337F33337F333301110BBBBB0333337F337FF3337F33330111B0BBBB
      0333337F337733337F333301110BBBBB0333337F337F33337F333301110BBBBB
      0333337F3F7F33337F333301E10BBBBB0333337F7F7F33337F333301EE0BBBBB
      0333337F777FFFFF7F3333000000000003333377777777777333}
    NumGlyphs = 2
  end
  object re: TRichEdit
    Left = 16
    Top = 296
    Width = 953
    Height = 465
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 4
  end
  object numNota: TEdit
    Left = 224
    Top = 266
    Width = 121
    Height = 21
    Enabled = False
    TabOrder = 3
  end
  object cbxRelatorios: TComboBox
    Left = 16
    Top = 176
    Width = 961
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 1
    OnChange = cbxRelatoriosChange
    Items.Strings = (
      'Inconsist'#234'ncias na importa'#231#227'o'
      'Inconsist'#234'ncias p'#243's-importa'#231#227'o'
      'Valores recalculados pelo sistema'
      'Notas finalizadas'
      'Inconsist'#234'ncias na Solicita'#231#227'o de Honor'#225'rios')
  end
  object nomePlan: TEdit
    Left = 136
    Top = 200
    Width = 817
    Height = 21
    TabOrder = 2
  end
  object mesRef: TComboBox
    Left = 136
    Top = 232
    Width = 97
    Height = 21
    Style = csDropDownList
    Enabled = False
    ItemHeight = 13
    TabOrder = 7
    Items.Strings = (
      'Janeiro'
      'Fevereiro'
      'Mar'#231'o'
      'Abril'
      'Maio'
      'Junho'
      'Julho'
      'Agosto'
      'Setembro'
      'Outubro'
      'Novembro'
      'Dezembro')
  end
  object anoRef: TComboBox
    Left = 240
    Top = 232
    Width = 97
    Height = 21
    Style = csDropDownList
    Enabled = False
    ItemHeight = 13
    TabOrder = 8
    Items.Strings = (
      '2011'
      '2012'
      '2013'
      '2014'
      '2015'
      '2016'
      '2017'
      '2018'
      '2019'
      '2020'
      '2021'
      '2022'
      '2023')
  end
  object sd: TSaveDialog
    Options = [ofHideReadOnly, ofPathMustExist, ofEnableSizing]
    Left = 512
    Top = 232
  end
  object ppMenu: TPopupMenu
    Left = 448
    Top = 248
    object Retornarplanilhaparadigitao1: TMenuItem
      Caption = 'Retornar planilha para digita'#231#227'o'
      OnClick = Retornarplanilhaparadigitao1Click
    end
  end
end
