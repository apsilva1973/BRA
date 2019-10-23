object frmEnviaGcpj: TfrmEnviaGcpj
  Left = 218
  Top = 29
  Width = 1254
  Height = 735
  Caption = 'Envia dados para GCPJ'
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
    Caption = 'Planilhas enviar GCPJ'
  end
  object DBGrid1: TDBGrid
    Left = 16
    Top = 96
    Width = 1201
    Height = 353
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
      end
      item
        Expanded = False
        FieldName = 'fgenviadogcpj'
        Width = 16
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'fgretornogcpj'
        Width = 16
        Visible = True
      end>
  end
  object rgFiltro: TRadioGroup
    Left = 16
    Top = 32
    Width = 1201
    Height = 57
    Columns = 3
    ItemIndex = 0
    Items.Strings = (
      'Todas as planilhas'
      'Somente planilhas para enviar'
      'Somente planilhas enviadas')
    TabOrder = 1
  end
  object re: TRichEdit
    Left = 16
    Top = 488
    Width = 1201
    Height = 202
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object BitBtn4: TBitBtn
    Left = 16
    Top = 456
    Width = 80
    Height = 25
    Caption = 'Enviar'
    TabOrder = 3
    OnClick = BitBtn4Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000130B0000130B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333303
      333333333333337FF3333333333333903333333333333377FF33333333333399
      03333FFFFFFFFF777FF3000000999999903377777777777777FF0FFFF0999999
      99037F3337777777777F0FFFF099999999907F3FF777777777770F00F0999999
      99037F773777777777730FFFF099999990337F3FF777777777330F00FFFFF099
      03337F773333377773330FFFFFFFF09033337F3FF3FFF77733330F00F0000003
      33337F773777777333330FFFF0FF033333337F3FF7F3733333330F08F0F03333
      33337F7737F7333333330FFFF003333333337FFFF77333333333000000333333
      3333777777333333333333333333333333333333333333333333}
    NumGlyphs = 2
  end
  object BitBtn2: TBitBtn
    Left = 96
    Top = 456
    Width = 80
    Height = 25
    Caption = 'Cancelar'
    TabOrder = 4
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
end
