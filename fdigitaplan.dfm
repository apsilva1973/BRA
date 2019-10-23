object frmDigitaPlan: TfrmDigitaPlan
  Left = 80
  Top = 134
  Width = 1362
  Height = 726
  Caption = 'Digita dados GCPJ'
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
    Width = 102
    Height = 13
    Caption = 'Planilhas para Digitar:'
  end
  object Label2: TLabel
    Left = 416
    Top = 184
    Width = 26
    Height = 13
    Caption = 'Nota:'
    Visible = False
  end
  object BitBtn4: TBitBtn
    Left = 8
    Top = 176
    Width = 80
    Height = 25
    Caption = 'Digitar'
    TabOrder = 4
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
  object DBGrid1: TDBGrid
    Left = 8
    Top = 24
    Width = 585
    Height = 145
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
        Width = 95
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'nomeescritorio'
        Title.Alignment = taCenter
        Title.Caption = 'Escrit'#243'rio'
        Width = 250
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'anomesreferencia'
        Title.Alignment = taCenter
        Title.Caption = 'Ano/M'#234's'
        Width = 70
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'sequencia'
        Title.Alignment = taCenter
        Title.Caption = 'Seq.'
        Width = 40
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
    Left = 168
    Top = 176
    Width = 80
    Height = 25
    Caption = 'Cancelar'
    TabOrder = 1
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
    Left = 8
    Top = 208
    Width = 1001
    Height = 361
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object BitBtn3: TBitBtn
    Left = 648
    Top = 0
    Width = 75
    Height = 25
    Caption = 'BitBtn3'
    TabOrder = 3
    Visible = False
    OnClick = BitBtn3Click
  end
  object BitBtn5: TBitBtn
    Left = 728
    Top = 0
    Width = 75
    Height = 25
    Caption = 'BitBtn5'
    TabOrder = 5
    Visible = False
    OnClick = BitBtn5Click
  end
  object processando: TEdit
    Left = 584
    Top = 178
    Width = 425
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    TabOrder = 6
  end
  object numNota_: TEdit
    Left = 456
    Top = 178
    Width = 121
    Height = 21
    TabOrder = 7
    Visible = False
  end
  object BitBtn6: TBitBtn
    Left = 808
    Top = 0
    Width = 75
    Height = 25
    Caption = 'bitbtn6'
    TabOrder = 8
    Visible = False
    OnClick = BitBtn6Click
  end
  object BitBtn7: TBitBtn
    Left = 456
    Top = 8
    Width = 75
    Height = 25
    Caption = 'BitBtn7'
    TabOrder = 9
    Visible = False
    OnClick = BitBtn7Click
  end
  object DBGrid2: TDBGrid
    Left = 592
    Top = 24
    Width = 241
    Height = 145
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgMultiSelect]
    PopupMenu = pmNotas
    TabOrder = 10
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
        FieldName = 'numeronota'
        Title.Alignment = taCenter
        Title.Caption = 'Nota'
        Width = 70
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'totaldeatos'
        Title.Alignment = taCenter
        Title.Caption = 'Tot.Atos'
        Width = 55
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'valortotal'
        Title.Alignment = taCenter
        Title.Caption = 'Valor Total'
        Width = 90
        Visible = True
      end>
  end
  object DBGrid3: TDBGrid
    Left = 832
    Top = 24
    Width = 177
    Height = 145
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgMultiSelect]
    TabOrder = 11
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'atosdigitados'
        Title.Alignment = taCenter
        Title.Caption = 'Digitados'
        Width = 55
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'valordigitado'
        Title.Alignment = taCenter
        Title.Caption = 'Vl.Digitado'
        Width = 90
        Visible = True
      end>
  end
  object BitBtn1: TBitBtn
    Left = 88
    Top = 176
    Width = 80
    Height = 25
    Caption = 'Excluir'
    TabOrder = 12
    OnClick = BitBtn1Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000130B0000130B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      333333333333333333FF33333333333330003333333333333777333333333333
      300033FFFFFF3333377739999993333333333777777F3333333F399999933333
      3300377777733333337733333333333333003333333333333377333333333333
      3333333333333333333F333333333333330033333F33333333773333C3333333
      330033337F3333333377333CC3333333333333F77FFFFFFF3FF33CCCCCCCCCC3
      993337777777777F77F33CCCCCCCCCC399333777777777737733333CC3333333
      333333377F33333333FF3333C333333330003333733333333777333333333333
      3000333333333333377733333333333333333333333333333333}
    NumGlyphs = 2
  end
  object pnlWb: TPanel
    Left = 8
    Top = 208
    Width = 1001
    Height = 417
    TabOrder = 13
    object wb: TWebBrowser
      Left = 1
      Top = 1
      Width = 999
      Height = 415
      Align = alClient
      TabOrder = 0
      OnNewWindow2 = wbNewWindow2
      OnDocumentComplete = wbDocumentComplete
      ControlData = {
        4C00000040670000E42A00000000000000000000000000000000000000000000
        000000004C000000000000000000000001000000E0D057007335CF11AE690800
        2B2E12620C000000000000004C0000000114020000000000C000000000000046
        8000000000000000000000000000000000000000000000000000000000000000
        00000000000000000100000000000000000000000000000000000000}
    end
  end
  object BitBtn8: TBitBtn
    Left = 256
    Top = 176
    Width = 80
    Height = 25
    Caption = 'Continuar'
    Enabled = False
    TabOrder = 14
    OnClick = BitBtn8Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000130B0000130B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      3333333333FFFFF3333333333999993333333333F77777FFF333333999999999
      3333333777333777FF33339993707399933333773337F3777FF3399933000339
      9933377333777F3377F3399333707333993337733337333337FF993333333333
      399377F33333F333377F993333303333399377F33337FF333373993333707333
      333377F333777F333333993333101333333377F333777F3FFFFF993333000399
      999377FF33777F77777F3993330003399993373FF3777F37777F399933000333
      99933773FF777F3F777F339993707399999333773F373F77777F333999999999
      3393333777333777337333333999993333333333377777333333}
    NumGlyphs = 2
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 6000
    OnTimer = Timer1Timer
    Left = 320
    Top = 184
  end
  object pmNotas: TPopupMenu
    Left = 360
    Top = 304
    object Retornarnotafinalizada1: TMenuItem
      Caption = 'Retornar nota finalizada'
      OnClick = Retornarnotafinalizada1Click
    end
    object ConferirAtosdaNotaComoGCPJ1: TMenuItem
      Caption = 'Conferir Atos da Nota Com o GCPJ'
      Enabled = False
      OnClick = ConferirAtosdaNotaComoGCPJ1Click
    end
    object Marcarnotacomofinalizada1: TMenuItem
      Caption = 'Marcar nota como finalizada'
      OnClick = Marcarnotacomofinalizada1Click
    end
  end
end
