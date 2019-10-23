object frmValidaPlan: TfrmValidaPlan
  Left = 12
  Top = 123
  Width = 1012
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
    Left = 416
    Top = 208
    Width = 26
    Height = 13
    Caption = 'Nota:'
  end
  object BitBtn2: TBitBtn
    Left = 256
    Top = 200
    Width = 80
    Height = 25
    Caption = 'Cancelar'
    TabOrder = 2
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
  object DBGrid1: TDBGrid
    Left = 16
    Top = 24
    Width = 961
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
  object BitBtn1: TBitBtn
    Left = 16
    Top = 200
    Width = 80
    Height = 25
    Caption = 'Validar'
    TabOrder = 1
    OnClick = BitBtn1Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
      555555555555555555555555555555555555555555FF55555555555559055555
      55555555577FF5555555555599905555555555557777F5555555555599905555
      555555557777FF5555555559999905555555555777777F555555559999990555
      5555557777777FF5555557990599905555555777757777F55555790555599055
      55557775555777FF5555555555599905555555555557777F5555555555559905
      555555555555777FF5555555555559905555555555555777FF55555555555579
      05555555555555777FF5555555555557905555555555555777FF555555555555
      5990555555555555577755555555555555555555555555555555}
    NumGlyphs = 2
  end
  object re: TRichEdit
    Left = 8
    Top = 232
    Width = 969
    Height = 337
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 3
  end
  object BitBtn3: TBitBtn
    Left = 648
    Top = 0
    Width = 75
    Height = 25
    Caption = 'BitBtn3'
    TabOrder = 4
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
    Top = 202
    Width = 393
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
  object numNota: TEdit
    Left = 456
    Top = 202
    Width = 121
    Height = 21
    TabOrder = 7
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
  object BitBtn8: TBitBtn
    Left = 96
    Top = 200
    Width = 80
    Height = 25
    Caption = 'Calc.Notas'
    TabOrder = 10
    OnClick = BitBtn8Click
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      333333333333333333333333333333333333333FFFFFFFFFFF33330000000000
      03333377777777777F33333003333330033333377FF333377F33333300333333
      0333333377FF33337F3333333003333303333333377FF3337333333333003333
      333333333377FF3333333333333003333333333333377FF33333333333330033
      3333333333337733333333333330033333333333333773333333333333003333
      33333333337733333F3333333003333303333333377333337F33333300333333
      03333333773333337F33333003333330033333377FFFFFF77F33330000000000
      0333337777777777733333333333333333333333333333333333}
    NumGlyphs = 2
  end
  object cbSemVolumetria: TCheckBox
    Left = 832
    Top = 176
    Width = 145
    Height = 17
    Caption = 'Sem c'#225'lculo de volumetria'
    TabOrder = 11
  end
  object BitBtn4: TBitBtn
    Left = 176
    Top = 200
    Width = 80
    Height = 25
    Caption = 'Excluir'
    TabOrder = 12
    OnClick = BitBtn4Click
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
end
