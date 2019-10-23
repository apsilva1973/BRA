object frmAdosDigitados: TfrmAdosDigitados
  Left = 269
  Top = 27
  Width = 912
  Height = 632
  Caption = 'frmAdosDigitados'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object lblNumeroNota: TLabel
    Left = 16
    Top = 24
    Width = 69
    Height = 13
    Caption = 'Atos digitados:'
  end
  object Label2: TLabel
    Left = 592
    Top = 552
    Width = 69
    Height = 13
    Caption = 'Total Digitado:'
  end
  object DBGrid1: TDBGrid
    Left = 16
    Top = 48
    Width = 857
    Height = 489
    DataSource = dmHonorarios.dsAtosDigitados
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
        FieldName = 'VALORCORRIGIDO'
        Title.Alignment = taCenter
        Title.Caption = 'Valor Digitado'
        Width = 150
        Visible = True
      end>
  end
  object BitBtn1: TBitBtn
    Left = 16
    Top = 552
    Width = 177
    Height = 25
    Caption = 'Marcar Ato Como N'#195'O DIGITADO'
    TabOrder = 1
    OnClick = BitBtn1Click
  end
  object BitBtn3: TBitBtn
    Left = 384
    Top = 552
    Width = 177
    Height = 25
    Caption = 'Sair'
    TabOrder = 2
    OnClick = BitBtn3Click
  end
  object totalDigitado: TEdit
    Left = 680
    Top = 544
    Width = 193
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    TabOrder = 3
  end
  object BitBtn2: TBitBtn
    Left = 192
    Top = 552
    Width = 193
    Height = 25
    Caption = 'Marcar Todos Como N'#195'O DIGITADOS'
    TabOrder = 4
    OnClick = BitBtn2Click
  end
end
