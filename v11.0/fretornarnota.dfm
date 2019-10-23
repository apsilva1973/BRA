object frmRetornarFinalizada: TfrmRetornarFinalizada
  Left = 374
  Top = 144
  Width = 291
  Height = 153
  Caption = 'Retornar Nota Finalizada'
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
    Top = 24
    Width = 79
    Height = 13
    Caption = 'N'#250'mero da nota:'
  end
  object numNota: TEdit
    Left = 16
    Top = 48
    Width = 249
    Height = 21
    TabOrder = 0
  end
  object BitBtn1: TBitBtn
    Left = 16
    Top = 80
    Width = 75
    Height = 25
    Caption = 'Ok'
    TabOrder = 1
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 96
    Top = 80
    Width = 75
    Height = 25
    Caption = 'Sair'
    ModalResult = 2
    TabOrder = 2
  end
end
