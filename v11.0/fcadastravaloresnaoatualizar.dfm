object frmCadastraValoresNaoAtualizar: TfrmCadastraValoresNaoAtualizar
  Left = 215
  Top = 162
  Width = 910
  Height = 563
  Caption = 'Cadastra Valores N'#227'o Atualizar'
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
    Top = 8
    Width = 158
    Height = 13
    Caption = 'Controle de Valores n'#227'o atualizar:'
  end
  object Label2: TLabel
    Left = 368
    Top = 272
    Width = 61
    Height = 13
    Caption = 'Identificador:'
  end
  object Label3: TLabel
    Left = 368
    Top = 310
    Width = 96
    Height = 13
    Caption = 'Tipo de Andamento:'
  end
  object Label4: TLabel
    Left = 368
    Top = 346
    Width = 57
    Height = 13
    Caption = 'Valor pagar:'
  end
  object Label6: TLabel
    Left = 360
    Top = 96
    Width = 152
    Height = 13
    Caption = 'Data de Refer'#234'ncia (a partit de):'
  end
  object DBGrid1: TDBGrid
    Left = 16
    Top = 32
    Width = 320
    Height = 120
    DataSource = dmHonorarios.dsTiposNaoAtualizar
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
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
        FieldName = 'Identificador'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'datadoato'
        Title.Alignment = taCenter
        Title.Caption = 'Data do Ato'
        Width = 95
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'dataCadastro'
        Title.Alignment = taCenter
        Title.Caption = 'Data de Cadastro'
        Width = 95
        Visible = True
      end>
  end
  object DBGrid2: TDBGrid
    Left = 16
    Top = 216
    Width = 320
    Height = 289
    DataSource = dmHonorarios.dsValoresNaoAtualizar
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'identificador'
        Title.Alignment = taCenter
        Title.Caption = 'Identificador'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'tipoandamento'
        Title.Alignment = taCenter
        Title.Caption = 'Andamento'
        Width = 150
        Visible = True
      end
      item
        Alignment = taRightJustify
        Expanded = False
        FieldName = 'valorpagar'
        Title.Alignment = taCenter
        Title.Caption = 'Valor pagar'
        Visible = True
      end>
  end
  object BitBtn1: TBitBtn
    Left = 352
    Top = 32
    Width = 105
    Height = 25
    Caption = 'Incluir'
    TabOrder = 2
    OnClick = BitBtn1Click
  end
  object BitBtn3: TBitBtn
    Left = 456
    Top = 32
    Width = 105
    Height = 25
    Caption = 'Excluir'
    TabOrder = 3
    OnClick = BitBtn3Click
  end
  object BitBtn2: TBitBtn
    Left = 344
    Top = 216
    Width = 105
    Height = 25
    Caption = 'Incluir'
    TabOrder = 4
    OnClick = BitBtn2Click
  end
  object BitBtn4: TBitBtn
    Left = 448
    Top = 216
    Width = 105
    Height = 25
    Caption = 'Excluir'
    TabOrder = 5
    OnClick = BitBtn4Click
  end
  object BitBtn5: TBitBtn
    Left = 552
    Top = 216
    Width = 105
    Height = 25
    Caption = 'Alterar'
    TabOrder = 6
    OnClick = BitBtn5Click
  end
  object DBEdit1: TDBEdit
    Left = 480
    Top = 270
    Width = 73
    Height = 21
    DataField = 'identificador'
    DataSource = dmHonorarios.dsValoresNaoAtualizar
    Enabled = False
    ReadOnly = True
    TabOrder = 7
  end
  object DBEdit2: TDBEdit
    Left = 480
    Top = 307
    Width = 177
    Height = 21
    CharCase = ecUpperCase
    DataField = 'tipoandamento'
    DataSource = dmHonorarios.dsValoresNaoAtualizar
    Enabled = False
    TabOrder = 8
  end
  object DBEdit3: TDBEdit
    Left = 480
    Top = 344
    Width = 121
    Height = 21
    DataField = 'valorpagar'
    DataSource = dmHonorarios.dsValoresNaoAtualizar
    Enabled = False
    TabOrder = 9
  end
  object BitBtn6: TBitBtn
    Left = 656
    Top = 216
    Width = 105
    Height = 25
    Caption = 'Cancelar'
    Enabled = False
    TabOrder = 10
    OnClick = BitBtn6Click
  end
  object BitBtn7: TBitBtn
    Left = 760
    Top = 216
    Width = 105
    Height = 25
    Caption = 'Gravar'
    Enabled = False
    TabOrder = 11
    OnClick = BitBtn7Click
  end
  object BitBtn8: TBitBtn
    Left = 16
    Top = 168
    Width = 857
    Height = 25
    Caption = 'Associa escrit'#243'rios'
    TabOrder = 12
    OnClick = BitBtn8Click
  end
  object BitBtn9: TBitBtn
    Left = 560
    Top = 32
    Width = 105
    Height = 25
    Caption = 'Alterar'
    TabOrder = 13
    OnClick = BitBtn9Click
  end
  object BitBtn10: TBitBtn
    Left = 664
    Top = 32
    Width = 105
    Height = 25
    Caption = 'Cancelar'
    Enabled = False
    TabOrder = 14
    OnClick = BitBtn10Click
  end
  object BitBtn11: TBitBtn
    Left = 768
    Top = 32
    Width = 105
    Height = 25
    Caption = 'Gravar'
    Enabled = False
    TabOrder = 15
    OnClick = BitBtn11Click
  end
  object DBEdit5: TDBEdit
    Left = 528
    Top = 94
    Width = 97
    Height = 21
    DataField = 'datadoato'
    DataSource = dmHonorarios.dsTiposNaoAtualizar
    Enabled = False
    TabOrder = 16
  end
end
