object dmgcpj_base_IX: Tdmgcpj_base_IX
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 435
  Top = 103
  Height = 367
  Width = 606
  object adoConn76: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 24
  end
  object dts76: TADODataSet
    Connection = adoConn76
    Parameters = <>
    Left = 192
    Top = 40
  end
  object adoCmd76: TADOCommand
    Connection = adoConn76
    Parameters = <>
    Left = 96
    Top = 24
  end
  object adoConn77: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 128
  end
  object dts77: TADODataSet
    Connection = adoConn77
    Parameters = <>
    Left = 192
    Top = 144
  end
  object adoCmd77: TADOCommand
    Connection = adoConn77
    Parameters = <>
    Left = 96
    Top = 128
  end
  object dtsPesq: TADODataSet
    Connection = adoConn77
    Parameters = <>
    Left = 304
    Top = 96
  end
end
