object dmgcpj_base_iii: Tdmgcpj_base_iii
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 409
  Top = 309
  Height = 150
  Width = 260
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 24
  end
  object adoCmd: TADOCommand
    Connection = adoConn
    Parameters = <>
    Left = 88
    Top = 24
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 168
    Top = 16
  end
end
