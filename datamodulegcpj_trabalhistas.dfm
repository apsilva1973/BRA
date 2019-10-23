object dmgcpj_trabalhistas: Tdmgcpj_trabalhistas
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 435
  Top = 116
  Height = 150
  Width = 215
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 24
    Top = 16
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 112
    Top = 16
  end
  object adoCmd: TADOCommand
    Connection = adoConn
    Parameters = <>
    Left = 72
    Top = 32
  end
end
