object dmgcpj_base_VIII: Tdmgcpj_base_VIII
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 494
  Top = 213
  Height = 150
  Width = 215
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 24
  end
  object adoCmd: TADOCommand
    Connection = adoConn
    Parameters = <>
    Left = 80
    Top = 48
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 112
    Top = 32
  end
end
