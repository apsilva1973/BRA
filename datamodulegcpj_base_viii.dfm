object dmgcpj_base_VIII: Tdmgcpj_base_VIII
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 494
  Top = 213
  Height = 298
  Width = 290
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
    Top = 48
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 144
    Top = 32
  end
  object adoConn_ate_2013: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 96
  end
  object adoCmd_ate_2013: TADOCommand
    Connection = adoConn_ate_2013
    Parameters = <>
    Left = 144
    Top = 128
  end
end
