object dmgcpj_base_X: Tdmgcpj_base_X
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 651
  Top = 239
  Height = 312
  Width = 486
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 76
    Top = 24
  end
  object adoCmd: TADOCommand
    Connection = adoConn
    Parameters = <>
    Left = 204
    Top = 24
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 340
    Top = 24
  end
  object adoConn_de_2016_ate_2017: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 68
    Top = 96
  end
  object adoCmd_de_2016_ate_2017: TADOCommand
    Connection = adoConn_de_2016_ate_2017
    Parameters = <>
    Left = 220
    Top = 96
  end
  object dts_de_2016_ate_2017: TADODataSet
    Connection = adoConn_de_2016_ate_2017
    Parameters = <>
    Left = 364
    Top = 96
  end
  object adoConn_ate_2015: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 92
    Top = 200
  end
  object adoCmd_ate_2015: TADOCommand
    Connection = adoConn_ate_2015
    Parameters = <>
    Left = 220
    Top = 200
  end
  object dts_ate_2015: TADODataSet
    Connection = adoConn_ate_2015
    Parameters = <>
    Left = 364
    Top = 200
  end
end
