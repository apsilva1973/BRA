object dmgcpj_base_ii: Tdmgcpj_base_ii
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 673
  Top = 321
  Height = 320
  Width = 391
  object adoConn24: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 24
  end
  object adoCmd24: TADOCommand
    Connection = adoConn24
    Parameters = <>
    Left = 88
    Top = 24
  end
  object dts24: TADODataSet
    Connection = adoConn24
    Parameters = <>
    Left = 168
    Top = 16
  end
  object adoConn88: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 16
    Top = 120
  end
  object adoCmd88: TADOCommand
    Connection = adoConn88
    Parameters = <>
    Left = 104
    Top = 120
  end
  object dts88: TADODataSet
    Connection = adoConn88
    Parameters = <>
    Left = 184
    Top = 112
  end
  object adoConnB7: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 24
    Top = 200
  end
  object adoCmdb7: TADOCommand
    Connection = adoConnB7
    Parameters = <>
    Left = 112
    Top = 200
  end
  object dtsB7: TADODataSet
    Connection = adoConnB7
    Parameters = <>
    Left = 192
    Top = 192
  end
end
