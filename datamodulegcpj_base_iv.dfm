object dmgcpj_base_iv: Tdmgcpj_base_iv
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 580
  Top = 103
  Height = 365
  Width = 550
  object adoConn50: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 40
  end
  object dts50: TADODataSet
    Connection = adoConn50
    Parameters = <>
    Left = 160
    Top = 40
  end
  object adoCmd50: TADOCommand
    Connection = adoConn50
    Parameters = <>
    Left = 104
    Top = 40
  end
  object dts66: TADODataSet
    Connection = adoConn66
    Parameters = <>
    Left = 168
    Top = 144
  end
  object adoConn66: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 144
  end
  object adoCmd66: TADOCommand
    Connection = adoConn66
    Parameters = <>
    Left = 104
    Top = 144
  end
  object dts3T: TADODataSet
    Connection = adoConn3T
    Parameters = <>
    Left = 168
    Top = 216
  end
  object adoConn3T: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 216
  end
  object adoCmd3T: TADOCommand
    Connection = adoConn3T
    Parameters = <>
    Left = 104
    Top = 216
  end
  object dtsMesu: TADODataSet
    Connection = adoConn3T
    Parameters = <>
    Left = 256
    Top = 216
  end
end
