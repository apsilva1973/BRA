object dmgcpcj_base_I: Tdmgcpcj_base_I
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 645
  Top = 146
  Height = 398
  Width = 469
  object adoConn57: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 16
  end
  object dts57: TADODataSet
    Connection = adoConn57
    Parameters = <>
    Left = 160
    Top = 16
  end
  object adoCmd57: TADOCommand
    Connection = adoConn57
    Parameters = <>
    Left = 104
    Top = 16
  end
  object ADOConnXII: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 96
  end
  object ADOCmdXII: TADOCommand
    Connection = ADOConnXII
    Parameters = <>
    Left = 104
    Top = 96
  end
  object dtsXII: TADODataSet
    Connection = ADOConnXII
    Parameters = <>
    Left = 168
    Top = 96
  end
  object adoConn75: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 184
  end
  object adoCmd75: TADOCommand
    Connection = adoConn75
    Parameters = <>
    Left = 104
    Top = 184
  end
  object dts75: TADODataSet
    Connection = adoConn75
    Parameters = <>
    Left = 176
    Top = 192
  end
end
