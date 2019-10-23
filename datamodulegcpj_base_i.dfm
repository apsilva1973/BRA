object dmgcpcj_base_I: Tdmgcpcj_base_I
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 618
  Top = 131
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
    Left = 248
    Top = 16
  end
  object adoCmd57: TADOCommand
    Connection = adoConn57
    Parameters = <>
    Left = 136
    Top = 16
  end
  object ADOConnXII: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 176
  end
  object ADOCmdXII: TADOCommand
    Connection = ADOConnXII
    Parameters = <>
    Left = 104
    Top = 176
  end
  object dtsXII: TADODataSet
    Connection = ADOConnXII
    Parameters = <>
    Left = 168
    Top = 176
  end
  object adoConn75: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 264
  end
  object adoCmd75: TADOCommand
    Connection = adoConn75
    Parameters = <>
    Left = 104
    Top = 264
  end
  object dts75: TADODataSet
    Connection = adoConn75
    Parameters = <>
    Left = 176
    Top = 272
  end
  object adoConn57_ate_2016: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 72
  end
  object dts57_ate_2016: TADODataSet
    Connection = adoConn57_ate_2016
    Parameters = <>
    Left = 256
    Top = 80
  end
  object adoCmd57_ate_2016: TADOCommand
    Connection = adoConn57_ate_2016
    Parameters = <>
    Left = 144
    Top = 72
  end
end
