object dmgcpcj_base_v: Tdmgcpcj_base_v
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 1
  Top = 106
  Height = 150
  Width = 215
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 24
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 144
    Top = 40
  end
  object adoCmd: TADOCommand
    Connection = adoConn
    Parameters = <>
    Left = 80
    Top = 48
  end
end
