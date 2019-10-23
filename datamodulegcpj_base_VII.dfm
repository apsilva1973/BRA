object dmgcpj_base_vii: Tdmgcpj_base_vii
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 192
  Top = 116
  Height = 139
  Width = 228
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
    Left = 144
    Top = 40
  end
end
