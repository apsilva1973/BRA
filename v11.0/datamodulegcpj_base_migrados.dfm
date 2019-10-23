object dmGcpj_migrados: TdmGcpj_migrados
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 435
  Top = 103
  Height = 150
  Width = 215
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 40
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 104
    Top = 48
  end
  object adoCmd: TADOCommand
    Connection = adoConn
    Parameters = <>
    Left = 152
    Top = 24
  end
end
