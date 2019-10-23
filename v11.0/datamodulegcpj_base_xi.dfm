object dmgcpcj_base_XI: Tdmgcpcj_base_XI
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 742
  Top = 153
  Height = 271
  Width = 456
  object adoConnNfiscal: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 32
    Top = 24
  end
  object adoConnNfiscdt: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 24
    Top = 96
  end
  object adoCmdNfiscdt: TADOCommand
    Connection = adoConnNfiscdt
    Parameters = <>
    Left = 120
    Top = 88
  end
  object dtsNfiscdt: TADODataSet
    Connection = adoConnNfiscdt
    Parameters = <>
    Left = 192
    Top = 96
  end
  object dtsNfiscal: TADODataSet
    Connection = adoConnNfiscal
    Parameters = <>
    Left = 216
    Top = 16
  end
  object adoCmdNfiscal: TADOCommand
    Connection = adoConnNfiscal
    Parameters = <>
    Left = 120
    Top = 16
  end
end
