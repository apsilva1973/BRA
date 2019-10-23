object dmHonorarios: TdmHonorarios
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 569
  Top = 131
  Height = 616
  Width = 474
  object adoConn: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 16
  end
  object adoCmd: TADOCommand
    Connection = adoConn
    Parameters = <>
    Left = 112
    Top = 16
  end
  object dts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 184
    Top = 24
  end
  object dtsOrg: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 40
    Top = 72
  end
  object adoDts: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 112
    Top = 72
  end
  object dtsEscritorios: TADODataSet
    Connection = adoConn
    OnNewRecord = dtsEscritoriosNewRecord
    Parameters = <>
    Left = 344
    Top = 16
  end
  object dsEscritorios: TDataSource
    DataSet = dtsEscritorios
    Left = 192
    Top = 80
  end
  object dtsAtosDigitados: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 256
    Top = 16
  end
  object dsAtosDigitados: TDataSource
    DataSet = dtsAtosDigitados
    Left = 280
    Top = 72
  end
  object adoConnRpt: TADOConnection
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 40
    Top = 152
  end
  object dtsRpt: TADODataSet
    Connection = adoConnRpt
    Parameters = <>
    Left = 136
    Top = 152
  end
  object dtsRptTotal: TADODataSet
    Connection = adoConnRpt
    Parameters = <>
    Left = 224
    Top = 160
  end
  object dtsAtosPendentes: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 296
    Top = 128
  end
  object dsAtosPendentes: TDataSource
    DataSet = dtsAtosPendentes
    Left = 320
    Top = 184
  end
  object dtsTiposNaoAtualizar: TADODataSet
    Connection = adoConn
    BeforePost = dtsTiposNaoAtualizarBeforePost
    BeforeDelete = dtsTiposNaoAtualizarBeforeDelete
    OnNewRecord = dtsTiposNaoAtualizarNewRecord
    Parameters = <>
    Left = 48
    Top = 256
  end
  object dsTiposNaoAtualizar: TDataSource
    DataSet = dtsTiposNaoAtualizar
    OnDataChange = dsTiposNaoAtualizarDataChange
    Left = 192
    Top = 256
  end
  object dtsValoresNaoAtualizar: TADODataSet
    Connection = adoConn
    BeforePost = dtsValoresNaoAtualizarBeforePost
    OnNewRecord = dtsValoresNaoAtualizarNewRecord
    Parameters = <>
    Left = 56
    Top = 336
  end
  object dsValoresNaoAtualizar: TDataSource
    DataSet = dtsValoresNaoAtualizar
    Left = 184
    Top = 336
  end
  object dtsEscritoriosIn: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 56
    Top = 400
  end
  object dtsEscritoriosOut: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 64
    Top = 480
  end
  object dsEscritoriosIn: TDataSource
    DataSet = dtsEscritoriosIn
    Left = 200
    Top = 400
  end
  object dsEscritoriosOut: TDataSource
    DataSet = dtsEscritoriosOut
    Left = 192
    Top = 480
  end
  object dtsAdvogados: TADODataSet
    Connection = adoConn
    Parameters = <>
    Left = 288
    Top = 320
  end
  object dsAdvogados: TDataSource
    DataSet = dtsAdvogados
    Left = 368
    Top = 320
  end
end
