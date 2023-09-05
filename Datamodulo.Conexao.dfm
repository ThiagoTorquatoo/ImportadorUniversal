object DataModuleConexao: TDataModuleConexao
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 280
  Width = 700
  object fdConexao: TFDConnection
    Params.Strings = (
      'Database=pgrolim'
      'Server=192.168.25.50'
      'Port=5432'
      'User_Name=postgres'
      'Password=1nf0tec'
      'DriverID=Ora'
      'CharacterSet=UTF8')
    LoginPrompt = False
    Left = 112
    Top = 88
  end
  object qrCliente: TFDQuery
    Connection = fdConexao
    Left = 256
    Top = 48
  end
  object qrComprasItens: TFDQuery
    Connection = fdConexao
    Left = 256
    Top = 176
  end
  object dsClientes: TDataSource
    DataSet = qrCliente
    Left = 408
    Top = 48
  end
  object dsCompras: TDataSource
    DataSet = qrCompra
    Left = 408
    Top = 104
  end
  object dsComprasItens: TDataSource
    DataSet = qrComprasItens
    Left = 408
    Top = 160
  end
  object FDPhysFBDriverLink: TFDPhysFBDriverLink
    VendorLib = 'C:\Users\Programador\Desktop\fbclient.dll'
    Left = 568
    Top = 56
  end
  object FDPhysMySQLDriverLink: TFDPhysMySQLDriverLink
    Left = 568
    Top = 104
  end
  object FDPhysPgDriverLink: TFDPhysPgDriverLink
    Left = 568
    Top = 152
  end
  object qrCompra: TFDQuery
    Connection = fdConexao
    SQL.Strings = (
      'SELECT '
      'CODIGO,'
      'COD_CLIENTE,'
      'DT_COMPRA,'
      'VALOR_TOTAL,'
      'OPERADOR,'
      'TIPO,'
      'SITUACAO,'
      'FORMA_PAGAMENTO,'
      'NUMERONOTA '
      'FROM CADCOMPRAS'
      'where COD_CLIENTE ='#39'8'#39)
    Left = 256
    Top = 128
  end
  object FDPhysOracleDriverLink: TFDPhysOracleDriverLink
    Left = 568
    Top = 224
  end
end
