object dmDados: TdmDados
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 790
  Width = 1022
  object wolferp: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=MSDASQL.1;Persist Security Info=False;Data Source=WOLF_' +
      'ERP;'
    LoginPrompt = False
    Left = 936
    Top = 40
  end
  object dsEstoque: TDataSource
    DataSet = qryEstoque
    Left = 112
    Top = 88
  end
  object qryEstoque: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM MOVIMENTACAO_ESTOQUE')
    Left = 40
    Top = 88
  end
  object qryProdutos: TADOQuery
    Connection = wolferp
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM PRODUTOS')
    Left = 40
    Top = 152
    object qryProdutosID: TAutoIncField
      FieldName = 'ID'
      ReadOnly = True
    end
    object qryProdutosCODIGO_BARRAS: TStringField
      FieldName = 'CODIGO_BARRAS'
    end
    object qryProdutosDESCRICAO: TStringField
      FieldName = 'DESCRICAO'
      Size = 100
    end
    object qryProdutosID_FORNECEDOR: TIntegerField
      FieldName = 'ID_FORNECEDOR'
    end
    object qryProdutosPRECO_CUSTO: TBCDField
      FieldName = 'PRECO_CUSTO'
      DisplayFormat = 'R$ #,##0.00'
      Precision = 15
      Size = 2
    end
    object qryProdutosPRECO_VENDA: TBCDField
      FieldName = 'PRECO_VENDA'
      DisplayFormat = 'R$ #,##0.00'
      Precision = 15
      Size = 2
    end
    object qryProdutosESTOQUE_ATUAL: TBCDField
      FieldName = 'ESTOQUE_ATUAL'
      Precision = 15
      Size = 3
    end
    object qryProdutosESTOQUE_MINIMO: TBCDField
      FieldName = 'ESTOQUE_MINIMO'
      Precision = 15
      Size = 3
    end
    object qryProdutosUNIDADE: TStringField
      FieldName = 'UNIDADE'
      Size = 5
    end
    object qryProdutosATIVO: TStringField
      FieldName = 'ATIVO'
      FixedChar = True
      Size = 1
    end
  end
  object dsProdutos: TDataSource
    DataSet = qryProdutos
    Left = 112
    Top = 152
  end
  object qryConsultas: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 40
    Top = 216
  end
  object qryExec: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 40
    Top = 280
  end
  object qryFinanceiro: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM FINANCEIRO ORDER BY DATA_VENCIMENTO DESC')
    Left = 40
    Top = 344
  end
  object qryRelatorios: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 40
    Top = 408
  end
  object qryClientes: TADOQuery
    Connection = wolferp
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CLIENTES')
    Left = 40
    Top = 24
    object qryClientesID: TAutoIncField
      FieldName = 'ID'
      ReadOnly = True
    end
    object qryClientesNOME: TStringField
      FieldName = 'NOME'
      Size = 100
    end
    object qryClientesCPF_CNPJ: TStringField
      FieldName = 'CPF_CNPJ'
    end
    object qryClientesTELEFONE: TStringField
      FieldName = 'TELEFONE'
    end
    object qryClientesEMAIL: TStringField
      FieldName = 'EMAIL'
      Size = 100
    end
    object qryClientesENDERECO: TStringField
      FieldName = 'ENDERECO'
      Size = 150
    end
    object qryClientesCIDADE: TStringField
      FieldName = 'CIDADE'
      Size = 50
    end
    object qryClientesLIMITE_CREDITO: TBCDField
      FieldName = 'LIMITE_CREDITO'
      DisplayFormat = 'R$ #,##0.00'
      Precision = 10
      Size = 2
    end
    object qryClientesDATA_CADASTRO: TDateTimeField
      FieldName = 'DATA_CADASTRO'
    end
  end
  object qryUsuarios: TADOQuery
    Connection = wolferp
    CursorType = ctStatic
    AfterScroll = qryUsuariosAfterScroll
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM USUARIOS')
    Left = 40
    Top = 520
    object qryUsuariosID: TAutoIncField
      FieldName = 'ID'
      ReadOnly = True
    end
    object qryUsuariosNOME_COMPLETO: TStringField
      FieldName = 'NOME_COMPLETO'
      Size = 100
    end
    object qryUsuariosSENHA: TStringField
      FieldName = 'SENHA'
      Size = 4
    end
    object qryUsuariosNIVEL_ACESSO: TStringField
      FieldName = 'NIVEL_ACESSO'
    end
    object qryUsuariosCPF: TStringField
      FieldName = 'CPF'
      Size = 14
    end
    object qryUsuariosTELEFONE: TStringField
      FieldName = 'TELEFONE'
    end
    object qryUsuariosATIVO: TStringField
      FieldName = 'ATIVO'
      FixedChar = True
      Size = 1
    end
    object qryUsuariosMATRICULA: TStringField
      FieldName = 'MATRICULA'
      FixedChar = True
      Size = 4
    end
    object qryUsuariosACESSO_VENDAS: TStringField
      FieldName = 'ACESSO_VENDAS'
      FixedChar = True
      Size = 1
    end
    object qryUsuariosACESSO_ESTOQUE: TStringField
      FieldName = 'ACESSO_ESTOQUE'
      FixedChar = True
      Size = 1
    end
    object qryUsuariosACESSO_FINANCEIRO: TStringField
      FieldName = 'ACESSO_FINANCEIRO'
      FixedChar = True
      Size = 1
    end
    object qryUsuariosACESSO_CADASTROS: TStringField
      FieldName = 'ACESSO_CADASTROS'
      FixedChar = True
      Size = 1
    end
    object qryUsuariosACESSO_RELATORIOS: TStringField
      FieldName = 'ACESSO_RELATORIOS'
      FixedChar = True
      Size = 1
    end
    object qryUsuariosACESSO_LOGS_SISTEMA: TStringField
      FieldName = 'ACESSO_LOGS_SISTEMA'
      FixedChar = True
      Size = 1
    end
    object qryUsuariosACESSO_DESENVOLVEDOR: TStringField
      FieldName = 'ACESSO_DESENVOLVEDOR'
      FixedChar = True
      Size = 1
    end
  end
  object dsUsuarios: TDataSource
    DataSet = qryUsuarios
    Left = 112
    Top = 520
  end
  object qryFornecedor: TADOQuery
    Connection = wolferp
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM FORNECEDOR')
    Left = 40
    Top = 464
    object qryFornecedorID: TAutoIncField
      FieldName = 'ID'
      ReadOnly = True
    end
    object qryFornecedorRAZAO_SOCIAL: TStringField
      FieldName = 'RAZAO_SOCIAL'
      Size = 100
    end
    object qryFornecedorNOME_FANTASIA: TStringField
      FieldName = 'NOME_FANTASIA'
      Size = 100
    end
    object qryFornecedorCNPJ: TStringField
      FieldName = 'CNPJ'
    end
    object qryFornecedorINSCRICAO_ESTADUAL: TStringField
      FieldName = 'INSCRICAO_ESTADUAL'
    end
    object qryFornecedorCONTATO_VENDEDOR: TStringField
      FieldName = 'CONTATO_VENDEDOR'
      Size = 50
    end
    object qryFornecedorTELEFONE: TStringField
      FieldName = 'TELEFONE'
    end
    object qryFornecedorEMAIL: TStringField
      FieldName = 'EMAIL'
      Size = 100
    end
    object qryFornecedorENDERECO: TStringField
      FieldName = 'ENDERECO'
      Size = 150
    end
    object qryFornecedorCIDADE: TStringField
      FieldName = 'CIDADE'
      Size = 50
    end
    object qryFornecedorDATA_CADASTRO: TDateTimeField
      FieldName = 'DATA_CADASTRO'
    end
  end
  object dsFornecedor: TDataSource
    DataSet = qryFornecedor
    Left = 112
    Top = 465
  end
  object dsRelatorios: TDataSource
    DataSet = qryRelatorios
    Left = 112
    Top = 408
  end
  object dsFinanceiro: TDataSource
    DataSet = qryFinanceiro
    Left = 111
    Top = 344
  end
  object dsConsultas: TDataSource
    DataSet = qryConsultas
    Left = 112
    Top = 216
  end
  object dsExec: TDataSource
    DataSet = qryExec
    Left = 112
    Top = 280
  end
  object dsClientes: TDataSource
    DataSet = qryClientes
    Left = 112
    Top = 24
  end
  object qryGenerica: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 40
    Top = 584
  end
  object qrySessoes: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'DELETE FROM SESSOES WHERE ID = '#39)
    Left = 40
    Top = 640
  end
  object qryMsg: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT MENSAGEM FROM MENSAGENS_SISTEMA WHERE LIDA = "N"')
    Left = 40
    Top = 704
  end
  object dsSessoes: TDataSource
    DataSet = qrySessoes
    Left = 112
    Top = 640
  end
  object dsGenerica: TDataSource
    DataSet = qryGenerica
    Left = 112
    Top = 584
  end
  object dsMsg: TDataSource
    DataSet = qryMsg
    Left = 112
    Top = 704
  end
  object qryConfig: TADOQuery
    Connection = wolferp
    Parameters = <
      item
        Name = 'Chave'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'SELECT VALOR FROM CONFIGURACOES_SISTEMA WHERE CHAVE = :Chave')
    Left = 208
    Top = 24
  end
  object dsConfig: TDataSource
    DataSet = qryConfig
    Left = 272
    Top = 24
  end
  object qryDevLogs: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 208
    Top = 88
  end
  object dsDevLogs: TDataSource
    DataSet = qryDevLogs
    OnDataChange = dsDevLogsDataChange
    Left = 272
    Top = 88
  end
  object qryDashboard: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 208
    Top = 152
  end
  object dsDashboard: TDataSource
    DataSet = qryDashboard
    Left = 272
    Top = 152
  end
  object qryEmpresa: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 208
    Top = 216
  end
  object dsEmpresa: TDataSource
    DataSet = qryEmpresa
    Left = 272
    Top = 216
  end
  object dsVendas: TDataSource
    DataSet = qryVendas
    Left = 272
    Top = 280
  end
  object qryVendas: TADOQuery
    Connection = wolferp
    CursorType = ctStatic
    AfterScroll = qryVendasAfterScroll
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM VENDAS')
    Left = 208
    Top = 280
  end
  object dsDiretrizes_wolferp: TDataSource
    DataSet = qryDiretrizes_wolferp
    Left = 272
    Top = 344
  end
  object qryDiretrizes_wolferp: TADOQuery
    Connection = wolferp
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM DIRETRIZES_WOLFERP')
    Left = 208
    Top = 344
    object qryDiretrizes_wolferpID: TAutoIncField
      FieldName = 'ID'
      ReadOnly = True
    end
    object qryDiretrizes_wolferpMODULO: TStringField
      FieldName = 'MODULO'
      Size = 30
    end
    object qryDiretrizes_wolferpCHAVE: TStringField
      FieldName = 'CHAVE'
      Size = 50
    end
    object qryDiretrizes_wolferpVALOR: TStringField
      FieldName = 'VALOR'
      Size = 255
    end
    object qryDiretrizes_wolferpTIPO: TStringField
      FieldName = 'TIPO'
      FixedChar = True
      Size = 2
    end
    object qryDiretrizes_wolferpDESCRICAO: TStringField
      FieldName = 'DESCRICAO'
      Size = 200
    end
    object qryDiretrizes_wolferpOPCOES: TMemoField
      FieldName = 'OPCOES'
      BlobType = ftMemo
    end
    object qryDiretrizes_wolferpDATA_ATUALIZACAO: TDateTimeField
      FieldName = 'DATA_ATUALIZACAO'
    end
  end
  object qryLogin: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT MATRICULA, SENHA, ATIVO FROM USUARIOS WHERE ATIVO = '#39'S'#39';')
    Left = 208
    Top = 406
  end
  object dsLogin: TDataSource
    DataSet = qryLogin
    Left = 272
    Top = 406
  end
  object dsItemVenda: TDataSource
    DataSet = qryItemVenda
    Left = 272
    Top = 465
  end
  object qryItemVenda: TADOQuery
    Connection = wolferp
    Parameters = <
      item
        Name = 'ID'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'SELECT * FROM VENDAS_ITENS WHERE ID_VENDA = :ID')
    Left = 208
    Top = 465
  end
  object qryOrcamentos: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM ORCAMENTOS')
    Left = 208
    Top = 522
  end
  object dsOrcamentos: TDataSource
    DataSet = qryOrcamentos
    Left = 272
    Top = 522
  end
  object dsAux: TDataSource
    Left = 272
    Top = 584
  end
  object qryAux: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 208
    Top = 584
  end
  object dsOrcamentosItens: TDataSource
    DataSet = qryOrcamentosItens
    Left = 272
    Top = 640
  end
  object qryOrcamentosItens: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM ORCAMENTOS_ITENS')
    Left = 208
    Top = 640
  end
  object qryPermissoes: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM PERMISSOES_ACESSO')
    Left = 208
    Top = 704
  end
  object dsPermissoes: TDataSource
    DataSet = qryPermissoes
    Left = 272
    Top = 704
  end
  object qryListaEmpresa: TADOQuery
    Connection = wolferp
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT ID, NOME_FANTASIA FROM EMPRESA ORDER BY NOME_FANTASIA')
    Left = 392
    Top = 24
  end
  object dsListaEmpresa: TDataSource
    DataSet = qryListaEmpresa
    Left = 464
    Top = 24
  end
  object dsReposicao: TDataSource
    DataSet = qryReposicao
    Left = 464
    Top = 88
  end
  object qryReposicao: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT '
      '  P.ID, '
      '  P.NOME, '
      '  P.ESTOQUE, '
      '  P.ESTOQUE_MINIMO,'
      '  P.PRECO_CUSTO,'
      '  (P.ESTOQUE_MINIMO - P.ESTOQUE) as SUGESTAO_COMPRA,'
      '  F.NOME_FANTASIA as FORNECEDOR'
      'FROM PRODUTOS P'
      'LEFT JOIN FORNECEDOR F ON F.ID = P.ID_FORNECEDOR'
      'WHERE P.ESTOQUE <= P.ESTOQUE_MINIMO'
      'ORDER BY F.NOME_FANTASIA, P.NOME')
    Left = 392
    Top = 88
  end
  object dsEtiquetas: TDataSource
    DataSet = qryEtiquetas
    Left = 464
    Top = 152
  end
  object qryEtiquetas: TADOQuery
    Connection = wolferp
    Parameters = <>
    Left = 392
    Top = 152
  end
  object dsPagar: TDataSource
    DataSet = qryPagar
    Left = 464
    Top = 216
  end
  object qryPagar: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM FINANCEIRO WHERE TIPO = '#39'D'#39)
    Left = 392
    Top = 216
  end
  object qryReceber: TADOQuery
    Connection = wolferp
    Parameters = <>
    SQL.Strings = (
      'SELECT '
      '  F.ID, '
      '  F.NUMERO_DOC, '
      '  C.NOME AS CLIENTE, '
      '  F.DATA_VENCIMENTO, '
      '  F.VALOR AS VALOR_ORIGINAL, '
      '  F.VALOR_PAGO, '
      '  (F.VALOR - F.VALOR_PAGO) AS RESTANTE,'
      '  F.STATUS'
      'FROM FINANCEIRO F'
      'LEFT JOIN CLIENTES C ON C.ID = F.ID_CLIENTE'
      'WHERE F.TIPO = '#39'R'#39' ')
    Left = 392
    Top = 280
  end
  object dsReceber: TDataSource
    DataSet = qryReceber
    Left = 464
    Top = 280
  end
end
