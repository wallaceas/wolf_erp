unit uDmDados;

interface

uses
  Windows, SysUtils, Classes, DB, ADODB, Dialogs;

type
  TdmDados = class(TDataModule)
    wolferp: TADOConnection;
    dsEstoque: TDataSource;
    qryEstoque: TADOQuery;
    qryProdutos: TADOQuery;
    dsProdutos: TDataSource;
    qryConsultas: TADOQuery;
    qryExec: TADOQuery;
    qryFinanceiro: TADOQuery;
    qryRelatorios: TADOQuery;
    qryClientes: TADOQuery;
    qryUsuarios: TADOQuery;
    dsUsuarios: TDataSource;
    qryFornecedor: TADOQuery;
    dsFornecedor: TDataSource;
    dsRelatorios: TDataSource;
    dsFinanceiro: TDataSource;
    dsConsultas: TDataSource;
    dsExec: TDataSource;
    dsClientes: TDataSource;
    qryGenerica: TADOQuery;
    qrySessoes: TADOQuery;
    qryMsg: TADOQuery;
    dsSessoes: TDataSource;
    dsGenerica: TDataSource;
    dsMsg: TDataSource;
    qryConfig: TADOQuery;
    dsConfig: TDataSource;
    qryDevLogs: TADOQuery;
    dsDevLogs: TDataSource;
    qryDashboard: TADOQuery;
    dsDashboard: TDataSource;
    qryProdutosID: TAutoIncField;
    qryProdutosCODIGO_BARRAS: TStringField;
    qryProdutosDESCRICAO: TStringField;
    qryProdutosID_FORNECEDOR: TIntegerField;
    qryProdutosPRECO_CUSTO: TBCDField;
    qryProdutosPRECO_VENDA: TBCDField;
    qryProdutosESTOQUE_ATUAL: TBCDField;
    qryProdutosESTOQUE_MINIMO: TBCDField;
    qryProdutosUNIDADE: TStringField;
    qryProdutosATIVO: TStringField;
    qryEmpresa: TADOQuery;
    dsEmpresa: TDataSource;
    qryClientesID: TAutoIncField;
    qryClientesNOME: TStringField;
    qryClientesCPF_CNPJ: TStringField;
    qryClientesTELEFONE: TStringField;
    qryClientesEMAIL: TStringField;
    qryClientesENDERECO: TStringField;
    qryClientesCIDADE: TStringField;
    qryClientesLIMITE_CREDITO: TBCDField;
    qryClientesDATA_CADASTRO: TDateTimeField;
    dsVendas: TDataSource;
    qryVendas: TADOQuery;
    qryFornecedorID: TAutoIncField;
    qryFornecedorRAZAO_SOCIAL: TStringField;
    qryFornecedorNOME_FANTASIA: TStringField;
    qryFornecedorCNPJ: TStringField;
    qryFornecedorINSCRICAO_ESTADUAL: TStringField;
    qryFornecedorCONTATO_VENDEDOR: TStringField;
    qryFornecedorTELEFONE: TStringField;
    qryFornecedorEMAIL: TStringField;
    qryFornecedorENDERECO: TStringField;
    qryFornecedorCIDADE: TStringField;
    qryFornecedorDATA_CADASTRO: TDateTimeField;
    dsDiretrizes_wolferp: TDataSource;
    qryDiretrizes_wolferp: TADOQuery;
    qryUsuariosID: TAutoIncField;
    qryUsuariosNOME_COMPLETO: TStringField;
    qryUsuariosSENHA: TStringField;
    qryUsuariosNIVEL_ACESSO: TStringField;
    qryUsuariosCPF: TStringField;
    qryUsuariosTELEFONE: TStringField;
    qryUsuariosATIVO: TStringField;
    qryUsuariosMATRICULA: TStringField;
    qryUsuariosACESSO_VENDAS: TStringField;
    qryUsuariosACESSO_ESTOQUE: TStringField;
    qryUsuariosACESSO_FINANCEIRO: TStringField;
    qryUsuariosACESSO_CADASTROS: TStringField;
    qryUsuariosACESSO_RELATORIOS: TStringField;
    qryUsuariosACESSO_LOGS_SISTEMA: TStringField;
    qryUsuariosACESSO_DESENVOLVEDOR: TStringField;
    qryDiretrizes_wolferpID: TAutoIncField;
    qryDiretrizes_wolferpMODULO: TStringField;
    qryDiretrizes_wolferpCHAVE: TStringField;
    qryDiretrizes_wolferpVALOR: TStringField;
    qryDiretrizes_wolferpTIPO: TStringField;
    qryDiretrizes_wolferpDESCRICAO: TStringField;
    qryDiretrizes_wolferpOPCOES: TMemoField;
    qryDiretrizes_wolferpDATA_ATUALIZACAO: TDateTimeField;
    qryLogin: TADOQuery;
    dsLogin: TDataSource;
    dsItemVenda: TDataSource;
    qryItemVenda: TADOQuery;
    qryOrcamentos: TADOQuery;
    dsOrcamentos: TDataSource;
    dsAux: TDataSource;
    qryAux: TADOQuery;
    dsOrcamentosItens: TDataSource;
    qryOrcamentosItens: TADOQuery;
    qryPermissoes: TADOQuery;
    dsPermissoes: TDataSource;
    qryListaEmpresa: TADOQuery;
    dsListaEmpresa: TDataSource;
    dsReposicao: TDataSource;
    qryReposicao: TADOQuery;
    dsEtiquetas: TDataSource;
    qryEtiquetas: TADOQuery;
    dsPagar: TDataSource;
    qryPagar: TADOQuery;
    qryReceber: TADOQuery;
    dsReceber: TDataSource;
    procedure qryUsuariosAfterScroll(DataSet: TDataSet);
    procedure dsDevLogsDataChange(Sender: TObject; Field: TField);
    procedure DataModuleCreate(Sender: TObject);
    procedure qryVendasAfterScroll(DataSet: TDataSet);
    //procedure dsReceberDataChange(Sender: TObject; Field: TField);
    //procedure dsProdutosStateChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    iIDUsuarioLogado: Integer;
    UsuarioLogadoID: Integer;
    UsuarioNome: String;
    UsuarioEhAdm: Boolean;
    UsuarioLogin: String;
    UsuarioNivel: String;
    SessaoID: Integer;
    function VerificarParametro(pChave: String): Boolean;
    procedure GravarLog(Tela, Acao, Descricao: String);
    procedure ExecSQL(pComando: String);
  end;

var
  dmDados: TdmDados;

implementation

uses uPrincipal;


{$R *.dfm}

procedure TdmDados.ExecSQL(pComando: String);
begin
  qryExec.Close;
  qryExec.SQL.Text := pComando;
  qryExec.ExecSQL;
end;

function TdmDados.VerificarParametro(pChave: String): Boolean;
begin
  qryConfig.Close;
  qryConfig.Parameters.ParamByName('Chave').Value := pChave;
  qryConfig.Open;
    if (qryConfig.IsEmpty) or (qryConfig.FieldByName('VALOR').AsString <> 'S') then
      Result := False
    else
      Result := True;
      qryConfig.Close;
end;

procedure TdmDados.DataModuleCreate(Sender: TObject);
begin
  wolferp.LoginPrompt := False;

  wolferp.ConnectionString := wolferp.ConnectionString + ';User ID=root;Password=';

  try
    wolferp.Connected := True;
  except
    on E: Exception do
      ShowMessage('Erro ao conectar no banco: ' + E.Message);
  end;
end;

procedure TdmDados.dsDevLogsDataChange(Sender: TObject; Field: TField);
begin
  if Assigned(frmPrincipal) then
  begin
    if (qryDevLogs.Active = False) or (qryDevLogs.IsEmpty) then
    begin
      frmPrincipal.mmoDetalheLog.Lines.Clear;
    end
    else
    begin
      frmPrincipal.mmoDetalheLog.Text := qryDevLogs.FieldByName('DESCRICAO').AsString;
    end;
  end;
end;

//procedure TdmDados.dsReceberDataChange(Sender: TObject; Field: TField);
//var
//  ValorOriginal, JurosCalculado, ValorTotal: Currency;
//  Vencimento: TDate;
//  DiasAtraso: Integer;
//begin
//  if dmDados.qryReceber.IsEmpty then Exit;
//
//  // Só calcula se estiver ABERTO
//  if dmDados.qryReceber.FieldByName('STATUS').AsString = 'ABERTO' then
//  begin
//    Vencimento    := dmDados.qryReceber.FieldByName('DATA_VENCIMENTO').AsDateTime;
//    ValorOriginal := dmDados.qryReceber.FieldByName('VALOR').AsCurrency;
//
//    // Calcula os dias (Data Atual - Data Vencimento)
//    DiasAtraso := Trunc(Date) - Trunc(Vencimento);
//
//    if DiasAtraso > 0 then
//    begin
//      // REGRAS DE NEGÓCIO (Exemplo: 2% Multa + 0.33% Juros ao dia)
//      JurosCalculado := (ValorOriginal * 0.02) + (ValorOriginal * 0.0033 * DiasAtraso);
//      ValorTotal     := ValorOriginal + JurosCalculado;
//
//      lblDiasAtraso.Caption      := IntToStr(DiasAtraso) + ' dias';
//      lblJurosMulta.Caption      := FormatFloat('R$ #,##0.00', JurosCalculado);
//      lblValorAtualizado.Caption := FormatFloat('R$ #,##0.00', ValorTotal);
//
//      pnlRodapeDireitoFinanceiro.Visible := True; // Mostra o painel
//    end
//    else
//    begin
//      pnlRodapeDireitoFinanceiro.Visible := False; // Esconde se não estiver atrasado
//    end;
//  end
//  else
//    pnlRodapeDireitoFinanceiro.Visible := False;
//end;

//procedure TdmDados.dsProdutosStateChange(Sender: TObject);
//var
//  Editando: Boolean;
//begin
//  Editando := dmDados.qryProdutos.State in [dsInsert, dsEdit];
//
//  btnNovoProdutoCadastros.Enabled := not Editando; // Só pode clicar em Novo se NÃO estiver editando
//  btnExcluirProdutoCadastros.Enabled  := not Editando; // Não pode excluir enquanto edita
//  btnBuscarProduto.Enabled   := not Editando; // Não deve buscar enquanto edita
//
//  btnSalvarProdutoCadastros.Enabled   := Editando; // Só pode salvar SE estiver editando
//  btnCancelarProdutoCadastros.Enabled := Editando;     // Só pode cancelar SE estiver editando
//
//  gridProdutosCadastros.Enabled := not Editando; // Bloquear a Grid para não mudar de registro enquanto edita
//end;

procedure TdmDados.GravarLog(Tela, Acao, Descricao: String);
begin
  with qryGenerica do
  begin
    Close;
    SQL.Text := 'INSERT INTO LOGS_SISTEMA (ID_USUARIO, LOGIN_USUARIO, TELA, ACAO, DESCRICAO) ' +
                'VALUES (:IdUser, :Login, :Tela, :Acao, :Desc)';

    Parameters.ParamByName('IdUser').Value := UsuarioLogadoID;
    Parameters.ParamByName('Login').Value  := UsuarioNome; 
    Parameters.ParamByName('Tela').Value   := Tela;
    Parameters.ParamByName('Acao').Value   := Acao;
    Parameters.ParamByName('Desc').Value   := Descricao;
    
    ExecSQL;
  end;
end;

procedure TdmDados.qryUsuariosAfterScroll(DataSet: TDataSet);
begin
  if Assigned(frmPrincipal) then
    frmPrincipal.edtNovaSenhaUsu.Clear;
end;

procedure TdmDados.qryVendasAfterScroll(DataSet: TDataSet);
begin
  // Só busca itens se houver uma venda selecionada
  if (not qryVendas.Active) or (qryVendas.IsEmpty) then
  begin
    qryItemVenda.Close;
    Exit;
  end;

  qryItemVenda.Close;
  qryItemVenda.SQL.Text := 'SELECT I.*, P.NOME as PRODUTO ' +
                             'FROM VENDAS_ITENS I ' +
                             'LEFT JOIN PRODUTOS P ON P.ID = I.ID_PRODUTO ' +
                             'WHERE I.ID_VENDA = :IDVenda';
  qryItemVenda.Parameters.ParamByName('IDVenda').Value := qryVendas.FieldByName('ID').AsInteger;
  qryItemVenda.Open;
end;
end.
