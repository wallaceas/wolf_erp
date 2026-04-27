unit uPrincipal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, ComCtrls, ExtCtrls, StdCtrls, Grids, DateUtils, uDmDados, DB,
  DBCtrls, Mask, IdHashMessageDigest, CheckLst, dxGDIPlusClasses, Jpeg, DBGrids,
  ExtDlgs, ADODB, DBClient, IniFiles, Printers;

type
  TfrmPrincipal = class(TForm)
    PageControl1: TPageControl;
    tsHome: TTabSheet;
    tsVendasDesativado: TTabSheet;
    tsProdutosDesativado: TTabSheet;
    tsEstoque: TTabSheet;
    pnlGridProdutos: TPanel;
    gridProdutos: TDBGrid;
    pnl2: TPanel;
    pnlTopo: TPanel;
    pnlRodape: TPanel;
    lblTitulo: TLabel;
    edtCodigo: TEdit;
    lblNomeProduto: TLabel;
    lblMensagem: TLabel;
    StringGrid1: TStringGrid;
    lblTextoTotal: TLabel;
    lblTotalVenda: TLabel;
    tsFinanceiro: TTabSheet;
    pgcFinanceiro: TPageControl;
    tsTabExtrato: TTabSheet;
    pnlFiltroFin: TPanel;
    lblDe: TLabel;
    dtpInicio: TDateTimePicker;
    lblAte: TLabel;
    dtpFim: TDateTimePicker;
    pnlTotalFin: TPanel;
    dbgrdDataControls: TDBGrid;
    grpPeriodoDeDias: TGroupBox;
    rgFiltroTipo: TRadioGroup;
    tsRelatorios: TTabSheet;
    tabLancamento: TTabSheet;
    pnlLancEsquerda: TPanel;
    pnlFeedback: TPanel;
    lblNovoLanc: TLabel;
    rgTipo: TRadioGroup;
    lblDescricao: TLabel;
    edtDescricao: TEdit;
    pnl3: TPanel;
    btnSalvarLancamento: TBitBtn;
    pnlDireitaVenc: TPanel;
    lblVencimento: TLabel;
    dtpVencimento: TDateTimePicker;
    pnlEsquerdaVal: TPanel;
    lblValor: TLabel;
    edtValor: TEdit;
    pnlCardResumo: TPanel;
    lblTipoGrande: TLabel;
    lblValorGigante: TLabel;
    pnlFundoRelatorios: TPanel;
    tsFornecedorDesativado: TTabSheet;
    tsUsuariosDesativado: TTabSheet;
    pnlListaUsers: TPanel;
    gridUsuarios: TDBGrid;
    pnlDadosUsuarios: TPanel;
    pnlIdUsu: TPanel;
    dbtxtIdUsu: TDBText;
    pnlLoginUsu: TPanel;
    pnlNmCompletoUsu: TPanel;
    pnlMatriculaUsu: TPanel;
    pnlNvSenhaUsu: TPanel;
    dbedtLoginUsu: TDBEdit;
    lblLoginUsu: TLabel;
    dbedtNmCompletoUsu: TDBEdit;
    lbl2: TLabel;
    dbedtMatriculaUsu: TDBEdit;
    lbl5: TLabel;
    lblIdUsu: TLabel;
    edtNovaSenhaUsu: TEdit;
    lblNvSenhaUsu: TLabel;
    pnlAdminAtivo: TPanel;
    dbchkAdminUsu: TDBCheckBox;
    dbchkAtivoUsu: TDBCheckBox;
    btnExcluirUsuario: TBitBtn;
    btnSalvarUsuario: TBitBtn;
    btnNovoUsuario: TBitBtn;
    pnl4: TPanel;
    pnlFiltros: TPanel;
    pnlEscolha: TPanel;
    cbbTipoRelatorio: TComboBox;
    lblEscolhaRelatorio: TLabel;
    grpPeriodo: TGroupBox;
    dtpRelInicio: TDateTimePicker;
    dtpRelFim: TDateTimePicker;
    grpFiltros: TGroupBox;
    lblDeRel: TLabel;
    lblAteRel: TLabel;
    btnGerar: TBitBtn;
    pnl5: TPanel;
    chklstFiltrosRel: TCheckListBox;
    pnlListaForn: TPanel;
    grid2: TDBGrid;
    pnlDadosForn: TPanel;
    pnlIdFornecedor: TPanel;
    lblIdFornecedor: TLabel;
    dbtxtIdFornecedor: TDBText;
    pnlRazaoSocial: TPanel;
    lblRazaoSocialFornecedor: TLabel;
    pnlNomeFantasia: TPanel;
    lblNomeFantasiaFornecedor: TLabel;
    dbedtNomeFantasiaFornecedor: TDBEdit;
    pnlCNPJ: TPanel;
    pnlTelefone: TPanel;
    pnlEMAIL: TPanel;
    btn1: TBitBtn;
    btnSalvarForn: TBitBtn;
    btnNovoForn: TBitBtn;
    pnl9: TPanel;
    dbedtRazaoSocialFornecedor: TDBEdit;
    lblCNPJFornecedor: TLabel;
    dbedtCNPJ: TDBEdit;
    lblTelefoneFornecedor: TLabel;
    dbedtTelefone: TDBEdit;
    lblEmailFornecedor: TLabel;
    dbedtEMAILFornecedor: TDBEdit;
    tsDesenvolvedor: TTabSheet;
    pcDevTools: TPageControl;
    tsDevSessoes: TTabSheet;
    tsDevLogs: TTabSheet;
    tsDevSQL: TTabSheet;
    tsDevVersao: TTabSheet;
    tmrSessao: TTimer;
    pnlInfoSessao: TPanel;
    pnlEncerrarSessao: TPanel;
    pnl10: TPanel;
    btnEncerrarSessao: TBitBtn;
    gridSessoesUsuarios: TDBGrid;
    btnEnviarAviso: TBitBtn;
    grpDetalhesSessao: TGroupBox;
    lblTempoConectado: TLabel;
    lblStatusConexao: TLabel;
    mmoLogsAcoes: TMemo;
    pnlDetalheUsuarioEBuscar: TPanel;
    pnlBuscar: TPanel;
    btnbuscar: TBitBtn;
    pnlAtualizarLista: TPanel;
    btnAtualizarLista: TBitBtn;
    edtMensagem: TEdit;
    pnlFiltroLogs: TPanel;
    dtpInicioDevLogs: TDateTimePicker;
    dtpFimDevLogs: TDateTimePicker;
    cbxFiltroAcao: TComboBox;
    edtBuscaDevLogs: TEdit;
    btnBuscarDevLogs: TBitBtn;
    gridDevLogs: TDBGrid;
    mmoDetalheLog: TMemo;
    lblDeDevLogs: TLabel;
    lblAteDevLogs: TLabel;
    lblAcaoDevLogs: TLabel;
    lblPalavraChaveDevLogs: TLabel;
    mmoVersao: TMemo;
    tsCadastros: TTabSheet;
    tsPermissoes: TTabSheet;
    tvPermissoes: TTreeView;
    btnAtualizarArvore: TBitBtn;
    pnlGeralHome: TPanel;
    tmrInicioHome: TTimer;
    pnldtBuscaSessao: TPanel;
    edtBuscaSessao: TEdit;
    pnlFiltroUsuario: TPanel;
    rgFiltroPorUsuario: TRadioGroup;
    pnlPesquisaProdutos: TPanel;
    pnlLblPesquisa: TPanel;
    pnlEdtBuscaProduto: TPanel;
    pnl14: TPanel;
    pnlBtnBuscaProduto: TPanel;
    lblPesquisaProduto: TLabel;
    edtBusca: TEdit;
    btnBuscaProduto: TBitBtn;
    pnlPCCadastro: TPanel;
    pcCadastro: TPageControl;
    tsGeralProdutos: TTabSheet;
    tsValoresProdutos: TTabSheet;
    tsEstoqueProdutos: TTabSheet;
    pnlValores: TPanel;
    pnlPrecoCusto: TPanel;
    dbedtPrecoCusto: TDBEdit;
    lblPrecoCusto: TLabel;
    pnlMargemEstimada: TPanel;
    lblMargemEstimada: TLabel;
    pnlPrecoVenda: TPanel;
    lblPrecoVenda: TLabel;
    dbedtPrecoVenda: TDBEdit;
    edtMargemEstimada: TEdit;
    pnl12: TPanel;
    dbimgValores: TDBImage;
    pnlBtnCarregarImagem: TPanel;
    btnCarregarImagem: TBitBtn;
    btnRemover: TBitBtn;
    pnlGeral: TPanel;
    pnlIdProduto: TPanel;
    dbedtIdProdutoGeral: TDBEdit;
    lblIdProdutoGeral: TLabel;
    pnlCodBarras: TPanel;
    dbedtCodigoBarras: TDBEdit;
    pnlDescricaoProduto: TPanel;
    dbedtDescricaoProduto: TDBEdit;
    lblCodigoBarras: TLabel;
    lblDescricaoGeral: TLabel;
    pnlEstoqueMinimo: TPanel;
    pnlEstoqueAtual: TPanel;
    dbedtEstoqueMinimo: TDBEdit;
    dbedtEstoqueAtual: TDBEdit;
    lblEstoqueMinimo: TLabel;
    lblEstoqueAtual: TLabel;
    pnlUnidade: TPanel;
    lblUnidade: TLabel;
    dbcbbUnidade: TDBComboBox;
    pnlCategoria: TPanel;
    lblCategoria: TLabel;
    dbcbbCategoria: TDBComboBox;
    pnl11: TPanel;
    pnlProdutoAtivo: TPanel;
    dbchkProdutoAtivo: TDBCheckBox;
    dlgOpenValores: TOpenPictureDialog;
    pcCadastros: TPageControl;
    tsProdutosCadastros: TTabSheet;
    tsClientesCadastros: TTabSheet;
    tsFornecedorCadastros: TTabSheet;
    tsUsuarioCadastros: TTabSheet;
    tsVendas: TTabSheet;
    tsCaixaDiario: TTabSheet;
    tsContasAPagar: TTabSheet;
    tsContasAReceber: TTabSheet;
    pcVendas: TPageControl;
    tsNovaVenda: TTabSheet;
    tsConsultarVendas: TTabSheet;
    pnlTopoClientes: TPanel;
    pnlRodapeClientes: TPanel;
    pnlDireitaClientes: TPanel;
    splClientesCadastros: TSplitter;
    gridClienteFornecedor: TDBGrid;
    pcLateralClientes: TPageControl;
    tsDadosGeraisClientes: TTabSheet;
    tsEnderecoClientes: TTabSheet;
    tsFinanceiroClientes: TTabSheet;
    pnlCodigoLateral: TPanel;
    lblCodigoLateral: TLabel;
    dbedtCodigoLateral: TDBEdit;
    pnlNomeLateral: TPanel;
    lblNomeLateral: TLabel;
    dbedtNomeLateral: TDBEdit;
    pnlCPFLateral: TPanel;
    pnlTelefoneLateral: TPanel;
    pnlEmailLateral: TPanel;
    lblCPFCNPJLateral: TLabel;
    lblTelefoneLateral: TLabel;
    lblEmailLateral: TLabel;
    dbedtCPFCNPJLateral: TDBEdit;
    dbedtTelefoneLateral: TDBEdit;
    dbedtEmailLateral: TDBEdit;
    pnlEnderecoLateral: TPanel;
    pnlEnderecoCliente: TPanel;
    lblEnderecoClienteCadastros: TLabel;
    pnlCidadeCliente: TPanel;
    lblCidadelateral: TLabel;
    dbedtEnderecoLateral: TDBEdit;
    dbedtCidadLateral: TDBEdit;
    pnlFinanceiroClientes: TPanel;
    pnlFinanceiroLateral: TPanel;
    lblLimiteCredito: TLabel;
    dbedtLimiteCredito: TDBEdit;
    pnlValorAPrazoLateral: TPanel;
    lblValorAPrazo: TLabel;
    dbedtValorAPrazo: TDBEdit;
    pnlFiltrarClienteVendedorVendas: TPanel;
    pnlAbaFornecedor: TPanel;
    pnlTopoFornecedor: TPanel;
    pnlRodapeFornecedor: TPanel;
    pnlLateralFornecedor: TPanel;
    lblPesquisarFornecedorCadastros: TLabel;
    edtFornecedorBuscar: TEdit;
    pcFornecedor: TPageControl;
    splFornecedor: TSplitter;
    gridFornecedorCadastros: TDBGrid;
    tsDadosGeraisFornecedorcadastros: TTabSheet;
    tsContatoFornecedor: TTabSheet;
    tsEnderecoFornecedor: TTabSheet;
    pnlIdDadosEmpFornecedor: TPanel;
    pnlRazaoSocialFornecedor: TPanel;
    pnlNomeFantasiaFornecedor: TPanel;
    pnlCnpjFornecedor: TPanel;
    pnlInscricaoEstadualFornecedor: TPanel;
    lblIdDadosEmpFornecedor: TLabel;
    dbedtIdDadosEmpFornecedor: TDBEdit;
    lblRZSocialFornecedor: TLabel;
    dbedtRZSocialFornecedor: TDBEdit;
    lblNomeFantasia: TLabel;
    dbedtDadosEmpNomeFantasiaFornecedor: TDBEdit;
    lblDadosEmpCNPJFornecedor: TLabel;
    lblInscricaoEstadualFornecedor: TLabel;
    dbedtDadosEmpCNPJFornecedor: TDBEdit;
    dbedtIncricaoEstadualFornecedor: TDBEdit;
    pnlContatoVendedorFornecedor: TPanel;
    lblContatoVendedorFornecedor: TLabel;
    dbedtContatoVendedorFornecedor: TDBEdit;
    pnlTelefoneFornecedor: TPanel;
    pnlEmailPerdidosFornecedor: TPanel;
    lblTelefonewhatsappFornecedor: TLabel;
    lblEmailPedidosFornecedor: TLabel;
    dbedtTelefonewhatsappFornecedor: TDBEdit;
    dbedtEmailPedidosFornecedor: TDBEdit;
    pnlBotaoBuscarFornecedor: TPanel;
    btnBuscarFornecedor: TBitBtn;
    pnlBotaoNovoFornecedor: TPanel;
    btnNovoFornecedor: TBitBtn;
    pnlBotaoSalvarFornecedor: TPanel;
    pnlBotaoCancelarFornecedor: TPanel;
    btnSalvarFornecedor: TBitBtn;
    btnCancelarFornecedor: TBitBtn;
    pnlBotaoExcluirFornecedor: TPanel;
    btnExcluirFornecedor: TBitBtn;
    pnlEnderecoFornecedor: TPanel;
    pnlCidadeFornecedor: TPanel;
    lblEnderecoFornecedor: TLabel;
    lblCidadeFornecedor: TLabel;
    dbedtEnderecoFornecedor: TDBEdit;
    dbedtCidadeFornecedor: TDBEdit;
    lblPesquisarClientecadastros: TLabel;
    edtClienteBuscar: TEdit;
    pnlBotaoBuscarCliente: TPanel;
    btnBuscarCliente: TBitBtn;
    pnl30: TPanel;
    pnlAbaUsuario: TPanel;
    pnlPesquisarUsuario: TPanel;
    pnlBotaoUsuarioCadastros: TPanel;
    pnlLateralUsuarioCadastros: TPanel;
    pnlProdutosCadastrosPrincipal: TPanel;
    pnlPesquisarprodutos: TPanel;
    pnlBotoesCadastroProduto: TPanel;
    pnlDireitaProdutoCadastros: TPanel;
    pcLateralProdutoCadastros: TPageControl;
    splProdutosCadastros: TSplitter;
    gridProdutosCadastros: TDBGrid;
    tsDadosGerais: TTabSheet;
    tsValores: TTabSheet;
    tsEstoqueProdutosCadastros: TTabSheet;
    lblPesquisarProdutoCadastros: TLabel;
    edtProdutoCadastros: TEdit;
    pnlBotaoBuscarProduto: TPanel;
    btnBuscarProdutoCadastros: TBitBtn;
    pnlBotaoNovoCliente: TPanel;
    pnlBotaoSalvarCliente: TPanel;
    pnlBotaoCancelarCliente: TPanel;
    pnlBotaoExcluirCliente: TPanel;
    pnlBotaoNovoProduto: TPanel;
    pnlBotaoSalvarProduto: TPanel;
    pnlBotaoCancelarProduto: TPanel;
    pnlBotaoExcluirProduto: TPanel;
    btnNovoClienteCadastros: TBitBtn;
    btnSalvarClienteCadastros: TBitBtn;
    btnCancelarClienteCadastros: TBitBtn;
    btnExcluirClienteCadastros: TBitBtn;
    btnNovoProdutoCadastros: TBitBtn;
    btnSalvarProdutoCadastros: TBitBtn;
    btnCancelarProdutoCadastros: TBitBtn;
    btnExcluirProdutoCadastros: TBitBtn;
    pnl47: TPanel;
    pnl48: TPanel;
    pnl49: TPanel;
    pnlIdProdutoCadastros: TPanel;
    pnlDescricaoProdutoCadastros: TPanel;
    pnlCodigoBarrasProduto: TPanel;
    pnlUnidadeProduto: TPanel;
    pnlCategoriaProduto: TPanel;
    lblIdMercadoriaCadastros: TLabel;
    dbedtIdMercadoriaCadastros: TDBEdit;
    lblDescricaoCadastros: TLabel;
    lblCodigoBarrasProduto: TLabel;
    lblUnidadeCadastros: TLabel;
    lblCategoriaCadastros: TLabel;
    dbedtDescricaoCadastros: TDBEdit;
    dbedtCodigoBarrasCadastros: TDBEdit;
    dbcbbUnidadeCadastros: TDBComboBox;
    dbcbbCategoriaCadastros: TDBComboBox;
    pnlPrecoCustoProduto: TPanel;
    pnlPrecoVendaProduto: TPanel;
    lblPrecoCustoCadastros: TLabel;
    lblPrecoVendaProdutosCadastros: TLabel;
    dbedtPrecoCustoCadastros: TDBEdit;
    dbedtPrecoVendaCadastros: TDBEdit;
    pnl57: TPanel;
    pnl58: TPanel;
    lbl3: TLabel;
    lbl6: TLabel;
    lblValorCompraFornecedor: TLabel;
    pnlBtnNovoUsuario: TPanel;
    pnlBtnSalvarUsuario: TPanel;
    pnlBtnCancelarUsuario: TPanel;
    pnlBtnExcluirUsuario: TPanel;
    pnl62: TPanel;
    btnNovoUsuarioCadastros: TBitBtn;
    btnSalvarUsuarioCadastros: TBitBtn;
    btnCancelarUsuarioCadastros: TBitBtn;
    btnExcluirUsuarioCadastros: TBitBtn;
    lblPesquisarUsuarioCadastros: TLabel;
    edtBuscarUsuario: TEdit;
    pnlBtnBuscarUsuario: TPanel;
    btnBuscarUsuarioCadastros: TBitBtn;
    ts1: TTabSheet;
    pnl35: TPanel;
    pnlEditorDesenvolvedor: TPanel;
    pnlTopolConfigGlobaisSistema: TPanel;
    pnlRodapeConfigSistema: TPanel;
    splConfigSistema: TSplitter;
    gridConfigSistema: TDBGrid;
    lblConfigGlobaisSistema: TLabel;
    pcConfSistema: TPageControl;
    tsTipo1ConfigSistema: TTabSheet;
    tsTipo2ConfigSistema: TTabSheet;
    tsTipo3ConfigSistema: TTabSheet;
    pnl59: TPanel;
    dbtxt1: TDBText;
    pcLateralUsuarioCadastros: TPageControl;
    tsDadosGeraisUsuarioCadastros: TTabSheet;
    tsFuncaoUsuarioCadastros: TTabSheet;
    tsPermissoesUsuarioCadastros: TTabSheet;
    pnlIdUsuarioCadastros: TPanel;
    pnlNomeCompletoUsuarioCadastros: TPanel;
    pnlSenhaUsuarioCadastros: TPanel;
    pnlMatriculaUsuarioCadastros: TPanel;
    pnlTelefoneUsuarioCadastros: TPanel;
    pnlNivelPermissaoCadastrosUsuarioCadastros: TPanel;
    lblIdUsuarioCadastros: TLabel;
    lblNomecompletoUsuarioCadastros: TLabel;
    lblSenhaUsuarioCadastros: TLabel;
    lblMatriculaUsuarioCadastros: TLabel;
    lblTelefoneUsuarioCadastros: TLabel;
    dbedtIdUsuarioCadastros: TDBEdit;
    dbedtNomeCompletoUsuarioCadastros: TDBEdit;
    dbedtSenhaUsuarioCadastros: TDBEdit;
    dbedtMatriculaUsuarioCadastros: TDBEdit;
    pnlAtivoUsuarioCadastros: TPanel;
    dbedtTelefoneUsuarioCadastros: TDBEdit;
    gridUsuarioCadastros: TDBGrid;
    lblAtivoUsuarioCadastros: TLabel;
    chkAtivoUsuarioCadastros: TCheckBox;
    pnlFuncaoUsuarioCadastros: TPanel;
    lblFuncaoUsuarioCadastros: TLabel;
    dbcbbFuncaoUsuarioCadastros: TDBComboBox;
    pnlEquipamentoUsuarioCadastros: TPanel;
    lblEquipamentoUsuarioCadastros: TLabel;
    dbcbbEquipamentoUsuarioCadastros: TDBComboBox;
    grpPermissoesCadastrosUsuario: TGroupBox;
    dbchkAcessoProdutosUsuarioCadastros: TDBCheckBox;
    dbchkAcessoClientesUsuarioCadastros: TDBCheckBox;
    dbchkAcessoFornecedorUsuarioCadastros: TDBCheckBox;
    dbchkAcessoUsuariosUsuarioCadastros: TDBCheckBox;
    pnlPermissoesVendas: TPanel;
    grpPermissoesVendasUsuario: TGroupBox;
    dbchkAcessoNovaVendaUsuarioCadastros: TDBCheckBox;
    dbchkAcessoConsultaVendasUsuarioCadastros: TDBCheckBox;
    pnlPermissoesFinanceiras: TPanel;
    grpPermissoesFinanceiraUsuario: TGroupBox;
    dbchkAcessoExtratoUsuarioCadastros: TDBCheckBox;
    dbchkAcessoLancamentoUsuarioCadastros: TDBCheckBox;
    dbchkAcessoCaixaDiarioUsuarioCadastros: TDBCheckBox;
    dbchkAcessoContasaPagarUsuarioCadastros: TDBCheckBox;
    dbchkAcessoContasaReceberUsuarioCadastros: TDBCheckBox;
    pnlPermissoesOutrosUsuario: TPanel;
    grpPermissoesOutrosUsuario: TGroupBox;
    dbchkAcessoMovEstoqueUsuarioCadastros: TDBCheckBox;
    dbchkAcessoRelatorioUsuarioCadastros: TDBCheckBox;
    tsOrcamento: TTabSheet;
    pnlOpcoesdeVenda: TPanel;
    grpDadosVendasVendas: TGroupBox;
    dblkcbbDadosVendaClienteVendas: TDBLookupComboBox;
    lblDadosVendaClienteVendas: TLabel;
    lblVendedorDadosVendaVendas: TLabel;
    pnlInformacaoProdutoVendas: TPanel;
    pnlInformacaoProdutoAdicionadoVendas: TPanel;
    gridItemVendaVendas: TDBGrid;
    pnlCodBarrasProdutoVendas: TPanel;
    lblCodBarrasProdutoVendas: TLabel;
    dbedtCodBarrasProdutoVendas: TDBEdit;
    pnlQuantidadeProdutoVendas: TPanel;
    pnlDescontoUnitarioProdutoVendas: TPanel;
    pnlImagemProdutoVendas: TPanel;
    pnlValorUnitarioeSubTotal: TPanel;
    lblQuantidadeProdutoVendas: TLabel;
    lblDescontoUnitarioVendas: TLabel;
    lblPrecoVendas: TLabel;
    dbedtQuantidadeProdutoVendas: TDBEdit;
    dbedtDescontoUnitarioVendas: TDBEdit;
    pnlBtnAdicionarProdutoVendas: TPanel;
    btnAdicionarProdutoVendas: TBitBtn;
    pnlFormadePagamento: TPanel;
    pnlPCFormadePagamento: TPanel;
    pnlTotalaPagar: TPanel;
    pnlFinalizarVenda: TPanel;
    pnlCancelarVenda: TPanel;
    rgFormadePagamentoVendas: TRadioGroup;
    grpTotalaPagarVendas: TGroupBox;
    lblTotalaPagarVendas: TLabel;
    btnFinalizarVenda: TBitBtn;
    btnCancelarVenda: TBitBtn;
    dsItensVenda: TClientDataSet;
    dsItensVendaID_PRODUTO: TIntegerField;
    dsItensVendaDESCRICAO: TStringField;
    dsItensVendaQTD: TFloatField;
    dsItensVendaUNITARIO: TCurrencyField;
    dsItensVendaTOTAL: TCurrencyField;
    dblkcbbVendedorVendas: TDBLookupComboBox;
    lblLimitedeCreditoVendas: TLabel;
    dsItensVendaDESCONTO: TCurrencyField;
    pnlFiltrosPesquisaVendas: TPanel;
    pnlGridsInformacaoVendas: TPanel;
    grpFiltrosPesquisaVendas: TGroupBox;
    dtpDeFiltrosPesquisaVenda: TDateTimePicker;
    edtNPedidoFiltrosPesquisaVendas: TEdit;
    chkVendaCanceladaFiltrosPesquisaVendas: TCheckBox;
    btnPesquisarFiltrosPesquisaVendas: TBitBtn;
    lblDeFiltrosPesquisaVendas: TLabel;
    lblAteFiltrosPesquisaVendas: TLabel;
    dtpAteFiltrosPesquisaVendas: TDateTimePicker;
    lblNPedidoFiltrosPesquisaVendas: TLabel;
    pnl68: TPanel;
    pnl70: TPanel;
    pnlConsultarVendasItensVendas: TPanel;
    pnlConsultarVendasVendas: TPanel;
    gridCabecalhoVendas: TDBGrid;
    gridItensVendas: TDBGrid;
    pnl63: TPanel;
    pnl72: TPanel;
    pnlInformacaoProdutoOrcamento: TPanel;
    pnl75: TPanel;
    pnl76: TPanel;
    pnlFormadePagamentoOrcamento: TPanel;
    pnl78: TPanel;
    pnl79: TPanel;
    pnl80: TPanel;
    pnl81: TPanel;
    pnlCodigoBarrasOrcamento: TPanel;
    pnlQuantidadeOrcamento: TPanel;
    pnl84: TPanel;
    pnl85: TPanel;
    gridItemVendaOrcamento: TDBGrid;
    rgFormaPagamentoOrcamento: TRadioGroup;
    grp3: TGroupBox;
    lblTotalaPagarVendasOrcamento: TLabel;
    btnSalvarOrcamentoVendas: TBitBtn;
    btnTransformarVendaVendas: TBitBtn;
    pgcDetalhePagtoOrcamento: TPageControl;
    tsCartaoCreditoOrcamento: TTabSheet;
    tsPixOrcamento: TTabSheet;
    tsBoletoOrcamento: TTabSheet;
    pgcDetalhePagto: TPageControl;
    tsCartaoCredito: TTabSheet;
    tsPix: TTabSheet;
    tsBoleto: TTabSheet;
    pnlBandeiraVendas: TPanel;
    lblBandeiraVendas: TLabel;
    cbxBandeira: TComboBox;
    pnlParcelamentoVendas: TPanel;
    lblParcelamentoVendas: TLabel;
    cbxParcelas: TComboBox;
    lblValorParcelasVendas: TLabel;
    pnlimgQRCode: TPanel;
    pnlGerarPicVendas: TPanel;
    pnlPixCopiaColaVendas: TPanel;
    imgQRCode: TImage;
    btnGerarPix: TBitBtn;
    edtCopiaCola: TEdit;
    tsCartaoDebito: TTabSheet;
    tsDinheiro: TTabSheet;
    pnlDataVencimentoVendas: TPanel;
    lblDtVencimento: TLabel;
    dtpVencimentoBoleto: TDateTimePicker;
    pnlBotaoGerarBoletoVendas: TPanel;
    btnGerarBoleto: TBitBtn;
    lblCodigoBarrasOrcamento: TLabel;
    lblQuantidadeOrcamento: TLabel;
    lblDescontoUnitarioOrcamento: TLabel;
    dbedtCodigoBarrasOrcamento: TDBEdit;
    dbedtQuantidadeOrcamento: TDBEdit;
    grpPesquisaOrcamento: TGroupBox;
    pnl65: TPanel;
    dbedtNumOrcamento: TDBEdit;
    btnPesquisarOrcamento: TBitBtn;
    btnLimparOrcamento: TBitBtn;
    dbedtDescontoUnitarioOrcamento: TDBEdit;
    btnAdicionarProdutoOrcamento: TBitBtn;
    pnlCodigoBoletoVendas: TPanel;
    edtBoletoGerado: TEdit;
    pnl7: TPanel;
    pnlMenuInicial: TPanel;
    pnlLogoEmpresaLogada: TPanel;
    imgLogoEmpresaLogada: TImage;
    pnlUsuarioePerfilLogado: TPanel;
    lblNomeEmpresa: TLabel;
    lblUsuarioEPerfilLogado: TLabel;
    pnlConfiguracao: TPanel;
    btnConfiguracao: TSpeedButton;
    pnlSairSistema: TPanel;
    btnSairSistema: TSpeedButton;
    pnlAcessoRapido: TPanel;
    pnlBotaoNovaVendaHome: TPanel;
    pnlBotaoOrcamentosHome: TPanel;
    pnlBotaoClientesHome: TPanel;
    pnlBotaoProdutosHome: TPanel;
    pnlBotaoFinanceiroHome: TPanel;
    pnlBotaoRelatorioHome: TPanel;
    btnAcessoRapidoNovaVenda: TBitBtn;
    btnAcessoRapidoOrcamentos: TBitBtn;
    btnAcessoRapidoClientes: TBitBtn;
    btnAcessoRapidoProdutos: TBitBtn;
    btnAcessoRapidoFinanciero: TBitBtn;
    btnAcessoRapidoRelatorios: TBitBtn;
    pnl8: TPanel;
    pnlInformacoesSaldoDiario: TPanel;
    pnlAvisoSistemaeTarefas: TPanel;
    pnlVendas: TPanel;
    pnlAReceber: TPanel;
    pnlAPagarHoje: TPanel;
    pnlEstoqueCritico: TPanel;
    pnlAvisoSistema: TPanel;
    lblCardVendasHojeHome: TLabel;
    lblCardAReceberHojeHome: TLabel;
    lblCardAPagarHojeHome: TLabel;
    lblCardEstoqueCriticoHome: TLabel;
    pnlTotalVendasResumo: TPanel;
    lblTotalVendasHome: TLabel;
    lblValorVendasResumoHome: TLabel;
    pnlTotalAReceberResumo: TPanel;
    lblTotalAReceberHome: TLabel;
    lblTotalAReceberResumoHome: TLabel;
    pnlTotalAPagarResumo: TPanel;
    lblTotalAPagarHome: TLabel;
    lblTotalAPagarResumoHome: TLabel;
    pnlEstoque: TPanel;
    mmoAvisosSistema: TMemo;
    pnl82: TPanel;
    btnAtualizarDashboard: TBitBtn;
    lbl7: TLabel;
    lblTotalEstoqueCriticoResumoHome: TLabel;
    pnlTarefasLembretesHome: TPanel;
    pnlAdicionarNomeTarefa: TPanel;
    pnlAdicionarTextoTarefas: TPanel;
    chklstTarefasLembretesHome: TCheckListBox;
    lblTarefasLembretesHome: TLabel;
    dbedtTarefasLembretesHome: TDBEdit;
    pnlBtnAdicionarTarefasLembreteHome: TPanel;
    btnAdicionarTarefasLembreteHome: TBitBtn;
    splTarefasHome: TSplitter;
    lblDescricaoProduto: TLabel;
    pnlPermissaoSupervisor: TPanel;
    grpPermissaoDeSupervisor: TGroupBox;
    dbchkAcessoGerenciaPermissoes: TDBCheckBox;
    pnlPermissoesDesenvolvedor: TPanel;
    pgcPermissoesDesenvolvedor: TPageControl;
    ts2: TTabSheet;
    pnlInformacaoExtraSistema: TPanel;
    tsUtilitarios: TTabSheet;
    pgcConfig: TPageControl;
    tsVisual: TTabSheet;
    pnl91: TPanel;
    pnl92: TPanel;
    pnlAparencia: TPanel;
    pnl94: TPanel;
    grpVisual: TGroupBox;
    rdgTema: TRadioGroup;
    chk1: TCheckBox;
    chk2: TCheckBox;
    pnl93: TPanel;
    pnl95: TPanel;
    grp2: TGroupBox;
    chk3: TCheckBox;
    chk4: TCheckBox;
    tsDispositivos: TTabSheet;
    pnl96: TPanel;
    pnlImpressorasDispositivos: TPanel;
    grpImpressorasDispositivos: TGroupBox;
    lblImpressorasPedidos: TLabel;
    lblImpressoraCupom: TLabel;
    pnl98: TPanel;
    cbxImpressoraA4: TComboBox;
    pnl99: TPanel;
    cbxImpressoraCupom: TComboBox;
    grpConfigCupomDispositivos: TGroupBox;
    tsDadosEmpresa: TTabSheet;
    lblMensagemRodapeCupom: TLabel;
    edtMensagemRodapeCupom: TEdit;
    rdgLarguraPapelDispositivos: TRadioGroup;
    pnl97: TPanel;
    pnl100: TPanel;
    pnl101: TPanel;
    pnl102: TPanel;
    lblNomeFantasiaUtilitarios: TLabel;
    edtNomeFantasiaUtilitarios: TEdit;
    pnl103: TPanel;
    lblEnderecoTelefoneUtilitarios: TLabel;
    edtEnderecoTelefoneUtilitarios: TEdit;
    pnl104: TPanel;
    mmoTextoPadraoOrcamentosUtilitarios: TMemo;
    lblTextoPadraoOrcamentosUtilitarios: TLabel;
    splTextoPadraoOrcamentosUtilitarios: TSplitter;
    lblDealhedoTmemo: TLabel;
    pnlPermissaoExcluirCadastros: TPanel;
    grpPermissaoExcluirCadastros: TGroupBox;
    dbchkExcluirCadastros: TDBCheckBox;
    lblSacWolfERP: TLabel;
    dbimgFotoProdutoVendas: TDBImage;
    pnlSalvarAlteracaoUtilitarios: TPanel;
    btnSalvarConfigUtilitarios: TBitBtn;
    pnlRestaurarPadroesUtilitarios: TPanel;
    btnRestaurarUtilitarios: TBitBtn;
    dbimgFotoProdutoOrcamento: TDBImage;
    lblPesquisarOrcamento: TLabel;
    pnlFotoProduto: TPanel;
    dbimgFotoProduto: TDBImage;
    pnlDescontoProdutoCadastros: TPanel;
    lblDescontoProduto: TLabel;
    dbedtDescontoProduto: TDBEdit;
    pgcEstoque: TPageControl;
    tsMovimentacao: TTabSheet;
    tsHistorico: TTabSheet;
    tsReposicao: TTabSheet;
    pnl1: TPanel;
    pnl13: TPanel;
    pnlConfirmarMovimentacaoEstoque: TPanel;
    pnl16: TPanel;
    pnlCodBarrasIdProduto: TPanel;
    lblCodBarrasIdProduto: TLabel;
    edtBuscaProdutoEstoque: TEdit;
    pnl17: TPanel;
    pnlOperacoes: TPanel;
    rdgOperacao: TRadioGroup;
    pnl19: TPanel;
    pnlQuantidadeEstoque: TPanel;
    lblQuantidadeEstoque: TLabel;
    edtQuantidadeEstoque: TEdit;
    pnlNotaFiscalDocumento: TPanel;
    lblNotaFiscalDocumento: TLabel;
    edtDocumentoEstoque: TEdit;
    pnlInformacaoSaldoAtualEstoque: TPanel;
    pnlInformaacaoMovimentoEstoque: TPanel;
    lblInformacaoSaldoAtualEstoque: TLabel;
    lblInformacaoMovimentoEstoque: TLabel;
    pnlInformacaoSaldoFinalEstoque: TPanel;
    lblInformacaoSaldoFinalEstoque: TLabel;
    pnlConfirmarMovimentoEstoque: TPanel;
    btnConfirmarMovimentoEstoque: TBitBtn;
    pnlTopoHistoricoEstoque: TPanel;
    pnl20: TPanel;
    pnll: TPanel;
    dtpDeHistoricoEstoque: TDateTimePicker;
    dtpAteHistoricoEstoque: TDateTimePicker;
    lblDeHistoricoEstoque: TLabel;
    lblAteHistoricoEstoque: TLabel;
    lblTipoMovimentoHistoricoEstoque: TLabel;
    cbxTipoMovimentoHistoricoEstoque: TComboBox;
    pnl21: TPanel;
    gridHistoricoEstoque: TDBGrid;
    btn2: TBitBtn;
    pnl18: TPanel;
    pnl22: TPanel;
    pnl23: TPanel;
    pnl24: TPanel;
    pnlGeraOrdemCompra: TPanel;
    pnl26: TPanel;
    gridReposicaoEstoque: TDBGrid;
    tsEtiquetas: TTabSheet;
    btnFiltrarReposicao: TBitBtn;
    lblFornecedorPreferencialReposicaoEstoque: TLabel;
    dblkcbbFiltroFornecedorReposicaoEstoque: TDBLookupComboBox;
    pnl6: TPanel;
    btnGerarPedidoCompraEstoque: TBitBtn;
    lblClienteOrcamentoVendas: TLabel;
    dblkcbbClienteOrcamento: TDBLookupComboBox;
    lbl1: TLabel;
    dblkcbbVendedorOrcamentosVendas: TDBLookupComboBox;
    splGridCabecalhoVndas: TSplitter;
    pnlImprimir2ViaConsultarVendas: TPanel;
    pnlCancelarVendaConsultarVendas: TPanel;
    btnCancelarVendaConsultarVendas: TBitBtn;
    btnImprimir2ViaConsultarVendas: TBitBtn;
    splGridItensVendasConsultarVendas: TSplitter;
    pnlEstoqueHome: TPanel;
    btnAcessoRapidoEstoque: TBitBtn;
    pnlModeloImpressaoEtiquetasEstoque: TPanel;
    pnlFiltrarProdutoEtiquetasEstoque: TPanel;
    pnl27: TPanel;
    pnlBuscarProdutoEtiquetasEstoque: TPanel;
    lblPesquisarProdutoEtiquetasEstoque: TLabel;
    gridFiltrarProdutoEtiquetasEstoque: TDBGrid;
    pnlBotoesAddRemoveLimpa: TPanel;
    dbedtBuscarProdutoEtiquetasEstoque: TDBEdit;
    gridFilaImpressaoEtiquetasEstoque: TDBGrid;
    pnl34: TPanel;
    rdgModeloImpressaoEtiquetasEstoque: TRadioGroup;
    pnlBotaoImprimirEtiquetasEstoque: TPanel;
    btnImprimirEtiquetasEstoque: TBitBtn;
    pnlBotaoBuscarProdutoEtiquetasEstoque: TPanel;
    btnBuscarProdutoEtiquetasEstoque: TBitBtn;
    btnAdicionarEtiquetasEstoque: TBitBtn;
    pnlEspaco1: TPanel;
    btnRemoveEtiquetasEstoque: TBitBtn;
    pnlEspaco2: TPanel;
    btnLimparEtiquetasEstoque: TBitBtn;
    dsEtiquetas: TClientDataSet;
    dsEtiquetasID_PRODUTO: TIntegerField;
    dsEtiquetasNOME: TStringField;
    dsEtiquetasBARRAS: TStringField;
    dsEtiquetasPRECO: TCurrencyField;
    dsEtiquetasCOPIAS: TIntegerField;
    pnl15: TPanel;
    pnl25: TPanel;
    pnl28: TPanel;
    pnl29: TPanel;
    rdgStatusContasAReceberFinanceiro: TRadioGroup;
    pnl31: TPanel;
    lblPesquisarClienteContaReceberFinanceiro: TLabel;
    edtClienteContasReceberFincanceiro: TEdit;
    grid3: TDBGrid;
    pnlVencimentoContasReceberFinanceiro: TPanel;
    lblDeVencimentoContasReceberFinanceiro: TLabel;
    dtpDeVencimentoContasReceberFinanceiro: TDateTimePicker;
    dtpAteVencimentoContasReceberFinanceiro: TDateTimePicker;
    lblAteVencimentoContasReceberFinanceiro: TLabel;
    pnlRodapeEsquerdoFinanceiro: TPanel;
    pnlRodapeDireitoFinanceiro: TPanel;
    btnNovoTituloManualContasReceberFinanceiro: TBitBtn;
    btnReceberTituloContasReceberFinanceiro: TBitBtn;
    btnGerarReciboContasReceberFinanceiro: TBitBtn;
    pnlEspaco1Financeiro: TPanel;
    pnlEspaco2Financeiro: TPanel;
    lblValorAtualizado: TLabel;
    pnl32: TPanel;
    lblJurosMulta: TLabel;
    pnl33: TPanel;
    lblDiasAtraso: TLabel;
    pnl36: TPanel;
    pnl37: TPanel;
    pnl38: TPanel;
    rdgStatusContasPagarFinanceiro: TRadioGroup;
    gridContasPagarFinanceiro: TDBGrid;
    pnl39: TPanel;
    lblPesquisarFornecedorContasPagar: TLabel;
    dblkcbbFiltrarFornecedorContasPagar: TDBLookupComboBox;
    pnl40: TPanel;
    lblDeContasPagarFornecedor: TLabel;
    dtpDeContasPagarFornecedor: TDateTimePicker;
    lblAteContasPagarFornecedor: TLabel;
    dtpAteContasPagarFornecedor: TDateTimePicker;
    pnlBotoesContasPagarFinanceiro: TPanel;
    pnlLabelContasPagarFinanceiro: TPanel;
    dbchkCancelarVendaUsuarioCadastros: TDBCheckBox;
    dbchkLiberarVendaAltaUsuarioCadastros: TDBCheckBox;
    btnNovoLancamentoContasPagarFinanceiro: TBitBtn;
    pnlEspaco1RodapeEsquerdoContasPagar: TPanel;
    btnBaixarContasContasPagarFinanceiro: TBitBtn;
    pnlEspaco2RodapeEsquerdoContasPagar: TPanel;
    btnEstornarPagamentoContasPagarFinanceiro: TBitBtn;
    lblTotalPagarPeriodoContasPagarFinanceiro: TLabel;
    pnlEspaco1RodapeDireitoContasPagar: TPanel;
    lblTotalSelecionadoContasPagarFinanceiro: TLabel;
    pnl41: TPanel;
    pnl42: TPanel;
    pnl43: TPanel;
    pnl44: TPanel;
    dsOrcamentosItens: TClientDataSet;
    dsOrcamentosItensID: TIntegerField;
    pnl45: TPanel;
    pnlProdutoSelecionadoMovimentacaoEstoque: TPanel;
    lblNomeProdutoSelecionado: TLabel;
    lblCodBarrasProdutoSelecionado: TLabel;
    lblUnidadeProdutoSelecionado: TLabel;
    pnl46: TPanel;
    lblPrecoAtualProdutoSelecionado: TLabel;
    lblPrecoVendaProdutoSelecionado: TLabel;
    pnl50: TPanel;
    pnlObservacoesEstoque: TPanel;
    lblObservacoesEstoque: TLabel;
    mmoObservacoesEstoque: TMemo;
    pnl51: TPanel;
    lblCustoUnitarioMovimentacaoEstoque: TLabel;
    edtCustoUnitarioMovimnetacaoEstoque: TEdit;
    pnl52: TPanel;
    dblkcbbFiltrarFornecedorMovimentacaoEstoque: TDBLookupComboBox;
    lblFiltrarFornecedorMovimentacaoEstoque: TLabel;
    splExpandirAreaObservacoesEstoque: TSplitter;
    lblOperacaoDica: TLabel;
    pnlSelecionados: TPanel;
    lblNomeProdutoSelecionadoMovimentacaoEstoque: TLabel;
    lblCodBarrasProdutoSelecionadoMovimentacaoEstoque: TLabel;
    lblUnidadeProdutoSelecionadoMovimentacaoEstoque: TLabel;
    pnl53: TPanel;
    lblPrecoAtualProdutoSelecionadoMovimentacaoEstoque: TLabel;
    lblPrecoVendaProdutoSelecionadoMovimentacaoEstoque: TLabel;
    pnl54: TPanel;
    btnBuscarProdutoMovimentacaoEstoque: TBitBtn;
    procedure btnFinanceiroLateralClick(Sender: TObject);
    procedure btnHomeLateralClick(Sender: TObject);
    procedure btnEstoqueLateralClick(Sender: TObject);
    procedure btnRelatoriosLateralClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    //procedure btnNovoClick(Sender: TObject);
    //procedure gridProdutosCellClick(Column: TColumn);
    //procedure btnSalvarClick(Sender: TObject);
    //procedure tsVendasDesativadoShow(Sender: TObject);
    //procedure edtBuscaEstoqueChange(Sender: TObject);
    //procedure dbgrdEstoqueDrawColumnCell(Sender: TObject; const Rect: TRect;
      //DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure rgTipoMovimentoClick(Sender: TObject);
    //procedure dbgrdEstoqueCellClick(Column: TColumn);
    //procedure btnConfirmarMovClick(Sender: TObject);
    procedure dbgrdDataControlsDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure rgTipoClick(Sender: TObject);
    procedure edtValorExit(Sender: TObject);
    procedure btnSalvarLancamentoClick(Sender: TObject);
    procedure pnlFeedbackResize(Sender: TObject);
    procedure edtValorKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edtDescricaoKeyPress(Sender: TObject; var Key: Char);
    procedure cbbTipoRelatorioChange(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnVendasLateralClick(Sender: TObject);
    procedure btnCadastrosLateralClick(Sender: TObject);
    //procedure btnNovoUsuarioClick(Sender: TObject);
    //procedure btnSalvarUsuarioClick(Sender: TObject);
    //procedure btnExcluirUsuarioClick(Sender: TObject);
    procedure btnFecharSistemaClick(Sender: TObject);
    procedure btnSairSistemaClick(Sender: TObject);
    //procedure btnNovoFornClick(Sender: TObject);
    procedure btnSalvarFornClick(Sender: TObject);
    //procedure btn1Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure tmrSessaoTimer(Sender: TObject);
    procedure btnEncerrarSessaoClick(Sender: TObject);
    procedure btnEnviarAvisoClick(Sender: TObject);
    procedure btnAtualizarListaClick(Sender: TObject);
    procedure btnbuscarClick(Sender: TObject);
    procedure rbApenasVendedoresClick(Sender: TObject);
    procedure rbApenasEstoqueClick(Sender: TObject);
    procedure rbTodosClick(Sender: TObject);
    //procedure grpDetalhesConexaoClick(Sender: TObject);
    procedure edtMensagemKeyPress(Sender: TObject; var Key: Char);
    procedure btnBuscarDevLogsClick(Sender: TObject);
    procedure gridDevLogsDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure tsDevVersaoShow(Sender: TObject);
    procedure tsPermissoesShow(Sender: TObject);
    procedure btnAtualizarArvoreClick(Sender: TObject);
    procedure btnVendaHomeClick(Sender: TObject);
    //procedure btnProdutoHomeClick(Sender: TObject);
    //procedure btnCadastrosHomeClick(Sender: TObject);
    procedure tmrInicioHomeTimer(Sender: TObject);
    procedure tsHomeShow(Sender: TObject);
    procedure gridSessoesUsuariosCellClick(Column: TColumn);
    //procedure btnFornecedorClick(Sender: TObject);
    //procedure btnUsuariosClick(Sender: TObject);
    //procedure edtMensagemChange(Sender: TObject);
    //procedure rgFiltroPorUsuarioClick(Sender: TObject);
    //procedure tsProdutosDesativadoContextPopup(Sender: TObject; MousePos: TPoint;
     // var Handled: Boolean);
    //procedure dbchkProdutoAtivoClick(Sender: TObject);
    procedure btnCarregarImagemClick(Sender: TObject);
    procedure btnRemoverClick(Sender: TObject);
    procedure btnBuscaProdutoClick(Sender: TObject);
    procedure gridProdutosDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure gridProdutosTitleClick(Column: TColumn);
    //procedure lblBemVindoClick(Sender: TObject);
    //procedure PageControl1Change(Sender: TObject);
    //procedure tsVendasContextPopup(Sender: TObject; MousePos: TPoint;
      //var Handled: Boolean);
    //procedure btnBuscarProdutoClick(Sender: TObject);
    //procedure btnBuscarFornecedorClick(Sender: TObject);
    procedure btnNovoProdutoCadastrosClick(Sender: TObject);
    procedure btnNovoUsuarioCadastrosClick(Sender: TObject);
    procedure btnExcluirProdutoCadastrosClick(Sender: TObject);
    procedure btnCancelarProdutoCadastrosClick(Sender: TObject);
    procedure btnSalvarProdutoCadastrosClick(Sender: TObject);
    procedure btnBuscarProdutoCadastrosClick(Sender: TObject);
    procedure btnNovoClienteCadastrosClick(Sender: TObject);
    procedure btnSalvarClienteCadastrosClick(Sender: TObject);
    procedure btnCancelarClienteCadastrosClick(Sender: TObject);
    procedure btnExcluirClienteCadastrosClick(Sender: TObject);
    procedure btnNovoFornecedorClick(Sender: TObject);
    procedure btnSalvarFornecedorClick(Sender: TObject);
    procedure btnCancelarFornecedorClick(Sender: TObject);
    procedure btnExcluirFornecedorClick(Sender: TObject);
    procedure btnBuscarUsuarioCadastrosClick(Sender: TObject);
    //procedure dbedtIdUsuarioCadastrosChange(Sender: TObject);
    procedure btnSalvarUsuarioCadastrosClick(Sender: TObject);
    procedure btnCancelarUsuarioCadastrosClick(Sender: TObject);
    procedure btnExcluirUsuarioCadastrosClick(Sender: TObject);
    //procedure dbedtCodBarrasProdutoVendasChange(Sender: TObject);
    procedure dbedtCodBarrasProdutoVendasKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnAdicionarProdutoVendasClick(Sender: TObject);
    procedure btnFinalizarVendaClick(Sender: TObject);
    procedure rgFormadePagamentoVendasClick(Sender: TObject);
    procedure cbxParcelasChange(Sender: TObject);
    procedure btnTransformarVendaVendasClick(Sender: TObject);
    procedure btnPesquisarFiltrosPesquisaVendasClick(Sender: TObject);
    procedure btnPesquisarOrcamentoClick(Sender: TObject);
    procedure btnAcessoRapidoNovaVendaClick(Sender: TObject);
    //procedure pnl7Click(Sender: TObject);
    procedure btnAcessoRapidoOrcamentosClick(Sender: TObject);
    procedure btnAcessoRapidoClientesClick(Sender: TObject);
    procedure btnAcessoRapidoProdutosClick(Sender: TObject);
    procedure btnAcessoRapidoFinancieroClick(Sender: TObject);
    procedure btnAcessoRapidoRelatoriosClick(Sender: TObject);
    procedure btnAtualizarDashboardClick(Sender: TObject);
    procedure btnConfiguracaoClick(Sender: TObject);
    procedure btnAdicionarTarefasLembreteHomeClick(Sender: TObject);
    procedure chklstTarefasLembretesHomeClick(Sender: TObject);
    procedure btnSalvarOrcamentoVendasClick(Sender: TObject);
    //procedure btnGerarBoletoClick(Sender: TObject);
    procedure edtBuscaProdutoEstoqueKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edtQuantidadeEstoqueChange(Sender: TObject);
    procedure gridHistoricoEstoqueDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure btnFiltrarReposicaoClick(Sender: TObject);
    procedure btnGerarPedidoCompraEstoqueClick(Sender: TObject);
    //procedure btnCancelarVendaConsultarVendasClick(Sender: TObject);
    //procedure dbedtNumOrcamentoChange(Sender: TObject);
    procedure btnAcessoRapidoEstoqueClick(Sender: TObject);
    procedure btnAdicionarEtiquetasEstoqueClick(Sender: TObject);
    //procedure btnLimparOrcamentoClick(Sender: TObject);
    procedure btnRemoveEtiquetasEstoqueClick(Sender: TObject);
    procedure btnLimparEtiquetasEstoqueClick(Sender: TObject);
    procedure btnImprimirEtiquetasEstoqueClick(Sender: TObject);
    procedure btnBuscarProdutoEtiquetasEstoqueClick(Sender: TObject);
    procedure rdgStatusContasPagarFinanceiroClick(Sender: TObject);
    procedure gridContasPagarFinanceiroDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure btnReceberTituloContasReceberFinanceiroClick(Sender: TObject);
    procedure btnGerarReciboContasReceberFinanceiroClick(Sender: TObject);
    procedure rdgOperacaoClick(Sender: TObject);
    procedure btnConfirmarMovimentoEstoqueClick(Sender: TObject);
    procedure edtBuscaProdutoEstoqueChange(Sender: TObject);
    //procedure edtNomeFantasiaUtilitariosChange(Sender: TObject);
    //procedure btnSalvarConfigUtilitariosClick(Sender: TObject);
    //procedure btnCancelarVendaClick(Sender: TObject);
    //procedure edtBuscaSessaoChange(Sender: TObject);
  private
    { Private declarations }
    FAcao: string;
    procedure ExportarParaCSV(DataSet: TDataSet; NomeArquivo: string);
    procedure AtualizarFeedbackVisual;
    procedure AplicarPermissoes(UsuarioID: Integer);
    procedure CalcularTotalVenda;
    function TextoParaValor(Texto: String): Currency;
    procedure PopularParcelas;
    procedure ImprimirTermicaRAW;
    procedure CalcularJurosMulta;
//    procedure SolicitarAutorizacao;
    procedure AtualizarSimulacao;

  public
    { Public declarations }
    iIDUsuarioLogado: Integer;
    procedure CarregarArvoreSistema;
    procedure CarregarPermissoesTela;
    function BoolToChar(Valor: Boolean): String;
    //procedure SolicitarAutorizacao;
//    procedure ImprimirFolhaA4;
//    procedure ImprimirGondola;
//    procedure CalcularTotalOrcamento;
    //function Solicitar(TituloAcao: String; EhPerigoso: Boolean; AC_EXCLUIR_CADASTROS: String): Boolean;
  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

uses uLogin, uAutenticacao;

{$R *.dfm}

procedure TfrmPrincipal.AtualizarSimulacao;
var
  SaldoAtual, QtdDigitada, SaldoFinal: Double;
  Operacao: Integer;
begin
  // Se não tem produto selecionado, não faz nada
  if lblNomeProdutoEstoque.Caption = 'NENHUM PRODUTO SELECIONADO' then Exit;

  // Pega os valores da tela convertendo texto para número de forma segura
  SaldoAtual  := StrToFloatDef(lblInformacaoSaldoAtualEstoque.Caption, 0);
  QtdDigitada := StrToFloatDef(edtQuantidadeEstoque.Text, 0);
  Operacao    := rdgOperacao.ItemIndex;

  // Calcula o Futuro baseado no RadioGroup
  if Operacao = 0 then
  begin
    // [0] Entrada: Soma
    SaldoFinal := SaldoAtual + QtdDigitada;
    lblInformacaoMovimentoEstoque.Caption := 'Movimento: + ' + FloatToStr(QtdDigitada);
  end
  else if (Operacao = 1) or (Operacao = 2) then
  begin
    // [1] Saída ou [2] Perda: Subtrai
    SaldoFinal := SaldoAtual - QtdDigitada;
    lblInformacaoMovimentoEstoque.Caption := 'Movimento: - ' + FloatToStr(QtdDigitada);
  end
  else
  begin
    // A quantidade digitada passa a ser o novo saldo exato
    SaldoFinal := QtdDigitada;
    lblInformacaoMovimentoEstoque.Caption := 'Movimento: Ajuste de Balanço';
  end;

  // Atualiza os Labels do GroupBox "Previsão"
  lblInformacaoSaldoAtualEstoque.Caption := 'Saldo Atual: ' + FloatToStr(SaldoAtual);
  lblInformacaoSaldoFinalEstoque.Caption := 'Saldo Final: ' + FloatToStr(SaldoFinal);

  // O ALERTA VISUAL (Isso salva vidas no suporte)
  if SaldoFinal < 0 then
    lblInformacaoSaldoFinalEstoque.Font.Color := clRed  // Ficou negativo, vermelho!
  else
    lblInformacaoSaldoFinalEstoque.Font.Color := clGreen; // Positivo, verde ok!
end;

procedure TfrmPrincipal.CalcularJurosMulta;
var
  DataVencimento: TDate;
  ValorOriginal, JurosCalculado, ValorAtualizado: Currency;
  DiasAtraso: Integer;
begin
  // Limpa o painel se não tiver nada selecionado
  if dmDados.qryReceber.IsEmpty then Exit;

  // Se já está recebido, não calcula juros, apenas mostra o que foi pago
  if dmDados.qryReceber.FieldByName('STATUS').AsString = 'RECEBIDO' then
  begin
    lblDiasAtraso.Caption := 'Conta Paga';
    lblJurosMulta.Caption := 'R$ 0,00';
    lblValorAtualizado.Caption := FormatFloat('R$ #,##0.00', dmDados.qryReceber.FieldByName('VALOR_PAGO').AsCurrency);
    Exit;
  end;

  DataVencimento := dmDados.qryReceber.FieldByName('DATA_VENCIMENTO').AsDateTime;
  ValorOriginal  := dmDados.qryReceber.FieldByName('VALOR_ORIGINAL').AsCurrency;
  
  // Calcula dias de atraso comparando com hoje
  DiasAtraso := Trunc(Now) - Trunc(DataVencimento);

  if DiasAtraso > 0 then
  begin
    lblDiasAtraso.Caption := IntToStr(DiasAtraso) + ' dia(s)';

    // 2% de Multa Fixa + 0,33% de Juros por dia de atraso
    JurosCalculado := (ValorOriginal * 0.02) + (ValorOriginal * 0.0033 * DiasAtraso);
    
    ValorAtualizado := ValorOriginal + JurosCalculado;
    
    // Pinta de vermelho suave para alertar que tem juros
    pnlRodapeDireitoFinanceiro.Color := $008080FF;
  end
  else
  begin
    // Em dia (ou a vencer)
    lblDiasAtraso.Caption := 'Em dia';
    JurosCalculado := 0;
    ValorAtualizado := ValorOriginal;
    pnlRodapeDireitoFinanceiro.Color := $00E1F5FE; // Azul clarinho normal
  end;

  // Atualiza os Labels na tela
  lblJurosMulta.Caption := FormatFloat('R$ #,##0.00', JurosCalculado);
  lblValorAtualizado.Caption := FormatFloat('R$ #,##0.00', ValorAtualizado);
end;

procedure TfrmPrincipal.ImprimirTermicaRAW;
var
  ListaComandos: TStringList;
  i: Integer;
  NomeImpressoraRede, NomeProd: String;
begin
  ListaComandos := TStringList.Create;
  try
    dsEtiquetas.First; // Volta para o primeiro item da lista

    while not dsEtiquetas.Eof do
    begin
      // Pega os dados da memória
      NomeProd := Copy(dsEtiquetas.FieldByName('NOME').AsString, 1, 25); // Corta nome longo

      // Se o usuário pediu 5 etiquetas desse item, repetimos o comando 5 vezes
      for i := 1 to dsEtiquetas.FieldByName('COPIAS').AsInteger do
      begin
        // Exemplo para 40mm x 40mm
        ListaComandos.Add('L');              // Início da Etiqueta
        ListaComandos.Add('D11');            // Densidade/Calor
        ListaComandos.Add('H10');            // Temperatura

        // Nome do Produto
        ListaComandos.Add('121100000200020' + NomeProd);

        // Preço (Fonte Grande - Rotação 1, Fonte 4, X=060, Y=020)
        // FormatFloat garante que saia "10,50" bonitinho
        ListaComandos.Add('141100000600020R$ ' + FormatFloat('0.00', dsEtiquetas.FieldByName('PRECO').AsCurrency));

        // Código de Barras (EAN13 - Tipo 1E)
        ListaComandos.Add('1E1204000150020' + dsEtiquetas.FieldByName('BARRAS').AsString);

        ListaComandos.Add('E');              // Fim da Etiqueta (Imprime!)
      end;

      dsEtiquetas.Next; // Vai para o próximo produto da lista
    end;

    // --- ENVIANDO PARA A IMPRESSORA ---
    // Salvar num arquivo e copiar para a porta
    ListaComandos.SaveToFile('C:\WolfERP\Etiqueta.txt');

    // Aqui usamos o DOS para mandar direto pra porta (LPT1 ou USB Compartilhada)
    // Supondo que a impressora esteja compartilhada como "Argox" no Windows
    // WinExec é antigo, mas funciona bem no Delphi 2006 para isso.
    WinExec(PChar('cmd.exe /C copy /b C:\WolfERP\Etiqueta.txt \\127.0.0.1\Argox'), SW_HIDE);

    ShowMessage('Enviado para impressora térmica!');

  finally
    ListaComandos.Free;
  end;
end;

procedure TfrmPrincipal.CarregarPermissoesTela;
begin
  dmDados.qryPermissoes.Close;
  dmDados.qryPermissoes.SQL.Text := 'SELECT * FROM PERMISSOES_ACESSO WHERE ID_USUARIO = :ID';
  dmDados.qryPermissoes.Parameters.ParamByName('ID').Value := dmDados.qryUsuarios.FieldByName('ID').AsInteger;
  dmDados.qryPermissoes.Open;

  if dmDados.qryPermissoes.IsEmpty then
  begin
    // Se não tem registro, desmarca tudo
    //Permissão Cadastro
    dbchkAcessoProdutosUsuarioCadastros.Checked := False;
    dbchkAcessoClientesUsuarioCadastros.Checked := False;
    dbchkAcessoFornecedorUsuarioCadastros.Checked := False;
    dbchkAcessoUsuariosUsuarioCadastros.Checked := False;

    //Permissão Vendas
    dbchkAcessoNovaVendaUsuarioCadastros.Checked := False;
    dbchkAcessoConsultaVendasUsuarioCadastros.Checked := False;
    dbchkCancelarVendaUsuarioCadastros.Checked := False;
    dbchkLiberarVendaAltaUsuarioCadastros.Checked := False;

    //Permissão Financeira
    dbchkAcessoExtratoUsuarioCadastros.Checked := False;
    dbchkAcessoLancamentoUsuarioCadastros.Checked := False;
    dbchkAcessoCaixaDiarioUsuarioCadastros.Checked := False;
    dbchkAcessoContasaPagarUsuarioCadastros.Checked := False;
    dbchkAcessoContasaReceberUsuarioCadastros.Checked := False;

    //Outros
    dbchkAcessoMovEstoqueUsuarioCadastros.Checked := False;
    dbchkAcessoRelatorioUsuarioCadastros.Checked := False;

    //Permissão Supervisor
    dbchkAcessoGerenciaPermissoes.Checked := False;

    //Permissão Excluir Cadastros
    dbchkExcluirCadastros.Checked := False;

    Exit;
  end;

    //Permissão Cadastro
    dbchkAcessoProdutosUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_CAD_PRODUTOS').AsString = 'S';
    dbchkAcessoClientesUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_CAD_CLIENTES').AsString = 'S';
    dbchkAcessoFornecedorUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_CAD_FORNECEDOR').AsString = 'S';
    dbchkAcessoUsuariosUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_CAD_USUARIOS').AsString = 'S';

    //Permissão Vendas
    dbchkAcessoNovaVendaUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_VENDAS_NOVA').AsString = 'S';
    dbchkAcessoConsultaVendasUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_VENDAS_CONSULTAR').AsString = 'S';
    dbchkCancelarVendaUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_CANCELAR_VENDAS').AsString = 'S';
    dbchkLiberarVendaAltaUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_LIBERAR_VENDA_ALTA').AsString = 'S';

    //Permissão Financeira
    dbchkAcessoExtratoUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_FIN_EXTRATO').AsString = 'S';
    dbchkAcessoLancamentoUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_FIN_LANCAMENTO').AsString = 'S';
    dbchkAcessoCaixaDiarioUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_FIN_CAIXA').AsString = 'S';
    dbchkAcessoContasaPagarUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_FIN_PAGAR').AsString = 'S';
    dbchkAcessoContasaReceberUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_FIN_RECEBER').AsString = 'S';

    //Outros
    dbchkAcessoMovEstoqueUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_ESTOQUE_MOV').AsString = 'S';
    dbchkAcessoRelatorioUsuarioCadastros.Checked := dmDados.qryPermissoes.FieldByName('AC_RELATORIOS').AsString = 'S';

    //Permissão Supervisor
    dbchkAcessoGerenciaPermissoes.Checked := dmDados.qryPermissoes.FieldByName('AC_GERENCIAR_PERMISSOES').AsString = 'S';

    //Permissão Excluir Cadastros
    dbchkExcluirCadastros.Checked :=  dmDados.qryPermissoes.FieldByName('AS_EXCLUIR_CADASTROS').AsString = 'S';
end;

function TfrmPrincipal.TextoParaValor(Texto: String): Currency;
begin
  // Remove "R$" e espaços
  Texto := StringReplace(Texto, 'R$', '', [rfReplaceAll]);
  Texto := Trim(Texto);

  // Remove o PONTO de milhar (ex: 1.200 -> 1200)
  Texto := StringReplace(Texto, '.', '', [rfReplaceAll]);

  Result := StrToCurrDef(Texto, 0);
end;

procedure TfrmPrincipal.PopularParcelas;
var
  i: Integer;
  TotalVenda, ValorParcela: Currency;
  TextoOpcao: String;
begin
  cbxParcelas.Items.Clear;
  lblValorParcelasVendas.Caption := '';

  TotalVenda := TextoParaValor(lblTotalaPagarVendas.Caption);

  if TotalVenda <= 0 then Exit;

  // Loop para gerar de 1x até 12x
  for i := 1 to 12 do
  begin
    ValorParcela := TotalVenda / i;

    // Se for acima de 6x, adiciona 5% de juros, por exemplo:
    // if i > 6 then ValorParcela := ValorParcela * 1.05;

    // Formata: "1x de R$ 100,00", "2x de R$ 50,00"
    TextoOpcao := IntToStr(i) + 'x de ' + FormatFloat('R$ #,##0.00', ValorParcela);
    
    cbxParcelas.Items.Add(TextoOpcao);
  end;

  // Seleciona a primeira opção (1x) automaticamente
  cbxParcelas.ItemIndex := 0;

  cbxParcelasChange(Self);
end;

procedure TfrmPrincipal.CalcularTotalVenda;
var
  vTotal: Currency;
  vClone: TClientDataSet;
begin
  vTotal := 0;
  
  if dsItensVenda.IsEmpty then
  begin
    lblTotalaPagarVendas.Caption := 'R$ 0,00';
    Exit;
  end;

  vClone := TClientDataSet.Create(nil);
  try
    vClone.CloneCursor(dsItensVenda, True);
    vClone.First;
    while not vClone.Eof do
    begin
      vTotal := vTotal + vClone.FieldByName('VALOR_TOTAL').AsCurrency;
      vClone.Next;
    end;
  finally
    vClone.Free;
  end;

  lblTotalaPagarVendas.Caption := FormatFloat('R$ #,##0.00', vTotal);
end;

procedure TfrmPrincipal.AplicarPermissoes(UsuarioID: Integer);
var
  qry: TADOQuery;
  PermitidoPagar, PermitidoReceber, PermitidoCaixa: Boolean;
begin
  qry := TADOQuery.Create(nil);
  try
    qry.Connection := dmDados.wolferp;
    qry.SQL.Text := 'SELECT * FROM USUARIOS WHERE ID = ' + IntToStr(UsuarioID);
    qry.Open;

    if not qry.IsEmpty then
    begin
      PermitidoPagar   := (qry.FieldByName('AC_FIN_PAGAR').AsString = 'S');
      PermitidoReceber := (qry.FieldByName('AC_FIN_RECEBER').AsString = 'S');
      PermitidoCaixa   := (qry.FieldByName('AC_FIN_CAIXA').AsString = 'S');

      tsContasAPagar.TabVisible   := PermitidoPagar;
      tsContasAReceber.TabVisible := PermitidoReceber;
      tsCaixaDiario.TabVisible   := PermitidoCaixa;

      tsFinanceiro.TabVisible := PermitidoPagar or PermitidoReceber or PermitidoCaixa;

      tsProdutosCadastros.TabVisible     := (qry.FieldByName('AC_CAD_PRODUTOS').AsString = 'S');
      tsClientesCadastros.TabVisible     := (qry.FieldByName('AC_CAD_CLIENTES').AsString = 'S');

      tsCadastros.TabVisible := tsProdutosCadastros.TabVisible or tsClientesCadastros.TabVisible;

    end;
  finally
    qry.Free;
  end;
end;

function GetComputerNetName: string;
var
  buffer: array[0..255] of char;
  size: dword;
begin
  size := 256;
  if GetComputerName(buffer, size) then
    Result := buffer
  else
    Result := '';
end;

procedure TfrmPrincipal.AtualizarFeedbackVisual;
var
  vValor: Double;
begin
  vValor := StrToFloatDef(edtValor.Text, 0);

  if rgTipo.ItemIndex = 0 then
  begin
    pnlFeedback.Color := $00E1F0FF;
    lblTipoGrande.Caption := 'ENTRADA (RECEITA)';
    lblTipoGrande.Font.Color := clGreen;
    lblValorGigante.Font.Color := clGreen;
  end
  else
  begin
    pnlFeedback.Color := $00E1E1FF;
    lblTipoGrande.Caption := 'SAÍDA (DESPESA)';
    lblTipoGrande.Font.Color := clRed;
    lblValorGigante.Font.Color := clRed;
  end;

  lblValorGigante.Caption := FormatFloat('R$ #,##0.00', vValor);
end;

function GerarMD5(const Texto: String): String;
var
  idmd5: TIdHashMessageDigest5;
begin
  idmd5 := TIdHashMessageDigest5.Create;
  try
    Result := idmd5.AsHex(idmd5.HashValue(Texto));
  finally
    idmd5.Free;
  end;
end;

procedure TfrmPrincipal.btnSalvarUsuarioCadastrosClick(Sender: TObject);
var
  vSenhaDigitada: String;
  IDUser: Integer;
begin
  if dmDados.qryUsuarios.FieldByName('LOGIN').AsString = '' then
  begin
    ShowMessage('O Login é obrigatório.');
    dbedtLoginUsu.SetFocus;
    Exit;
  end;

  vSenhaDigitada := Trim(edtNovaSenhaUsu.Text);

  if (dmDados.qryUsuarios.State = dsInsert) and (vSenhaDigitada = '') then
  begin
    ShowMessage('Para novos usuários, é obrigatório definir uma senha inicial.');
    edtNovaSenhaUsu.SetFocus;
    Exit;
  end;

  if vSenhaDigitada <> '' then
  begin
    dmDados.qryUsuarios.FieldByName('SENHA').AsString := GerarMD5(vSenhaDigitada);
  end;

  try
    dmDados.qryUsuarios.Post;
    edtNovaSenhaUsu.Clear;
    ShowMessage('Usuário salvo com sucesso!');
  except
    on E: Exception do
      ShowMessage('Erro ao salvar: ' + E.Message);
  end;


  IDUser := dmDados.qryUsuarios.FieldByName('ID').AsInteger;

  // Verifica se já existe registro na tabela permissoes_acesso para este ID
  dmDados.qryAux.Close;
  dmDados.qryAux.SQL.Text := 'SELECT ID FROM PERMISSOES_ACESSO WHERE ID_USUARIO = ' + IntToStr(IDUser);
  dmDados.qryAux.Open;

  if dmDados.qryAux.IsEmpty then
  begin
    dmDados.ExecSQL('INSERT INTO PERMISSOES_ACESSO (ID_USUARIO) VALUES (' + IntToStr(IDUser) + ')');
  end;

    // Função auxiliar BoolToChar: Se Checked retorna 'S', se não 'N'
  dmDados.ExecSQL('UPDATE PERMISSOES_ACESSO SET ' +

    //Permissão Cadastros
    'AC_CAD_PRODUTOS = ' + QuotedStr(BoolToChar(dbchkAcessoProdutosUsuarioCadastros.Checked)) + ', ' +
    'AC_CAD_CLIENTES = ' + QuotedStr(BoolToChar(dbchkAcessoClientesUsuarioCadastros.Checked)) + ', ' +
    'AC_CAD_FORNECEDOR  = ' + QuotedStr(BoolToChar(dbchkAcessoFornecedorUsuarioCadastros.Checked)) + ', ' +
    'AC_CAD_USUARIOS = ' + QuotedStr(BoolToChar(dbchkAcessoUsuariosUsuarioCadastros.Checked)) + ', ' +

    //Permissão Vendas
    'AC_VENDAS_NOVA = ' + QuotedStr(BoolToChar(dbchkAcessoNovaVendaUsuarioCadastros.Checked)) + ', ' +
    'AC_VENDAS_CONSULTAR = ' + QuotedStr(BoolToChar(dbchkAcessoConsultaVendasUsuarioCadastros.Checked)) + ', ' +
    'AC_CANCELAR_VENDA = ' + QuotedStr(BoolToChar(dbchkCancelarVendaUsuarioCadastros.Checked)) + ', ' +
    'AC_LIBERAR_VENDA_ALTA = ' + QuotedStr(BoolToChar(dbchkLiberarVendaAltaUsuarioCadastros.Checked)) + ', ' +

    //Permissão Financeiro
    'AC_FIN_EXTRATO = ' + QuotedStr(BoolToChar(dbchkAcessoExtratoUsuarioCadastros.Checked)) + ', ' +
    'AC_FIN_LANCAMENTO = ' + QuotedStr(BoolToChar(dbchkAcessoLancamentoUsuarioCadastros.Checked)) + ', ' +
    'AC_FIN_CAIXA = ' + QuotedStr(BoolToChar(dbchkAcessoCaixaDiarioUsuarioCadastros.Checked)) + ', ' +
    'AC_FIN_PAGAR = ' + QuotedStr(BoolToChar(dbchkAcessoContasaPagarUsuarioCadastros.Checked)) + ', ' +
    'AC_FIN_RECEBER = ' + QuotedStr(BoolToChar(dbchkAcessoContasaReceberUsuarioCadastros.Checked)) + ', ' +

    //Outros
    'AC_ESTOQUE_MOV = ' + QuotedStr(BoolToChar(dbchkAcessoMovEstoqueUsuarioCadastros.Checked)) + ', ' +
    'AC_RELATORIOS = ' + QuotedStr(BoolToChar(dbchkAcessoRelatorioUsuarioCadastros.Checked)) + ', ' +

    //Permissão Supervisor
    'AC_GERENCIAR_PERMISSOES = ' + QuotedStr(BoolToChar(dbchkAcessoGerenciaPermissoes.Checked)) + ', ' +

    //Permissões Excluir Cadastros
    'AC_EXCLUIR_CADASTROS = ' + QuotedStr(BoolToChar(dbchkExcluirCadastros.Checked)) +

    ' WHERE ID_USUARIO = ' + IntToStr(IDUser));

  ShowMessage('Usuário e Permissões salvos com sucesso!');
end;

// Função auxiliar para limpar o código acima
function TfrmPrincipal.BoolToChar(Valor: Boolean): String;
begin
  if Valor then Result := 'S' else Result := 'N';
end;

procedure TfrmPrincipal.btnTransformarVendaVendasClick(Sender: TObject);
var
  ID_Orcamento, ID_NovaVenda: Integer;
begin
  if dmDados.qryOrcamentos.IsEmpty then Exit;
  if dmDados.qryOrcamentos.FieldByName('STATUS').AsString = 'APROVADO' then
  begin
    ShowMessage('Este orçamento já foi transformado em venda!');
    Exit;
  end;

  ID_Orcamento := dmDados.qryOrcamentos.FieldByName('ID').AsInteger;

  dmDados.wolferp.BeginTrans;
  try
    // Criar a Venda (Copia dados do cabeçalho do orçamento)
    dmDados.ExecSQL('INSERT INTO VENDAS (DATA_VENDA, ID_CLIENTE, ID_VENDEDOR, TOTAL_VENDA, FORMA_PAGTO) ' +
                    'SELECT NOW(), ID_CLIENTE, ID_VENDEDOR, TOTAL_ORCAMENTO, 0 ' + // 0=Dinheiro (padrão inicial)
                    'FROM ORCAMENTOS WHERE ID = ' + IntToStr(ID_Orcamento));

    // Pegar o ID da Venda nova
    dmDados.qryAux.Close;
    dmDados.qryAux.SQL.Text := 'SELECT LAST_INSERT_ID() as ID';
    dmDados.qryAux.Open;
    ID_NovaVenda := dmDados.qryAux.FieldByName('ID').AsInteger;

    // Copiar os Itens (Do Orçamento para Venda)
    dmDados.ExecSQL('INSERT INTO VENDAS_ITENS (ID_VENDA, ID_PRODUTO, QTD, UNITARIO, TOTAL) ' +
                    'SELECT ' + IntToStr(ID_NovaVenda) + ', ID_PRODUTO, QTD, UNITARIO, TOTAL ' +
                    'FROM ORCAMENTOS_ITENS WHERE ID_ORCAMENTO = ' + IntToStr(ID_Orcamento));

    // Baixar Estoque (Iterar sobre os itens do orçamento para baixar um a um)
    // Nota: Poderíamos fazer via SQL puro, mas via loop temos controle se o estoque ficar negativo
    dmDados.qryOrcamentosItens.First;
    while not dmDados.qryOrcamentosItens.Eof do
    begin
       dmDados.ExecSQL('UPDATE PRODUTOS SET ESTOQUE = ESTOQUE - ' +
                       StringReplace(dmDados.qryOrcamentosItens.FieldByName('QTD').AsString, ',', '.', [rfReplaceAll]) +
                       ' WHERE ID = ' + dmDados.qryOrcamentosItens.FieldByName('ID_PRODUTO').AsString);
       dmDados.qryOrcamentosItens.Next;
    end;

    // Atualizar Status do Orçamento
    dmDados.ExecSQL('UPDATE ORCAMENTOS SET STATUS = "APROVADO" WHERE ID = ' + IntToStr(ID_Orcamento));

    // CONFIRMA TUDO
    dmDados.wolferp.CommitTrans;

    ShowMessage('Orçamento transformado na Venda Nº ' + IntToStr(ID_NovaVenda) + ' com sucesso!');

    // Atualiza as grids
    btnPesquisarOrcamento.Click;

  except
    on E: Exception do
    begin
      // SE DER ERRO, DESFAZ TUDO
      dmDados.wolferp.RollbackTrans;
      ShowMessage('Erro ao converter: ' + E.Message);
    end;
  end;
end;

//procedure TfrmPrincipal.btn1Click(Sender: TObject);
//begin
//  if dmDados.qryFornecedor.IsEmpty then Exit;
//
//  if MessageDlg('Tem certeza que deseja excluir o fornecedor ' +
//     dmDados.qryFornecedor.FieldByName('NOME_FANTASIA').AsString + '?',
//     mtConfirmation, [mbYes, mbNo], 0) = mrYes then
//  begin
//    try
//      dmDados.qryFornecedor.Delete;
//    except
//      on E: Exception do
//        ShowMessage('Não foi possível excluir. Verifique se existem Produtos ligados a este Fornecedor.');
//    end;
//  end;
//end;

procedure TfrmPrincipal.btnNovoClienteCadastrosClick(Sender: TObject);
begin
  dmDados.qryClientes.Insert;

  if dbedtNomeLateral.CanFocus then
    dbedtNomeLateral.SetFocus;

  // dmDados.qryClientes.FieldByName('CLIENTES').AsFloat := 0;
end;

procedure TfrmPrincipal.btnSalvarClienteCadastrosClick(Sender: TObject);
begin
    if dmDados.qryClientes.State in [dsInsert, dsEdit] then
    try
      dmDados.qryClientes.Post;
      ShowMessage('Cliente salvo com sucesso!');

      edtClienteBuscar.Clear;
    except
      on E: Exception do
      begin
        ShowMessage('Erro ao gravar: ' + E.Message);
      end;
    end;
end;

procedure TfrmPrincipal.btnExcluirClienteCadastrosClick(Sender: TObject);
var
  NomeCliente : String;
begin
  // Verifica se tem algo selecionado
  if dmDados.qryClientes.IsEmpty then Exit;

  // CHAMA A TELA DE SEGURANÇA
  if frmAutenticacao.Solicitar('Confirmar Exclusão do Cliente?', True, 'AC_EXCLUIR_CADASTROS') then
  begin
    // Se o código chegou aqui, é porque a senha estava certa E tinha permissão
    dmDados.qryClientes.Delete;
    ShowMessage('Registro excluído com sucesso.');
  end;
  // Se cancelou ou errou a senha, nada acontece.

  if not dmDados.qryClientes.IsEmpty then
  begin
    NomeCliente := dmDados.qryClientes.FieldByName('NOME').AsString;

    if MessageDlg('Tem certeza que deseja excluir: ' + NomeCliente + '?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dmDados.qryClientes.Delete;

      ShowMessage('Cliente excluído com sucesso.');
    end;
  end
  else
  begin
    ShowMessage('Não há cliente selecionado para excluir.');
  end;
end;

procedure TfrmPrincipal.btnExcluirFornecedorClick(Sender: TObject);
var
  NomeFornecedor: String;
begin
    // Verifica se tem algo selecionado
  if dmDados.qryFornecedor.IsEmpty then Exit;

  // Parâmetros: ("Texto da Tela", EhPerigoso=True, "Coluna_No_Banco")
  if frmAutenticacao.Solicitar('Confirmar Exclusão do Fornecedor?', True, 'AC_EXCLUIR_CADASTROS') then
  begin
    // Se o código chegou aqui, é porque a senha estava certa E tinha permissão
    dmDados.qryFornecedor.Delete;
    ShowMessage('Registro excluído com sucesso.');
  end;
  // Se cancelou ou errou a senha, nada acontece.


  if not dmDados.qryFornecedor.IsEmpty then
  begin
    NomeFornecedor := dmDados.qryFornecedor.FieldByName('NOME').AsString;

    if MessageDlg('Tem certeza que deseja excluir: ' + NomeFornecedor + '?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dmDados.qryFornecedor.Delete;

      ShowMessage('Fornecedor excluído com sucesso.');
    end;
  end
  else
  begin
    ShowMessage('Não há fornecedor selecionado para excluir.');
  end;
end;

procedure TfrmPrincipal.btnExcluirProdutoCadastrosClick(Sender: TObject);
var
  NomeProduto: String;
begin
  // Verifica se tem algo selecionado
  if dmDados.qryProdutos.IsEmpty then Exit;

  // Parâmetros: ("Texto da Tela", EhPerigoso=True, "Coluna_No_Banco")
  if frmAutenticacao.Solicitar('Confirmar Exclusão do Produto?', True, 'AC_EXCLUIR_CADASTROS') then
  begin
    // Se o código chegou aqui, é porque a senha estava certa E tinha permissão
    dmDados.qryProdutos.Delete;
    ShowMessage('Registro excluído com sucesso.');
  end;
  // Se cancelou ou errou a senha, nada acontece.

  if not dmDados.qryProdutos.IsEmpty then
  begin
    NomeProduto := dmDados.qryProdutos.FieldByName('NOME').AsString;

    if MessageDlg('Tem certeza que deseja excluir: ' + NomeProduto + '?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dmDados.qryProdutos.Delete;

      ShowMessage('Produto excluído com sucesso.');
    end;
  end
  else
  begin
    ShowMessage('Não há produto selecionado para excluir.');
  end;
end;

procedure TfrmPrincipal.btnExcluirUsuarioCadastrosClick(Sender: TObject);
begin
  // Verifica se tem algo selecionado
  if dmDados.qryUsuarios.IsEmpty then Exit;

  // Parâmetros: ("Texto da Tela", EhPerigoso=True, "Coluna_No_Banco")
  if frmAutenticacao.Solicitar('Confirmar Exclusão do Usuário?', True, 'AC_EXCLUIR_CADASTROS') then
  begin
    // Se o código chegou aqui, é porque a senha estava certa E tinha permissão
    dmDados.qryUsuarios.Delete;
    ShowMessage('Registro excluído com sucesso.');
  end;
  // Se cancelou ou errou a senha, nada acontece.


  if MessageDlg('Tem certeza que deseja excluir este usuário?',
     mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    dmDados.qryUsuarios.Delete;
  end;
end;

//procedure TfrmProdutos.btnBuscarUsuarioCadastrosClick(Sender: TObject);
//begin
//  dmDados.qryProdutos.Close;
//  dmDados.qryProdutos.SQL.Clear;
//
//  if Trim(edtProdutoCadastros.Text) = '' then
//  begin
//    dmDados.qryProdutos.SQL.Add('SELECT * FROM PRODUTOS');
//  end
//  else
//  begin
//    // Dica: Use Parameters para evitar erros de SQL e segurança
//    dmDados.qryProdutos.SQL.Add('SELECT * FROM PRODUTOS WHERE ID = :Codigo');
//    dmDados.qryProdutos.Parameters.ParamByName('Codigo').Value := edtProdutoCadastros.Text;
//  end;
//
//  dmDados.qryProdutos.Open;
//
//  if dmDados.qryProdutos.IsEmpty and (Trim(edtProdutoCadastros.Text) <> '') then
//    ShowMessage('Nenhum produto encontrado com o código ' + edtProdutoCadastros.Text);
//end;

procedure TfrmPrincipal.btnNovoProdutoCadastrosClick(Sender: TObject);
begin
  dmDados.qryProdutos.Insert;

  if dbedtDescricaoCadastros.CanFocus then
    dbedtDescricaoCadastros.SetFocus;

  // dmDados.qryProdutos.FieldByName('ESTOQUE').AsFloat := 0;
end;

//procedure TfrmPrincipal.btnProdutoHomeClick(Sender: TObject);
//begin
//  PageControl1.ActivePage := tsProdutos;
//end;
//
//procedure TfrmPrincipal.btnUsuariosClick(Sender: TObject);
//begin
//  PageControl1.ActivePage := tsUsuarios;
//end;
//
//procedure TfrmPrincipal.btnUsuariosHomeClick(Sender: TObject);
//begin
//  PageControl1.ActivePage := tsUsuarios;
//end;

procedure TfrmPrincipal.btnAcessoRapidoFinancieroClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsFinanceiro;
end;

procedure TfrmPrincipal.btnAcessoRapidoOrcamentosClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsVendas;
end;

procedure TfrmPrincipal.btnAtualizarDashboardClick(Sender: TObject);
begin
  // ATUALIZAR VENDAS DE HOJE
  if dmDados.qryAux.Active then dmDados.qryAux.Close;
  dmDados.qryAux.SQL.Clear;
  dmDados.qryAux.SQL.Add('SELECT SUM(TOTAL_VENDA) AS TOTAL');
  dmDados.qryAux.SQL.Add('FROM VENDAS');
  dmDados.qryAux.SQL.Add('WHERE DATA_VENDA >= :Hoje');
  
  dmDados.qryAux.Parameters.ParamByName('Hoje').Value := Date; 
  dmDados.qryAux.Open;
  
  if dmDados.qryAux.FieldByName('TOTAL').IsNull then
    lblValorVendasResumoHome.Caption := 'R$ 0,00'
  else
    lblValorVendasResumoHome.Caption := FormatFloat('R$ #,##0.00', dmDados.qryAux.FieldByName('TOTAL').AsCurrency);

  // ATUALIZAR A RECEBER
  if dmDados.qryAux.Active then dmDados.qryAux.Close;
  dmDados.qryAux.SQL.Clear;
  dmDados.qryAux.SQL.Add('SELECT SUM(VALOR) AS TOTAL');
  dmDados.qryAux.SQL.Add('FROM FINANCEIRO');
  dmDados.qryAux.SQL.Add('WHERE DATA_VENCIMENTO = :Hoje');
  dmDados.qryAux.SQL.Add('AND STATUS = ''ABERTO''');
  dmDados.qryAux.SQL.Add('AND TIPO = ''R''');
  
  dmDados.qryAux.Parameters.ParamByName('Hoje').Value := Date;
  dmDados.qryAux.Open;
  
  if dmDados.qryAux.FieldByName('TOTAL').IsNull then
    lblTotalAReceberResumoHome.Caption := 'R$ 0,00'
  else
    lblTotalAReceberResumoHome.Caption := FormatFloat('R$ #,##0.00', dmDados.qryAux.FieldByName('TOTAL').AsCurrency);

  // ATUALIZAR ESTOQUE CRÍTICO
  if dmDados.qryAux.Active then dmDados.qryAux.Close;
  dmDados.qryAux.SQL.Clear;
  dmDados.qryAux.SQL.Add('SELECT COUNT(*) AS QTD');
  dmDados.qryAux.SQL.Add('FROM PRODUTOS');
  dmDados.qryAux.SQL.Add('WHERE ESTOQUE <= ESTOQUE_MINIMO');
  
  dmDados.qryAux.Open;

  lblTotalEstoqueCriticoResumoHome.Caption := dmDados.qryAux.FieldByName('QTD').AsString + ' Produtos';

  // Muda a cor do painel para alertar visualmente
  if dmDados.qryAux.FieldByName('QTD').AsInteger > 0 then
    pnlEstoque.Color := $008080FF // Vermelho Claro (Alerta)
  else
    pnlEstoque.Color := clBtnFace; // Cor Padrão (Tudo OK)

  // ATUALIZA TAREFAS / LEMBRETE
  chklstTarefasLembretesHome.Clear;
  dmDados.qryAux.Close;
  dmDados.qryAux.SQL.Text := 'SELECT ID, DESCRICAO, CONCLUIDO FROM TAREFAS_USUARIO ' +
                             'WHERE ID_USUARIO = :User AND CONCLUIDO = "N" ORDER BY ID DESC';
  dmDados.qryAux.Parameters.ParamByName('User').Value := dmDados.iIDUsuarioLogado;
  dmDados.qryAux.Open;

  while not dmDados.qryAux.Eof do
  begin
    // Adiciona o texto e guarda o ID da tarefa dentro do "Object" do item
    chklstTarefasLembretesHome.AddItem(dmDados.qryAux.FieldByName('DESCRICAO').AsString,
                       TObject(dmDados.qryAux.FieldByName('ID').AsInteger));
    dmDados.qryAux.Next;
  end;
end;

procedure TfrmPrincipal.btnAcessoRapidoClientesClick(Sender: TObject);
begin
  pcCadastros.ActivePage := tsClientesCadastros;
end;

procedure TfrmPrincipal.btnAcessoRapidoEstoqueClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsEstoque;
end;

procedure TfrmPrincipal.btnAcessoRapidoNovaVendaClick(Sender: TObject);
begin
  pcVendas.ActivePage := tsNovaVenda;
end;

procedure TfrmPrincipal.btnAcessoRapidoProdutosClick(Sender: TObject);
begin
  pcCadastros.ActivePage := tsProdutosCadastros;
end;

procedure TfrmPrincipal.btnAcessoRapidoRelatoriosClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsRelatorios;
end;

procedure TfrmPrincipal.btnAdicionarEtiquetasEstoqueClick(Sender: TObject);
begin
  // Verifica se tem produto selecionado na esquerda
  if dmDados.qryProdutos.IsEmpty then Exit;

  // Verifica se esse produto JÁ ESTÁ na fila
  if dsEtiquetas.Locate('ID_PRODUTO', dmDados.qryProdutos.FieldByName('ID').AsInteger, []) then
  begin
    // SE JÁ EXISTE: Entra em modo de edição e soma +1 na cópia
    dsEtiquetas.Edit;
    dsEtiquetas.FieldByName('COPIAS').AsInteger := dsEtiquetas.FieldByName('COPIAS').AsInteger + 1;
    dsEtiquetas.Post;
  end
  else
  begin
    // SE NÃO EXISTE: Cria uma nova linha (Append)
    dsEtiquetas.Append;
    // Copia os dados do Banco para a Memória
    dsEtiquetas.FieldByName('ID_PRODUTO').AsInteger := dmDados.qryProdutos.FieldByName('ID').AsInteger;
    dsEtiquetas.FieldByName('NOME').AsString        := dmDados.qryProdutos.FieldByName('NOME').AsString;
    dsEtiquetas.FieldByName('PRECO').AsCurrency     := dmDados.qryProdutos.FieldByName('PRECO_VENDA').AsCurrency;
    // Começa com 1 etiqueta
    dsEtiquetas.FieldByName('COPIAS').AsInteger     := 1;
    dsEtiquetas.Post;
  end;
end;

procedure TfrmPrincipal.btnAdicionarProdutoVendasClick(Sender: TObject);
var
  vTotalItem: Currency;
  vQtd: Double;
  vUnit: Currency;
begin
  if dmDados.qryProdutos.IsEmpty then
  begin
    ShowMessage('Pesquise um produto primeiro!');
    dbedtCodBarrasProdutoVendas.SetFocus;
    Exit;
  end;

  vQtd := StrToFloatDef(dbedtQuantidadeProdutoVendas.Text, 1);
  vUnit := dmDados.qryProdutos.FieldByName('PRECO_VENDA').AsCurrency;

  vUnit := vUnit - StrToFloatDef(dbedtDescontoUnitarioVendas.Text, 0);

  vTotalItem := vQtd * vUnit;

  dsItensVenda.Append;
  dsItensVenda.FieldByName('ID_PRODUTO').AsInteger  := dmDados.qryProdutos.FieldByName('ID').AsInteger;
  dsItensVenda.FieldByName('DESCRICAO').AsString    := dmDados.qryProdutos.FieldByName('NOME').AsString;
  dsItensVenda.FieldByName('QUANTIDADE').AsFloat    := vQtd;
//  dsItensVenda.FieldByName('DESCONTO').AsCurrency := vDesconto;
  dsItensVenda.FieldByName('VALOR_UNIT').AsCurrency := vUnit;
  dsItensVenda.FieldByName('VALOR_TOTAL').AsCurrency:= vTotalItem;
  dsItensVenda.Post;

  CalcularTotalVenda;

  dbedtCodBarrasProdutoVendas.Clear;
  dbedtQuantidadeProdutoVendas.Text := '1';
//  lblDescricaoProduto.Caption := '...';
  lblPrecoVendas.Caption := '0,00';
  dbimgFotoProdutoVendas.Picture := nil;
  dmDados.qryProdutos.Close;

  dbedtCodBarrasProdutoVendas.SetFocus;
end;

procedure TfrmPrincipal.btnAdicionarTarefasLembreteHomeClick(Sender: TObject);
begin
  if Trim(dbedtTarefasLembretesHome.Text) = '' then Exit;

  dmDados.ExecSQL('INSERT INTO TAREFAS_USUARIO (ID_USUARIO, DESCRICAO, DATA_CRIACAO) VALUES (' +
                  IntToStr(dmDados.iIDUsuarioLogado) + ', ' +
                  QuotedStr(dbedtTarefasLembretesHome.Text) + ', NOW())');

  dbedtTarefasLembretesHome.Clear;
end;

procedure TfrmPrincipal.btnAtualizarArvoreClick(Sender: TObject);
var
  NoRaiz, NoFilho: TTreeNode;
  i: Integer;
  NomeAba: String;
begin
  tvPermissoes.Items.Clear;

  NoRaiz := tvPermissoes.Items.Add(nil, 'WOLF ERP (Sistema Completo)');

  for i := 0 to PageControl1.PageCount - 1 do
  begin
    if (PageControl1.Pages[i].Name <> 'tsLogin') and
       (PageControl1.Pages[i].Name <> 'tsDesenvolvedor') then
    begin
      NomeAba := PageControl1.Pages[i].Caption;

      NoFilho := tvPermissoes.Items.AddChild(NoRaiz, NomeAba);

      tvPermissoes.Items.AddChild(NoFilho, 'Comp. Name: ' + PageControl1.Pages[i].Name);
      tvPermissoes.Items.AddChild(NoFilho, 'Index: ' + IntToStr(i));
    end;
  end;

  tvPermissoes.FullExpand;
end;

procedure TfrmPrincipal.btnAtualizarListaClick(Sender: TObject);
begin
  dmDados.qrySessoes.Close;
  dmDados.qrySessoes.SQL.Text := 'SELECT * FROM SESSOES ORDER BY DATA_LOGIN DESC';
  dmDados.qrySessoes.Open;
end;

procedure TfrmPrincipal.btnBuscaProdutoClick(Sender: TObject);
begin
  dmDados.qryProdutos.Close;
  dmDados.qryProdutos.SQL.Clear;
  dmDados.qryProdutos.SQL.Add('SELECT * FROM PRODUTOS');
  dmDados.qryProdutos.SQL.Add('WHERE (NOME LIKE :Busca) OR (COD_BARRAS = :Codigo)');
  dmDados.qryProdutos.Parameters.ParamByName('Buscar').Value := '%' + edtBusca.Text + '%';
  dmDados.qryProdutos.Parameters.ParamByName('Codigo').Value := edtBusca.Text;
  dmDados.qryProdutos.Open;
end;

procedure TfrmPrincipal.btnbuscarClick(Sender: TObject);
begin
  dmDados.qrySessoes.Close;
  dmDados.qrySessoes.SQL.Clear;

  dmDados.qrySessoes.SQL.Add('SELECT S.*, U.NIVEL_ACESSO, U.NOME_COMPLETO ');
  dmDados.qrySessoes.SQL.Add('FROM SESSOES S ');
  dmDados.qrySessoes.SQL.Add('INNER JOIN USUARIOS U ON S.ID_USUARIO = U.ID ');
  dmDados.qrySessoes.SQL.Add('WHERE 1=1 ');

  case rgFiltroPorUsuario.ItemIndex of
    1: dmDados.qrySessoes.SQL.Add('AND U.NIVEL_ACESSO = "VENDEDOR"');
    2: dmDados.qrySessoes.SQL.Add('AND U.NIVEL_ACESSO = "ESTOQUISTA"');
  end;

  // Filtro por Texto (Nome ou IP)
  if Trim(edtBuscaSessao.Text) <> '' then
  begin
    dmDados.qrySessoes.SQL.Add('AND (U.NOME_COMPLETO LIKE :Busca OR S.IP LIKE :Busca)');
    dmDados.qrySessoes.Parameters.ParamByName('Buscar').Value := '%' + edtBuscaSessao.Text + '%';
  end;

  dmDados.qrySessoes.Open;
end;

//begin
//  dmDados.qrySessoes.Close;
//  dmDados.qrySessoes.SQL.Clear;
//
//  dmDados.qrySessoes.SQL.Add('SELECT S.*, U.LOGIN, U.NIVEL_ACESSO, U.NOME_COMPLETO ');
//  dmDados.qrySessoes.SQL.Add('FROM SESSOES S ');
//  dmDados.qrySessoes.SQL.Add('INNER JOIN USUARIOS U ON S.ID_USUARIO = U.ID ');
//
//  case rgFiltroPorUsuario.ItemIndex of
//    0:
//    begin
//      dmDados.qrySessoes.SQL.Add('WHERE U.NIVEL_ACESSO = ''VENDEDOR''');
//    end;
//
//    1:
//    begin
//      dmDados.qrySessoes.SQL.Add('WHERE U.NIVEL_ACESSO = ''ESTOQUISTA''');
//    end;
//
//    2:
//    begin
//      // Não adiciona WHERE nenhum, para trazer tudo.
//    end;
//  end;
//
//  try
//    dmDados.qrySessoes.Open;
//  except
//    on E: Exception do
//      ShowMessage('Erro ao filtrar: ' + E.Message);
//  end;
//end;

procedure TfrmPrincipal.btnBuscarDevLogsClick(Sender: TObject);
begin
  with dmDados.qryDevLogs do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT * FROM LOGS_SISTEMA WHERE DATA_HORA BETWEEN :D1 AND :D2');

    if cbxFiltroAcao.Text <> 'TODOS' then
      SQL.Add('AND ACAO = ' + QuotedStr(cbxFiltroAcao.Text));

    if Trim(edtBuscaDevLogs.Text) <> '' then
      SQL.Add('AND (DESCRICAO LIKE :Busca OR LOGIN_USUARIO LIKE :Busca)');

    SQL.Add('ORDER BY DATA_HORA DESC');

    Parameters.ParamByName('D1').Value := Int(dtpInicioDevLogs.Date);
    Parameters.ParamByName('D2').Value := Int(dtpFimDevLogs.Date) + 0.99999;

    if Trim(edtBuscaDevLogs.Text) <> '' then
      Parameters.ParamByName('Busca').Value := '%' + edtBuscaDevLogs.Text + '%';

    Open;
  end;
end;

procedure TfrmPrincipal.btnBuscarProdutoCadastrosClick(Sender: TObject);
begin
  if Trim(edtBusca.Text) = '' then
  begin
    ShowMessage('Digite o código do produto para pesquisar.');
    edtProdutoCadastros.SetFocus;
    Exit;
  end;

  dmDados.qryProdutos.Close;
  dmDados.qryProdutos.SQL.Clear;

  // Ajuste "ID" e "CODIGO_BARRAS" conforme o nome exato no seu banco
  dmDados.qryProdutos.SQL.Add('SELECT * FROM PRODUTOS');
  dmDados.qryProdutos.SQL.Add('WHERE ID = :Codigo OR CODIGO_BARRAS = :Codigo');

  // Passa o parâmetro
  dmDados.qryProdutos.Parameters.ParamByName('Codigo').Value := edtBusca.Text;

  //Abre
  dmDados.qryProdutos.Open;

  //Verifica se achou
  if dmDados.qryProdutos.IsEmpty then
  begin
    ShowMessage('Produto não encontrado!');
    // Opcional: Reabrir a lista completa se não achar
    // dmDados.qryProdutos.Close;
    // dmDados.qryProdutos.SQL.Text := 'SELECT * FROM PRODUTOS';
    // dmDados.qryProdutos.Open;
  end;
  // NOTA: Se achou, os campos laterais (Descrição, Preço, etc.)
  // já estarão mostrando os dados corretos automaticamente.
end;

procedure TfrmPrincipal.btnBuscarProdutoEtiquetasEstoqueClick(Sender: TObject);
begin
  // 1Limpa a consulta anterior
  dmDados.qryEtiquetas.Close;
  dmDados.qryEtiquetas.SQL.Clear;

  // Monta o SQL Inteligente
  dmDados.qryEtiquetas.SQL.Add('SELECT ID, NOME, COD_BARRAS, PRECO_VENDA');
  dmDados.qryEtiquetas.SQL.Add('FROM PRODUTOS');

  // Verifica se o usuário digitou algo
  if Trim(dbedtBuscarProdutoEtiquetasEstoque.Text) <> '' then
  begin
    // Tenta converter para número: se for número, busca por ID ou Barras.
    // Se der erro (texto), busca por Nome.
    try
      StrToInt(dbedtBuscarProdutoEtiquetasEstoque.Text);
      // Se chegou aqui, é número: Busca Exata (Código ou Barras)
      dmDados.qryEtiquetas.SQL.Add('WHERE ID = :Busca OR COD_BARRAS = :Busca');
    except
      // Se deu erro, é texto: Busca Aproximada (LIKE)
      dmDados.qryEtiquetas.SQL.Add('WHERE NOME LIKE :BuscaLike');
    end;
  end
  else
  begin
    // Se vazio, traz os top 50 para não travar a tela
    dmDados.qryEtiquetas.SQL.Add('LIMIT 50');
  end;

  // Passa os Parâmetros com segurança
  if Trim(dbedtBuscarProdutoEtiquetasEstoque.Text) <> '' then
  begin
    // Verifica qual parâmetro o SQL está esperando
    if dmDados.qryEtiquetas.Parameters.FindParam('Busca') <> nil then
      dmDados.qryEtiquetas.Parameters.ParamByName('Busca').Value := dbedtBuscarProdutoEtiquetasEstoque.Text
    else
      dmDados.qryEtiquetas.Parameters.ParamByName('BuscaLike').Value := '%' + dbedtBuscarProdutoEtiquetasEstoque.Text + '%';
  end;

  // Abre a Grid
  dmDados.qryEtiquetas.Open;

  // Se achou algo, foca na grid para o usuário já poder navegar com setas
  if not dmDados.qryEtiquetas.IsEmpty then
    gridFiltrarProdutoEtiquetasEstoque.SetFocus;
end;

procedure TfrmPrincipal.btnBuscarUsuarioCadastrosClick(Sender: TObject);
begin
  if Trim(edtBuscarUsuario.Text) = '' then
  begin
    ShowMessage('Digite o código do Usuário para pesquisar.');
    edtBuscarUsuario.SetFocus;
    Exit;
  end;

  dmDados.qryUsuarios.Close;
  dmDados.qryUsuarios.SQL.Clear;

  dmDados.qryUsuarios.SQL.Add('SELECT * FROM USUARIOS');
  dmDados.qryUsuarios.SQL.Add('WHERE ID = :Codigo OR MATRICULA = :Matricula');

  dmDados.qryUsuarios.Parameters.ParamByName('Codigo').Value := edtBuscarUsuario.Text;

  dmDados.qryUsuarios.Open;

  if dmDados.qryUsuarios.IsEmpty then
  begin
    ShowMessage('Usuário não encontrado!');
    // dmDados.qryUsuarios.Close;
    // dmDados.qryUsuarios.SQL.Text := 'SELECT * FROM USUARIOS';
    // dmDados.qryUsuarios.Open;
  end;
end;

procedure TfrmPrincipal.btnCarregarImagemClick(Sender: TObject);
begin
  if not (dmDados.qryProdutos.State in [dsEdit, dsInsert]) then
  begin
    dmDados.qryProdutos.Edit;
  end;

  if dlgOpenValores.Execute then
  begin
    try
      dbimgValores.Picture.LoadFromFile(dlgOpenValores.FileName);
    except
      ShowMessage('Erro ao carregar a imagem. Verifique se é um arquivo válido (JPG, BMP ou PNG).');
    end;
  end;
end;

procedure TfrmPrincipal.btnConfiguracaoClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsUtilitarios;
end;

procedure TfrmPrincipal.btnConfirmarMovimentoEstoqueClick(Sender: TObject);
var
  SaldoFinalCalculado: Double;
  NovoCusto: Currency;
  IDProduto: String;
begin
  IDProduto := edtBuscaProdutoEstoque.Text; // O ID que você pegou na pesquisa
  if IDProduto = '' then
  begin
    ShowMessage('Selecione um produto primeiro!');
    Exit;
  end;

  if StrToFloatDef(edtQuantidadeEstoque.Text, 0) <= 0 then
  begin
    ShowMessage('A quantidade deve ser maior que zero!');
    edtQuantidadeEstoque.SetFocus;
    Exit;
  end;

  // Extrai o "Saldo Final" limpo do Label de previsão
  // O Label está escrito "Saldo Final: 95". Precisamos arrancar o texto para sobrar só o número.
  SaldoFinalCalculado := StrToFloatDef(StringReplace(lblInformacaoSaldoFinalEstoque.Caption, 'Saldo Final: ', '', [rfReplaceAll]), 0);

  // Se for ficar negativo, dá um aviso (você pode até bloquear se quiser)
  if SaldoFinalCalculado < 0 then
  begin
    if MessageDlg('Atenção: O estoque ficará NEGATIVO. Deseja realmente confirmar?', mtWarning, [mbYes, mbNo], 0) = mrNo then
      Exit;
  end;

  // Inicia a Gravação (Transação para não dar erro pela metade)
  dmDados.wolferp.BeginTrans;
  try
    // Atualiza o cadastro do Produto com o novo estoque (que o Label calculou)
    dmDados.ExecSQL('UPDATE PRODUTOS SET ESTOQUE = ' + StringReplace(FloatToStr(SaldoFinalCalculado), ',', '.', [rfReplaceAll]) +
                    ' WHERE ID = ' + IDProduto);

    // Grava o histórico na tabela de movimentação (Kardex)
    dmDados.ExecSQL('INSERT INTO MOVIMENTACAO_ESTOQUE (DATA_MOV, ID_PRODUTO, TIPO, QTD, ID_USUARIO, MOTIVO, DOCUMENTO) VALUES (' +
                    'NOW(), ' +
                    IDProduto + ', ' +
                    IntToStr(rdgOperacao.ItemIndex) + ', ' + // Guarda 0, 1, 2 ou 3 para saber o que foi
                    StringReplace(edtQuantidadeEstoque.Text, ',', '.', [rfReplaceAll]) + ', ' +
                    IntToStr(dmDados.iIDUsuarioLogado) + ', ' +
                    QuotedStr(mmoObservacoesEstoque.Text) + ', ' +
                    QuotedStr(edtDocumentoEstoque.Text) + ')');

    dmDados.wolferp.CommitTrans;
    ShowMessage('Movimento realizado com sucesso!');

    SaldoFinalCalculado := StrToFloatDef(StringReplace(lblPrevisaoFinal.Caption, 'Saldo Final: ', '', [rfReplaceAll]), 0);
  NovoCusto := StrToCurrDef(edtCustoUnitario.Text, 0);

  dmDados.wolferp.BeginTrans;
  try
    // A. Grava o histórico na tabela de movimentação
    // Dica: Se a tabela MOVIMENTACAO_ESTOQUE tiver a coluna CUSTO_UNITARIO, grave nela também!
    dmDados.ExecSQL('INSERT INTO MOVIMENTACAO_ESTOQUE (DATA_MOV, ID_PRODUTO, TIPO, QTD, ID_USUARIO, MOTIVO, DOCUMENTO) VALUES (...)');

    // B. Atualiza o cadastro do Produto
    if rdgOperacao.ItemIndex = 0 then // SE FOR ENTRADA (COMPRA)
    begin
      // Atualiza Estoque E PREÇO DE CUSTO
      dmDados.ExecSQL('UPDATE PRODUTOS SET ' +
                      'ESTOQUE = ' + StringReplace(FloatToStr(SaldoFinalCalculado), ',', '.', [rfReplaceAll]) + ', ' +
                      'PRECO_CUSTO = ' + StringReplace(FloatToStr(NovoCusto), ',', '.', [rfReplaceAll]) +
                      ' WHERE ID = ' + IDProduto);
    end
    else
    begin
      // Se for Saída/Perda, atualiza SÓ O ESTOQUE (não mexe no custo)
      dmDados.ExecSQL('UPDATE PRODUTOS SET ' +
                      'ESTOQUE = ' + StringReplace(FloatToStr(SaldoFinalCalculado), ',', '.', [rfReplaceAll]) +
                      ' WHERE ID = ' + IDProduto);
    end;

    dmDados.wolferp.CommitTrans;
    ShowMessage('Movimento realizado com sucesso!');

    // Limpa a tela para o próximo produto
    edtQuantidadeEstoque.Clear;
    edtDocumentoEstoque.Clear;
    mmoObservacoesEstoque.Clear;
    lblInformacaoSaldoAtualEstoque.Caption := FloatToStr(SaldoFinalCalculado); // O Atual agora é o Final
    AtualizarSimulacao; // Roda a simulação para zerar os labels
    edtBuscaProdutoEstoque.SetFocus; // Volta o cursor para bipar o próximo produto

  except
    on E: Exception do
    begin
      dmDados.wolferp.RollbackTrans;
      ShowMessage('Erro ao gravar movimentação: ' + E.Message);
    end;
  end;
end;

//procedure TfrmPrincipal.btnConfirmarMovClick(Sender: TObject);
//var
//  vIDProd : Integer;
//  vQtd : Double;
//  vTipo : String;
//  vSQLUpdate: String;
//begin
//  // wallace 05/12/2025
//  if Trim(edtQtdMov.Text) = '' then
//  begin
//    ShowMessage('Informe a Quantidade.');
//    Exit;
//  end;
//
//  if dmDados.qryProdutos.IsEmpty then
//  begin
//    ShowMessage('Selecione um produto na lista.');
//    Exit;
//  end;
//
//  try
//    vQtd := StrToFloat(edtQtdMov.Text);
//  except
//    ShowMessage('Quantidade inválida.');
//    Exit;
//  end;
//
//  if rgTipoMovimento.ItemIndex = 0 then
//    vTipo := 'E'
//  else
//    vTipo := 'S';
//
//  vIDProd := dmDados.qryProdutos.FieldByName('ID').AsInteger;
//
//  dmDados.wolferp.BeginTrans;
//  try
//    with dmDados.qryExec do
//    begin
//      Close;
//      SQL.Clear;
//      SQL.Add('INSERT INTO MOVIMENTO (ID_PRODUTO, DATA, TIPO, QTD, OBS)');
//      SQL.Add('VALUES (:Id, NOW, :Tipo, :Qtd, :Obs)');
//
//      Parameters.ParamByName('Id').Value   := vIDProd;
//      Parameters.ParamByName('Tipo').Value := vTipo;
//      Parameters.ParamByName('Qtd').Value  := vQtd;
//      Parameters.ParamByName('Obs').Value  := mmoMotivo.Text;
//      ExecSQL;
//    end;
//
//    if vTipo = 'E' then
//      vSQLUpdate := 'UPDATE PRODUTOS SET ESTOQUE = ESTOQUE + :Qtd WHERE ID = :Id'
//    else
//      vSQLUpdate := 'UPDATE PRODUTOS SET ESTOQUE = ESTOQUE - :Qtd WHERE ID = :Id';
//
//    with dmDados.qryExec do
//    begin
//      Close;
//      SQL.Text := vSQLUpdate;
//      Parameters.ParamByName('Qtd').Value := vQtd;
//      Parameters.ParamByName('ID').Value := vIDProd;
//      ExecSQL;
//    end;
//
//    dmDados.wolferp.CommitTrans;
//
//    ShowMessage('Estoque atualizado com sucesso!');
//
//    dmDados.qryProdutos.Close;
//    dmDados.qryprodutos.Open;
//
//    dmDados.qryProdutos.Locate('ID', vIDProd, []);
//
//    lblSaldoAtual.Caption := 'Saldo Atual: ' + dmDados.qryProdutos.FieldByName('ESTOQUE').AsString;
//
//    edtQtdMov.Clear;
//    mmoMotivo.Clear;
//
//  except
//    on E : Exception do
//    begin
//      dmDados.wolferp.RollbackTrans;
//      ShowMessage('Erro ao movimentar estoque: ' + E.Message);
//    end;
//  end;
//end;

procedure TfrmPrincipal.btnEncerrarSessaoClick(Sender: TObject);
var
  vIDSession: Integer;
begin
  vIDSession := dmDados.qrySessoes.FieldByName('ID').AsInteger;

  if vIDSession = dmDados.UsuarioLogadoID then
  begin
    ShowMessage('Você não pode derrubar a si mesmo!');
    Exit;
  end;

  if MessageDlg('Derrubar este usuário?', mtWarning, [mbYes, mbNo], 0) = mrYes then
  begin
    dmDados.ExecSQL('DELETE FROM SESSOES WHERE ID = ' + IntToStr(vIDSession));

    dmDados.qrySessoes.Close;
    dmDados.qrySessoes.Open;
  end;
end;

procedure TfrmPrincipal.btnEnviarAvisoClick(Sender: TObject);
begin
if Trim(edtMensagem.Text) = '' then Exit;

  with dmDados.qryGenerica do
  begin
    Close;
    SQL.Text := 'INSERT INTO MENSAGENS_SISTEMA (MENSAGEM, DESTINO) VALUES (:Msg, "TODOS")';
    Parameters.ParamByName('Msg').Value := edtMensagem.Text;
    ExecSQL;
  end;

  edtMensagem.Clear;
  ShowMessage('Mensagem enviada para a fila de transmissão.');
end;

procedure TfrmPrincipal.btnReceberTituloContasReceberFinanceiroClick(Sender: TObject);
var
  IDConta: Integer;
  ValorParaBaixar: Currency;
begin
  if dmDados.qryReceber.IsEmpty then Exit;

  if dmDados.qryReceber.FieldByName('STATUS').AsString = 'RECEBIDO' then
  begin
    ShowMessage('Este título já foi recebido!');
    Exit;
  end;

  // Pega o valor atualizado do Label (retirando o 'R$ ' e formatando)
  ValorParaBaixar := StrToCurrDef(StringReplace(lblValorAtualizado.Caption, 'R$ ', '', [rfReplaceAll]), 0);

  // Pergunta de Confirmação "Rica" (Mostrando o valor)
  if MessageDlg('Confirma o recebimento no valor de ' + FormatFloat('R$ #,##0.00', ValorParaBaixar) + '?', 
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    IDConta := dmDados.qryReceber.FieldByName('ID').AsInteger;

    try
      // Executa o UPDATE no banco
      dmDados.ExecSQL('UPDATE FINANCEIRO SET ' +
                      'STATUS = ''RECEBIDO'', ' +
                      'DATA_PAGAMENTO = NOW(), ' +
                      'VALOR = ' + StringReplace(FloatToStr(ValorParaBaixar), ',', '.', [rfReplaceAll]) +
                      ' WHERE ID = ' + IntToStr(IDConta));

      ShowMessage('Título baixado com sucesso!');
      
      // Atualiza a Grid

    except
      on E: Exception do ShowMessage('Erro ao baixar título: ' + E.Message);
    end;
  end;
end;

procedure TfrmPrincipal.btnRelatoriosLateralClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsCadastros;
end;

procedure TfrmPrincipal.btnCancelarClienteCadastrosClick(Sender: TObject);
begin
    if dmDados.qryClientes.State in [dsInsert, dsEdit] then
  begin
    dmDados.qryClientes.Cancel;

    ShowMessage('Operação cancelada.');
  end;
end;

procedure TfrmPrincipal.btnCancelarFornecedorClick(Sender: TObject);
begin
    if dmDados.qryFornecedor.State in [dsInsert, dsEdit] then
  begin
    dmDados.qryFornecedor.Cancel;

    ShowMessage('Operação cancelada.');
  end;
end;

procedure TfrmPrincipal.btnCancelarProdutoCadastrosClick(Sender: TObject);
begin
  if dmDados.qryProdutos.State in [dsInsert, dsEdit] then
  begin
    dmDados.qryProdutos.Cancel;

    ShowMessage('Operação cancelada.');
  end;
end;

procedure TfrmPrincipal.btnCancelarUsuarioCadastrosClick(Sender: TObject);
begin
    if dmDados.qryUsuarios.State in [dsInsert, dsEdit] then
  begin
    dmDados.qryUsuarios.Cancel;

    ShowMessage('Operação cancelada.');
  end;
end;

//procedure TfrmPrincipal.btnCancelarVendaClick(Sender: TObject);
//begin
//
//  // Exige autorização específica de cancelamento
//  if frmAutenticacao.SolicitarAutorizacao then //('AC_CANCELAR_VENDA')
//  begin
//     CancelarVendaAtual;
//  end;
//end;

//procedure TfrmPrincipal.btnExcluirUsuarioCadastrosClick(Sender: TObject);
//var
//  NomeUsuario: String;
//begin
//  if not dmDados.qryUsuarios.IsEmpty then
//  begin
//    NomeUsuario := dmDados.qryUsuarios.FieldByName('NOME').AsString;
//
//    if MessageDlg('Tem certeza que deseja excluir: ' + NomeUsuario + '?',
//       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
//    begin
//      dmDados.qryUsuarios.Delete;
//
//      ShowMessage('Usuário excluído com sucesso.');
//    end;
//  end
//  else
//  begin
//    ShowMessage('Não há usuário selecionado para excluir.');
//  end;
//
//  // --- Confirmação de usuário ---
//  // A tela abre, pede senha, verifica no banco e devolve True ou False
//  if SolicitarLiberacao('Exclusão do Usuário: ' + NomeUsuario) then
//  begin
//    // Se passou, executa a exclusão
//    try
//      dmDados.qryUsuarios.Delete;
//      ShowMessage('Usuário excluído com sucesso.');
//    except
//      on E: Exception do ShowMessage('Erro: ' + E.Message);
//    end;
//  end
//  else
//  begin
//    // Se fechou a tela ou errou a senha
//    ShowMessage('Operação cancelada.');
//  end;
//end;

procedure TfrmPrincipal.btnFecharSistemaClick(Sender: TObject);
begin
  if MessageDlg('Deseja realmente sair do sistema?',
     mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    Application.Terminate;
  end;
end;

procedure TfrmPrincipal.btnCadastrosLateralClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsVendas;
end;

procedure TfrmPrincipal.btnGerarClick(Sender: TObject);
var
  vSQL: String;
  vArquivo: String;
begin
  case cbbTipoRelatorio.ItemIndex of
    0:
    begin
      vSQL := 'SELECT DATA_EMISSAO, VALOR_TOTAL FROM VENDA_CAB ' +
              'WHERE DATA_EMISSAO BETWEEN :D1 AND :D2';
      vArquivo := 'C:\Relatorio_Vendas.csv';
    end;
    1:
    begin
      vSQL := 'SELECT DESCRICAO, ESTOQUE, PRECO_VENDA FROM PRODUTOS ORDER BY DESCRICAO';
      vArquivo := 'C:\Relatorio_Estoque.csv';
    end;
    2:
    begin
      vSQL := 'SELECT DESCRICAO, ESTOQUE FROM PRODUTOS WHERE ESTOQUE <= 5';
      vArquivo := 'C:\Relatorio_Reposicao.csv';
    end;
  end;

  with dmDados.qryRelatorios do
  begin
    Close;
    SQL.Text := vSQL;

    if cbbTipoRelatorio.ItemIndex = 0 then
    begin

      Parameters.ParamByName('D1').Value := Int(dtpRelInicio.Date);
      Parameters.ParamByName('D2').Value := Int(dtpRelFim.Date) + 0.99999;
    end;

    Open;

    if IsEmpty then
    begin
      ShowMessage('Nenhum dado encontrado!');
      Exit;
    end;
  end;

  try
    ExportarParaCSV(dmDados.qryRelatorios, vArquivo);
  except
    on E: Exception do
      ShowMessage('Erro ao exportar. Verifique se o arquivo já está aberto no Excel.');
  end;

begin
  case cbbTipoRelatorio.ItemIndex of
    0: // Vendas (Index 0)
    begin
      vSQL := 'SELECT DATA_EMISSAO, VALOR_TOTAL FROM VENDA_CAB ' +
              'WHERE DATA_EMISSAO BETWEEN :D1 AND :D2';
      vArquivo := 'C:\Relatorio_Vendas.csv';
    end;

    1:
    begin
      vSQL := 'SELECT DESCRICAO, ESTOQUE, PRECO_VENDA FROM PRODUTOS ORDER BY DESCRICAO';
      vArquivo := 'C:\Relatorio_Estoque.csv';
    end;

    2:
    begin
      vSQL := 'SELECT DESCRICAO, ESTOQUE FROM PRODUTOS WHERE ESTOQUE <= 5';
      vArquivo := 'C:\Relatorio_Reposicao.csv';
    end;
  end;

  with dmDados.qryRelatorios do
  begin
    Close;
    SQL.Text := vSQL;

    if cbbTipoRelatorio.ItemIndex = 0 then
    begin
      Parameters.ParamByName('D1').Value := Int(dtpRelInicio.Date);
      Parameters.ParamByName('D2').Value := Int(dtpRelFim.Date) + 0.99999;
    end;

    Open; // Executa no banco

    if IsEmpty then
    begin
      ShowMessage('Nenhum dado encontrado para o período/filtro selecionado.');
      Exit;
    end;
  end;

  try
    ExportarParaCSV(dmDados.qryRelatorios, vArquivo);
  except
    on E: Exception do
      ShowMessage('Erro ao criar arquivo. Verifique se ele já está aberto no Excel.');
  end;
end;
end;

procedure TfrmPrincipal.btnGerarPedidoCompraEstoqueClick(Sender: TObject);
var
  IdPedidoGerado: Integer;
  IdFornecedorPadrao: Integer;
begin
  if dmDados.qryReposicao.IsEmpty then
  begin
    ShowMessage('Não há produtos precisando de reposição no momento.');
    Exit;
  end;

  // Pergunta se o usuário quer gerar para um fornecedor específico ou "DIVERSOS"
  // Se não selecionou fornecedor no combo, pede para selecionar ou usa um genérico
  if dblkcbbFiltroFornecedorReposicaoEstoque.KeyValue = Null then
  begin
    if MessageDlg('Nenhum fornecedor selecionado. Deseja gerar um Pedido Genérico?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      Exit;
    IdFornecedorPadrao := 0;
  end
  else
    IdFornecedorPadrao := dblkcbbFiltroFornecedorReposicaoEstoque.KeyValue;

  dmDados.wolferp.BeginTrans;
  try
    // Cria o Cabeçalho do Pedido - PEDIDOS_COMPRA
    dmDados.ExecSQL('INSERT INTO PEDIDOS_COMPRA (DATA_EMISSAO, ID_FORNECEDOR, STATUS, TOTAL_PREVISTO) VALUES (' +
                    'NOW(), ' +
                    IntToStr(IdFornecedorPadrao) + ', ' +
                    QuotedStr('ABERTO') + ', 0)');
    
    // Pega o ID gerado
    dmDados.qryAux.Close;
    dmDados.qryAux.SQL.Text := 'SELECT LAST_INSERT_ID() as ID';
    dmDados.qryAux.Open;
    IdPedidoGerado := dmDados.qryAux.FieldByName('ID').AsInteger;

    // Percorre a Grid de Reposição e insere nos Itens
    dmDados.qryReposicao.First; // Volta para o primeiro item da grid
    while not dmDados.qryReposicao.Eof do
    begin
      // Só insere se a sugestão for maior que zero
      if dmDados.qryReposicao.FieldByName('SUGESTAO_COMPRA').AsFloat > 0 then
      begin
        dmDados.ExecSQL('INSERT INTO ITENS_PEDIDOS_COMPRA (ID_PEDIDO, ID_PRODUTO, QTD_SUGERIDA, CUSTO_ATUAL) VALUES (' +
                        IntToStr(IdPedidoGerado) + ', ' +
                        dmDados.qryReposicao.FieldByName('ID').AsString + ', ' +
                        StringReplace(dmDados.qryReposicao.FieldByName('SUGESTAO_COMPRA').AsString, ',', '.', [rfReplaceAll]) + ', ' +
                        StringReplace(dmDados.qryReposicao.FieldByName('PRECO_CUSTO').AsString, ',', '.', [rfReplaceAll]) + ')');
      end;
      
      dmDados.qryReposicao.Next;
    end;

    dmDados.wolferp.CommitTrans;

    ShowMessage('Ordem de Compra Nº ' + IntToStr(IdPedidoGerado) + ' gerada com sucesso!' + #13 +
                'Acesse o menu "Compras" para imprimir ou enviar ao fornecedor.');

  except
    on E: Exception do
    begin
      dmDados.wolferp.RollbackTrans;
      ShowMessage('Erro ao gerar pedido: ' + E.Message);
    end;
  end;
end;

procedure TfrmPrincipal.btnGerarReciboContasReceberFinanceiroClick(Sender: TObject);
begin
  if dmDados.qryReceber.IsEmpty then Exit;

  if dmDados.qryReceber.FieldByName('STATUS').AsString <> 'RECEBIDO' then
  begin
    ShowMessage('Você só pode gerar recibo de títulos já recebidos.');
    Exit;
  end;

  // Filtra a Query do Relatório para o ID selecionado
  dmDados.qryRelatorios.Close;
  dmDados.qryRelatorios.SQL.Text := 'SELECT * FROM FINANCEIRO WHERE ID = ' + dmDados.qryReceber.FieldByName('ID').AsString;
  dmDados.qryRelatorios.Open;

  // Chama o gerador
  //frxRelatorioRecibo.ShowReport;
end;

procedure TfrmPrincipal.btnHomeLateralClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsHome;
end;

procedure TfrmPrincipal.btnImprimirEtiquetasEstoqueClick(Sender: TObject);
//var
//  i: Integer;
//begin
//  if dsEtiquetas.IsEmpty then
//  begin
//    ShowMessage('A fila de impressão está vazia.');
//    Exit;
//  end;
//
//  // Verifica qual modelo o usuário escolheu no RadioGroup
//  case rdgModeloImpressao.ItemIndex of
//    0: // TÉRMICA (Zebra/Argox)
//       begin
//         dsEtiquetas.First;
//         while not dsEtiquetas.Eof do
//         begin
//           // Loop das cópias: Se o usuário pediu 10, imprime 10 vezes
//           for i := 1 to dsEtiquetas.FieldByName('COPIAS').AsInteger do
//           begin
//              // Chama sua função de imprimir PPLA/PPLB passando os dados atuais
//              ImprimirEtiquetaTermica(
//                dsEtiquetas.FieldByName('NOME').AsString,
//                dsEtiquetas.FieldByName('PRECO').AsCurrency
//              );
//           end;
//           dsEtiquetas.Next;
//         end;
//       end;
//
//    1: // FOLHA A4 (Relatório FastReport/Fortes)
//       begin
//         // No relatório, você liga o MasterData direto no 'dsEtiquetas'
//         // O relatório vai imprimir exatamente o que está na grid
//         relatorioEtiquetas.ShowReport;
//       end;
//  end;

begin
  // Se tem algo para imprimir
  if dsEtiquetas.IsEmpty then
  begin
    ShowMessage('A lista de etiquetas está vazia. Adicione produtos primeiro.');
    Exit;
  end;

  // Qual impressora usar
//  case rdgModeloImpressaoEtiquetasEstoque.ItemIndex of
//    0: ImprimirTermicaRAW;   // Chama a rotina "Difícil" (Código direto pra impressora)
//    1: ImprimirFolhaA4;      // Chama a rotina "Fácil" (Gerador de Relatório)
//    2: ImprimirGondola;      // Chama outro relatório específico
//  end;
end;

procedure TfrmPrincipal.btnLimparEtiquetasEstoqueClick(Sender: TObject);
begin
  dsEtiquetas.EmptyDataSet; // Comando especial do ClientDataSet que limpa tudo
end;

procedure TfrmPrincipal.btnVendaHomeClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsVendas;
end;

//procedure TfrmPrincipal.btnNovoClick(Sender: TObject);
//begin
//  FAcao := 'INSERIR';
//  edtID.Clear;
//  edtNome.Clear;
//  edtPreco.Text := '0,00';
//  chkRetornavel.Checked := False;
//
//  edtNome.SetFocus;
//end;

//procedure TfrmPrincipal.btnNovoFornClick(Sender: TObject);
//begin
//  dmDados.qryFornecedor.Insert;
//
//  if dbedtNomeFantasiaFornecedor.CanFocus then
//    dbedtNomeFantasiaFornecedor.SetFocus;
//end;

procedure TfrmPrincipal.btnNovoFornecedorClick(Sender: TObject);
begin
  dmDados.qryFornecedor.Insert;

  if dbedtRZSocialFornecedor.CanFocus then
    dbedtRZSocialFornecedor.SetFocus;

  // dmDados.qryFornecedor.FieldByName('ESTOQUE').AsFloat := 0;
end;

procedure TfrmPrincipal.btnNovoUsuarioCadastrosClick(Sender: TObject);
begin
  dmDados.qryProdutos.Insert;

  if edtBuscarUsuario.CanFocus then
    edtBuscarUsuario.SetFocus;
end;

procedure TfrmPrincipal.btnPesquisarFiltrosPesquisaVendasClick(Sender: TObject);
var
  SQL: String;
begin
  SQL := 'SELECT V.ID, V.DATA_VENDA, V.TOTAL_VENDA, C.NOME as CLIENTE, ' +
         'CASE WHEN V.STATUS = "C" THEN "CANCELADA" ELSE "NORMAL" END AS STATUS_DESC ' +
         'FROM VENDAS V ' +
         'LEFT JOIN CLIENTES C ON C.ID = V.ID_CLIENTE ' +
         'WHERE V.DATA_VENDA BETWEEN :DtIni AND :DtFim ';

  if Trim(edtNPedidoFiltrosPesquisaVendas.Text) <> '' then
  begin
    SQL := SQL + ' AND (V.ID = ' + QuotedStr(edtNPedidoFiltrosPesquisaVendas.Text) +
                 ' OR C.NOME LIKE ' + QuotedStr('%' + edtNPedidoFiltrosPesquisaVendas.Text + '%') + ')';
  end;

  if not chkVendaCanceladaFiltrosPesquisaVendas.Checked then
    SQL := SQL + ' AND V.STATUS <> "C" ';

  SQL := SQL + ' ORDER BY V.ID DESC';

  dmDados.qryVendas.Close;
  dmDados.qryVendas.SQL.Text := SQL;
  dmDados.qryVendas.Parameters.ParamByName('DE').Value  := Trunc(dtpDeFiltrosPesquisaVenda.Date);
//  dmDados.qryVendas.Parameters.ParamByName('DE').Value := StartOfDay(dtpDeFiltrosPesquisaVenda.Date);
  dmDados.qryVendas.Parameters.ParamByName('ATE').Value := Trunc(dtpAteFiltrosPesquisaVendas.Date) + 0.99999;
//  dmDados.qryVendas.Parameters.ParamByName('ATE').Value := EndOfDay(dtpAteFiltrosPesquisaVendas.Date);
  dmDados.qryVendas.Open;

  if dmDados.qryVendas.IsEmpty then
    dmDados.qryItemVenda.Close;
end;

procedure TfrmPrincipal.btnPesquisarOrcamentoClick(Sender: TObject);
var
  ID : Integer;
begin
  // 1. Carrega o Cabeçalho
  dmDados.qryOrcamentos.Close;
  dmDados.qryOrcamentos.SQL.Text := 'SELECT * FROM ORCAMENTOS WHERE ID = ' + IntToStr(ID);
  dmDados.qryOrcamentos.Open;

  if dmDados.qryOrcamentos.IsEmpty then Exit;

  // Preenche os campos visuais
  dbedtNumOrcamento.Text := IntToStr(ID);
  dblkcbbClienteOrcamento.KeyValue := dmDados.qryOrcamentos.FieldByName('CLIENTE').AsInteger;
  dblkcbbVendedorOrcamentosVendas.KeyValue := dmDados.qryOrcamentos.FieldByName('VENDEDOR').AsInteger;
  rgFormaPagamentoOrcamento.ItemIndex := dmDados.qryOrcamentos.FieldByName('VENDEDOR').AsInteger;
  // ... preencha vendedor, forma pagto, etc.

  // Carrega os Itens
  // Temos que pegar do banco e jogar para a Memória (ClientDataSet) para permitir edição
  dmDados.qryOrcamentosItens.Close;
  dmDados.qryOrcamentosItens.SQL.Text := 'SELECT * FROM ORCAMENTOS_ITENS WHERE ID_ORCAMENTO = ' + IntToStr(ID);
  dmDados.qryOrcamentosItens.Open;

  dsOrcamentosItens.EmptyDataSet; // Limpa a tela

  while not dmDados.qryOrcamentosItens.Eof do
  begin
    dsOrcamentosItens.Append;
    dsOrcamentosItens.FieldByName('ID_PRODUTO').AsInteger := dmDados.qryOrcamentosItens.FieldByName('ID_PRODUTO').AsInteger;
    // Dica: Faça um Join na query acima para já trazer o NOME do produto
    dsOrcamentosItens.FieldByName('PRODUTO').AsString     := dmDados.qryOrcamentosItens.FieldByName('DESCRICAO_PRODUTO').AsString;
    dsOrcamentosItens.FieldByName('QTD').AsFloat          := dmDados.qryOrcamentosItens.FieldByName('QTD').AsFloat;
    dsOrcamentosItens.FieldByName('VALOR_UNIT').AsCurrency:= dmDados.qryOrcamentosItens.FieldByName('UNITARIO').AsCurrency;
    dsOrcamentosItens.FieldByName('VALOR_TOTAL').AsCurrency:= dmDados.qryOrcamentosItens.FieldByName('TOTAL').AsCurrency;
    dsOrcamentosItens.Post;

    dmDados.qryOrcamentosItens.Next;
  end;

//  CalcularTotalOrcamento;
  // Atualiza o Label Total
end;

//procedure TfrmPrincipal.btnNovoUsuarioClick(Sender: TObject);
//begin
//  dmDados.qryUsuarios.Insert;
//
//  dmDados.qryUsuarios.FieldByName('ATIVO').AsString := 'S';
//  dmDados.qryUsuarios.FieldByName('ADM').AsString   := 'N';
//
//  edtNovaSenhaUsu.Clear;
//  dbedtLoginUsu.SetFocus;
//end;

procedure TfrmPrincipal.btnFiltrarReposicaoClick(Sender: TObject);
begin
  dmDados.qryReposicao.Close;
  // Se quiser filtrar por fornecedor específico:
  if dblkcbbFiltroFornecedorReposicaoEstoque.Text <> '' then
  begin
     dmDados.qryReposicao.SQL.Text := '...SQL ACIMA... AND P.ID_FORNECEDOR = ' + IntToStr(dblkcbbFiltroFornecedorReposicaoEstoque.KeyValue);
  end;
  dmDados.qryReposicao.Open;
end;

procedure TfrmPrincipal.btnFinalizarVendaClick(Sender: TObject);
var
  IDVendaGerada: Integer;
  TotalVenda: Currency;
begin
  TotalVenda := StrToCurr(lblTotalaPagarVendas.Caption);

  // Se a venda for maior que R$ 5.000, pede gerente
  if TotalVenda > 5000 then
  begin
    // Parâmetros: ("Texto", EhPerigoso=False, "Coluna_No_Banco")
    // False aqui deixa o texto Azul, indicando que é um procedimento padrão, não um erro.
    if not frmAutenticacao.Solicitar('AUTORIZAR VENDA DE ALTO VALOR', False, 'AC_LIBERAR_VENDA_ALTA') then
    begin
      ShowMessage('Venda não autorizada.');
      Exit; // Para tudo se não autorizar
    end;
  end;

  //GravarVendaNoBanco;

  if dsItensVenda.IsEmpty then
  begin
    ShowMessage('A cesta está vazia!');
    Exit;
  end;
  if dblkcbbDadosVendaClienteVendas.Text = '' then
  begin
    ShowMessage('Selecione um Cliente!');
    grpDadosVendasVendas.SetFocus;
    Exit;
  end;

  try
    dmDados.qryVendas.Close;
    dmDados.qryVendas.SQL.Text := 'SELECT * FROM VENDAS WHERE 1=2';
    dmDados.qryVendas.Open;
    
    dmDados.qryVendas.Append;
    dmDados.qryVendas.FieldByName('DATA_VENDA').AsDateTime := Now;
    dmDados.qryVendas.FieldByName('ID_CLIENTE').AsInteger  := dblkcbbDadosVendaClienteVendas.KeyValue;
    dmDados.qryVendas.FieldByName('ID_VENDEDOR').AsInteger := dmDados.iIDUsuarioLogado;
    dmDados.qryVendas.FieldByName('VALOR_TOTAL').AsCurrency := StrToCurr(StringReplace(lblTotalaPagarVendas.Caption, 'R$ ', '', [rfReplaceAll]));
    dmDados.qryVendas.FieldByName('FORMA_PAGTO').AsInteger := rgFormadePagamentoVendas.ItemIndex;
    dmDados.qryVendas.Post;

    IDVendaGerada := dmDados.qryVendas.FieldByName('ID').AsInteger; 

    dsItensVenda.First;
    while not dsItensVenda.Eof do
    begin
      dmDados.ExecSQL('INSERT INTO VENDAS_ITENS (ID_VENDA, ID_PRODUTO, QUANTIDADE, VALOR_UNITARIO, SUBTOTAL) VALUES (' +
                      IntToStr(IDVendaGerada) + ', ' +
                      dsItensVenda.FieldByName('ID_PRODUTO').AsString + ', ' +
                      StringReplace(dsItensVenda.FieldByName('QUANTIDADE').AsString, ',', '.', [rfReplaceAll]) + ', ' +
                      StringReplace(dsItensVenda.FieldByName('VALOR_UNIT').AsString, ',', '.', [rfReplaceAll]) + ', ' +
                      StringReplace(dsItensVenda.FieldByName('VALOR_TOTAL').AsString, ',', '.', [rfReplaceAll]) + ')');

      dmDados.ExecSQL('UPDATE PRODUTOS SET ESTOQUE = ESTOQUE - ' +
                      StringReplace(dsItensVenda.FieldByName('QTD').AsString, ',', '.', [rfReplaceAll]) +
                      ' WHERE ID = ' + dsItensVenda.FieldByName('ID_PRODUTO').AsString);

      dsItensVenda.Next;
    end;

    ShowMessage('Venda Nº ' + IntToStr(IDVendaGerada) + ' finalizada com sucesso!');

    dsItensVenda.EmptyDataSet;
    dblkcbbDadosVendaClienteVendas.KeyValue := Null;
    lblTotalaPagarVendas.Caption := 'R$ 0,00';
    rgFormadePagamentoVendas.ItemIndex := 0;
    dbedtCodBarrasProdutoVendas.SetFocus;

  except
    on E: Exception do
      ShowMessage('Erro ao finalizar venda: ' + E.Message);
  end;
end;

procedure TfrmPrincipal.btnFinanceiroLateralClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsFinanceiro;
end;

procedure TfrmPrincipal.btnVendasLateralClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsRelatorios;
end;

procedure TfrmPrincipal.btnRemoveEtiquetasEstoqueClick(Sender: TObject);
begin
  if not dsEtiquetas.IsEmpty then
  begin
    dsEtiquetas.Delete;
  end;
end;

procedure TfrmPrincipal.btnRemoverClick(Sender: TObject);
begin
  if not (dmDados.qryProdutos.State in [dsEdit, dsInsert]) then
    dmDados.qryProdutos.Edit;

  dbimgValores.Picture := nil;
end;

procedure TfrmPrincipal.btnSairSistemaClick(Sender: TObject);
begin
  if MessageDlg('Deseja realmente sair do sistema?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    Hide;

    if frmLogin.ShowModal = mrOk then
    begin
       Show;
       FormShow(Self);
    end
    else
    begin
       Application.Terminate;
    end;
  end;

//  if MessageDlg('Deseja realmente sair do sistema?',
//     mtConfirmation, [mbYes, mbNo], 0) = mrYes then
//  begin
//    Application.Terminate;
//  end;
end;

//procedure TfrmPrincipal.btnSalvarClick(Sender: TObject);
//var
//  vSQL: string;
//begin
//begin
//  if Trim(edtNome.Text) = '' then
//  begin
//    ShowMessage('Digite o nome do produto.');
//    edtNome.SetFocus;
//    Exit;
//  end;
//
//  if FAcao = 'INSERIR' then
//    vSQL := 'INSERT INTO PRODUTOS (DESCRICAO, PRECO_VENDA, RETORNAVEL, ESTOQUE) VALUES (:Desc, :Preco, :Ret, 0)'
//  else if FAcao = 'EDITAR' then
//    vSQL := 'UPDATE PRODUTOS SET DESCRICAO = :Desc, PRECO_VENDA = :Preco, RETORNAVEL = :Ret WHERE ID = :ID'
//  else
//  begin
//    ShowMessage('Clique em Novo ou selecione um registro primeiro.');
//    Exit;
//  end;
//
//  with dmDados.qryExec do
//  begin
//    Close;
//    SQL.Text := vSQL;
//
//    Parameters.ParamByName('Desc').Value := edtNome.Text;
//
//    try
//      Parameters.ParamByName('Preco').Value := StrToFloat(edtPreco.Text);
//    except
//      ShowMessage('Preço inválido.');
//      Exit;
//    end;
//
//    if chkRetornavel.Checked then
//      Parameters.ParamByName('Ret').Value := 'S'
//    else
//      Parameters.ParamByName('Ret').Value := 'N';
//
//    if FAcao = 'EDITAR' then
//      Parameters.ParamByName('ID').Value := StrToInt(edtID.Text);
//
//    ExecSQL;
//  end;
//
//  ShowMessage('Salvo com sucesso!');
//
//  dmDados.qryProdutos.Close;
//  dmDados.qryProdutos.Open;
//
//  btnNovoClick(nil);
//end;
//end;


procedure TfrmPrincipal.btnSalvarFornClick(Sender: TObject);
begin
  if Trim(dmDados.qryFornecedor.FieldByName('NOME_FANTASIA').AsString) = '' then
  begin
    ShowMessage('O Nome Fantasia é obrigatório.');
    if dbedtNomeFantasiaFornecedor.CanFocus then dbedtNomeFantasiaFornecedor.SetFocus;
    Exit;
  end;

  if Trim(dmDados.qryFornecedor.FieldByName('CNPJ').AsString) = '' then
  begin
    ShowMessage('Por favor, informe o CNPJ do fornecedor.');
    if dbedtCNPJ.CanFocus then dbedtCNPJ.SetFocus;
    Exit;
  end;

  try
    dmDados.qryFornecedor.Post;
    ShowMessage('Fornecedor salvo com sucesso!');
  except
    on E: Exception do
      ShowMessage('Erro ao salvar: ' + E.Message);
  end;
end;

procedure TfrmPrincipal.btnSalvarFornecedorClick(Sender: TObject);
begin
  if dmDados.qryFornecedor.State in [dsInsert, dsEdit] then
  // Verifica se o usuário está realmente inserindo ou editando algo
  begin
    try
      dmDados.qryFornecedor.Post;
      ShowMessage('Fornecedor salvo com sucesso!');

      edtFornecedorBuscar.Clear;
    except
      on E: Exception do
      begin
        ShowMessage('Erro ao gravar: ' + E.Message);
      end;
    end;
  end;
end;

procedure TfrmPrincipal.btnEstoqueLateralClick(Sender: TObject);
begin
  PageControl1.ActivePage := tsEstoque;
end;

procedure TfrmPrincipal.CarregarArvoreSistema;
begin
  CarregarArvoreSistema
end;

procedure TfrmPrincipal.cbbTipoRelatorioChange(Sender: TObject);
var
  PrecisaData: Boolean;
begin
  PrecisaData := (cbbTipoRelatorio.ItemIndex = 0);

  dtpRelInicio.Enabled := PrecisaData;
  dtpRelFim.Enabled    := PrecisaData;

  if not PrecisaData then
  begin
    dtpRelInicio.Color := clBtnFace;
    dtpRelFim.Color    := clBtnFace;
  end
  else
  begin
    dtpRelInicio.Color := clWindow;
    dtpRelFim.Color    := clWindow;
  end;
end;

procedure TfrmPrincipal.cbxParcelasChange(Sender: TObject);
begin
  if cbxParcelas.ItemIndex = -1 then
  begin
    lblValorParcelasVendas.Caption := '...';
    Exit;
  end;

  lblValorParcelasVendas.Caption := cbxParcelas.Text;

  lblValorParcelasVendas.Font.Color := clBlue;
  lblValorParcelasVendas.Font.Style := [fsBold];
end;

procedure TfrmPrincipal.chklstTarefasLembretesHomeClick(Sender: TObject);
var
  Idx, IdTarefa: Integer;
begin
  Idx := chklstTarefasLembretesHome.ItemIndex;
  if Idx < 0 then Exit;

  // Se o usuário marcou como feito
  if chklstTarefasLembretesHome.Checked[Idx] then
  begin
    // Recupera o ID que guardamos no carregamento
    IdTarefa := Integer(chklstTarefasLembretesHome.Items.Objects[Idx]);

    // Atualiza no banco como concluído
    dmDados.ExecSQL('UPDATE TAREFAS_USUARIO SET CONCLUIDO = "S" WHERE ID = ' + IntToStr(IdTarefa));

    // Remove da lista visualmente
    chklstTarefasLembretesHome.Items.Delete(Idx);
  end;
end;

procedure TfrmPrincipal.dbedtCodBarrasProdutoVendasKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
var
  Jpg: TJPEGImage;
  Stream: TMemoryStream;
begin
  if Key = VK_RETURN then
  begin
    if Trim(dbedtCodBarrasProdutoVendas.Text) = '' then Exit;

    dmDados.qryProdutos.Close;
    dmDados.qryProdutos.SQL.Text := 'SELECT * FROM PRODUTOS WHERE CODIGO_BARRAS = :Cod OR ID = :Cod';
    dmDados.qryProdutos.Parameters.ParamByName('Cod').Value := dbedtCodBarrasProdutoVendas.Text;
    dmDados.qryProdutos.Open;

    if not dmDados.qryProdutos.IsEmpty then
    begin
      lblDescricaoProduto.Caption := dmDados.qryProdutos.FieldByName('NOME').AsString;
      lblPrecoVendas.Caption    := FormatFloat('R$ #,##0.00', dmDados.qryProdutos.FieldByName('PRECO_VENDA').AsCurrency);

      if not dmDados.qryProdutos.FieldByName('FOTO').IsNull then
      begin
        Stream := TMemoryStream.Create;
        try
          Jpg := TJPEGImage.Create;
          try
            TBlobField(dmDados.qryProdutos.FieldByName('FOTO')).SaveToStream(Stream);
            Stream.Position := 0;
            Jpg.LoadFromStream(Stream);
            dbimgFotoProdutoVendas.Picture.Assign(Jpg);
          finally
            Jpg.Free;
          end;
        finally
          Stream.Free;
        end;
      end
      else
        dbimgFotoProdutoVendas.Picture := nil; // Limpa se não tiver foto

      dbedtQuantidadeProdutoVendas.Text := '1';
      dbedtQuantidadeProdutoVendas.SetFocus;
      dbedtQuantidadeProdutoVendas.SelectAll;
    end
    else
    begin
      ShowMessage('Produto não encontrado!');
      dbedtCodBarrasProdutoVendas.SelectAll;
    end;
  end;
end;

procedure TfrmPrincipal.dbgrdDataControlsDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  vTipo: string;
begin
  vTipo := dmDados.qryFinanceiro.FieldByName('TIPO').AsString;

  if vTipo = 'D' then
    dbgrdDataControls.Canvas.Font.Color := clRed
  else if vTipo = 'R' then
    dbgrdDataControls.Canvas.Font.Color := clGreen;

  if gdSelected in State then
  begin
    dbgrdDataControls.Canvas.Brush.Color := clHighlight;
    dbgrdDataControls.Canvas.Font.Color := clWhite;
  end;

  dbgrdDataControls.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

//procedure TfrmPrincipal.dbgrdEstoqueCellClick(Column: TColumn);
//begin
//  lblProdutoSelecionado.Caption := dmDados.qryProdutos.FieldByName('DESCRICAO').AsString;
//  lblSaldoAtual.Caption := 'Saldo Atual: ' + dmDados.qryProdutos.FieldByName('ESTOQUE').AsString;
//
//  edtQtdMov.Text := '';
//  mmoMotivo.Lines.Clear;
//  edtQtdMov.SetFocus;
//end;

//procedure TfrmPrincipal.dbgrdEstoqueDrawColumnCell(Sender: TObject;
//  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
//var
//  vEstoque: Double;
//begin
//  if not dmDados.qryEstoque.FieldByName('ESTOQUE').IsNull then
//  begin
//    vEstoque := dmDados.qryEstoque.FieldByName('ESTOQUE').AsFloat;
//
//    if vEstoque < 0 then
//      dbgrdEstoque.Canvas.Font.Color := clRed
//    else if vEstoque <= 5 then
//      dbgrdEstoque.Canvas.Font.Color := clMaroon
//    else
//      dbgrdEstoque.Canvas.Font.Color := clBlack;
//  end;
//
//  if gdSelected in State then
//  begin
//    dbgrdEstoque.Canvas.Brush.Color := clHighlight;
//    dbgrdEstoque.Canvas.Font.Color  := clHighlightText;
//  end;
//
//  dbgrdEstoque.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//end;

procedure TfrmPrincipal.gridHistoricoEstoqueDrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  //aguardando código...
end;

procedure TfrmPrincipal.gridContasPagarFinanceiroDrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  StatusDaConta: String;
  DataVencimento: TDate;
begin
  if not dmDados.qryPagar.IsEmpty then
  begin
    StatusDaConta  := dmDados.qryPagar.FieldByName('STATUS').AsString;
    DataVencimento := dmDados.qryPagar.FieldByName('DATA_VENCIMENTO').AsDateTime;

    // Se NÃO estiver selecionada pelo mouse (para não estragar a cor de seleção do Windows)
    if not (gdSelected in State) then
    begin
      if StatusDaConta = 'PAGO' then
        gridContasPagarFinanceiro.Canvas.Font.Color := clGreen
        // Verde de PAGO
      else if (StatusDaConta = 'ABERTO') and (Trunc(DataVencimento) < Trunc(Date)) then
        gridContasPagarFinanceiro.Canvas.Font.Color := clRed;
        // Vermelho de ATRASADO
    end;
  end;

  // Manda o Delphi desenhar o texto com a cor que escolhemos acima
  gridContasPagarFinanceiro.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmPrincipal.gridDevLogsDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  vAcao: string;
begin
  vAcao := dmDados.qryDevLogs.FieldByName('ACAO').AsString;

  if (vAcao = 'EXCLUIR') or (vAcao = 'ERRO') then
  begin
    gridDevLogs.Canvas.Font.Color := clRed;
    gridDevLogs.Canvas.Font.Style := [fsBold];
  end
  else if (vAcao = 'LOGIN') then
  begin
    gridDevLogs.Canvas.Font.Color := clBlue;
  end
  else
  begin
    gridDevLogs.Canvas.Font.Color := clBlack;
  end;

  if (gdSelected in State) then 
  begin
    gridDevLogs.Canvas.Brush.Color := clHighlight;
    gridDevLogs.Canvas.Font.Color := clWhite;
  end;

  gridDevLogs.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

//procedure TfrmProdutos.gridProdutosDrawColumnCell(Sender: TObject; const Rect: TRect;
//  DataCol: Integer; Column: TColumn; State: gridProdutosDrawState);
//begin
//  if not (gdSelected in State) then
//  begin
//    if (dmDados.qryProdutos.RecNo mod 2 = 1) then
//      gridProdutos.Canvas.Brush.Color := $00F0F0F0
//    else
//      gridProdutos.Canvas.Brush.Color := clWhite;
//  end;
//
//  gridProdutos.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//end;

procedure TfrmPrincipal.gridProdutosDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if not (gdSelected in State) then
  begin
    if (dmDados.qryProdutos.RecNo mod 2 = 1) then
      gridProdutos.Canvas.Brush.Color := $00F0F0F0
    else
      gridProdutos.Canvas.Brush.Color := clWhite;
  end;

  gridProdutos.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmPrincipal.gridProdutosTitleClick(Column: TColumn);
begin
  dmDados.qryProdutos.Sort := Column.FieldName + ' ASC';

  dmDados.qryProdutos.Close;
  dmDados.qryProdutos.SQL.Text := 'SELECT * FROM PRODUTOS ORDER BY ' + Column.FieldName;
  dmDados.qryProdutos.Open;
end;

procedure TfrmPrincipal.gridSessoesUsuariosCellClick(Column: TColumn);
var
  vDataLogin, vUltimoPing: TDateTime;
  vMinutosOcioso: Integer;
  vUserID: Integer;
begin
  if dmDados.qrySessoes.IsEmpty then Exit;

  vDataLogin  := dmDados.qrySessoes.FieldByName('DATA_LOGIN').AsDateTime;
  vUltimoPing := dmDados.qrySessoes.FieldByName('ULTIMO_PING').AsDateTime; 
  vUserID     := dmDados.qrySessoes.FieldByName('ID_USUARIO').AsInteger;

  lblTempoConectado.Caption := 'Tempo Conectado: ' + FormatDateTime('hh:nn', Now - vDataLogin);

  vMinutosOcioso := MinutesBetween(Now, vUltimoPing);

  if vMinutosOcioso < 2 then
  begin
    lblStatusConexao.Caption := 'Status: ONLINE (Ativo)';
    lblStatusConexao.Font.Color := clGreen;
  end
  else if vMinutosOcioso < 30 then
  begin
    lblStatusConexao.Caption := 'Status: AUSENTE (' + IntToStr(vMinutosOcioso) + ' min)';
    lblStatusConexao.Font.Color := clOlive;
  end
  else
  begin
    lblStatusConexao.Caption := 'Status: OCIOSO / DESCONECTADO';
    lblStatusConexao.Font.Color := clRed;
  end;

  mmoLogsAcoes.Clear;
  with dmDados.qryGenerica do
  begin
    Close;
    SQL.Text := 'SELECT DATA_HORA, TELA, ACAO FROM LOGS_SISTEMA ' +
                'WHERE ID_USUARIO = :ID ORDER BY ID DESC LIMIT 10';
    Parameters.ParamByName('ID').Value := vUserID;
    Open;

    while not Eof do
    begin
      mmoLogsAcoes.Lines.Add('[' + FormatDateTime('hh:nn', FieldByName('DATA_HORA').AsDateTime) + '] ' +
                             FieldByName('TELA').AsString + ' - ' + FieldByName('ACAO').AsString);
      Next;
    end;
  end;
end;


procedure TfrmPrincipal.edtBuscaProdutoEstoqueChange(Sender: TObject);
begin
  with dmDados.qryEstoque do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT ID, DESCRICAO, ESTOQUE FROM PRODUTOS');

    SQL.Add('WHERE DESCRICAO LIKE ' + QuotedStr('%' + edtBuscaProdutoEstoque.Text + '%'));

    SQL.Add('ORDER BY DESCRICAO');
    Open;
  end;
end;

procedure TfrmPrincipal.edtBuscaProdutoEstoqueKeyDown(Sender: TObject;
var Key: Word; Shift: TShiftState);
begin
  // Se o usuário apertou a tecla ENTER
  if Key = VK_RETURN then
  begin
    if Trim(edtBuscaProdutoEstoque.Text) = '' then Exit;

    dmDados.qryAux.Close;
    dmDados.qryAux.SQL.Clear;
    dmDados.qryAux.SQL.Add('SELECT ID, NOME, ESTOQUE FROM PRODUTOS');
    
    // Tenta buscar por ID/Código de Barras, se der erro, busca por Nome
    try
      StrToInt(edtBuscaProdutoEstoque.Text);
      dmDados.qryAux.SQL.Add('WHERE ID = :Busca OR COD_BARRAS = :Busca');
    except
      dmDados.qryAux.SQL.Add('WHERE NOME LIKE :BuscaLike');
    end;

    // Passa os parâmetros
    if dmDados.qryAux.Parameters.FindParam('Busca') <> nil then
      dmDados.qryAux.Parameters.ParamByName('Busca').Value := edtBuscaProdutoEstoque.Text
    else
      dmDados.qryAux.Parameters.ParamByName('BuscaLike').Value := '%' + edtBuscaProdutoEstoque.Text + '%';

    dmDados.qryAux.Open;

    // Se encontrou o produto, preenche a tela
    if not dmDados.qryAux.IsEmpty then
    begin
      // Guarda o ID (Pode ser num Label invisível na tela chamado lblIDProdutoOculto)
      lblNomeProdutoEstoque.Caption := dmDados.qryAux.FieldByName('ID').AsString;
      
      // Mostra visualmente para o usuário
      lblNomeProdutoEstoque.Caption := dmDados.qryAux.FieldByName('NOME').AsString;
      lblInformacaoSaldoAtualEstoque.Caption := dmDados.qryAux.FieldByName('ESTOQUE').AsString;
      
      // Foca no campo de Quantidade para ele já digitar
      edtQuantidadeEstoque.SetFocus;
      
      // Chama a simulação para zerar a previsão
      AtualizarSimulacao; 
    end
    else
    begin
      ShowMessage('Produto não encontrado!');
      edtBuscaProdutoEstoque.SelectAll;
    end;
  end;
end;

procedure TfrmPrincipal.edtValorExit(Sender: TObject);
begin
  try
    edtValor.Text := FormatFloat('#,##0,00', StrToFloatDef(edtValor.Text, 0));
  except
    edtValor.Text := '0,00';
  end;
end;

procedure TfrmPrincipal.edtDescricaoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TfrmPrincipal.edtMensagemKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Key := #0;
    btnEnviarAviso.Click;
  end;
end;

procedure TfrmPrincipal.edtQuantidadeEstoqueChange(Sender: TObject);
begin
  AtualizarSimulacao;
end;

procedure TfrmPrincipal.edtValorKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  AtualizarFeedbackVisual;
end;

procedure TfrmPrincipal.ExportarParaCSV(DataSet: TDataSet; NomeArquivo: string);
var
  Lista: TStringList;
  Linha: string;
  i: Integer;
begin
  Lista := TStringList.Create;
  try
    Linha := '';
    for i := 0 to DataSet.FieldCount - 1 do
      Linha := Linha + DataSet.Fields[i].DisplayName + ';';
    Lista.Add(Linha);

    DataSet.First;
    while not DataSet.Eof do
    begin
      Linha := '';
      for i := 0 to DataSet.FieldCount - 1 do
      begin

        Linha := Linha + '"' + DataSet.Fields[i].AsString + '";';
      end;
      Lista.Add(Linha);
      DataSet.Next;
    end;

    Lista.SaveToFile(NomeArquivo);
    ShowMessage('Sucesso! Arquivo gerado em: ' + NomeArquivo);
  finally
    Lista.Free;
  end;
end;

procedure TfrmPrincipal.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (ssCtrl in Shift) and (ssShift in Shift) and (Key = Ord('U')) then
  begin
    if not tsDesenvolvedor.TabVisible then
    begin
      if InputBox('Acesso Restrito', 'Digite a senha de Dev:', '') = 'wolf123' then
      begin
        tsDesenvolvedor.TabVisible := True;
        pcDevTools.ActivePage := tsDesenvolvedor;
        ShowMessage('Modo Desenvolvedor ATIVADO!');
      end;
    end
    else
    begin
      tsDesenvolvedor.TabVisible := False;
      pcDevTools.ActivePage := tsHome;
    end;
  end;
  if Key = VK_F5 then
  begin
    btnConfirmarMovimentoEstoque.Click;
  end;
end;

procedure TfrmPrincipal.FormShow(Sender: TObject);
var
  i: Integer;
  Stream: TMemoryStream;
  Jpg: TJPEGImage;
begin
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := False;

  PageControl1.ActivePage := tsHome;

  // Carrega dados da empresa
  //dmDados.qryEmpresa.Close;
  //dmDados.qryEmpresa.SQL.Text := 'SELECT * FROM EMPRESA WHERE ID = 1';
  //dmDados.qryEmpresa.Open;

  dmDados.qryEmpresa.Close; 
  dmDados.qryEmpresa.SQL.Clear;
  dmDados.qryEmpresa.SQL.Add('SELECT * FROM EMPRESA WHERE ID = 1');
  dmDados.qryEmpresa.Open;

  if iIDUsuarioLogado > 0 then
  begin
    AplicarPermissoes(iIDUsuarioLogado);
  end
  else
  begin
    tsFinanceiro.TabVisible := False;
    tsEstoque.TabVisible := False;
  end;

  if not dmDados.qryEmpresa.IsEmpty then
  begin
    lblNomeEmpresa.Caption := dmDados.qryEmpresa.FieldByName('NOME_FANTASIA').AsString;

    if not dmDados.qryEmpresa.FieldByName('LOGO').IsNull then
    begin
      Stream := TMemoryStream.Create;
      try
        Jpg := TJPEGImage.Create;
        try
          try
            TBlobField(dmDados.qryEmpresa.FieldByName('LOGO')).SaveToStream(Stream);
            Stream.Position := 0;

            Jpg.LoadFromStream(Stream);
            imgLogoEmpresaLogada.Picture.Assign(Jpg);
          except
          end;
        finally
          Jpg.Free;
        end;
      finally
        Stream.Free;
      end;
    end;
  end;


  for i := 0 to pgcDetalhePagto.PageCount - 1 do
    pgcDetalhePagto.Pages[i].TabVisible := False;

  pgcDetalhePagto.ActivePage := tsDinheiro;


  // Verifica se EU (Logado) posso ver a aba de permissões
  // Usamos uma Query auxiliar para checar minha própria permissão na hora
  dmDados.qryAux.Close;
  dmDados.qryAux.SQL.Text := 'SELECT AC_GERENCIAR_PERMISSOES FROM PERMISSOES_ACESSO WHERE ID_USUARIO = :ID';
  dmDados.qryAux.Parameters.ParamByName('ID').Value := dmDados.iIDUsuarioLogado;
  dmDados.qryAux.Open;

  if (not dmDados.qryAux.IsEmpty) and (dmDados.qryAux.FieldByName('AC_GERENCIAR_PERMISSOES').AsString = 'S') then
  begin
    tsPermissoesUsuarioCadastros.TabVisible := True;
  end
  else
  begin
    tsPermissoesUsuarioCadastros.TabVisible := False;
  end;
end;

//procedure TfrmPrincipal.FormShow(Sender: TObject);
//var
//  i: Integer;
//  Stream: TMemoryStream;
//  Jpg: TJPEGImage;
//begin
//  for i := 0 to PageControl1.PageCount - 1 do
//    PageControl1.Pages[i].TabVisible := False;
//    PageControl1.ActivePage := tsHome;
//
//  dmDados.qryEmpresa.Close;
//  dmDados.qryEmpresa.SQL.Text := 'SELECT * FROM EMPRESA WHERE ID = 1';
//  dmDados.qryEmpresa.Open;
//
//  PageControl1.ActivePage := tsHome;
//  if iIDUsuarioLogado > 0 then
//  begin
//    AplicarPermissoes(iIDUsuarioLogado);
//  end
//  else
//  begin
//    tsFinanceiro.TabVisible := False;
//    tsEstoque.TabVisible := False;
//  end;
//
//  if not dmDados.qryEmpresa.IsEmpty then
//  begin
//    lblNomeEmpresa.Caption := dmDados.qryEmpresa.FieldByName('NOME_FANTASIA').AsString;
//
//    if not dmDados.qryEmpresa.FieldByName('LOGO').IsNull then
//    begin
//      Stream := TMemoryStream.Create;
//      Jpeg := TJPEGImage.Create;
//
//        TBlobField(dmDados.qryEmpresa.FieldByName('LOGO')).SaveToStream(Stream);
//        Stream.Position := 0;
//
//        Jpg.LoadFromStream(Stream);
//        imgLogoEmpresa.Picture.Assign(Jpg);
//
//      except
//        // Se der erro na imagem, não faz nada (deixa vazio)
//      end;
//      Stream.Free;
//      Jpg.Free;
//    end;
//  end;
//end;


procedure TfrmPrincipal.pnlFeedbackResize(Sender: TObject);
begin
  pnlCardResumo.Left := (pnlFeedback.Width - pnlCardResumo.Width) div 2;
  pnlCardResumo.Top := (pnlFeedback.Height - pnlCardResumo.Height) div 2;
end;

//procedure TfrmPrincipal.rgTipoClick(Sender: TObject);
//var
//  vValor: Double;
//begin
//  vValor := StrToFloatDef(edtValor.Text, 0);
//
//  if rgTipo.ItemIndex = 0 then
//  begin
//    pnlFeedback.Color := $00E1F0FF;
//    lblTipoGrande.Caption := 'ENTRADA DE VALOR';
//    lblTipoGrande.Font.Color := clGreen;
//    lblValorGigante.Font.Color := clGreen;
//  end
//  else
//  begin
//    pnlFeedback.Color := $00E1E1FF;
//    lblTipoGrande.Caption := 'SAÍDA DE VALOR';
//    lblTipoGrande.Font.Color := clRed;
//    lblValorGigante.Font.Color := clRed;
//  end;

//  lblValorGigante.Caption := FormatFloat('R$ #,##0.00', vValor);
//end;

procedure TfrmPrincipal.rbApenasEstoqueClick(Sender: TObject);
begin
  dmDados.qrySessoes.Close;
  dmDados.qrySessoes.SQL.Text :=
     'SELECT S.* FROM SESSOES S ' +
     'INNER JOIN USUARIOS U ON S.ID_USUARIO = U.ID ' +
     'WHERE U.NIVEL_ACESSO = "ESTOQUE"';
  dmDados.qrySessoes.Open;
end;

procedure TfrmPrincipal.rbApenasVendedoresClick(Sender: TObject);
begin
  dmDados.qrySessoes.Close;
  dmDados.qrySessoes.SQL.Text := 
     'SELECT S.* FROM SESSOES S ' +
     'INNER JOIN USUARIOS U ON S.ID_USUARIO = U.ID ' +
     'WHERE U.NIVEL_ACESSO = "VENDEDOR"';
  dmDados.qrySessoes.Open;
end;

procedure TfrmPrincipal.rbTodosClick(Sender: TObject);
begin
  dmDados.qrySessoes.Close;
  dmDados.qrySessoes.SQL.Text := 
     'SELECT S.* FROM SESSOES S ' +
     'INNER JOIN USUARIOS U ON S.ID_USUARIO = U.ID ' +
     'WHERE U.NIVEL_ACESSO = "TODOS"';
  dmDados.qrySessoes.Open;
end;

procedure TfrmPrincipal.rdgOperacaoClick(Sender: TObject);
begin
  // Roda a simulação que já criamos antes
  AtualizarSimulacao;

  // Lógica Dinâmica de Interface
  case rdgOperacao.ItemIndex of
    0: // ENTRADA
    begin
      // Habilita campos de compra
      edtCustoUnitarioMovimnetacaoEstoque.Enabled := True;
      dblkcbbFiltrarFornecedorMovimentacaoEstoque.Enabled := True;
      edtCustoUnitarioMovimnetacaoEstoque.Color := clWindow;
      dblkcbbFiltrarFornecedorMovimentacaoEstoque.Color := clWindow;
      
      lblOperacaoDica.Caption := 'Preencha o Custo e o Fornecedor para a Nota Fiscal.';
    end;
    
    1, 2, 3: // SAÍDA, PERDA ou AJUSTE
    begin
      // Desabilita campos que não fazem sentido na saída
      edtCustoUnitarioMovimnetacaoEstoque.Enabled := False;
      dblkcbbFiltrarFornecedorMovimentacaoEstoque.Enabled := False;
      
      // Muda a cor para cinza para dar feedback visual de que está bloqueado
      edtCustoUnitarioMovimnetacaoEstoque.Color := clBtnFace;
      dblkcbbFiltrarFornecedorMovimentacaoEstoque.Color := clBtnFace;
      
      // Limpa os valores para não gravar sujeira no banco
      dblkcbbFiltrarFornecedorMovimentacaoEstoque.KeyValue := Null;

      if rdgOperacao.ItemIndex = 2 then
        lblOperacaoDica.Caption := 'ATENÇÃO: Requer observação e senha de supervisor.'
      else
        lblOperacaoDica.Caption := '';
    end;
  end;
end;

procedure TfrmPrincipal.rdgStatusContasPagarFinanceiroClick(Sender: TObject);
begin
  dmDados.qryPagar.Close;
  dmDados.qryPagar.SQL.Clear;
  dmDados.qryPagar.SQL.Add('SELECT * FROM FINANCEIRO WHERE TIPO = ''D''');

  case rdgStatusContasPagarFinanceiro.ItemIndex of
    1: dmDados.qryPagar.SQL.Add('AND STATUS = ''ABERTO''');
    2: dmDados.qryPagar.SQL.Add('AND STATUS = ''PAGO''');
    3: dmDados.qryPagar.SQL.Add('AND STATUS = ''ABERTO'' AND DATA_VENCIMENTO < :Hoje');
  end;

  if rdgStatusContasPagarFinanceiro.ItemIndex = 3 then
    dmDados.qryPagar.Parameters.ParamByName('Hoje').Value := Trunc(Date);

  dmDados.qryPagar.Open;
end;

procedure TfrmPrincipal.rgFormadePagamentoVendasClick(Sender: TObject);
begin
  // 0=Dinheiro, 1=Débito, 2=Crédito, 3=PIX, 4=Boleto
  
  case rgFormadePagamentoVendas.ItemIndex of
    0, 1: // Dinheiro ou Débito (Não precisa de detalhes)
    begin
      pgcDetalhePagto.ActivePage := tsDinheiro;
    end;

    2: // Cartão de Crédito
    begin
      pgcDetalhePagto.ActivePage := tsCartaoCredito;
      PopularParcelas; // Chama função para calcular juros/parcelas
    end;

    3: // PIX
    begin
      pgcDetalhePagto.ActivePage := tsPix;
      // Dica: Limpe o QR Code anterior se houver
      imgQRCode.Picture := nil;
      edtCopiaCola.Clear;
    end;

    4: // Boleto
    begin
      pgcDetalhePagto.ActivePage := tsBoleto;
      dtpVencimentoBoleto.Date := Date + 3; // Sugere 3 dias de prazo
    end;
  end;
end;

procedure TfrmPrincipal.rgTipoClick(Sender: TObject);
begin
  if rgTipo.ItemIndex = 0 then
  begin
    btnSalvarLancamento.Font.Color := clGreen;
    btnSalvarLancamento.Caption := 'CONFIRMAR ENTRADA (RECEITA)';
    edtValor.Font.Color := clGreen;
  end
  else
  begin
    btnSalvarLancamento.Font.Color := clRed;
    btnSalvarLancamento.Caption := 'CONFIRMAR SAÍDA (DESPESA)';
    edtValor.Font.Color := clRed;
  end;

  AtualizarFeedbackVisual;
end;

procedure TfrmPrincipal.rgTipoMovimentoClick(Sender: TObject);
begin
  if rdgOperacao.ItemIndex = 0 then
  begin
    btnConfirmarMovimentoEstoque.Caption := 'CONFIRMAR ENTRADA';
    btnConfirmarMovimentoEstoque.Font.Color := clGreen;
  end
  else
  begin
    btnConfirmarMovimentoEstoque.Caption := 'CONFRIMAR SAÍDA';
    btnConfirmarMovimentoEstoque.Font.Color := clRed;
  end;
end;

procedure TfrmPrincipal.tmrInicioHomeTimer(Sender: TObject);
var
  vTotalVendas, vContasReceber: Currency;
  vEstoqueBaixo, vAReceber: Integer;
begin
  tmrInicioHome.Enabled := False;

//  lblBemVindo.Caption := 'Olá, ' + dmDados.UsuarioNome + '! ' +
//                         FormatDateTime('dddd, dd "de" mmmm "de" yyyy', Date);

  with dmDados.qryGenerica do
  begin
    Close;
    SQL.Text := 'SELECT SUM(VALOR_TOTAL) FROM VENDA_CAB ' +
                'WHERE DATA_EMISSAO >= CURRENT_DATE AND STATUS = ''F''';
    Open;
    vTotalVendas := Fields[0].AsCurrency;
    lblValorVendasResumoHome.Caption := FormatFloat('R$ #,##0.00', vTotalVendas);

    Close;
    SQL.Text := 'SELECT COUNT(*) FROM PRODUTOS WHERE ESTOQUE_ATUAL <= ESTOQUE_MINIMO';
    Open;
    vEstoqueBaixo := Fields[0].AsInteger;
    lblTotalEstoqueCriticoResumoHome.Caption := IntToStr(vEstoqueBaixo);

    Close;
    SQL.Text := 'SELECT COUNT(*) FROM VENDA_CAB WHERE STATUS = ''A''';
    Open;
    vAReceber := Fields[0].AsInteger;
    lblTotalAReceberResumoHome.Caption := IntToStr(vAReceber);
  end;

  with dmDados.qryDashboard do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT "ESTOQUE" AS TIPO, DESCRICAO AS DETALHE, ');
    SQL.Add('CONCAT("Saldo: ", CAST(ESTOQUE_ATUAL AS CHAR)) AS OBS, "CRITICO" AS NIVEL ');
    SQL.Add('FROM PRODUTOS WHERE ESTOQUE_ATUAL <= ESTOQUE_MINIMO');

    SQL.Add('UNION ALL');

    SQL.Add('SELECT "FINANCEIRO", DESCRICAO, ');
    SQL.Add('CONCAT("Valor: R$ ", CAST(VALOR AS CHAR)), "MEDIO" ');
    SQL.Add('FROM FINANCEIRO WHERE TIPO="R" AND STATUS="PENDENTE" AND DATA_VENCIMENTO <= CURRENT_DATE');

    Open;
  end;

//  if not dmDados.qryDashboard.IsEmpty then
//  begin
//     mmoAvisosSistema.Columns[0].Width := 100;
//     mmoAvisosSistema.Columns[1].Width := 300;
//  end;
end;

procedure TfrmPrincipal.tmrSessaoTimer(Sender: TObject);
var
  FUltimaMsgID: Integer;
begin
  with dmDados.qryGenerica do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT ID, MENSAGEM FROM MENSAGENS_SISTEMA');
      SQL.Add('WHERE ID > :UltimoID AND DESTINO = ''TODOS''');
      SQL.Add('ORDER BY ID ASC');
      Parameters.ParamByName('UltimoID').Value := FUltimaMsgID;
      Open;

      while not Eof do
      begin
        ShowMessage('MENSAGEM DO ADMIN: ' + FieldByName('MENSAGEM').AsString);

        FUltimaMsgID := FieldByName('ID').AsInteger;
        Next;
      end;

      dmDados.qryExec.Close;
      dmDados.qryExec.SQL.Text := 'UPDATE SESSOES SET ULTIMO_PING = NOW() WHERE ID = :SessionID';
      dmDados.qryExec.Parameters.ParamByName('SessionID').Value := dmDados.SessaoID;
      dmDados.qryExec.ExecSQL;

      if Fields[0].AsInteger = 0 then
      begin
        tmrSessao.Enabled := False;
        ShowMessage('Sessão encerrada pelo Administrador. O sistema será fechado.');
        Application.Terminate;
      end;
    end;
end;

procedure TfrmPrincipal.tsDevVersaoShow(Sender: TObject);
var
  vCaminhoArquivo: String;
begin
  vCaminhoArquivo := ExtractFilePath(Application.ExeName) + 'changelog.txt';

  if FileExists(vCaminhoArquivo) then
    mmoVersao.Lines.LoadFromFile(vCaminhoArquivo)
  else
    mmoVersao.Lines.Text := 'Arquivo de histórico (changelog.txt) não encontrado.';
end;

procedure TfrmPrincipal.tsHomeShow(Sender: TObject);
begin
  tmrInicioHome.Enabled := True;
end;

procedure TfrmPrincipal.tsPermissoesShow(Sender: TObject);
begin
  CarregarArvoreSistema;
end;

//procedure TfrmPrincipal.tsVendasDesativadoShow(Sender: TObject);
//begin
//  StringGrid1.Cells[0,0] := 'Cód.';
//  StringGrid1.Cells[1,0] := 'Descrição do Produto';
//  StringGrid1.Cells[2,0] := 'Qtd.';
//  StringGrid1.Cells[3,0] := 'Unitário (R$)';
//  StringGrid1.Cells[4,0] := 'Total (R$)';
//
//  StringGrid1.ColWidths[0] := 80;
//  StringGrid1.ColWidths[1] := 350;
//  StringGrid1.ColWidths[2] := 50;
//  StringGrid1.ColWidths[3] := 90;
//  StringGrid1.ColWidths[4] := 90;
//
//  StringGrid1.Rows[1].Clear;
//
//  lblTotalVenda.Caption := '0,00';
//  edtCodigo.SetFocus;
//end;

procedure TfrmPrincipal.btnSalvarLancamentoClick(Sender: TObject);
var
  vMsg: string;
  vTipoBD: string;
begin
  if StrToFloatDef(edtValor.Text, 0) <= 0 then
  begin
    ShowMessage('O valor precisa ser maior que zero.');
    edtValor.SetFocus;
    Exit;
  end;

  if Trim(edtDescricao.Text) = '' then
  begin
    ShowMessage('A descrição é obrigatória.');
    edtDescricao.SetFocus;
    Exit;
  end;

  if rgTipo.ItemIndex = 0 then
  begin
    vMsg := 'Confirmar a ENTRADA de ' + lblValorGigante.Caption + '?';
    vTipoBD := 'R';
  end
  else
  begin
    vMsg := 'Confirmar o PAGAMENTO (Saída) de ' + lblValorGigante.Caption + '?';
    vTipoBD := 'D';
  end;

  if MessageDlg(vMsg, mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    Exit;

  with dmDados.qryExec do
  begin
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO FINANCEIRO (DESCRICAO, VALOR, TIPO, DATA_VENCIMENTO, PAGO_SN)');
    SQL.Add('VALUES (:Desc, :Val, :Tipo, :Dt, :Pago)');
    
    Parameters.ParamByName('Desc').Value := edtDescricao.Text;
    Parameters.ParamByName('Val').Value  := StrToFloat(edtValor.Text);
    Parameters.ParamByName('Tipo').Value := vTipoBD;
    Parameters.ParamByName('Dt').Value   := dtpVencimento.Date;
    Parameters.ParamByName('Pago').Value := 'S';
    
    try
      ExecSQL;

      edtDescricao.Clear;
      edtValor.Text := '0,00';
      edtDescricao.SetFocus;

      AtualizarFeedbackVisual;
    except
      on E: Exception do
        ShowMessage('Erro crítico ao salvar: ' + E.Message);
    end;
  end;
end;

procedure TfrmPrincipal.btnSalvarOrcamentoVendasClick(Sender: TObject);
var
  idOrcamento: Integer;
begin
  // Validações básicas
  if dsOrcamentosItens.IsEmpty then Exit;

  // Verifica se é Edição ou Novo
  if dbedtNumOrcamento.Text = 'NOVO' then
  begin
    // Pega o ID gerado e joga na variável idOrcamento
  end
  else
  begin
    idOrcamento := StrToInt(dbedtNumOrcamento.Text);
    // UPDATE ORCAMENTOS SET ... WHERE ID = idOrcamento
    // DELETE FROM ORCAMENTOS_ITENS WHERE ID_ORCAMENTO = idOrcamento (Limpa itens velhos para gravar os novos)
  end;

  // Grava os Itens do CDS para a tabela ORCAMENTOS_ITENS
  dsOrcamentosItens.First;
  while not dsOrcamentosItens.Eof do
  begin
     // INSERT INTO ORCAMENTOS_ITENS ...
     dsOrcamentosItens.Next;
  end;

  ShowMessage('Orçamento salvo com sucesso!');
end;

procedure TfrmPrincipal.btnSalvarProdutoCadastrosClick(Sender: TObject);
begin
  if dmDados.qryProdutos.State in [dsInsert, dsEdit] then
  // Verifica se o usuário está realmente inserindo ou editando algo
  begin
    try
      dmDados.qryProdutos.Post;
      ShowMessage('Mercadoria salva com sucesso!');

      edtProdutoCadastros.Clear;
    except
      on E: Exception do
      begin
        ShowMessage('Erro ao gravar: ' + E.Message);
      end;
    end;
  end;
end;

end.
