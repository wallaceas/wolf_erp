unit uAutenticacao;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DB, ADODB, dxGDIPlusClasses;

type
  TfrmAutenticacao = class(TForm)
    pnlLogoWolfERP: TPanel;
    pnlConfirmacaoAcao: TPanel;
    imgLogoWolfERP: TImage;
    pnlInformacaoConfirmaAcao: TPanel;
    lblUsuarioAutenticacao: TLabel;
    edtUsuarioAutenticacao: TEdit;
    lblSenhaAutenticacao: TLabel;
    edtSenhaAutenticacao: TEdit;
    pnlNomeAcaoRealizada: TPanel;
    btnConfirmarAcao: TBitBtn;
    btnCancelarAcao: TBitBtn;
    lblConfirmarAcao: TLabel;
    lblAcaoRealizada: TLabel;
    procedure FormShow(Sender: TObject);
    //procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnConfirmarAcaoClick(Sender: TObject);
  private
    { Private declarations }
    FPermissaoExigida: String;
  public
    { Public declarations }
    // Função de Classe: Faz tudo sozinha (Cria, Configura, Mostra e Destrói)
    function Solicitar(TituloAcao: String; EhPerigoso: Boolean; AC_EXCLUIR_CADASTROS: String): Boolean;
  end;

var
  frmAutenticacao: TfrmAutenticacao;

implementation

uses uDmDados, uPrincipal;

{$R *.dfm}

function TfrmAutenticacao.Solicitar(TituloAcao: String; EhPerigoso: Boolean; AC_EXCLUIR_CADASTROS: String): Boolean;
begin
  Result := False;
  frmAutenticacao := TfrmAutenticacao.Create(nil);
  try
    // Configura o Visual dinamicamente
    frmAutenticacao.lblAcaoRealizada.Caption := TituloAcao;
    frmAutenticacao.FPermissaoExigida := AC_EXCLUIR_CADASTROS;

    if EhPerigoso then
      frmAutenticacao.lblAcaoRealizada.Font.Color := clRed  // Ações de Exclusão
    else
      frmAutenticacao.lblAcaoRealizada.Font.Color := clGreen; // Ações de Finalizar

    if frmAutenticacao.ShowModal = mrOk then
      Result := True;
  finally
    frmAutenticacao.Free;
  end;
end;

procedure TfrmAutenticacao.FormShow(Sender: TObject);
begin
  edtUsuarioAutenticacao.Clear;
  edtSenhaAutenticacao.Clear;
  edtUsuarioAutenticacao.SetFocus;
end;

procedure TfrmAutenticacao.btnConfirmarAcaoClick(Sender: TObject);
var
  qry: TADOQuery;
begin
  if (Trim(edtUsuarioAutenticacao.Text) = '') or (Trim(edtSenhaAutenticacao.Text) = '') then
  begin
    ShowMessage('Informe matrícula e senha.');
    edtUsuarioAutenticacao.SetFocus;
    Exit;
  end;

  qry := TADOQuery.Create(nil);
  try
    qry.Connection := dmDados.wolferp;
    qry.SQL.Clear;
    
    // Consulta Cruzada: Usuário existe? Senha bate? Tem a permissão específica?
    // Ajuste "PERMISSOES_ACESSO" e "ID_USUARIO" conforme seu banco real
    qry.SQL.Add('SELECT U.NOME FROM USUARIOS U');
    qry.SQL.Add('INNER JOIN PERMISSOES_ACESSO P ON P.ID_USUARIO = U.ID');
    qry.SQL.Add('WHERE U.MATRICULA = :Matricula AND U.SENHA = :Senha');

    // Aqui verifica a permissão dinâmica
    if Trim(FPermissaoExigida) <> '' then
       qry.SQL.Add('AND P.' + FPermissaoExigida + ' = ''S''');

    qry.Parameters.ParamByName('Matricula').Value := edtUsuarioAutenticacao.Text;
    qry.Parameters.ParamByName('Senha').Value := edtSenhaAutenticacao.Text;
    qry.Open;

    if not qry.IsEmpty then
    begin
      ModalResult := mrOk;
    end
    else
    begin
      Application.MessageBox('Acesso Negado ou Sem Permissão!', 'Segurança', MB_ICONWARNING + MB_OK);
      edtSenhaAutenticacao.Clear;
      edtSenhaAutenticacao.SetFocus;
      ModalResult := mrNone;
    end;
  finally
    qry.Free;
  end;
end;

end.


//unit uAutenticacao;
//
//interface
//
//uses
//  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
//  Dialogs, StdCtrls, Buttons, ExtCtrls, DB, ADODB, dxGDIPlusClasses;
//
//type
//  TfrmAutenticacao = class(TForm)
//    pnlLogoWolfERP: TPanel;
//    pnlConfirmacaoAcao: TPanel;
//    imgLogoWolfERP: TImage;
//    pnlInformacaoConfirmaAcao: TPanel;
//    lblUsuarioAutenticacao: TLabel;
//    edtUsuarioAutenticacao: TEdit;
//    lblSenhaAutenticacao: TLabel;
//    edtSenhaAutenticacao: TEdit;
//    pnlNomeAcaoRealizada: TPanel;
//    btnConfirmarAcao: TBitBtn;
//    btnCancelarAcao: TBitBtn;
//    lblConfirmarAcao: TLabel;
//    lblAcaoRealizada: TLabel;
//    procedure FormShow(Sender: TObject);
//    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
//    procedure btnConfirmarAcaoClick(Sender: TObject);
//  private
//    { Private declarations }
//    function VerificarCredenciais: Boolean;
//  public
//    { Public declarations }
//    UsuarioAutorizador: String;
//  end;
//
//var
//  frmAutenticacao: TfrmAutenticacao;
//
//  function SolicitarLiberacao(DescricaoRotina: String): Boolean;
//
//implementation
//
//uses uDmDados, uPrincipal;
//
//{$R *.dfm}
//
//
//function SolicitarLiberacao(DescricaoRotina: String): Boolean;
//var
//  Tela: TfrmAutenticacao;
//begin
//  Result := False;
//  Tela := TfrmAutenticacao.Create(nil);
//  try
//    Tela.lblRotina.Caption := DescricaoRotina;
//    // Se o usuário clicar no Check Verde (ModalResult = mrOk), retorna True
//    if Tela.ShowModal = mrOk then
//      Result := True;
//  finally
//    Tela.Free;
//  end;
//end;
//
//// -----------------------------------------------------------------------------
//// TituloAcao: O que vai aparecer escrito (ex: "EXCLUIR PRODUTO")
//// EhPerigoso: Se TRUE, o texto fica VERMELHO. Se FALSE, fica AZUL/VERDE.
//// PermissaoBanco: O nome da coluna no banco (ex: 'AC_EXCLUIR')
//// -----------------------------------------------------------------------------
//class function TfrmAutenticacao.Solicitar(TituloAcao: String; EhPerigoso: Boolean; AC_EXCLUIR_CADASTRO: String): Boolean;
//begin
//  Result := False;
//  frmAutenticacao := TfrmAutenticacao.Create(nil);
//  try
//    // 1. Configura o Visual dinamicamente
//    frmAutenticacao.lblAcao.Caption := TituloAcao;
//    frmAutenticacao.FPermissaoExigida := PermissaoBanco;
//
//    if EhPerigoso then
//      frmAutenticacao.lblAcao.Font.Color := clRed  // Ações de Exclusão
//    else
//      frmAutenticacao.lblAcao.Font.Color := clBlue; // Ações de Finalizar/Desconto
//
//    // 2. Mostra a tela e espera (ShowModal)
//    if frmAutenticacao.ShowModal = mrOk then
//      Result := True;
//  finally
//    frmAutenticacao.Free;
//  end;
//end;
//
//procedure TfrmAutenticacao.FormShow(Sender: TObject);
//begin
//  edtUsuarioAutenticacao.Clear;
//  edtSenhaAutenticacao.Clear;
//  edtUsuarioAutenticacao.SetFocus;
//end;
//
//procedure TfrmAutenticacao.btnConfirmarClick(Sender: TObject);
//var
//  qry: TADOQuery;
//begin
//  if (Trim(edtUsuario.Text) = '') or (Trim(edtSenha.Text) = '') then
//  begin
//    ShowMessage('Informe usuário e senha.');
//    edtUsuario.SetFocus;
//    Exit;
//  end;
//
//  qry := TADOQuery.Create(nil);
//  try
//    qry.Connection := dmDados.ConexaoDB;
//    qry.SQL.Clear;
//    
//    // Consulta Cruzada: Usuário existe? Senha bate? Tem a permissão específica?
//    // Ajuste "PERMISSOES_ACESSO" e "ID_USUARIO" conforme seu banco real
//    qry.SQL.Add('SELECT U.NOME FROM USUARIOS U');
//    qry.SQL.Add('INNER JOIN PERMISSOES_ACESSO P ON P.ID_USUARIO = U.ID');
//    qry.SQL.Add('WHERE U.LOGIN = :Login AND U.SENHA = :Senha');
//    
//    // Aqui verifica a permissão dinâmica (ex: AND P.AC_EXCLUIR = 'S')
//    if Trim(FPermissaoExigida) <> '' then
//       qry.SQL.Add('AND P.' + FPermissaoExigida + ' = ''S''');
//
//    qry.Parameters.ParamByName('Login').Value := edtUsuario.Text;
//    qry.Parameters.ParamByName('Senha').Value := edtSenha.Text; // Use MD5 se tiver
//    qry.Open;
//
//    if not qry.IsEmpty then
//    begin
//      // SUCESSO!
//      ModalResult := mrOk; 
//    end
//    else
//    begin
//      Application.MessageBox('Acesso Negado ou Sem Permissão!', 'Segurança', MB_ICONWARNING + MB_OK);
//      edtSenha.Clear;
//      edtSenha.SetFocus;
//      ModalResult := mrNone; // Mantém a tela aberta para tentar de novo
//    end;
//  finally
//    qry.Free;
//  end;
//end;
//
////procedure TfrmAutenticacao.btnConfirmarAcaoClick(Sender: TObject);
////begin
////  if VerificarCredenciais then
////  begin
////    ModalResult := mrOk; // Fecha a tela com sucesso
////  end
////  else
////  begin
////    ShowMessage('Usuário ou Senha inválidos!');
////    edtSenhaAutenticacao.Clear;
////    edtSenhaAutenticacao.SetFocus;
////    // Não fecha a tela, dá outra chance
////    ModalResult := mrNone;
////  end;
////end;
////
////procedure TfrmAutenticacao.FormKeyDown(Sender: TObject; var Key: Word;
////  Shift: TShiftState);
////begin
////    if Key = VK_RETURN then
////    btnConfirmarAcao.Click;
////  if Key = VK_ESCAPE then
////    ModalResult := mrCancel;
////end;
////
////procedure TfrmAutenticacao.FormShow(Sender: TObject);
////begin
////  edtUsuarioAutenticacao.Clear;
////  edtSenhaAutenticacao.Clear;
////  edtUsuarioAutenticacao.SetFocus;
////end;
////
////function TfrmAutenticacao.VerificarCredenciais: Boolean;
////var
////  qryUsuarios: TADOQuery;
////begin
////  Result := False;
////
////  if (Trim(edtUsuarioAutenticacao.Text) = '') or (Trim(edtSenhaAutenticacao.Text) = '') then Exit;
////
////  qryUsuarios := TADOQuery.Create(nil);
////  try
////    qryUsuarios.Connection := dmDados.wolferp;
////    qryUsuarios.SQL.Add('SELECT MATRICULA FROM USUARIOS WHERE MATRICULA = :Matricula AND SENHA = :Senha');
////    qryUsuarios.Parameters.ParamByName('Matricula').Value := edtUsuarioAutenticacao.Text;
////    qryUsuarios.Parameters.ParamByName('Senha').Value := edtSenhaAutenticacao.Text;
////
////    qryUsuarios.Open;
////
////    if not qryUsuarios.IsEmpty then
////    begin
////      Result := True;
////      UsuarioAutorizador := qryUsuarios.FieldByName('NOME').AsString;
////    end;
////  finally
////    qryUsuarios.Free;
////  end;
////end;
//
//end.
