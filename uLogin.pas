unit uLogin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, JPEG, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, uDmDados, IdHashMessageDigest, IdGlobal,
  dxGDIPlusClasses, ADODB, DB, ComCtrls, DBCtrls;

type
  TfrmLogin = class(TForm)
    pnlMarca: TPanel;
    imgLogo: TImage;
    lblVersao: TLabel;
    btnCancelar: TBitBtn;
    btnEntrar: TBitBtn;
    edtMatricula: TEdit;
    edtSenha: TEdit;
    lblSenha: TLabel;
    lblUsuario: TLabel;
    btnTrocaSenhaLogin: TBitBtn;
    grpSelecioneEmpresa: TGroupBox;
    pnlTrocaSenhaLogin: TPanel;
    lblTrocaSenhaLogin: TLabel;
    pnlInfoNovaSenhaLogin: TPanel;
    pnlNovaSenhaLogin: TPanel;
    lblNovaSenhaLogin: TLabel;
    edtNovaSenhaLogin: TEdit;
    pnlConfirmarSenhaLogin: TPanel;
    lblConfirmaSenhaLogin: TLabel;
    edtConfirmaSenha: TEdit;
    pnlDataExpiracaoLogin: TPanel;
    lblDataExpiracao: TLabel;
    dtpExpiracao: TDateTimePicker;
    pnlConfirmarECancelarTrocaSenhaLogin: TPanel;
    btnConfirmarTroca: TBitBtn;
    btnCancelarTroca: TBitBtn;
    dblkcbbEmpresa: TDBLookupComboBox;
    procedure pnlMarcaMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnFecharClick(Sender: TObject);
    procedure btnEntrarClick(Sender: TObject);
    procedure edtSenhaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edtMatriculaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edtMatriculaKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    //procedure btnInfoLoginClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnTrocaSenhaLoginClick(Sender: TObject);
    procedure btnConfirmarTrocaClick(Sender: TObject);
    procedure btnCancelarTrocaClick(Sender: TObject);
  private
    { Private declarations }
    iIDUsuarioLogado: Integer;
    fIDUsuarioTroca: Integer;
    function GerarMD5(Texto: String): String;
  public
    { Public declarations }
  end;

var
  frmLogin: TfrmLogin;

implementation

uses uPrincipal;

//uses uDmDados, uPrincipal;

{$R *.dfm}

function ObterVersaoDoExecutavel: String;
var
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
begin
  Result := 'v1.0.0 (Dev)';

  VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
  
  if VerInfoSize > 0 then
  begin
    GetMem(VerInfo, VerInfoSize);
    try
      if GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo) then
      begin
        if VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize) then
        begin
          Result := 'v' + IntToStr(VerValue^.dwFileVersionMS shr 16) + '.' +
                          IntToStr(VerValue^.dwFileVersionMS and $FFFF) + '.' +
                          IntToStr(VerValue^.dwFileVersionLS shr 16);
                          // + '.' + IntToStr(VerValue^.dwFileVersionLS and $FFFF);
        end;
      end;
    finally
      FreeMem(VerInfo, VerInfoSize);
    end;
  end;
end;

function TfrmLogin.GerarMD5(Texto: String): String;
var
  IdMD5: TIdHashMessageDigest5;
begin
  IdMD5 := TIdHashMessageDigest5.Create;
  try
    Result := IdMD5.AsHex(IdMD5.HashValue(Texto));
  finally
    IdMD5.Free;
  end;
end;

procedure TfrmLogin.btnCancelarClick(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TfrmLogin.btnCancelarTrocaClick(Sender: TObject);
begin
  pnlTrocaSenhaLogin.Visible := False;
  edtSenha.SetFocus;
end;

procedure TfrmLogin.btnConfirmarTrocaClick(Sender: TObject);
begin
  // Validações da Nova Senha
  if Trim(edtNovaSenhaLogin.Text) = '' then
  begin
    ShowMessage('A nova senha não pode ser vazia.');
    edtNovaSenhaLogin.SetFocus;
    Exit;
  end;

  if edtNovaSenhaLogin.Text <> edtConfirmaSenha.Text then
  begin
    ShowMessage('A confirmação da senha não confere!');
    edtConfirmaSenha.Clear;
    edtConfirmaSenha.SetFocus;
    Exit;
  end;

  // Atualiza no Banco de Dados
  try

    dmDados.ExecSQL('UPDATE USUARIOS SET SENHA = ' + QuotedStr(edtNovaSenhaLogin.Text) +
                  ', VALIDADE_SENHA = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', dtpExpiracao.Date)) + 
                  ' WHERE ID = ' + IntToStr(fIDUsuarioTroca));

//    dmDados.ExecSQL('UPDATE USUARIOS SET SENHA = ' + QuotedStr(edtNovaSenhaLogin.Text) +
//                    ', VALIDADE_SENHA = :DataValidade ' + // Assumindo que você tem campo de validade
//                    ' WHERE ID = ' + IntToStr(fIDUsuarioTroca),
//                    [dtpExpiracao.Date]); // Passando a data como parâmetro se sua função ExecSQL suportar, senão monte a string

    // Exemplo alternativo se seu ExecSQL for simples:
    // dmDados.ExecSQL('UPDATE USUARIOS SET SENHA = ' + QuotedStr(edtNovaSenhaLogin.Text) +
    //                 ', DATA_EXPIRACAO = ' + QuotedStr(FormatDateTime('yyyy-mm-dd', dtpExpiracao.Date)) +
    //                 ' WHERE ID = ' + IntToStr(fIDUsuarioTroca));

    ShowMessage('Senha alterada com sucesso! Utilize a nova senha para entrar.');
    
    // Fecha o painel e atualiza o campo de senha principal
    pnlTrocaSenhaLogin.Visible := False;
    edtSenha.Text := edtNovaSenhaLogin.Text;
    btnEntrar.SetFocus;

  except
    on E: Exception do
      ShowMessage('Erro ao alterar senha: ' + E.Message);
  end;
end;

procedure TfrmLogin.btnEntrarClick(Sender: TObject);
var
  vMatricula: String;
  qryLogin: TADOQuery;
begin

  // Obrigatório selecionar empresa
  if dblkcbbEmpresa.KeyValue = Null then
  begin
    ShowMessage('Por favor, selecione a Empresa para logar.');
    dblkcbbEmpresa.SetFocus;
    Exit;
  end;

  vMatricula := Trim(edtMatricula.Text);

  // Se o usuário digitou apenas 3 números (ex: "105"), adiciona o zero à esquerda (vira "0105")
  if Length(vMatricula) = 3 then
  begin
    vMatricula := '0' + vMatricula;
    edtMatricula.Text := vMatricula;
  end;

  if Length(vMatricula) <> 4 then
  begin
    ShowMessage('A Matrícula deve ter 4 dígitos (Ex: 0000).');
    edtMatricula.SetFocus;
    Exit;
  end;

  if Length(Trim(edtSenha.Text)) <> 4 then
  begin
    ShowMessage('A Senha deve conter exatamente 4 dígitos!');
    edtSenha.SetFocus;
    Exit;
  end;

  with dmDados.qryGenerica do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT ID, NOME_COMPLETO, NIVEL_ACESSO, ATIVO, SENHA');
    SQL.Add('FROM USUARIOS WHERE MATRICULA = :Matricula');

    Parameters.ParamByName('Matricula').Value := vMatricula;
    Open;

    if IsEmpty then
    begin
      ShowMessage('Matrícula ' + vMatricula + ' não encontrada.');
      edtMatricula.SetFocus;
      Exit;
    end;

    if FieldByName('SENHA').AsString <> edtSenha.Text then
    begin
      ShowMessage('Senha incorreta.');
      edtSenha.Clear;
      edtSenha.SetFocus;
      Exit;
    end;

    if FieldByName('ATIVO').AsString = 'N' then
    begin
      ShowMessage('Acesso negado. Usuário inativo.');
      Exit;
    end;

    dmDados.UsuarioLogadoID := FieldByName('ID').AsInteger;
    dmDados.UsuarioNome     := FieldByName('NOME_COMPLETO').AsString;
    
    if FieldByName('NIVEL_ACESSO').AsString = 'ADMIN' then
       dmDados.UsuarioEhAdm := True
    else
       dmDados.UsuarioEhAdm := False;

    with dmDados.qryGenerica do
    begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO SESSOES (ID_USUARIO, DATA_LOGIN, ULTIMO_PING) ');
      SQL.Add('VALUES (:IdUser, NOW(), NOW())');
      
      Parameters.ParamByName('IdUser').Value := dmDados.UsuarioLogadoID;
      // Parameters.ParamByName('IP').Value := '127.0.0.1';
      ExecSQL;

      Close;
      SQL.Text := 'SELECT LAST_INSERT_ID() AS ID_GERADO';
      Open;

      dmDados.SessaoID := FieldByName('ID_GERADO').AsInteger;
    end;

    ModalResult := mrOk; 
  end;
//  if SenhaCorreta then
//  begin
//  // Passa o ID para o formulário principal antes de fechar o login
//    frmPrincipal.iIDUsuarioLogado := qryLogin.FieldByName('ID').AsInteger;
//
//    ModalResult := mrOk; // Fecha o login e abre o principal
//  end;

  if (Trim(edtMatricula.Text) = '') or (Trim(edtSenha.Text) = '') then
  begin
    ShowMessage('Por favor, informe Usuário e Senha.');
    edtMatricula.SetFocus;
    Exit;
  end;

  qryLogin := TADOQuery.Create(nil);
  try
    qryLogin.connection := dmDados.wolferp;

    qryLogin.SQL.Clear;
    // Trazemos apenas ID e NOME. Não precisamos das permissões agora.
    // O Menu Principal carregará as permissões depois usando o ID.
    qryLogin.SQL.Add('SELECT ID, NOME_COMPLETO, SENHA FROM USUARIOS');
    qryLogin.SQL.Add('WHERE MATRICULA = :Matricula AND SENHA = :Senha');

    // 4. Parâmetros
    qryLogin.Parameters.ParamByName('Matricula').Value := edtMatricula.Text;
    qryLogin.Parameters.ParamByName('Senha').Value   := edtSenha.Text;

    qryLogin.Open;

    if not qryLogin.IsEmpty then
    begin
      frmPrincipal.iIDUsuarioLogado := qryLogin.FieldByName('ID').AsInteger;

      // Opcional: Passar o nome para exibir "Bem vindo, João"
      // frmPrincipal.lblNomeUsuario.Caption := qryLogin.FieldByName('NOME').AsString;

      ModalResult := mrOk;
    end
    else
    begin
      ShowMessage('Usuário ou Senha incorretos.');
      edtSenha.Clear;
      edtSenha.SetFocus;
    end;

  finally
    qryLogin.Free;
  end;
end;

//var
//  vSenhaHash: String;
//begin
//  if Trim(edtMatricula.Text) = '' then
//  begin
//    ShowMessage('Por favor, digite seu usuário.');
//    edtMatricula.SetFocus;
//    Exit;
//  end;
//
//  if Trim(edtSenha.Text) = '' then
//  begin
//    //ShowMessage('Por favor, digite sua senha.');
//    edtSenha.SetFocus;
//    Exit;
//  end;
//
//  if (UpperCase(edtMatricula.Text) = 'ADMIN') and (edtSenha.Text = '123') then
//  begin
//    dmDados.UsuarioLogadoID := 0;
//    dmDados.UsuarioNome     := 'Administrador';
//    dmDados.UsuarioEhAdm    := True;
//
//    ModalResult := mrOk;
//    Exit;
//  end;
//
//  vSenhaHash := GerarMD5(edtSenha.Text);
//
//  with dmDados.qryEstoque do
//  begin
//    Close;
//    SQL.Clear;
//    SQL.Add('SELECT ID, LOGIN, ADM FROM USUARIOS');
//    SQL.Add('WHERE LOGIN = :User AND SENHA_HASH = :Pass AND ATIVO = ''S''');
//
//    Parameters.ParamByName('User').Value := edtMatricula.Text;
//    Parameters.ParamByName('Pass').Value := vSenhaHash;
//    Open;
//
//    if IsEmpty then
//    begin
//      ShowMessage('Usuário ou senha incorretos!');
//      edtSenha.Clear;
//      edtSenha.SetFocus;
//    end
//    else
//    begin
//      dmDados.UsuarioLogadoID := FieldByName('ID').AsInteger;
//      dmDados.UsuarioNome     := FieldByName('LOGIN').AsString;
//      dmDados.UsuarioEhAdm    := (FieldByName('ADM').AsString = 'S');
//
//      ModalResult := mrOk;
//    end;
//  end;
//end;

procedure TfrmLogin.btnFecharClick(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TfrmLogin.btnTrocaSenhaLoginClick(Sender: TObject);
begin
  // 1. Valida se o usuário preencheu os campos da tela de login
  if (Trim(edtMatricula.Text) = '') or (Trim(edtSenha.Text) = '') then
  begin
    ShowMessage('Para alterar a senha, primeiro digite sua Matricula e Senha ATUAL nos campos de login.');
    edtMatricula.SetFocus;
    Exit;
  end;

  // Verifica no Banco se essa credencial está correta
  dmDados.qryAux.Close;
  dmDados.qryAux.SQL.Text := 'SELECT ID FROM USUARIOS WHERE MATRICULA = :Matricula AND SENHA = :Senha';
  dmDados.qryAux.Parameters.ParamByName('Matricula').Value := edtMatricula.Text;
  dmDados.qryAux.Parameters.ParamByName('Senha').Value := edtSenha.Text;
  dmDados.qryAux.Open;

  if not dmDados.qryAux.IsEmpty then
  begin
    // Libera o painel de troca
    fIDUsuarioTroca := dmDados.qryAux.FieldByName('ID').AsInteger;
    
    // Limpa os campos do painel
    edtNovaSenhaLogin.Clear;
    edtConfirmaSenha.Clear;
    dtpExpiracao.Date := Date + 60; // Sugere 60 dias de validade

    // Mostra o painel centralizado ou onde você desenhou
    pnlTrocaSenhaLogin.Visible := True;
    pnlTrocaSenhaLogin.BringToFront;
    edtNovaSenhaLogin.SetFocus;
  end
  else
  begin
    ShowMessage('Credenciais inválidas! Você precisa digitar sua senha atual corretamente para poder trocá-la.');
  end;
end;

procedure TfrmLogin.edtSenhaKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    btnEntrar.Click;
  end;
end;

procedure TfrmLogin.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    btnEntrar.Click;
  end;
end;

procedure TfrmLogin.FormShow(Sender: TObject);
begin
  edtMatricula.Clear;
  edtSenha.Clear;

  if edtMatricula.CanFocus then
    edtMatricula.SetFocus;

  lblVersao.Caption := 'Versão: ' + ObterVersaoDoExecutavel;

  // 1. Carrega as empresas do banco
  dmDados.qryListaEmpresa.Close;
  dmDados.qryListaEmpresa.Open;

  // Se tiver apenas uma empresa (cenário comum), já seleciona ela sozinha
  if dmDados.qryListaEmpresa.RecordCount = 1 then
  begin
    // Seleciona o ID da primeira empresa encontrada
    dblkcbbEmpresa.KeyValue := dmDados.qryListaEmpresa.FieldByName('ID').AsInteger;
  end;
  
  // Código anterior de focar no usuario...
  edtMatricula.SetFocus;
end;

procedure TfrmLogin.edtMatriculaKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    btnEntrar.Click;
  end;
end;

procedure TfrmLogin.edtMatriculaKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['0'..'9', #8]) then
  begin
    Key := #0;
    Beep;
  end;

  if Key = #13 then
  begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
  
end;

procedure TfrmLogin.pnlMarcaMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  Perform(WM_SYSCOMMAND, $F012, 0);
end;

end.
