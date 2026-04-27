program Project1;

uses
  ExceptionLog,
  Forms,
  Controls,
  uDmDados in 'uDmDados.pas' {dmDados: TDataModule},
  uLogin in 'uLogin.pas' {frmLogin},
  uPrincipal in 'uPrincipal.pas' {frmPrincipal},
  uAutenticacao in 'uAutenticacao.pas' {frmAutenticacao};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'WOLF ERP';
  Application.CreateForm(TdmDados, dmDados);
  Application.CreateForm(TfrmLogin, frmLogin);
  Application.CreateForm(TfrmAutenticacao, frmAutenticacao);
  frmLogin := TfrmLogin.Create(nil);
  try
    if frmLogin.ShowModal = mrOk then
    begin
      frmLogin.Free;
      Application.CreateForm(TfrmPrincipal, frmPrincipal);
      Application.Run;
    end
    else
    begin
      Application.Terminate;
    end;
  finally
//    frmLogin.Free;
  end;
end.
