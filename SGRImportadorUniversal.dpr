program SGRImportadorUniversal;

uses
  MidasLib,
  Vcl.Forms,
  Datamodulo.Conexao in 'Datamodulo.Conexao.pas' {DataModuleConexao: TDataModule},
  uImportarDireto in 'uImportarDireto.pas' {frmImportarDireto},
  uAjuda in 'uAjuda.pas' {frmAjuda},
  Funcoes.Logger in 'Funcoes.Logger.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmImportarDireto, frmImportarDireto);
  Application.Title := 'in.Pulse - Importação de clientes e compras';
  Application.Run;
end.
