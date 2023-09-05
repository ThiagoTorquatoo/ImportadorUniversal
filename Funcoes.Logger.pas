unit Funcoes.Logger;

interface

type
  TLogger = class
  private
    FArquivoLog: TextFile;
    class var FInstancia: TLogger;
    constructor Create;
  public
    class function ObterInstancia: TLogger;
    class function NewInstance: TObject; override;
    procedure RegistrarLog(pTexto: string; pMetodo: string; pExcecao: string);
    destructor Destroy; override;
  end;

implementation

uses
  Forms, SysUtils;

{ TLogger }

constructor TLogger.Create;
var
  lDiretorioAplicacao: string;
begin
  lDiretorioAplicacao := ExtractFilePath(Application.ExeName) + 'Logs\' ;

  if not DirectoryExists(lDiretorioAplicacao) then
  begin
    ForceDirectories(lDiretorioAplicacao);
  end;

  AssignFile(FArquivoLog, lDiretorioAplicacao + 'LogImportadorUniversal.txt');

  if not FileExists(lDiretorioAplicacao + 'LogImportadorUniversal.txt') then
  begin
    Rewrite(FArquivoLog);
    CloseFile(FArquivoLog);
  end;
end;

destructor TLogger.Destroy;
begin
  FInstancia.Free;
  inherited;
end;

class function TLogger.NewInstance: TObject;
begin
  if not Assigned(FInstancia) then
    FInstancia := TLogger(inherited NewInstance);

  Result := FInstancia;
end;

class function TLogger.ObterInstancia: TLogger;
begin
  Result := TLogger.Create;
end;

procedure TLogger.RegistrarLog(pTexto: string; pMetodo: string; pExcecao: string);
var
  lDataHora: string;
begin
  Append(FArquivoLog);
  lDataHora := FormatDateTime('[dd/mm/yyyy hh:nn:ss] ', Now);
  WriteLn(FArquivoLog, 'Data/Hora...........: ' + lDataHora);
  WriteLn(FArquivoLog, 'Formulário..........: ' + pTexto);
  WriteLn(FArquivoLog, 'Metodo..............: ' + pMetodo);
  WriteLn(FArquivoLog, 'Erro................: ' + pExcecao);
  WriteLn(FArquivoLog, StringOfChar('-', 70));

  CloseFile(FArquivoLog);
end;

end.
