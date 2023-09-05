unit Datamodulo.Conexao;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.PG,
  FireDAC.Phys.PGDef, FireDAC.VCLUI.Wait, Vcl.StdCtrls, Data.DB,
  FireDAC.Comp.Client, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, Vcl.Grids, Vcl.DBGrids, 
  FireDAC.Phys.FB, FireDAC.Phys.FBDef, System.IniFiles, FireDAC.Phys.MySQLDef,
  FireDAC.Phys.MySQL, FireDAC.Phys.IBBase, Funcoes.Logger,
  FireDAC.Phys.OracleDef, FireDAC.Phys.Oracle;

type

  TDBConfINI = class(TComponent)
	private
    fLocalLIB: string;
    fDriverID: string;
		fHostName: string;
		fDataBase: string;
		fPorta: string;
    fPaswd: string;
    fUserName: string;
	public
		constructor Create(AOwner: TComponent); override;
		procedure LeArquivo;

    property LocalLIB: String read fLocalLIB;
    property DriverID: String read fDriverID;
	  property HostNameWhats: String read fHostName;
		property DataBaseWhats: String read fDataBase;
		property PortaWhatsApp: String read fPorta;
		property UserNameWhats: String read fUserName;
    property PaswdWhats: String read fPaswd;
	end;

  TDataModuleConexao = class(TDataModule)
    fdConexao: TFDConnection;
    qrCliente: TFDQuery;
    qrComprasItens: TFDQuery;
    dsClientes: TDataSource;
    dsCompras: TDataSource;
    dsComprasItens: TDataSource;
    FDPhysFBDriverLink: TFDPhysFBDriverLink;
    FDPhysMySQLDriverLink: TFDPhysMySQLDriverLink;
    FDPhysPgDriverLink: TFDPhysPgDriverLink;
    qrCompra: TFDQuery;
    FDPhysOracleDriverLink: TFDPhysOracleDriverLink;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
    fDBConfINI: TDBConfINI;

    procedure CriarConexao;
  public
    { Public declarations }
    Function GetDataSouceClientes:TDataSource;
    Function GetDataSouceCompras:TDataSource;
    Function GetDataSouceComprasItens:TDataSource;

    procedure OpenDadosClientes(pSql:String);
    procedure OpenDadosCompras(pSql:String);
    procedure OpenDadosComprasItens(pSql:String);

    procedure CloseDadosClientes;
    procedure CloseDadosCompras;
    procedure CloseDadosComprasItens;

    Procedure OpenBase(StringConexao:String);
  end;

var
  DataModuleConexao: TDataModuleConexao;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TDM }

procedure TDataModuleConexao.CloseDadosClientes;
begin
  qrCliente.Close;
end;

procedure TDataModuleConexao.CloseDadosCompras;
begin
  qrCompra.Close;
end;

procedure TDataModuleConexao.CloseDadosComprasItens;
begin
  qrComprasItens.Close;
end;

procedure TDataModuleConexao.CriarConexao;
begin
  try
    fDBConfINI.Free;
    fDBConfINI := TDBConfINI.Create(Self);
    if (Assigned(fDBConfINI)) then
    begin
      if fDBConfINI.fDriverID = 'FB' then
      begin
        FDPhysFBDriverLink.VendorLib := fDBConfINI.fLocalLIB;
      end else
      if fDBConfINI.fDriverID = 'MySQL' then
      begin
        FDPhysMySQLDriverLink.VendorLib := fDBConfINI.fLocalLIB;
      end else
      if fDBConfINI.fDriverID = 'Ora' then
      begin
        FDPhysOracleDriverLink.VendorLib := fDBConfINI.fLocalLIB;
      end else
      if fDBConfINI.fDriverID = 'PG' then
      begin
        FDPhysPGDriverLink.VendorLib := fDBConfINI.fLocalLIB;
      end;

      fdConexao.Connected := False;
      fdConexao.Params.Values['DriverID'] := fDBConfINI.fDriverID;
      fdConexao.Params.Values['Server'] := fDBConfINI.fHostName;
      fdConexao.Params.Values['Port'] := fDBConfINI.fPorta;
      fdConexao.Params.Values['Database'] := fDBConfINI.fDataBase;
      fdConexao.Params.Values['User_name'] := fDBConfINI.fUserName;
      fdConexao.Params.Values['Password'] := fDBConfINI.fPaswd;
      fdConexao.Connected := True;
    end;
  except
    on E: Exception do
    begin
      TLogger.ObterInstancia.RegistrarLog('DataModuleCreate',
                                          'CriarConexao - NÃO CONSEGUIU CONECTAR' +
                                          ' NA BASE PELO MOTIVO ABAIXO:',
                                          E.Message);
      Application.Terminate;
    end;
  end;
end;

procedure TDataModuleConexao.DataModuleCreate(Sender: TObject);
begin
  CriarConexao;
end;

function TDataModuleConexao.GetDataSouceClientes: TDataSource;
begin
  Result := dsClientes;
end;

function TDataModuleConexao.GetDataSouceCompras: TDataSource;
begin
  result := dsCompras;
end;

function TDataModuleConexao.GetDataSouceComprasItens: TDataSource;
begin
  result := dsComprasItens;
end;

procedure TDataModuleConexao.OpenBase(StringConexao: String);
begin
  fdConexao.Close;
  fdConexao.Open();
end;

procedure TDataModuleConexao.OpenDadosClientes(pSql: String);
begin
  qrCliente.close;
  qrCliente.SQL.clear;

  //if Pos('WHERE',UpperCase(pSql)) <= 0 then
  // pSql := pSql + ' WHERE 1=2 ';

  qrCliente.SQL.Text := pSql;
  qrCliente.Open;
end;

procedure TDataModuleConexao.OpenDadosCompras(pSql: String);
begin
  qrCompra.Close;
  qrCompra.SQL.Text := pSql;
  qrCompra.Open;
end;

procedure TDataModuleConexao.OpenDadosComprasItens(pSql: String);
begin
  qrComprasItens.Close;
  qrComprasItens.SQL.Text := pSql;
  qrComprasItens.Open;
end;

{ TDBConfINI }

constructor TDBConfINI.Create(AOwner: TComponent);
begin
  inherited;
  LeArquivo;
end;

procedure TDBConfINI.LeArquivo;
var
	lExiste: Boolean;
	lCaminhoIni: String;
	lIni: TInifile;
	lValor: String;
begin
	lCaminhoIni := ExtractFilePath(Application.ExeName) + 'conexaoimportador.ini';
	lExiste := FileExists(lCaminhoIni);

	if not lExiste then
	begin
    raise Exception.Create('Não foi possível localizar o arquivo de ' +
                           'informações do banco de dados. ' +
                           'A aplicação será fechada.');
		Application.Terminate;
	end;

	lIni := TInifile.Create(lCaminhoIni);
	try
    lValor := lIni.ReadString('LOCALLIB', 'locallib', '');
    if (lValor <> '') then
			fLocalLIB := lValor;
    lValor := lIni.ReadString('DRIVERID', 'driverid', '');
		if (lValor <> '') then
			fDriverID := lValor;
    lValor := lIni.ReadString('HOSTNAME', 'hostname', '');
		if (lValor <> '') then
			fHostName := lValor;
    lValor := lIni.ReadString('DATABASE', 'database', '');
		if (lValor <> '') then
			fDataBase := lValor;
    lValor := lIni.ReadString('PORTA', 'porta', '');
		if (lValor <> '') then
			fPorta := lValor;
		lValor := lIni.ReadString('USUARIO', 'username', '');
		if (lValor <> '') then
			fUserName := lValor;
    lValor := lIni.ReadString('SENHA', 'password', '');
		if (lValor <> '') then
			fPaswd := lValor;
	finally
		lIni.Free;
	end;
end;

end.
