unit uImportarDireto;

interface

uses
  Winapi.Windows,Winapi.Messages,
  VCL.forms, System.SysUtils, VCL.dialogs, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.StdCtrls, Vcl.DBCtrls, Vcl.Controls, Vcl.Mask, Vcl.Grids, Vcl.DBGrids,
  Vcl.Buttons, System.Classes, Datasnap.DBClient, Data.DB, Data.Win.ADODB,
  Variants, uAjuda, System.IniFiles, Datamodulo.Conexao, System.UITypes,
  DBXJSON, JSON, REST.Types, IPPeerClient, REST.Client, Data.Bind.Components,
  Data.Bind.ObjectScope, IdTCPConnection, IdTCPClient, IdHTTP, IdBaseComponent,
  IdComponent, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL,
  IdSSLOpenSSL;

type
  TfrmImportarDireto = class(TForm)
    pb: TProgressBar;
    cdsConfig: TClientDataSet;
    cdsConfigSTRINGCONEXAO: TStringField;
    cdsConfigQUEBRAR_ARQUIVO: TStringField;
    dsConfig: TDataSource;
    cdsConfigCLIENTE_SQL: TMemoField;
    cdsConfigCOMPRA_SQL: TMemoField;
    qrUpdateCliente: TADOQuery;
    qrInsertCliente: TADOQuery;
    qrInsertCompra: TADOQuery;
    conexaoSGR: TADOConnection;
    qrInsertCompraItem: TADOQuery;
    cdsConfigCOMPRA_IT_SQL: TMemoField;
    cdsConfigATUALIZACAO_HORA: TTimeField;
    cdsConfigATUALIZACAO_TODODIA: TStringField;
    cdsConfigDESLIGA_COMPUTADOR: TStringField;
    cdsConfigATUALIZACAO_DATAULTCOMPRA: TStringField;
    cdsConfigVALIDA_CNPJ: TStringField;
    cdsConfigVALIDA_ERP: TStringField;
    cdsConfigLIMPA_COMPRAS: TStringField;
    qrPesquisaCli: TADOQuery;
    StatusBar1: TStatusBar;
    qrUpdateClienteBackup: TADOQuery;
    tblParametros: TADOTable;
    dsParametros: TDataSource;
    cdsConfigIMPORTAR_CLIENTES: TStringField;
    cdsConfigIMPORTAR_COMPRAS: TStringField;
    cdsConfigIMPORTAR_CLIENTE_EXISTENTE: TStringField;
    cdsConfigCriar_Agenda_ComprasDIA: TBooleanField;
    cdsConfigINAT_CLI_OPER: TStringField;
    cdsConfigOPER_INAT_CLI: TStringField;
    cdsConfigSOBRESCREVER_COMPRAS: TStringField;
    cdsConfigUTILIZAHOLDING: TStringField;
    pnlBotoes: TPanel;
    Label1: TLabel;
    lbClienteTotal: TLabel;
    Label3: TLabel;
    lbTotalClientes: TLabel;
    btnAjuda: TSpeedButton;
    btnImportar: TBitBtn;
    btnTrayIcon: TBitBtn;
    btnSalvar: TBitBtn;
    pgc: TPageControl;
    tsImportacao: TTabSheet;
    grp4: TGroupBox;
    grdClientes: TDBGrid;
    tsCliente: TTabSheet;
    grp2: TGroupBox;
    DBMemo1: TDBMemo;
    tsCompras: TTabSheet;
    grp3: TGroupBox;
    DBMemo2: TDBMemo;
    tsCompasItem: TTabSheet;
    grp5: TGroupBox;
    dbmmo1: TDBMemo;
    tsLog: TTabSheet;
    log: TMemo;
    TabSheet1: TTabSheet;
    grpCampos: TGroupBox;
    chkRAZAO: TDBCheckBox;
    chkFantasia: TDBCheckBox;
    chkRAZAO2: TDBCheckBox;
    chkRAZAO3: TDBCheckBox;
    chkRAZAO4: TDBCheckBox;
    chkRAZAO5: TDBCheckBox;
    chkRAZAO6: TDBCheckBox;
    chkRAZAO7: TDBCheckBox;
    chkRAZAO8: TDBCheckBox;
    chkRAZAO9: TDBCheckBox;
    chkRAZAO10: TDBCheckBox;
    chkRAZAO11: TDBCheckBox;
    chkRAZAO12: TDBCheckBox;
    chkRAZAO13: TDBCheckBox;
    chkRAZAO14: TDBCheckBox;
    chkRAZAO15: TDBCheckBox;
    chkRAZAO16: TDBCheckBox;
    chkRAZAO17: TDBCheckBox;
    DBCheckBox1: TDBCheckBox;
    chkIncluiClienteCampanha: TDBCheckBox;
    DBCheckBox2: TDBCheckBox;
    DBCheckBox3: TDBCheckBox;
    DBCheckBox5: TDBCheckBox;
    DBCheckBox4: TDBCheckBox;
    dbcAtivarDesativarClientes: TDBCheckBox;
    ckClienteInativoOper: TDBCheckBox;
    edOperadorClienteInativo: TDBEdit;
    DBCheckBox8: TDBCheckBox;
    dbcUtilizaHolding: TDBCheckBox;
    dbcImportarMarcas: TDBCheckBox;
    DBCheckBox7: TDBCheckBox;
    tsBanco: TTabSheet;
    grp1: TGroupBox;
    edStringConexao: TDBMemo;
    grp6: TGroupBox;
    lbAtualizacaoHora: TLabel;
    lbAtualizacaoDia: TLabel;
    lbDataUltCOmpra: TLabel;
    edAtualizacaoHora: TDBEdit;
    cbAtualizacaoDia: TComboBox;
    dbchkDesligaComputador: TDBCheckBox;
    chkValidaCNPJ: TDBCheckBox;
    chkValidaERP: TDBCheckBox;
    cbxDataUltimaCompra: TComboBox;
    chkLimpaBanco: TDBCheckBox;
    chkAtualizarObservacoesCliente: TDBCheckBox;
    cdsConfigATUALIZAR_OBS_CLIENTE_IMPORT: TStringField;
    tbsAPI: TTabSheet;
    chkUtilizaAPI: TDBCheckBox;
    cdsConfigUTILIZA_API: TStringField;
    edtURL: TDBEdit;
    lblUrl: TLabel;
    lblUsuario: TLabel;
    edtUsuario: TDBEdit;
    lblSenha: TLabel;
    edtSenha: TDBEdit;
    cdsConfigURL_API: TStringField;
    cdsConfigUSUARIO_API: TStringField;
    cdsConfigSENHA_API: TStringField;
    idSSL: TIdSSLIOHandlerSocketOpenSSL;
    IdHTTP: TIdHTTP;
    RESTClient: TRESTClient;
    RESTResponse: TRESTResponse;
    RESTRequest: TRESTRequest;
    procedure btnConectarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure trayIconClick(Sender: TObject);
    procedure btnTrayIconClick(Sender: TObject);
    procedure tmrImportacaoTimer(Sender: TObject);
    procedure btnSalvarClick(Sender: TObject);
    procedure cbAtualizacaoDiaChange(Sender: TObject);
    procedure cbxDataUltimaCompraChange(Sender: TObject);
    procedure btnAjudaClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure pgcChange(Sender: TObject);
    procedure Importar;
    procedure ObterDadosAPI;
  private
    CodigoCliente: Integer;
    pArquivoConexao, pCaminhoPrograma, pATUALIZACAO_HORA, pATUALIZACAO_TODODIA, pDESLIGA_COMPUTADOR: string;
    FFecharAutomaticamente: Boolean;
    bExecucaoBatch : Boolean;
    FDatataModuloConexao: TDataModuleConexao;

    property FecharAutomaticamente: Boolean read FFecharAutomaticamente
      write FFecharAutomaticamente;

    procedure Criar_Agenda_ComprasDIA;
    function conectar(StringConexao: string): Boolean;
    function insertOrUpdateCliente(AQuery: TADOQuery; COD_ERP, RAZAO, FANTASIA,
      CPF_CNPJ, END_RUA, CIDADE, BAIRRO, ESTADO, CEP, FONE1, EMAIL, OPERADOR,
      COD_UNIDADE, SALDO_DISPONIVEL, POTENCIAL, DATA_ULT_COMPRA, SEGMENTO,
      OBS_CLIENTES, OBS_ADMIN, ATIVO, COD_ERP_HOLDING,
      VENCIMENTO_LIMITE_CREDITO: string): string;
    function existeCliente(CLIENTE, CPF_CNPJ: string): Boolean;
    function retornaCampo(campo: string): string;
    function CalculaCnpjCpf(Numero: string): Boolean;
    function RetornaUnidade(Valor: string): string;
    function RetornaSegmento(Valor: string): string;
    function RetornaOperador(Valor: string): string; overload;
    function RetornaOperador(Valor: string; COD_ERP_CLIENTE: string;
                             ANomeQuery: string): string; overload; //
    function insertCompra(CLIENTE, Data, Valor, DESCRICAO, FORMA_PGTO: string;
      OPERADOR: String = '' ; TIPO : STring = ''; SITUACAO: string = ''): string;
    procedure insertCompraItem(NOTA, CODPROD, DESCRICAO, Qtd, UN_MEDIDA, VALOR_UN, DESCONTO: string);
    procedure atualizaParametros;
    function retornaSoNumero(Valor: string): string;
    procedure AlteraDataUltimaCompra;
    procedure showMessageDesenv(texto: string);
    procedure organizaCampanhas;
    procedure insereListaFones;
    procedure limpaComprasItensCompras;
      // procedure RepoeLigacoesFidelizadas;
    function refidelizaCotados: Boolean;
    function GetFileVersion(const FileName: string): string;
    function vazio(texto: string): Boolean;
    procedure AtualizaDataCompra;
    function ExecSql(xsql: string; Tipo: Integer = 0): TADOQuery;
    procedure SepararDDD(Fone: string; out DDD: string; out Telefone: string);
    function ObterLogDataSet(oDataSet: TDataSet): string;
    function ObterValorDataSet(oDataSet: TDataSet; const pPosicao: Integer): string; overload;
    function ObterValorDataSet(oDataSet: TDataSet; const pNome: String): string; overload;
    procedure SetValorParametro(oQuery: TADOQuery; const pPosicao: integer; const pValue: String);
    procedure AlteraDataUltimaCompraCNPJ;
    procedure InsereLog(pTexto: String; pTipo: TMsgDlgType = mtInformation);
    procedure atualizarClienteHolding(COD_ERP, COD_ERP_HOLDING, DATA_ULT_COMPRA,
      OPERADOR: string);
    function Iif(Condicao: Boolean; Verdadeiro, Falso: Variant): Variant;
    procedure insertOrUpdateMarcaCliente(CLIENTE, COD_ERP, SEGMENTO, CODMARCA,
      OPERADORM, DATAULTIMACOMPRAM, UNIDADE: string);
     procedure SobrescreverCompra(const pCliente, pDescricaoCompra: string);
    procedure DeletaCompra(pCompra: Integer);
    procedure PopularcdsConfig(const psArquivo: string);
    procedure DefinirMidia(poQuery: TADOQuery);
    function GetPrioridadeAgenda(poQuery: TADOQuery): string;
    function GetCampanha(poQuery: TADOQuery): string;
    function GetOPERADOR(poQuery: TADOQuery; const psOperador: string): string;
    function getUtilizaRegua: Boolean;
    function fParametrosSistema: String;
    procedure ConfigurarAjuda;
    procedure ExecutarFuncao(AFuncao: string);
  public
      { Public declarations }
    procedure ExecutaFormShow(pExecucaoBatch: Boolean);
  end;

var
  frmImportarDireto: TfrmImportarDireto;

implementation

{$R *.dfm}

procedure TfrmImportarDireto.AlteraDataUltimaCompra;
var
  qrUltCompra: TADOQuery;
begin
  try
    try
      if tblParametros.FieldByName('ATUALIZA_DATA_PREFIXO_CNPJ').AsString = 'S' then
      begin
        AlteraDataUltimaCompraCNPJ;
      end
      else
      begin
        qrUltCompra := TADOQuery.Create(nil);
        qrUltCompra.Connection := conexaoSGR;
        qrUltCompra.SQL.Clear;
        qrUltCompra.SQL.Add('UPDATE ');
        qrUltCompra.SQL.Add('	clientes cli');
        qrUltCompra.SQL.Add('SET cli.DATA_ULT_COMPRA = ');
        qrUltCompra.SQL.Add('   (');
        qrUltCompra.SQL.Add('	SELECT');
        qrUltCompra.SQL.Add('		MAX(c.DATA)');
        qrUltCompra.SQL.Add('	FROM');
        qrUltCompra.SQL.Add('		compras c');
        qrUltCompra.SQL.Add('	WHERE');
        qrUltCompra.SQL.Add('		c.CLIENTE = cli.CODIGO');
        qrUltCompra.SQL.Add('   )');
        qrUltCompra.ExecSQL; // ATUALIZA DATA ULTIMA COMPRA PELO ARQUIVO DE COMPRAS
      end
    except
      on e: Exception do
      begin
        InsereLog('Erro ao atualizar datas últimas compras dos clientes. - ' +
                      e.ClassName + ' - ' + e.message, mtError);
        InsereLog('Comando: ' + qrUltCompra.Sql.Text);
      end;
    end;
  finally
    FreeAndNil(qrUltCompra);
  end;
end;

procedure TfrmImportarDireto.AtualizaDataCompra ;
begin
  with TADOQuery.Create(nil) do
  try
    Connection := conexaoSGR;
    SQL.Add('UPDATE clientes');
    SQL.Add('SET DATA_ULT_COMPRA = NULL');
    SQL.Add('WHERE DATA_ULT_COMPRA <= ''1900-01-01''');
    ExecSQL;
  finally
    Free;
  end;
end;

procedure TfrmImportarDireto.atualizaParametros;
begin
  pATUALIZACAO_HORA := cdsConfigATUALIZACAO_HORA.AsString;
  pATUALIZACAO_TODODIA := cdsConfigATUALIZACAO_TODODIA.AsString;
  pDESLIGA_COMPUTADOR := cdsConfigDESLIGA_COMPUTADOR.AsString;
end;

procedure TfrmImportarDireto.btnAjudaClick(Sender: TObject);
var
  oAJuda: tstrings;
const
  AJUDA_CLIENTES = 'Ordem dos campos da importacão de Clientes: ' + sLineBreak +  sLineBreak +
    '0  COD_ERP  ' + sLineBreak +
    '1  RAZAO  ' + sLineBreak +
    '2  FANTASIA  ' + sLineBreak +
    '3  CPF_CNPJ   ' + sLineBreak +
    '4  END_RUA  ' + sLineBreak +
    '5  CIDADE   ' + sLineBreak +
    '6  BAIRRO  ' + sLineBreak +
    '7  ESTADO  ' + sLineBreak +
    '8  CEP  ' + sLineBreak +
    '9  FONE1 ' + sLineBreak +
    '10 EMAIL   ' + sLineBreak +
    '11 OPERADOR  ' + sLineBreak +
    '12 COD_UNIDADE   ' + sLineBreak +
    '13 SALDO_DISPONIVEL    ' + sLineBreak +
    '14 POTENCIAL  ' + sLineBreak +
    '15 DATA_ULT_COMPRA  ' + sLineBreak +
    '16 SEGMENTO  ' + sLineBreak +
    '17 OBS_CLIENTES   ' + sLineBreak +
    '18 ATIVO    ' + sLineBreak +
    '19 HOLDING     ' + sLineBreak +
    '20 COD_MIDIA   ' + sLineBreak +
//      '20 CODIGO MARCA ' + sLineBreak +
//      '21 OPERADOR MARCA ' + sLineBreak +
//      '22 DATA ULTIMA COMPRA MARCA ' + sLineBreak +
//      '23 UNIDADE MARCA ' + sLineBreak + sLineBreak +
    ' - Para campos não localizados na view do cliente, informar NULL na coluna referente.';

  AJUDA_COMPRAS = 'Ordem dos campos da importacão de Compras: ' + sLineBreak +  sLineBreak +
    '0 COD_ERP_COMPRA ' + sLineBreak +
    '1 COD_ERP_CLIENTE' + sLineBreak +
    '2 DATA ' + sLineBreak +
    '3 DESCRICAO ' + sLineBreak +
    '4 VALOR TOTAL ' + sLineBreak +
    '5 COD OPERADOR' + sLineBreak +
    '6 TIPO {NF = NOTA FISCAL | PD = PEDIDO}' + sLineBreak +
    '7 SITUACAO {C = CANCELADO | F = FATURADO}' + sLineBreak +
    '8 FORMA PAGTO ' + sLineBreak +
    '9 NUMERONOTA ' + sLineBreak +
    ' - Os campos COD OPERADOR, TIPO e SITUAÇAO sao de uso da regua.';

  AJUDA_ITENS_COMPRAS = 'Ordem dos campos da importacão de Itens de Compras: ' + sLineBreak +  sLineBreak +
     '0 COD_ERP_ITEM ' + sLineBreak +
     '1 COD_ERP_COMPRA' + sLineBreak +
     '2 DESCRICAO ' + sLineBreak +
     '3 QUANTIDADE ' + sLineBreak +
     '5 UNIDADE ' + sLineBreak +
     '6 VALOR UNITARIO ' + sLineBreak +
     '7 DESCONTO ';
begin
  oAJuda := TStringList.Create;
  frmAjuda := TfrmAjuda.Create(nil);

  try
    if pgc.ActivePage = tsCliente then
    begin
      oAJuda.Text := AJUDA_CLIENTES
    end else
    if pgc.ActivePage = tsCompras then
    begin
      oAJuda.Text := AJUDA_COMPRAS;
    end else
    begin
      oAJuda.Text := AJUDA_ITENS_COMPRAS;
    end;

    frmAjuda.showDlg(oAJuda);

  finally
    FreeAndNil(frmAjuda);
    FreeAndNil(oAJuda);
  end;
  //
end;

procedure TfrmImportarDireto.btnConectarClick(Sender: TObject);
var
  aMsg: string;
begin
  aMsg := '';


  if conectar(edStringConexao.Text) then
    aMsg := 'CONEXÃO COM BANCO: OK'
  else
    Exit;

  try
//    FDatataModuloConexao.OpenDadosClientes(cdsConfigCLIENTE_SQL.AsString);
    FDatataModuloConexao.qrCliente.SQL.Clear;
    FDatataModuloConexao.qrCompra.SQL.Clear;
    FDatataModuloConexao.qrComprasItens.SQL.Clear;
    FDatataModuloConexao.qrCliente.SQL.Text :=
      cdsConfigCLIENTE_SQL.AsString + ' WHERE 1=2 ';
    FDatataModuloConexao.qrCompra.SQL.Text := cdsConfigCOMPRA_SQL.AsString;
    FDatataModuloConexao.qrComprasItens.SQL.Text :=
      cdsConfigCOMPRA_IT_SQL.AsString;

    aMsg := aMsg + #13 + 'CLIENTES: OK';
  except
    on e: Exception do
    begin
      aMsg := aMsg + #13 + 'CLIENTES: ERRO ' + e.ClassName + ' ' + e.Message;
    end;
  end;
  try
    FDatataModuloConexao.OpenDadosCompras(StringReplace(cdsConfigCOMPRA_SQL.AsString, '/*CODIGO*/', QuotedStr('-1'), [rfReplaceAll, rfIgnoreCase]));
    FDatataModuloConexao.CloseDadosCompras;
    aMsg := aMsg + #13 + 'COMPRAS: OK';
  except
    on e: Exception do
    begin
      aMsg := aMsg + #13 + 'COMPRAS: ERRO ' + e.ClassName + ' ' + e.Message;
    end;
  end;

  try
    FDatataModuloConexao.OpenDadosComprasItens(StringReplace(cdsConfigCOMPRA_IT_SQL.AsString, '/*CODIGO*/', QuotedStr('-1'), [rfReplaceAll, rfIgnoreCase]));
    FDatataModuloConexao.CloseDadosComprasItens;
    aMsg := aMsg + #13 + 'ITENS: OK';
  except
    on e: Exception do
    begin
      aMsg := aMsg + #13 + 'ITENS: ERRO ' + e.ClassName + ' ' + e.Message;
    end;
  end;

  if not vazio(aMsg) then
    ShowMessage(aMsg);
end;

procedure TfrmImportarDireto.btnImportarClick(Sender: TObject);
begin
  Importar;
end;

procedure TfrmImportarDireto.btnSalvarClick(Sender: TObject);
begin
  try
    if cdsConfig.State in ([dsEdit, dsInsert]) then
      cdsConfig.Post;

    cdsConfig.SaveToFile(pCaminhoPrograma + pArquivoConexao, dfXML);

    atualizaParametros;
    cdsConfig.Edit;
  except
    on e: Exception do
    begin
      ShowMessage('Problema ao abrir os parametros' + #13 + e.ClassName + #13 + e.Message);
      Abort;
    end;
  end;
end;

procedure TfrmImportarDireto.btnTrayIconClick(Sender: TObject);
begin
  frmImportarDireto.Hide;
end;

function TfrmImportarDireto.GetFileVersion(const FileName: string): string;
var
  Major, Minor, Release, Build: Integer;
  Zero: DWORD; // set to 0 by GetFileVersionInfoSize
  VersionInfoSize: DWORD;
  PVersionData: pointer;
  PFixedFileInfo: PVSFixedFileInfo;
  FixedFileInfoLength: UINT;

begin
  Result := '0';
  Major := 0;
  Minor := 0;
  Release := 0;
  Build := 0;
  VersionInfoSize := GetFileVersionInfoSize(pChar(FileName), Zero);
  if VersionInfoSize = 0 then
    Exit;
  PVersionData := AllocMem(VersionInfoSize);
  try
    if GetFileVersionInfo(pChar(FileName), 0, VersionInfoSize, PVersionData) = False then
      Exit;
    if VerQueryValue(PVersionData, '', pointer(PFixedFileInfo), FixedFileInfoLength) = False then
      Exit;
    Major := PFixedFileInfo^.dwFileVersionMS shr 16;
    Minor := PFixedFileInfo^.dwFileVersionMS and $FFFF;
    Release := PFixedFileInfo^.dwFileVersionLS shr 16;
    Build := PFixedFileInfo^.dwFileVersionLS and $FFFF;
  finally
    FreeMem(PVersionData);
  end;

  if (Major or Minor or Release or Build) <> 0 then
  begin
    Result := 'Versão: ' + IntToStr(Minor) + '.' +
      IntToStr(Release) + '.' + IntToStr(Build);
  end;
   // Result := IntToStr(Major) +'.'+ IntToStr(Minor) +'.'+ IntToStr(Release) +'.'+ IntToStr(Build);
end;

procedure TfrmImportarDireto.cbxDataUltimaCompraChange(Sender: TObject);
begin
  cdsConfigATUALIZACAO_DATAULTCOMPRA.AsInteger := cbxDataUltimaCompra.ItemIndex;
end;

function TfrmImportarDireto.conectar(StringConexao: string): Boolean;
begin
  Result := True;
  try
    FDatataModuloConexao.OpenBase(StringConexao);
  except
    on e: Exception do
    begin
      Result := False;
      ShowMessage(E.Message);
    end;
  end;
end;

procedure TfrmImportarDireto.ConfigurarAjuda;
begin
  if (pgc.ActivePage = tsCliente) or (pgc.ActivePage = tsCompras) or
   (pgc.ActivePage = tsCompasItem) then
  begin
    btnAjuda.Visible := true;
  end else
  begin
    btnAjuda.Visible := false;
  end;
end;

procedure TfrmImportarDireto.cbAtualizacaoDiaChange(Sender: TObject);
begin
  cdsConfigATUALIZACAO_TODODIA.AsInteger := cbAtualizacaoDia.ItemIndex;
end;

function TfrmImportarDireto.existeCliente(CLIENTE, CPF_CNPJ: string): Boolean;
begin
  with TADOQuery.Create(nil) do
  try
    Connection := conexaoSGR;
    SQL.Add('SELECT CODIGO FROM clientes');
    SQL.Add('WHERE COD_ERP = :COD_ERP');

    if tblParametros.FieldByName('IMP_POR_CPFNCPJ_CLIENTE').AsString = 'S' then
    begin
      if not vazio(CPF_CNPJ) then
      begin
        SQL.Add('OR CPF_CNPJ = :CPF_CNPJ');
        Parameters[1].Value := trim(CPF_CNPJ);
      end;
    end;

    Parameters[0].Value := trim(CLIENTE);
    Open;
    Result := not IsEmpty;

    CodigoCliente := fieldbyname('CODIGO').AsInteger;
  finally
    Free;
  end;
end;

procedure TfrmImportarDireto.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  if tblParametros.State in [dsedit, dsinsert] then
    tblParametros.Post;

  if cdsConfig.State in ([dsEdit, dsInsert]) then
    cdsConfig.Post;

  cdsConfig.SaveToFile(pCaminhoPrograma + pArquivoConexao, dfXML);
  Application.Terminate;
end;

procedure TfrmImportarDireto.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  Canclose := True;
  if not bExecucaoBatch then
  begin
    case MessageDlg('Deseja salvar os parâmetros informados?', mtConfirmation, mbYesNoCancel, 0) of
      mrYes : btnSalvar.OnClick(btnSalvar);
      mrCancel : CanClose := False;
    end;
  end;
end;

procedure TfrmImportarDireto.FormCreate(Sender: TObject);
begin
  FDatataModuloConexao := TDataModuleConexao.Create(Self);

  Self.Caption := Self.Caption;

  if (fParametrosSistema() = '') then
  begin
    Application.Terminate;
    Exit;
  end;

  pArquivoConexao := 'MANUAL';

  pCaminhoPrograma := ExtractFilePath(Application.ExeName);
  conexaoSGR.Connected := False;
  try
    conexaoSGR.ConnectionString := 'Provider=MSDASQL.1;Persist Security Info=False;Data Source=crm_sgr';
    conexaoSGR.Connected := True;
  except
    on e: Exception do
    begin
      ShowMessage('Problema ao conectar com o SGR' + #13 + e.ClassName + #13 + e.Message);
      Application.Terminate;
    end;
  end;

  StatusBar1.Panels[0].Text := GetFileVersion(Application.ExeName);

  if FecharAutomaticamente then
  begin
    frmImportarDireto.ExecutaFormShow(False);
    btnImportarClick(frmImportarDireto.btnImportar);
    Application.Terminate;
  end;
end;

procedure TfrmImportarDireto.FormShow(Sender: TObject);
begin
  ExecutaFormShow(False);
  ConfigurarAjuda;
end;

function TfrmImportarDireto.fParametrosSistema: String;
var
	lI: SmallInt;
	lAuxResult: String;
begin
	for lI := 1 to ParamCount do
	begin
//    if ParamStr(lI) = 'MANUAL' then
//    begin
//
//    end;

    if ParamStr(lI) = 'FA' then
    begin
      //FA = Fechar automaticamente aplicação
      FecharAutomaticamente := True;
    end;

		lAuxResult := lAuxResult + ParamStr(lI);
	end;

	Result := lAuxResult;
end;

procedure TfrmImportarDireto.insereListaFones;
begin
  try
    conexaoSGR.Execute('CALL PR_INSERE_FONES_CAMPANHAS()');
  except
    on e: Exception do
    begin
      log.Lines.Add('Problemas ao inserir lista de fones' + #13 + e.ClassName + #13 + e.Message + #13);
    end;
  end;
end;

function TfrmImportarDireto.insertCompra(CLIENTE, Data, Valor, DESCRICAO, FORMA_PGTO: string;
  OPERADOR: String = '' ; TIPO : STring = ''; SITUACAO: string = ''): string;
var
  xValor: string;
begin
  Result := '';

  with qrInsertCompra do
  try
    xValor := StringReplace(Valor, ',', '.', []);
    Close;
    SetValorParametro(qrInsertCompra, 0, Trim(Cliente));
    SetValorParametro(qrInsertCompra, 1, Trim(Data));
    SetValorParametro(qrInsertCompra, 2, Trim(xValor));
    SetValorParametro(qrInsertCompra, 3, Trim(DESCRICAO));
    SetValorParametro(qrInsertCompra, 4, Trim(FORMA_PGTO));

    if (tblParametros.FieldByName('IMPORTAR_OPERADOR_COMPRA').AsString = 'S') and
      (trim(OPERADOR) <> '') then
    begin
      SetValorParametro(qrInsertCompra, 5, RetornaOperador(OPERADOR));
    end;

    if getUtilizaRegua then
    begin
      SetValorParametro(qrInsertCompra, 6, Trim(TIPO));
      SetValorParametro(qrInsertCompra, 7, Trim(SITUACAO));
    end;

    qrInsertCompra.Parameters.ParamByName('FATURADO').Value := 'S';

    ExecSQL;

    with TADOQuery.Create(nil) do
    try
      Connection := conexaoSGR;
      SQL.Add('SELECT MAX(CODIGO) FROM compras WHERE CLIENTE = ' + QuotedStr(CLIENTE));
      Open;
      Result := Fields[0].AsString;
    finally
      Free;
    end;
  except
    on e: Exception do
    begin
      log.Lines.Add('Problemas ao inserir compras' + #13 + e.ClassName + #13 +
                    e.Message + #13);
    end;
  end;
end;

procedure TfrmImportarDireto.insertCompraItem(NOTA, CODPROD, DESCRICAO, Qtd, UN_MEDIDA, VALOR_UN, DESCONTO: string);
var
  xQtd, xVALOR_UN, xDESCONTO: string;
begin
  with qrInsertCompraItem do
  try
    xQtd := StringReplace(Qtd, ',', '.', []);
    xVALOR_UN := StringReplace(VALOR_UN, ',', '.', []);
    xDESCONTO := StringReplace(DESCONTO, ',', '.', []);

    if xDESCONTO = EmptyStr then
    begin
      xDESCONTO := '0';
    end;

    if xVALOR_UN = EmptyStr then
    begin
      xVALOR_UN := '0';
    end;

    Close;
    SetValorParametro(qrInsertCompraItem, 0, Trim(NOTA));
  	SetValorParametro(qrInsertCompraItem, 1, Trim(CODPROD));
  	SetValorParametro(qrInsertCompraItem, 2, Trim(DESCRICAO));
  	SetValorParametro(qrInsertCompraItem, 3, Trim(xQtd));
  	SetValorParametro(qrInsertCompraItem, 4, Trim(UN_MEDIDA));
	  SetValorParametro(qrInsertCompraItem, 5, Trim(xVALOR_UN));
  	SetValorParametro(qrInsertCompraItem, 6, Trim(xDESCONTO));
    ExecSQL;
  except
    on e: Exception do
    begin
      log.Lines.Add('Problemas ao inserir itens' + #13 + e.ClassName + #13 + e.Message + #13);
    end;
  end;
end;

function TfrmImportarDireto.insertOrUpdateCliente(AQuery: TADOQuery; COD_ERP,
  RAZAO, FANTASIA, CPF_CNPJ, END_RUA, CIDADE, BAIRRO, ESTADO, CEP, FONE1, EMAIL,
  OPERADOR, COD_UNIDADE, SALDO_DISPONIVEL, POTENCIAL, DATA_ULT_COMPRA, SEGMENTO,
  OBS_CLIENTES, OBS_ADMIN, ATIVO, COD_ERP_HOLDING,
  VENCIMENTO_LIMITE_CREDITO: string): string;
var
  AREA1: string;
  CodigoGrupo: Integer;
  lQueryCampanha: TADOQuery;
  lQueryInsertCampanha: TADOQuery;
  lObservacao: string;
begin
  CodigoGrupo := 0;
  Result := '';

  qrPesquisaCli.Close;
  qrPesquisaCli.SQL.Text := 'SELECT RAZAO, FANTASIA,	' +
    ' CAST(CONCAT(AREA1, FONE1) AS CHARACTER(15)) AS FONE1, 	CPF_CNPJ, ' +
    ' END_RUA, ATIVO,	BAIRRO,	CIDADE,	CEP,	ESTADO,	EMAIL,	SEGMENTO,' +
    ' OBS_ADMIN,SALDO_DISPONIVEL,	POTENCIAL FROM clientes ';

  if CodigoCliente > 0 then
  begin
    qrPesquisaCli.SQL.add(' WHERE CODIGO = ' + IntToStr(CodigoCliente))
  end else
  begin
    qrPesquisaCli.SQL.add(' WHERE COD_ERP = ' +
      QuotedStr(FDatataModuloConexao.GetDataSouceClientes.DataSet.Fields[0].AsString));
  end;

  qrPesquisaCli.Open;

  if (AQuery.Name = 'qrUpdateCliente') and
    ((cdsConfigIMPORTAR_CLIENTES.AsString = 'N')
      and (cdsConfigIMPORTAR_CLIENTE_EXISTENTE.AsString = 'N')) then
  begin
    Result := retornaCampo('CODIGO');
    Exit;
  end;

  // razao
  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('RAZAO').AsString <> 'S' then
    begin
      RAZAO := '';
    end;
  end;

  if vazio(RAZAO) then
  begin
    // RAZAO := retornaCampo('RAZAO');
    RAZAO := qrPesquisaCli.FieldByName('RAZAO').AsString;
    log.Lines.Add(RAZAO);
    log.Lines.Add(' RAZÃO Inválida, foram mantidos os valores originais para este campo' + #13);
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('FANTASIA').AsString <> 'S' then
    begin
      FANTASIA := '';
    end;
  end;

  if vazio(FANTASIA) then
  begin
    // FANTASIA := retornaCampo('FANTASIA');
    FANTASIA := qrPesquisaCli.FieldByName('FANTASIA').AsString;
    log.Lines.Add(FANTASIA);
    log.Lines.Add('FANTASIA Inválida, foram mantidos os valores originais para este campo' + #13);
  end;

  //showMessageDesenv(END_RUA);
  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if (tblParametros.FieldByName('ENDERECO').AsString <> 'S') then
    begin
      END_RUA := qrPesquisaCli.FieldByName('END_RUA').AsString;
    end;
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if (tblParametros.FieldByName('ATIVAR_DESATIVAR_CLIENTE').AsString <> 'S') then
    begin
      ATIVO := qrPesquisaCli.FieldByName('ATIVO').AsString;
    end;
  end;

  //showMessageDesenv(END_RUA);
  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('BAIRRO').AsString <> 'S' then
    begin
      BAIRRO := qrPesquisaCli.FieldByName('BAIRRO').AsString;
    end;
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('CIDADE').AsString <> 'S' then
    begin
      CIDADE := qrPesquisaCli.FieldByName('CIDADE').AsString;
    end;
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('CEP').AsString <> 'S' then
    begin
      CEP := qrPesquisaCli.FieldByName('CEP').AsString;
    end;
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('ESTADO').AsString <> 'S' then
    begin
      ESTADO := qrPesquisaCli.FieldByName('ESTADO').AsString;
    end;
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('EMAIL').AsString <> 'S' then
    begin
      EMAIL := qrPesquisaCli.FieldByName('EMAIL').AsString;
    end;
  end;

   // fone
  if (AQuery.Name = 'qrUpdateCliente') and
    (tblParametros.FieldByName('TELEFONE').AsString <> 'S') then
  begin
    //FONE1 :=
    SepararDDD(qrPesquisaCli.FieldByName('FONE1').AsString, AREA1, FONE1)
  end else
  if (tblParametros.FieldByName('TELEFONE').AsString = 'S') then
  begin
   // FONE1 := qrPesquisaCli.FieldByName('FONE1').AsString;
    FONE1 := retornaSoNumero(FONE1);
    AREA1 := Copy(FONE1, 1, 2);
    FONE1 := Copy(FONE1, 3, Length(FONE1));
  end;

  // showMessageDesenv('Area: ' + AREA1 + ' Fone: ' + FONE1);
//  if (Length(AREA1 + FONE1) < 10) and
//    (tblParametros.FieldByName('TELEFONE').AsString = 'S') then
//  begin
//    log.Lines.Add(ObterLogDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet));
//    log.Lines.Add('TELEFONE Inválido, foram mantidos os valores originais para este campo' + #13);
//  end else
//  if (Length(AREA1 + FONE1) >= 10) and
//    (tblParametros.FieldByName('TELEFONE').AsString = 'S') then
//  begin
//    AREA1 := Copy(FONE1, 1, 2);
//    FONE1 := Copy(FONE1, 3, Length(FONE1));
//  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('CNPJ').AsString = 'S' then
    begin
      CPF_CNPJ := CPF_CNPJ;
    end else
    begin
      CPF_CNPJ := qrPesquisaCli.FieldByName('CPF_CNPJ').AsString;
    end;
  end;

  if cdsConfigVALIDA_CNPJ.AsString = 'S' then
  begin
    if not CalculaCnpjCpf(CPF_CNPJ) then
    begin
      CPF_CNPJ := qrPesquisaCli.FieldByName('CPF_CNPJ').AsString;
      log.Lines.Add(CPF_CNPJ);
      log.Lines.Add('CPF/CNPJ Inválido, foram mantidos os valores originais para este campo' + #13);
    end;
  end;

  COD_UNIDADE := RetornaUnidade(qrPesquisaCli.FieldByName('SEGMENTO').AsString);

  if cdsConfigATUALIZAR_OBS_CLIENTE_IMPORT.AsString = 'S' then
  begin
    if OBS_ADMIN = '' then
    begin
      lObservacao := qrPesquisaCli.FieldByName('OBS_ADMIN').AsString;
    end else
    begin
      lObservacao := OBS_ADMIN;
    end;
  end;

  if (Trim(COD_ERP) = Trim(COD_ERP_HOLDING)) or
    (cdsConfigUTILIZAHOLDING.AsString = 'N') then
  begin
    COD_ERP_HOLDING := EmptyStr;
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('SEGMENTO').AsString <> 'S' then
    begin
      SEGMENTO := qrPesquisaCli.FieldByName('SEGMENTO').AsString
    end else
    begin
      SEGMENTO := RetornaSegmento(SEGMENTO);
    end;
  end else
  begin
    SEGMENTO := RetornaSegmento(SEGMENTO);
  end;

  // showMessageDesenv('SEGMENTO: ' + SEGMENTO);
  if (AQuery.Name = 'qrUpdateCliente') then
  begin
    OPERADOR := RetornaOperador(OPERADOR, COD_ERP, AQuery.Name)
  end else
  begin
    OPERADOR := RetornaOperador(OPERADOR);
  end;

  if (cdsConfigINAT_CLI_OPER.AsString = 'S') then
  begin
    if cdsConfigOPER_INAT_CLI.AsString = trim(OPERADOR) then
    begin
      ATIVO := 'NAO';
    end;
  end;
   // showMessageDesenv('OPERADOR: ' + OPERADOR);

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('SALDO').AsString <> 'S' then
    begin
      SALDO_DISPONIVEL :=
        FloatToStr(qrPesquisaCli.FieldByName('SALDO_DISPONIVEL').AsFloat);
    end;
  end;

  SALDO_DISPONIVEL := FloatToStr(StrToFloatDef(SALDO_DISPONIVEL, 0));

  if DATA_ULT_COMPRA = '' then
  begin
    DATA_ULT_COMPRA := 'NULL';
  end;

  if VENCIMENTO_LIMITE_CREDITO = '' then
  begin
    VENCIMENTO_LIMITE_CREDITO := 'NULL';
  end;

  if vazio(SALDO_DISPONIVEL) then
  begin
    SALDO_DISPONIVEL := '0';
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if tblParametros.FieldByName('POTENCIAL').AsString <> 'S' then
    begin
      POTENCIAL := qrPesquisaCli.FieldByName('POTENCIAL').AsString;
    end;
  end;

  if AQuery.Name = 'qrUpdateCliente' then
  begin
    if vazio(POTENCIAL) then
    begin
      POTENCIAL := '0';
    end;
  end;

  showMessageDesenv(COD_ERP + #13 + RAZAO + #13 + FANTASIA + #13 +
    CPF_CNPJ + #13 + END_RUA + #13 + CIDADE + #13 + BAIRRO + #13 +
    ESTADO + #13 + CEP + #13 + AREA1 + #13 + FONE1 + #13 + EMAIL + #13 +
    OPERADOR + #13 + COD_UNIDADE + #13 + SALDO_DISPONIVEL + #13 +
    POTENCIAL + #13 + DATA_ULT_COMPRA + #13 + SEGMENTO + #13 +
    OBS_CLIENTES + #13 + ATIVO + #13);

  AQuery.Close;
  AQuery.SQL.Clear;
  AQuery.SQL.Add('UPDATE');
  AQuery.SQL.Add(' CLIENTES ');
  AQuery.SQL.Add(' SET DATA_ULT_COMPRA = NULL, ');
  AQuery.SQL.Add(' VENCIMENTO_LIMITE_CREDITO = NULL ');
  AQuery.SQL.Add(' WHERE' );
  AQuery.SQL.Add(' CODIGO = :CODIGO');
  AQuery.Parameters.ParamByName('CODIGO').Value := CodigoCliente;
  AQuery.ExecSQL;

  AQuery.Close;
  try
    AQuery.SQL.Clear;
    if AQuery.Name = 'qrUpdateCliente' then
    begin
      AQuery.SQL.Add('UPDATE IGNORE');
      AQuery.SQL.Add(' CLIENTES ');
      AQuery.SQL.Add(' SET');
      AQuery.SQL.Add(' COD_ERP = :COD_ERP,');
      AQuery.SQL.Add(' RAZAO = :RAZAO,');
      AQuery.SQL.Add(' FANTASIA = :FANTASIA,');
      AQuery.SQL.Add(' CPF_CNPJ = :CPF_CNPJ,');
      AQuery.SQL.Add(' END_RUA = :END_RUA,');
      AQuery.SQL.Add(' CIDADE = :CIDADE,');
      AQuery.SQL.Add(' BAIRRO = :BAIRRO,');
      AQuery.SQL.Add(' ESTADO = :ESTADO,');
      AQuery.SQL.Add(' CEP = :CEP,');
      AQuery.SQL.Add(' AREA1= :AREA1,');
      AQuery.SQL.Add(' FONE1 = :FONE1,');
      AQuery.SQL.Add(' EMAIL = :EMAIL,');
      AQuery.SQL.Add(' OPERADOR = :OPERADOR,');
      AQuery.SQL.Add(' COD_UNIDADE = :COD_UNIDADE,');
      AQuery.SQL.Add(' SALDO_DISPONIVEL = :SALDO_DISPONIVEL,');
      AQuery.SQL.Add(' POTENCIAL = :POTENCIAL,');
      AQuery.SQL.Add(' DATA_ULT_COMPRA = :DATA_ULT_COMPRA, ');
      AQuery.SQL.Add(' VENCIMENTO_LIMITE_CREDITO = :VENCIMENTO_LIMITE_CREDITO, ');
      AQuery.SQL.Add(' SEGMENTO = :SEGMENTO,');
      AQuery.SQL.Add(' ATIVO = :ATIVO,');

      if lObservacao <> '' then
      begin
        AQuery.SQL.Add('OBS_ADMIN = :OBS_ADMIN,');
      end;
      AQuery.SQL.Add(' OBS_CLIENTES = :OBS_CLIENTES,');

      AQuery.SQL.Add(' COD_MIDIA = :COD_MIDIA');
      AQuery.SQL.Add(' WHERE' );
      AQuery.SQL.Add(' CODIGO = :CODIGO');
    end else
    begin
      AQuery.SQL.Add( 'INSERT IGNORE INTO CLIENTES ' + #13 +
        '(RAZAO, FANTASIA, CPF_CNPJ, END_RUA, CIDADE, BAIRRO, ' + #13 +
        'ESTADO, CEP, AREA1, FONE1, EMAIL, OPERADOR, COD_ERP, ' + #13 +
        'COD_UNIDADE, SALDO_DISPONIVEL, POTENCIAL, DATA_ULT_COMPRA, ' + #13 +
        'SEGMENTO, OBS_CLIENTES, ATIVO, COD_MIDIA, OBS_ADMIN, ' + #13 +
        'VENCIMENTO_LIMITE_CREDITO) ' + #13 +
        'VALUES ' + #13 +
        '(:RAZAO, :FANTASIA, :CPF_CNPJ, :END_RUA, :CIDADE, :BAIRRO, ' + #13 +
        ':ESTADO, :CEP, :AREA1, :FONE1, :EMAIL, :OPERADOR, :COD_ERP, '  + #13 +
        ':COD_UNIDADE, :SALDO_DISPONIVEL, :POTENCIAL, ' + #13 +
        ':DATA_ULT_COMPRA, :SEGMENTO, :ATIVO, ' + #13 +
        ':OBS_ADMIN, :OBS_CLIENTES, :COD_MIDIA, :VENCIMENTO_LIMITE_CREDITO)');
    end;

    AQuery.Parameters.ParamByName('COD_ERP').Value := trim(COD_ERP);
    AQuery.Parameters.ParamByName('OPERADOR').Value := trim(OPERADOR);

    if ATIVO = 'SIM' then
    begin
      AQuery.Parameters.ParamByName('ATIVO').Value := 'SIM'
    end else
    if ATIVO = 'NAO' then
    begin
      AQuery.Parameters.ParamByName('ATIVO').Value := 'NAO';
    end;

    AQuery.Parameters.ParamByName('CPF_CNPJ').Value := trim(CPF_CNPJ);
    AQuery.Parameters.ParamByName('RAZAO').Value := trim(RAZAO);
    AQuery.Parameters.ParamByName('FANTASIA').Value := trim(FANTASIA);
    AQuery.Parameters.ParamByName('END_RUA').Value := trim(END_RUA);
    AQuery.Parameters.ParamByName('CIDADE').Value := Trim(CIDADE);
    AQuery.Parameters.ParamByName('BAIRRO').Value := trim(BAIRRO);
    AQuery.Parameters.ParamByName('ESTADO').Value := trim(ESTADO);
    AQuery.Parameters.ParamByName('CEP').Value := trim(CEP);

    if Trim(AREA1) = '' then
    begin
      AQuery.Parameters.ParamByName('AREA1').Value := '0';
    end else
    begin
      AQuery.Parameters.ParamByName('AREA1').Value := Trim(AREA1);
    end;

    AQuery.Parameters.ParamByName('FONE1').Value := Trim(FONE1);
    AQuery.Parameters.ParamByName('EMAIL').Value := trim(EMAIL);
    AQuery.Parameters.ParamByName('COD_UNIDADE').Value := trim(COD_UNIDADE);
    AQuery.Parameters.ParamByName('SALDO_DISPONIVEL').Value :=
      Trim(SALDO_DISPONIVEL);
    AQuery.Parameters.ParamByName('POTENCIAL').Value := trim(POTENCIAL);

    if DATA_ULT_COMPRA = 'NULL' then
    begin
      AQuery.Parameters.ParamByName('DATA_ULT_COMPRA').Value := NULL;
    end else
    begin
      AQuery.Parameters.ParamByName('DATA_ULT_COMPRA').Value := DATA_ULT_COMPRA;
    end;

    if VENCIMENTO_LIMITE_CREDITO = 'NULL' then
    begin
      AQuery.Parameters.ParamByName('VENCIMENTO_LIMITE_CREDITO').Value := NULL;
    end else
    begin
      AQuery.Parameters.ParamByName('VENCIMENTO_LIMITE_CREDITO').Value :=
        VENCIMENTO_LIMITE_CREDITO;
    end;

    AQuery.Parameters.ParamByName('SEGMENTO').Value := trim(SEGMENTO);

    if AQuery.Name = 'qrUpdateCliente' then
    begin
      if lObservacao <> '' then
      begin
        AQuery.Parameters.ParamByName('OBS_ADMIN').Value := lObservacao;
      end;
    end;
    AQuery.Parameters.ParamByName('OBS_CLIENTES').Value := OBS_CLIENTES;

    AQuery.Parameters.ParamByName('COD_MIDIA').Value := '0';

    if AQuery.Name = 'qrUpdateCliente' then
    begin
      AQuery.Parameters.ParamByName('CODIGO').Value := CodigoCliente;
    end;

    DefinirMidia(AQuery);

    AQuery.ExecSQL;
  except
    on e: Exception do
    begin
      log.Lines.Add('Problema ao ' + AQuery.Name + 'alterar cliente ' + #13 +
        e.ClassName + #13 + e.Message);
      log.Lines.Add(
        ObterLogDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet));
    end;
  end;

  Result := retornaCampo('CODIGO');

  if FDatataModuloConexao.GetDataSouceClientes.DataSet.FindField('AREA_DESCRICAO') <> nil then
  begin
    if FDatataModuloConexao.GetDataSouceClientes.DataSet.fieldbyname('AREA_DESCRICAO').AsString <> '' then
    begin
       with TADOQuery.Create(nil) do
      try
        Connection := conexaoSGR;
        SQL.Text := ' SELECT CODIGO FROM grupos WHERE DESCRICAO = ' +
          QuotedStr(FDatataModuloConexao.GetDataSouceClientes.DataSet.fieldbyname('AREA_DESCRICAO').AsString);
        Open;
        CodigoGrupo := FieldByName('CODIGO').AsInteger;
        Close;
      finally
        Free;
      end;

      if CodigoGrupo = 0 then
      begin
        with TADOQuery.Create(nil) do
        try
          Connection := conexaoSGR;
          SQL.Text := ' SELECT max(CODIGO) as CODIGO FROM grupos ';
          Open;
          CodigoGrupo := FieldByName('CODIGO').AsInteger + 1;
          Close;
          SQL.Text := 'INSERT INTO grupos values(' + IntToStr(CodigoGrupo) + ',' +
            QuotedStr(FDatataModuloConexao.GetDataSouceClientes.DataSet.fieldbyname('AREA_DESCRICAO').AsString) + ')';
          ExecSQL;
        finally
          Free;
        end;
      end;
    end;

    if CodigoGrupo > 0 then
    begin
      with TADOQuery.Create(nil) do
      try
        Connection := conexaoSGR;
        SQL.Text := ' update clientes set GRUPO = ' + IntToStr(CodigoGrupo) +
          ' where CODIGO = ' + Result;
        ExecSQL;
      finally
        Free;
      end;
    end;
  end;

  if (Result <> '') and (trim(OPERADOR) <> '') then
  begin
    lQueryCampanha := TADOQuery.Create(nil);;
    lQueryCampanha.Connection := conexaoSGR;

    try
      lQueryCampanha.Close;
      lQueryCampanha.SQL.Clear;
      lQueryCampanha.SQL.Add('UPDATE campanhas_clientes SET OPERADOR = ' + trim(OPERADOR) +
        ' WHERE CLIENTE = ' + Result + ' AND CONCLUIDO = "NAO" ' +
        'AND OPERADOR <> -2 AND COALESCE(FIDELIZA,'''') <> ''S'' ');
      lQueryCampanha.ExecSQL;
    finally
      FreeAndNil(lQueryCampanha);
    end;

    if AQuery.Name <> 'qrUpdateCliente' then
    begin
      if tblParametros.FieldByName('INCLUI_CLIENTE_CAMP_PRINC_IMP').AsString = 'S' then
      begin
        FONE1 := AREA1 + FONE1;
        lQueryInsertCampanha := TADOQuery.Create(nil);
        try
          lQueryInsertCampanha.Connection := conexaoSGR;
          try
            lQueryInsertCampanha.SQL.Clear;
            lQueryInsertCampanha.SQL.Add('INSERT IGNORE INTO campanhas_clientes (CLIENTE, ' +
             'CAMPANHA, CONCLUIDO, DT_AGENDAMENTO, OPERADOR, FONE1, ' +
             'AGENDA) VALUES');
            lQueryInsertCampanha.SQL.Add('(' + Result + ', ' + GetCampanha(AQuery) +
            ' , ''NAO'', NOW(), ' +
            trim(GetOPERADOR(AQuery, OPERADOR)) + ', ' + QuotedStr(trim(FONE1))
              + ', ' +
            GetPrioridadeAgenda(AQuery) + ' )');
            showMessageDesenv(AQuery.SQL.Text);
            lQueryInsertCampanha.ExecSQL;
          except
            on e: Exception do
            begin
              log.Lines.Add('Problema ao inserir CAMPANHAS ' + #13 +
                e.ClassName + #13 + e.Message);
//                log.Lines.Add(
//                  ObterLogDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet));
            end;
          end;
        finally
          FreeAndNil(lQueryInsertCampanha);
        end;
      end;
    end;

    if (COD_ERP_HOLDING <> '') then
    begin
      if StrToIntDef(COD_ERP_HOLDING,0) > 0 then
      begin
        atualizarClienteHolding(Trim(COD_ERP),
                                Trim(COD_ERP_HOLDING),
                                DATA_ULT_COMPRA,
                                OPERADOR);
      end;
    end;
  end;
end;

procedure TfrmImportarDireto.limpaComprasItensCompras;
begin
  with TADOQuery.Create(nil) do
  try
    Connection := conexaoSGR;
    SQL.Add('DELETE FROM itens_compra');
    ExecSQL;
    SQL.Clear;
    SQL.Add('DELETE FROM compras');
    ExecSQL;
  finally
    Free;
  end;
end;

procedure TfrmImportarDireto.organizaCampanhas;
begin
  try
    conexaoSGR.Execute('CALL PR_ORGANIZA_CAMPANHAS()');
  except
    on e: Exception do
    begin
      log.Lines.Add('Problemas ao organizar campanhas' + #13 +
        e.ClassName + #13 + e.Message + #13);
    end;
  end;
end;

procedure TfrmImportarDireto.pgcChange(Sender: TObject);
begin
  ConfigurarAjuda;
end;

function TfrmImportarDireto.ExecSql(xsql: string; Tipo: Integer = 0): TADOQuery;
begin
  Result := nil;
  try
    if tipo = 0 then
    begin
      Result := TADOQuery.Create(nil);
      Result.Connection := conexaoSGR;
      Result.SQL.Text := xsql;
      Result.Open;
    end else
    begin
      conexaoSGR.Execute(xsql);
      Result := nil;
    end;
  except
    on e: exception do
      log.Lines.Add('Erro sql: ' + xsql + ' - ' + e.Message);
  end;
end;

function TfrmImportarDireto.refidelizaCotados: Boolean;
var
  query, Query2: TADOQuery;

  function GetTransferenciaClientes(Cliente: integer): string;
  begin

    with ExecSql(' SELECT a.DATA_HORA FROM transferencia_clientes a INNER JOIN '
      + ' transferencia_clientes_itens b ON a.CODIGO = b.CODIGO WHERE b.CLIENTE = ' + IntToStr(Cliente)
      + ' ORDER BY a.DATA_HORA DESC LIMIT 1 ') do
    begin
      if IsEmpty then
        result := 'NULL'
      else
        result := QuotedStr(FormatDateTime('yyyy-mm-dd HH:NN:SS', FieldByName('DATA_HORA').AsDateTime));
      Free;
    end;

  end;

begin
  Result := False;

  if ExecSql(' SELECT CARTEIRA_FIXA_OPERADOR FROM parametros ')
    .FieldByName('CARTEIRA_FIXA_OPERADOR').AsString = 'S' then
  begin
    Exit;
  end;

  try
    query := TADOQuery.Create(nil);
    query.Connection := conexaoSGR;
    Query2 := TADOQuery.Create(nil);
    Query2.Connection := conexaoSGR;

    Query.SQL.Clear;
    Query.SQL.ADD('SELECT');
    Query.SQL.ADD('	cli.CODIGO          ');
    Query.SQL.ADD('FROM                   ');
    Query.SQL.ADD('	clientes cli,       ');
    Query.SQL.ADD('	campanhas_clientes cc, ');
    Query.SQL.ADD('	resultados r        ');
    Query.SQL.ADD('WHERE                  ');
    Query.SQL.ADD('	cc.CLIENTE = cli.CODIGO  ');
    Query.SQL.ADD('	AND r.FIDELIZARCOTACAO = "SIM" ');
    Query.SQL.ADD('	AND cli.OPERADOR <> -2   ');
    Query.SQL.ADD('	AND r.CODIGO = cc.RESULTADO ');
    Query.SQL.ADD('GROUP BY        ');
    Query.SQL.ADD('	cli.CODIGO    ');
    Query.Open;

    while not Query.Eof do
    begin
      Query2.Close;
      Query2.SQL.Clear;
      Query2.SQL.Text := 'CALL PR_REPOE_LIGACOES_FIDELIZADAS(' + Query.Fields[0].AsString
        + ',9999,' + GetTransferenciaClientes(Query.Fields[0].AsInteger) + ')';
      try
        Query2.ExecSQL;
      except
        on E: Exception do
        begin
          log.Lines.Add('Importar RepoeLigacoesFidelizadas' + #13 + e.Message);
        end;
      end;

      Query.Next;
    end;

    Result := True;
  finally
    FreeAndNil(query);
    FreeAndNil(Query2);
  end;
end;

function TfrmImportarDireto.retornaCampo(campo: string): string;
begin
  with TADOQuery.Create(nil) do
  try
    Connection := conexaoSGR;
    SQL.Add('SELECT');
    SQL.Add(campo);
    SQL.Add('FROM vw_clientes');
    SQL.Add('WHERE COD_ERP = :CAMPO');
    Parameters[0].Value := FDatataModuloConexao.GetDataSouceClientes.DataSet.Fields[0].AsString;
    Open;
    Result := Fields[0].AsString;
  finally
    Free;
  end;
end;

function TfrmImportarDireto.RetornaUnidade(Valor: string): string;
begin
  if tblParametros.FieldByName('UNIDADE').AsString = 'S' then
  begin
    with TADOQuery.Create(nil) do
    try
      Connection := conexaoSGR;
      SQL.Text := 'SELECT COD_UNIDADE FROM SEGMENTOS WHERE CODIGO = '
        + IntToStr(StrToIntDef(Valor, 0));
      Open;

      if FieldByName('COD_UNIDADE').AsString <> '' then
      begin
        close;
        SQL.Text := 'SELECT COD_UNIDADE FROM clientes WHERE COD_ERP = :COD_ERP';
        Parameters[0].Value :=  FDatataModuloConexao.qrCliente.FieldByName('COD_ERP').AsString ;
        //FDatataModuloConexao.GetDataSouceClientes.DataSet.Fields[0].AsString;
        Open;
      end;

      if FieldByName('COD_UNIDADE').AsString = '' then
      begin
        Result := '1'
      end else
      begin
        Result := Fields[0].AsString;
      end;

      close;
    finally
      Free;
    end;
  end else
  begin
    with TADOQuery.Create(nil) do
    try
      Connection := conexaoSGR;
      SQL.Text := 'SELECT COD_UNIDADE FROM clientes WHERE COD_ERP = :COD_ERP';
      Parameters[0].Value := FDatataModuloConexao.GetDataSouceClientes.DataSet.Fields[0].AsString;
      Open;

      if IsEmpty then
      begin
        Result := '1'
      end else
      begin
        Result := Fields[0].AsString;
      end;

      close;
    finally
      Free;
    end;
  end;
end;

procedure TfrmImportarDireto.SepararDDD(Fone: string; out DDD, Telefone: string);

begin
  if Length(Fone) > 9 then
  begin
    DDD := Copy(fone, 1, 2);
    delete( fone, 1, 2);
    Telefone := Copy(Fone, 1, length(Fone));
  end;
end;

procedure TfrmImportarDireto.showMessageDesenv(texto: string);
begin
  if FileExists(pCaminhoPrograma + 'desenvolvimento.cfg') then
    ShowMessage(texto);
end;

procedure TfrmImportarDireto.tmrImportacaoTimer(Sender: TObject);
begin
  {if (pATUALIZACAO_TODODIA <> '8') then
    if (StrToInt(pATUALIZACAO_TODODIA) = 0) or (DayOfWeek(now) = StrToInt(pATUALIZACAO_TODODIA)) then
      if FormatDateTime('hh:mm', now) = Copy(pATUALIZACAO_HORA, 1, 5) then
      begin
        btnImportarClick(Self);
            //trayIcon.Hint := 'in.Pulse - Atualização concluída as ' + DateToStr(now);
        if pDESLIGA_COMPUTADOR = 'S' then
          WinExec('cmd /c shutdown -s -t 600', SW_HIDE); // 10 minutos
      end;   }
end;

procedure TfrmImportarDireto.trayIconClick(Sender: TObject);
begin
  frmImportarDireto.Show;
end;

function TfrmImportarDireto.vazio(texto: string): Boolean;
begin
  Result := (Length(trim(texto)) = 0) or (texto = 'NULL');
end;

function TfrmImportarDireto.CalculaCnpjCpf(Numero: string): Boolean;
var
  i, d, b, Digito: byte;
  Soma: Integer;
  CNPJ: Boolean;
  DgPass, DgCalc: string;

  function ApenasNumerosStr(pStr: string): string;
  var
    i: Integer;
  begin
    Result := '';
    for i := 1 to Length(pStr) do
    begin
      if (CharInSet(pStr[i], ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'])) then
      begin
        Result := Result + pStr[i];
      end;
    end;
  end;

  function IIf(pCond: Boolean; pTrue, pFalse: Variant): Variant;
  begin
    if pCond then
    begin
      Result := pTrue
    end else
    begin
      Result := pFalse;
    end;
  end;

begin
  if vazio(Numero) then
  begin
    Result := True;
    Exit;
  end;

  Result := False;
  Numero := ApenasNumerosStr(Numero);
   // Caso o número não seja 11 (CPF) ou 14 (CNPJ), aborta
  case Length(Numero) of
    11:
      CNPJ := False;
    14:
      CNPJ := True;
  else
    Exit;
  end;
   // Separa o número do digito
  DgCalc := '';
  DgPass := Copy(Numero, Length(Numero) - 1, 2);
  Numero := Copy(Numero, 1, Length(Numero) - 2);
   // Calcula o digito 1 e 2
  for d := 1 to 2 do
  begin
    b := IIf(d = 1, 2, 3); // byTE
    Soma := IIf(d = 1, 0, STRTOINTDEF(DgCalc, 0) * 2);
    for i := Length(Numero) downto 1 do
    begin
      Soma := Soma + (Ord(Numero[i]) - Ord('0')) * b;
      Inc(b);
      if (b > 9) and CNPJ then
        b := 2;
    end;
    Digito := 11 - Soma mod 11;
    if Digito >= 10 then
      Digito := 0;
    DgCalc := DgCalc + Chr(Digito + Ord('0'));
  end;
  Result := DgCalc = DgPass;
end;

function TfrmImportarDireto.RetornaOperador(Valor: string): string;
begin
  if vazio(Valor) then
  begin
    Result := '0'
  end else
  begin
    with TADOQuery.Create(nil) do
    try
      Connection := conexaoSGR;
      SQL.Text := 'SELECT CODIGO FROM operadores WHERE CODIGO_ERP = :CODIGO_ERP';
      Parameters[0].Value := trim(Valor);
      Open;
      if IsEmpty then
        Result := '-2'
      else
        Result := Fields[0].AsString;
    finally
      Free;
    end;
  end;
end;

function TfrmImportarDireto.RetornaOperador(Valor, COD_ERP_CLIENTE: string;
  ANomeQuery: string): string;
var
  oQuery: TADOQuery;
  lQueryCampanha: TADOQuery;
begin
  oQuery := TADOQuery.Create(nil);
  try
    try
      oQuery.Connection := conexaoSGR;
      oQuery.SQL.Text := 'SELECT * FROM OPERADORES WHERE CODIGO_ERP = ' + QuotedStr(Trim(Valor));
      oQuery.Open;

      if ANomeQuery <> 'qrUpdateCliente' then
      begin
        oQuery.Connection := conexaoSGR;
        oQuery.SQL.Text := 'SELECT * FROM OPERADORES WHERE CODIGO_ERP = ' + QuotedStr(Trim(Valor));
        oQuery.Open;
        if (oQuery.FieldByName('ATIVO').AsString = 'NAO')
          and (oQuery.FieldByName('CODIGO').AsInteger <> -2) then
        begin
          oQuery.close;
          oQuery.sql.Text :=
                        'SELECT ' + #13 +
                        '	CLI.CODIGO, ' + #13 +
                        '	TRC.OPERADOR_DESTINO  ' + #13 +
                        'FROM ' + #13 +
                        '	CLIENTES CLI  ' + #13 +
                        'INNER JOIN (SELECT MAX(CODIGO) AS MAX_CODIGO, CLIENTE ' + #13 +
                        'FROM transferencia_clientes_itens GROUP BY CLIENTE) TCI ON ' + #13 +
                        '	TCI.CLIENTE = CLI.CODIGO ' + #13 +
                        'INNER JOIN transferencia_clientes TRC ON ' + #13 +
                        '	TRC.CODIGO = TCI.MAX_CODIGO ' + #13 +
                        'WHERE ' + #13 +
                        '	CLI.COD_ERP = ' + Trim(COD_ERP_CLIENTE);
          oQuery.Open;

          if not(oQuery.IsEmpty) then
          begin
            Result := oQuery.FieldByName('OPERADOR_DESTINO').AsString;
          end else
          begin
            Result := '0';
          end;
        end else
        begin
          Result := RetornaOperador(Valor);
        end;
      end else
      begin
        lQueryCampanha := TADOQuery.Create(nil);
        try
          lQueryCampanha.Connection := conexaoSGR;
          lQueryCampanha.SQL.Text := 'SELECT * FROM CAMPANHAS_CLIENTES CA ' +
            'LEFT JOIN CLIENTES CC ON CA.CLIENTE = CC.CODIGO ' +
            'LEFT JOIN CAMPANHAS C ON C.CODIGO = CC.COD_CAMPANHA ' +
            'WHERE CC.COD_ERP = ' + COD_ERP_CLIENTE;
          lQueryCampanha.Open;
          if lQueryCampanha.FieldByName('TIPO').AsString = 'ATIVOS' then
          begin
            Result := RetornaOperador(Valor);
          end else
          if lQueryCampanha.FieldByName('AGENDA').AsString < '0' then
          begin
            Result := lQueryCampanha.FieldByName('OPERADOR').AsString;
          end else
          if lQueryCampanha.FieldByName('AGENDA').AsString >= '0' then
          begin
            Result := '0';
          end;
        finally
          FreeAndNil(lQueryCampanha);
        end;
      end;
    finally
      FreeAndNil(oQuery);
    end;
  except
    on E: Exception do
    begin
      log.Lines.Add('Erro ao retorno operador do cliente - ' +
        COD_ERP_CLIENTE + ' - ' + e.ClassName + ' - ' + e.Message );
    end;
  end;

end;

function TfrmImportarDireto.RetornaSegmento(Valor: string): string;
begin
  with TADOQuery.Create(nil) do
  try
    Connection := conexaoSGR;
    SQL.Text := 'SELECT CODIGO FROM segmentos WHERE NOME = :NOME';
    Parameters[0].Value := trim(Valor);
    Open;
    if IsEmpty then
    begin
      Close;
      SQL.Text := 'INSERT INTO segmentos (NOME) VALUES (:NOME)';
      Parameters[0].Value := trim(Valor);
      ExecSQL;

      Close;
      SQL.Text := 'SELECT MAX(CODIGO) FROM segmentos';
      Open;
    end;
    Result := Fields[0].AsString;
  finally
    Free;
  end;
end;

function TfrmImportarDireto.retornaSoNumero(Valor: string): string;
var
  i: Integer;

//  if not (CharInSet(Key,['0'..'9',#8]) then key := #0;
begin // retorna somente numeros
  for i := 1 to Length(Valor) do
  begin
    if CharInSet(Valor[i], ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']) then
    begin
      Result := Result + Valor[i];
    end;
  end;

 { if Length(Result) = 11 then
    Result := Copy(Result, 2, 10)}
end;

procedure TfrmImportarDireto.ExecutaFormShow(pExecucaoBatch: Boolean);
begin
  pgc.ActivePage := tsImportacao;
  bExecucaoBatch := pExecucaoBatch;
  cdsConfig.Close;
  cdsConfig.CreateDataSet;
  cdsConfig.Append;
  if FileExists(pCaminhoPrograma + pArquivoConexao) then
  begin
    PopularcdsConfig(pCaminhoPrograma + pArquivoConexao);
//    cdsConfig.LoadFromFile(pCaminhoPrograma + pArquivoConexao);
//    cdsConfig.Open;
//    cdsConfig.Edit;
  end else
  begin
    cdsConfigATUALIZACAO_TODODIA.AsString := '8';
    cdsConfigDESLIGA_COMPUTADOR.AsString := 'N';
    cdsConfigATUALIZACAO_DATAULTCOMPRA.AsString := '0';
    cdsConfigVALIDA_CNPJ.AsString := 'S';
    cdsConfigVALIDA_ERP.AsString := 'S';
    cdsConfigLIMPA_COMPRAS.AsString := 'N';
    cdsConfigINAT_CLI_OPER.AsString := 'N';
    cdsConfigOPER_INAT_CLI.AsString := '';
    cdsConfigUTILIZAHOLDING.AsString := 'N';
    cdsConfigSOBRESCREVER_COMPRAS.AsString := 'N';
    cdsConfigATUALIZAR_OBS_CLIENTE_IMPORT.AsString := 'N';
    cdsConfigUTILIZA_API.AsString := 'N';
  end;
  atualizaParametros;
  cbAtualizacaoDia.ItemIndex := cdsConfigATUALIZACAO_TODODIA.AsInteger;
  cbxDataUltimaCompra.ItemIndex := cdsConfigATUALIZACAO_DATAULTCOMPRA.AsInteger;

  tblParametros.Open;

  if cdsConfigIMPORTAR_CLIENTES.IsNull then
  begin
    cdsConfigIMPORTAR_CLIENTES.AsAnsiString :=
      tblParametros.FieldByName('IMPORTAR_CLIENTES').AsAnsiString;
  end;

  if cdsConfigIMPORTAR_COMPRAS.IsNull then
  begin
    cdsConfigIMPORTAR_COMPRAS.AsAnsiString :=
      tblParametros.FieldByName('IMPORTAR_COMPRAS').AsAnsiString;
  end;

  if cdsConfigIMPORTAR_CLIENTE_EXISTENTE.IsNull then
  begin
    cdsConfigIMPORTAR_CLIENTE_EXISTENTE.AsAnsiString :=
      tblParametros.FieldByName('IMPORTAR_CLIENTE_EXISTENTE').AsAnsiString;
  end;

  if cdsConfigUTILIZA_API.AsString = 'N' then
  begin
    tbsAPI.TabVisible := False;
  end;
end;

procedure TfrmImportarDireto.ExecutarFuncao(AFuncao: string);
begin
  try
    conexaoSGR.Execute('CALL ' + AFuncao + '()');
  except
    on e: Exception do
    begin
      log.Lines.Add('Problemas ao organizar executar a função:' + AFuncao + #13 +
        e.ClassName + #13 + e.Message + #13);
    end;
  end;
end;

procedure TfrmImportarDireto.Criar_Agenda_ComprasDIA;
begin
  if cdsConfigCriar_Agenda_ComprasDIA.Value then
  begin
    with ExecSql(' select a.CODIGO, CONCAT(a.AREA1, a.FONE1) AS FONE1 from clientes a where a.OPERADOR = -2 '
      + ' and exists(select 1 from compras b where DATA >= cast(ADDDATE(now(), -1000) as DATE) and b.CLIENTE = a.CODIGO) '
      + ' and not exists(select 1 from campanhas_clientes c where c.CLIENTE = a.CODIGO and c.CONCLUIDO = ''NAO'') ') do
    try
      while not eof do
      begin
        conexaoSGR.Execute('update clientes set ATIVO = ''SIM'' where CODIGO = ' + fieldbyname('CODIGO').AsString);

        conexaoSGR.Execute('INSERT IGNORE INTO campanhas_clientes (CLIENTE, CAMPANHA, CONCLUIDO, DT_AGENDAMENTO, OPERADOR, FONE1, AGENDA) VALUES'
          + '(' + fieldbyname('CODIGO').AsString + ', 1 , ''NAO'', ''2000/01/15 10:00'', -2, ' + QuotedStr(trim(fieldbyname('FONE1').AsString)) + ', -250 )');
        next;
      end;
    finally
      free;
    end;

    conexaoSGR.Execute('update clientes a set ATIVO = ''NAO'' where a.ATIVO = ''SIM'' and a.OPERADOR = -2 '
      + ' and not exists(select 1 from compras b where DATA >= cast(ADDDATE(now(), -7) as DATE) and b.CLIENTE = a.CODIGO) ');
  end;
end;

procedure TfrmImportarDireto.ObterDadosAPI;
var
  lUrl: string;
  lJson: TJSONObject;
  lTeste: string;
begin
  lUrl := cdsConfigURL_API.AsString;
  try
    RESTClient.BaseURL := lUrl;
    RESTRequest.Method := TRESTRequestMethod.rmPOST;
    RESTRequest.ClearBody;
    RESTRequest.AddBody('{"user" : "' + cdsConfigUSUARIO_API.AsString +
      '","pwd": " ' + cdsConfigSENHA_API.AsString + '"}',
                         ContentTypeFromString('application/json'));
    RESTRequest.Execute;

    if (RESTResponse.StatusCode = 200) then
    begin
      if Assigned(RESTResponse.JSONValue) then
      begin
        lJson := nil;
        try
          lJson := TJSONObject.ParseJSONValue(RESTResponse.JSONValue.ToJSON) as TJSONObject;
          lTeste := TJSONString(lJson.Get('token').JsonValue as TJsonString).Value;
        finally
          lJson.Free;
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      log.Lines.Add('Erro: Erro ao obter o Token');
    end;
  end;
end;

function TfrmImportarDireto.ObterLogDataSet(oDataSet: TDataSet): string;
var
  I: Integer;
begin
  for I := 0 to Pred(oDataSet.Fields.Count) do
  begin
    Result := Result + ', ' + oDataSet.Fields[i].AsString;
  end;
end;

function TfrmImportarDireto.ObterValorDataSet(oDataSet: TDataSet;
  const pPosicao: Integer): string;
begin
  try
    Result := oDataSet.Fields[pPosicao].AsString;
  except
    Result := EmptyStr;
    log.Lines.Add('Erro: ObterValorDataSet: ' + pPosicao.ToString);
  end;
end;

procedure TfrmImportarDireto.SetValorParametro(oQuery: TADOQuery;
  const pPosicao: integer; const pValue: String);
begin
  try
    oQuery.Parameters[pPosicao].Value := trim(pValue);
  except
    log.Lines.Add('Erro: SetValorParametro: ' + pPosicao.ToString);
  end;
end;

procedure TfrmImportarDireto.AlteraDataUltimaCompraCNPJ;
var
  qrUltCompraCNPJ : TADOQuery;
begin
  try
    try
      qrUltCompraCNPJ := TADOQuery.Create(nil);
      qrUltCompraCNPJ.Connection := conexaoSGR;
      qrUltCompraCNPJ.SQL.Clear;

      qrUltCompraCNPJ.SQL.Add('UPDATE ');
      qrUltCompraCNPJ.SQL.Add('	clientes cli');
      qrUltCompraCNPJ.SQL.Add('SET cli.DATA_ULT_COMPRA = ');
      qrUltCompraCNPJ.SQL.Add('   (');
      qrUltCompraCNPJ.SQL.Add('	SELECT');
      qrUltCompraCNPJ.SQL.Add('	  MAX(c.DATA_ULT_COMPRA)');
      qrUltCompraCNPJ.SQL.Add('	FROM');
      qrUltCompraCNPJ.SQL.Add('		clientes c');
      qrUltCompraCNPJ.SQL.Add('	where substring(cpf_cnpj,1,8) = (select substring(cpf_cnpj,1,8) from clientes where codigo = cli.CODIGO)');
      qrUltCompraCNPJ.SQL.Add('   )');
      qrUltCompraCNPJ.SQL.Add('   , cli.operador = ');
      qrUltCompraCNPJ.SQL.Add('   (');
      qrUltCompraCNPJ.SQL.Add('	SELECT');
      qrUltCompraCNPJ.SQL.Add('	  c.OPERADOR');
      qrUltCompraCNPJ.SQL.Add('	FROM');
      qrUltCompraCNPJ.SQL.Add('		clientes c');
      qrUltCompraCNPJ.SQL.Add('	where substring(cpf_cnpj,1,8) = (select substring(cpf_cnpj,1,8) from clientes where codigo = cli.CODIGO)');
      qrUltCompraCNPJ.SQL.Add('   )');

      showmessagedesenv(qrUltCompraCNPJ.SQL.text);

      qrUltCompraCNPJ.ExecSQL;
    except
      on e: Exception do
      begin
        InsereLog('Erro ao atualizar datas últimas compras dos clientes. - ' +
                      e.ClassName + ' - ' + e.message, mtError);
        InsereLog('Comando: ' + qrUltCompraCNPJ.Sql.Text);
      end;
    end;
  finally
    FreeAndNil(qrUltCompraCNPJ);
  end;
end;

procedure TfrmImportarDireto.InsereLog(pTexto: String; pTipo: TMsgDlgType);
var
  vTipoNome : String;
begin
  vTipoNome := 'ERRO';
  case pTipo of
    TMsgDlgType.mtWarning: vTipoNome := 'ATENÇÃO';
    TMsgDlgType.mtError: vTipoNome := 'ERRO';
  end;

  Log.Lines.Add('(' + TimeToStr(Time) + ') ' + vTipoNome + ' - ' + pTexto);
  Application.ProcessMessages;
end;

procedure TfrmImportarDireto.atualizarClienteHolding(COD_ERP, COD_ERP_HOLDING,
  DATA_ULT_COMPRA, OPERADOR: string);

  function existeHolding: Boolean;
  var
    oQuery: TADOQuery;
  begin
    oQuery := TADOQuery.Create(nil);

    try
      try
        oQuery.Connection := conexaoSGR;
        oQuery.SQL.Text :=
          'SELECT * FROM CLIENTES WHERE COD_ERP = ' + COD_ERP_HOLDING;
        oQuery.Open;

        Result := not(oQuery.IsEmpty);
      finally
        FreeAndNil(oQuery);
      end;
    except
      on E: exception do
        raise Exception.Create(E.Message);
    end;
  end;
var
  oQuery: TADOQuery;
  iPos_ini: Integer;
  sObs_temp: string;
begin
  if not(existeHolding) then
  begin
    log.Lines.Add('Holding ' + COD_ERP_HOLDING + ' não localizada.');
    exit;
  end;
  oQuery := TADOQuery.Create(nil);
  try
    try
      oQuery.Connection := conexaoSGR;
      oQuery.SQL.Text :=
            'UPDATE CLIENTES SET ATIVO = ''NAO'' WHERE COD_ERP = ' + QuotedStr(COD_ERP);
      oQuery.ExecSQL;

      oQuery.SQL.Text := 'SELECT OBS_ADMIN FROM CLIENTES WHERE COD_ERP = ' + QuotedStr(COD_ERP_HOLDING);
      oQuery.Open;

      if not(oQuery.IsEmpty) then
      begin
        sObs_temp := Trim(oQuery.FieldByName('OBS_ADMIN').AsString);

        iPos_Ini := Pos('Comprou pelo ERP: <', sOBS_TEMP);

        if iPos_Ini > 0 then
          sOBS_TEMP := copy(sOBS_TEMP, 1, iPos_Ini + 18) + TRIM(COD_ERP) +
                Copy(sOBS_TEMP, iPos_Ini + 18 + Pos('>', Copy(sOBS_TEMP, iPos_Ini + 18, 12)) - 1, Length(sOBS_TEMP))
        else if Pos('Comprou pelo ERP: ', sOBS_TEMP) > 0 then
          sOBS_TEMP := 'Comprou pelo ERP: <' + Trim(COD_ERP) + '>'
        else
          sOBS_TEMP := sOBS_TEMP + ' Comprou pelo ERP: <' + Trim(COD_ERP) + '>';

        oQuery.SQL.Text := 'UPDATE IGNORE clientes' + ' SET OPERADOR = ' +
                  Iif(vazio(OPERADOR), 'OPERADOR', OPERADOR) +
                  ' , DATA_ULT_COMPRA = ' + QuotedStr(FormatDateTime('YYYY-MM-DD', StrToDateDef(DATA_ULT_COMPRA, 0))) +
                  ' , OBS_ADMIN = ' + QuotedStr(sOBS_TEMP) + ' , ATIVO = ''SIM'' '
                  + ' WHERE COD_ERP = ' + QuotedStr(Trim(COD_ERP_HOLDING));

        oQuery.ExecSQL;

      end;

    finally
      FreeAndNil(oQuery);
    end;
  except
    on E: Exception do
      raise Exception.Create(E.Message);
  end;
end;

function TfrmImportarDireto.Iif(Condicao: Boolean; Verdadeiro, Falso: Variant): Variant;
begin
  if Condicao then
    Result := Verdadeiro
  else
    Result := Falso;
end;

procedure TfrmImportarDireto.Importar;
var
  aNF, aCLIENTE, CodigoCliente: string;
  cont: Integer;
  qrAux: TADOQuery;
  lDataUltimaCompra: TDateTime;
begin
  if not bExecucaoBatch then
  begin
    case MessageDlg('Deseja salvar os parâmetros informados?', mtConfirmation, mbYesNoCancel, 0) of
      mrYes : btnSalvar.OnClick(btnSalvar);
      mrCancel : Exit;
    end;
  end;

  try
    if cdsConfig.State in ([dsEdit, dsInsert]) then
      cdsConfig.Post;

    cdsConfig.Edit;

    if not conectar(edStringConexao.Text) then
      Abort;

   //trayIcon.Hint         := 'in.Pulse - Iniciando atualização';
    pb.Position := 0;
    cont := 0;
    pgc.ActivePage := tsImportacao;

    if cdsConfigUTILIZA_API.AsString = 'S' then
    begin
      ObterDadosAPI;
    end;

    try
      log.Lines.Clear;
      log.Lines.Add('Atualização iniciada: ' + DateTimeToStr(now) + #13);
      FDatataModuloConexao.OpenDadosClientes(cdsConfigCLIENTE_SQL.AsString);
      grdClientes.DataSource := FDatataModuloConexao.GetDataSouceClientes;
    except
      on e: Exception do
      begin
        ShowMessage('Problemas ao abrir clientes' + #13 + e.ClassName + #13 + e.Message);
        Abort;
      end;
    end;

    if FDatataModuloConexao.GetDataSouceClientes.DataSet.IsEmpty then
    begin
      InsereLog('Nenhum cliente encontrado.');
      Abort;
    end;

    if cdsConfigLIMPA_COMPRAS.AsString = 'S' then
    begin
      limpaComprasItensCompras;
    end;

    pb.Max := FDatataModuloConexao.GetDataSouceClientes.DataSet.RecordCount;
    lbTotalClientes.Caption := IntToStr(pb.Max);
    FDatataModuloConexao.GetDataSouceClientes.DataSet.DisableControls;
    FDatataModuloConexao.GetDataSouceClientes.DataSet.First;
    while not FDatataModuloConexao.GetDataSouceClientes.DataSet.eof do
    begin
      if vazio(FDatataModuloConexao.GetDataSouceClientes.DataSet.Fields[0].AsString)
        and (cdsConfigVALIDA_ERP.AsString = 'S') then
      begin
        log.Lines.Add(ObterLogDataSet(
          FDatataModuloConexao.GetDataSouceClientes.DataSet) +
          ' CÓDIGO ERP INVÁLIDO' + #13);
      end else
      begin
        // atualiza clientes
        // query: TADOQuery; 0 COD_ERP, 1 RAZAO, 2 FANTASIA, 3 CPF_CNPJ, 4 END_RUA, 5 CIDADE, 6 BAIRRO, 7 ESTADO,
        // 8 CEP, 9 FONE1, 10 EMAIL, 11 OPERADOR, 12 COD_UNIDADE, 13 SALDO_DISPONIVEL, 14 POTENCIAL,
        // 15 DATA_ULT_COMPRA, 16 SEGMENTO, 17 OBS_CLIENTES 18 ATIVO 19 COD_ERP_HOLDING
        if ExisteCliente(ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 0),
          ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 3)) then
        begin
          qrAux := qrUpdateCliente
        end else
        begin
          qrAux := qrInsertCliente;
        end;

        //Atenção o CAMPO (DATAULTIMACOMPRA), DEVE SER RESOLVIDO NO SQL, pois
        //em alguns casos (BANCOS) da erro de conversão na aplicação
        //trazendo certo no sql não terá problemas ao gravar esse campo

        aCLIENTE := insertOrUpdateCliente(qrAux,
            FDatataModuloConexao.qrCliente.FieldByName('COD_ERP').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('RAZAO').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('FANTASIA').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('CPF_CNPJ').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('END_RUA').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('CIDADE').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('BAIRRO').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('ESTADO').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('CEP').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('FONE1').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('EMAIL').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('OPERADOR').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('COD_UNIDADE').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('SALDO_DISPONIVEL').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('POTENCIAL').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('DATAULTIMACOMPRA').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('SEGMENTO').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('OBS_CLIENTE').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('OBS_ADMIN').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('ATIVO').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('HOLDING').AsString,
            FDatataModuloConexao.qrCliente.FieldByName('VENCIMENTO_LIMITE_CREDITO').AsString);

        if (tblParametros.FieldByName('HABILITAR_CADASTRO_MARCAS').AsString = 'S') then
        begin
          if Trim(ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet,
                                    20)) <> EmptyStr then
          begin
            //cod marca
            //Codigo da Marca
            //operador da marca
            //data última compra marca
            //unidade
            insertOrUpdateMarcaCliente(
              aCLIENTE,
              ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 0),
              ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 16),
              ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 20),
              ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 21),
              ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 22),
              ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 23));
          end;
        end;

        if (not vazio(aCLIENTE)) and (cdsConfigIMPORTAR_COMPRAS.AsString = 'S') then
        begin
          try
            CodigoCliente := ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 0);
            CodigoCliente := StringReplace(CodigoCliente, 'A', '', [rfReplaceAll, rfIgnoreCase]);
            FDatataModuloConexao.OpenDadosCompras(
              StringReplace(cdsConfigCOMPRA_SQL.AsString, '/*CODIGO*/', QuotedStr(CodigoCliente), [rfReplaceAll, rfIgnoreCase]));
            showMessageDesenv('Compras: ' + #13 + IntToStr(FDatataModuloConexao.GetDataSouceCompras.DataSet.RecordCount));
            Application.ProcessMessages;
            FDatataModuloConexao.GetDataSouceCompras.DataSet.First;
          except
            on e: Exception do
            begin
              log.Lines.Add('Problemas ao abrir compras ' + #13 + e.ClassName + #13 + e.Message + #13 +
                ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, 0) + #13);
            end;
          end;

          while not FDatataModuloConexao.GetDataSouceCompras.DataSet.eof do
          begin
            SobrescreverCompra(aCLIENTE,
              FDatataModuloConexao.qrCompra.FieldByName('DESCRICAO').AsString);

            // 0 COMPRA 1 CLIENTE, 2 Data, 3 Valor, 4 DESCRICAO(coloquei o codigo da compra), 5 FORMA_PGTO
            // 6 OPERADOR, 7 TIPO, 8 SITUACAO,
            aNF := insertCompra(aCLIENTE,
              FDatataModuloConexao.qrCompra.FieldByName('DT_COMPRA').AsString,
              FDatataModuloConexao.qrCompra.FieldByName('VALOR_TOTAL').AsString,
              FDatataModuloConexao.qrCompra.FieldByName('DESCRICAO').AsString,
              FDatataModuloConexao.qrCompra.FieldByName('FORMA_PAGAMENTO').AsString,
              FDatataModuloConexao.qrCompra.FieldByName('OPERADOR').AsString,
              FDatataModuloConexao.qrCompra.FieldByName('TIPO').AsString,
              FDatataModuloConexao.qrCompra.FieldByName('SITUACAO').AsString);

            if not vazio(aNF) then
            begin
              try
                FDatataModuloConexao.OpenDadosComprasItens(
                StringReplace(cdsConfigCOMPRA_IT_SQL.AsString, '/*CODIGO*/',
                  QuotedStr(ObterValorDataSet(FDatataModuloConexao.GetDataSouceCompras.DataSet, 0)),[rfReplaceAll, rfIgnoreCase])
                );

                log.Lines.Add('Itens Compra: '+
                  StringReplace(cdsConfigCOMPRA_IT_SQL.AsString, '/*CODIGO*/',
                  QuotedStr(ObterValorDataSet(FDatataModuloConexao.GetDataSouceCompras.DataSet, 0)),
                  [rfReplaceAll, rfIgnoreCase])
                );

                while not FDatataModuloConexao.GetDataSouceComprasItens.DataSet.eof do
                begin
                  // NOTA, 1 CODPROD, 2 DESCRICAO, 3 QDT, 4 UN_MEDIDA, 5 VALOR_UN, 6 DESCONTO
                  if not vazio(aNF) then
                  begin
                    insertCompraItem(aNF,
                     FDatataModuloConexao.qrComprasItens.FieldByName('ID_COMPRA').AsString,
                     FDatataModuloConexao.qrComprasItens.FieldByName('DESCRICAO').AsString,
                     FDatataModuloConexao.qrComprasItens.FieldByName('QUANTIDADE').AsString,
                     FDatataModuloConexao.qrComprasItens.FieldByName('UNIDADE').AsString,
                     FDatataModuloConexao.qrComprasItens.FieldByName('VALOR_UNITARIO').AsString,
                     FDatataModuloConexao.qrComprasItens.FieldByName('DESCONTO').AsString);
                  end;
                  FDatataModuloConexao.GetDataSouceComprasItens.DataSet.Next;
                end;
              except
                on e: Exception do
                begin
                  log.Lines.Add('Problemas ao abrir itens compras ' + #13 + e.ClassName + #13 + e.Message + #13 +
                  FDatataModuloConexao.GetDataSouceCompras.DataSet.Fields[0].AsString + #13);
                end;
              end;
            end;
            FDatataModuloConexao.GetDataSouceCompras.DataSet.Next;
            Application.ProcessMessages;
          end;
        end;
      end;
      Inc(cont);
      lbClienteTotal.Caption := IntToStr(cont);
      FDatataModuloConexao.GetDataSouceClientes.DataSet.Next;
      pb.Position := pb.Position + 1;
      Application.ProcessMessages;
    end;

    insereListaFones;

    with TADOQuery.Create(nil) do
    try
      Connection := conexaoSGR;
      try
        Close;
        SQL.Text := 'DELETE FROM itens_compra WHERE NOT EXISTS (SELECT c.CODIGO FROM compras c WHERE c.CODIGO = NOTA ) OR NOTA IS NULL';
        ExecSQL; // DELETA ITENS QUE NÃO EXISTAM NA TABELA DE COMPRAS
      except
        on e: Exception do
        begin
          log.Lines.Add('Problemas ao DELETA ITENS QUE NÃO EXISTAM NA TABELA DE COMPRAS' + #13 + e.ClassName + #13 + e.Message + #13);
        end;
      end;
      try
        Close;
        SQL.Text := 'DELETE FROM compras WHERE NOT EXISTS (SELECT c.CODIGO FROM clientes c WHERE c.CODIGO = CLIENTE )';
        ExecSQL; // DELETA COMPRAS QUE NÃO EXISTAM NA TABELA DE CLIENTES
      except
        on e: Exception do
        begin
          log.Lines.Add('Problemas ao DELETA COMPRAS QUE NÃO EXISTAM NA TABELA DE CLIENTES' + #13 + e.ClassName + #13 + e.Message + #13);
        end;
      end;

    finally
      Free;
    end;

    if cdsConfigIMPORTAR_COMPRAS.AsString = 'S' then
    begin
      if (cdsConfigATUALIZACAO_DATAULTCOMPRA.AsInteger = 0) then
      begin

        with TADOQuery.Create(nil) do
        try
          Connection := conexaoSGR;
          SQL.Clear;
          SQL.Add('CALL PR_ATUALIZA_DT_ULT_COMPRA_CLI');
          showMessageDesenv('PR_ATUALIZA_DT_ULT_COMPRA_CLI');
          ExecSQL;

          Free;
        except
          on e: Exception do
          begin
            log.Lines.Add(e.ClassName + ' ' + e.Message + #13);
            Free;
          end;
        end;
      end else
      begin
//        showMessageDesenv('AlteraDataUltimaCompra');
//        AlteraDataUltimaCompra;
        showMessageDesenv('-- organizaCampanhas');
        log.Lines.Add('organizaCampanhas');
        organizaCampanhas;
      end;

      showMessageDesenv('organizaAtualizacao');
      log.Lines.Add('-- organizaAtualizacao');
      ExecutarFuncao('PR_ORGANIZA_ATUALIZACAO');

      showMessageDesenv('AtualizaDataCompra');
      log.Lines.Add('AtualizaDataCompra');
      AtualizaDataCompra;

      showMessageDesenv('refidelizaCotados');
      log.Lines.Add('-- refidelizaCotados');
      refidelizaCotados;

      showMessageDesenv('AtualizaPeriodoRecompra');
      log.Lines.Add('-- AtualizaPeriodoRecompra');
      ExecutarFuncao('PR_ATUALIZAR_PERIODO_RECOMPRA');
    end;

    try
      conexaoSGR.Execute(' UPDATE clientes a SET POTENCIAL = COALESCE((select AVG(VALOR) AS VALOR from compras '
        + ' WHERE CLIENTE = a.CODIGO AND OPERADOR IS NULL) ,0) where a.POTENCIAL <= 0 ');
    except
      on E: Exception do
      begin
        log.Lines.Add(' update POTENCIAL cliente ' + #13 + e.ClassName + #13 + e.Message);
      end;
    end;

    Criar_Agenda_ComprasDIA;

    log.Lines.Add('Atualizando data de atualizacao cliente.');
    try
      conexaoSGR.Execute('UPDATE parametros set DATA_IMPORTACAO_CLIENTES_COMPRAS = CURRENT_TIMESTAMP() ');
    except
      on E: Exception do
      begin
        log.Lines.Add('Erro ao atualizar data da ult. importação - ' + e.ClassName + ' - ' + e.Message);
        log.Lines.Add('Comando:  UPDATE parametros set DATA_IMPORTACAO_CLIENTES_COMPRAS = CURRENT_TIMESTAMP() ');
      end;
    end;

    if tblParametros.FieldByName('HABILITAR_CADASTRO_MARCAS').AsString = 'S' then
    begin
      InsereLog('Atualizando operadores dos clientes conforme cadastro de marcas.');
      conexaoSGR.Execute('CALL PR_ATUALIZA_OPERADOR_MARCAS');

      InsereLog('Atualizando carteira dos clientes conforme data da última compra do cadastro de marcas.');
      conexaoSGR.Execute('CALL PR_ATUALIZA_CARTEIRA_CLIENTES_MARCAS');
    end;

   // pb.Position := 0;
    log.Lines.Add('Atualização encerrada: ' + DateTimeToStr(now));
    //if vazio(log.Text) then
    log.Lines.SaveToFile(pCaminhoPrograma + 'ErrosImportacao' + FormatDateTime('ddmmyy', date) + '.txt');
    FDatataModuloConexao.GetDataSouceClientes.DataSet.enableControls;
    FDatataModuloConexao.GetDataSouceClientes.DataSet.Close;
  finally
    if tblParametros.State in [dsedit, dsinsert] then
      tblParametros.Post;
  end;
end;

procedure TfrmImportarDireto.insertOrUpdateMarcaCliente(CLIENTE, COD_ERP,
  SEGMENTO, CODMARCA, OPERADORM, DATAULTIMACOMPRAM, UNIDADE: string);

  function existeMarca: Boolean;
  var
    oQuery: TADOQuery;
  begin
    oQuery := TADOQuery.Create(nil);
    try
      oQuery.Connection := conexaoSGR;
      oQuery.SQL.Text :=
        'SELECT * FROM MARCAS WHERE COD_ERPMARCA = ' + CODMARCA;
      oQuery.Open;

      Result := not(oQuery.IsEmpty);
    finally
      FreeAndNil(oQuery);
    end;
  end;

  function clientePossuiMarca: Boolean;
  var
    oQuery: TADOQuery;

  begin
    oQuery := TADOQuery.Create(nil);
    try
      oQuery.Connection := conexaoSGR;
      oQuery.SQL.Text :=
                        'SELECT ' +
                        '	CLI.CODIGO ' +
                        'FROM ' +
                        '	clientes_marcas CLM ' +
                        'INNER JOIN CLIENTES CLI ON ' +
                        '	CLI.CODIGO = CLM.CLIENTE ' +
                        'INNER JOIN MARCAS MAR ON ' +
                        ' MAR.CODIGO = CLM.MARCA ' +
                        'WHERE ' +
                        '	CLI.COD_ERP = :COD_ERP ' +
                        '	AND MAR.COD_ERPMARCA = :COD_ERPMARCA ';
      oQuery.Parameters.ParamByName('COD_ERP').Value := COD_ERP;
      oQuery.Parameters.ParamByName('COD_ERPMARCA').Value := CODMARCA;
      oQuery.Open;

      Result := not(oQuery.IsEmpty);
    finally
      FreeAndNil(oQuery);
    end;
  end;

  function getCodigoMarca: Integer;
  var
    oQuery: TADOQuery;
  begin
    oQuery := TADOQuery.Create(nil);
    try
      oQuery.Connection := conexaoSGR;
      oQuery.SQL.Text := 'SELECT CODIGO FROM MARCAS WHERE COD_ERPMARCA = ' + CODMARCA;
      oQuery.Open;

      Result := oQuery.FieldByName('CODIGO').AsInteger;
    finally
      FreeAndNil(oQuery);
    end;
  end;

var
  oQuery: TADOQuery;
  oData: tdatetime;
const
  INSERT_MARCA = 'INSERT INTO MARCAS(DESCRICAO, COD_ERPMARCA, UNIDADE) ' +
                  ' VALUES(:DESCRICAO, :COD_ERPMARCA, :UNIDADE)';

  UPDATE_MARCA = 'UPDATE MARCAS SET DESCRICAO = :DESCRICAO WHERE COD_ERPMARCA = :COD_ERPMARCA';

begin
  oQuery := TADOQuery.Create(nil);

  try
    try
      oQuery.Connection := conexaoSGR;
      if not existeMarca then
      begin
        oQuery.SQL.Text := INSERT_MARCA;
        oQuery.Parameters.ParamByName('DESCRICAO').Value := SEGMENTO;
        oQuery.Parameters.ParamByName('COD_ERPMARCA').Value := CODMARCA;
        oQuery.Parameters.ParamByName('UNIDADE').Value := UNIDADE;

        oQuery.ExecSQL;

      end
      else
      begin
        oQuery.SQL.Text := UPDATE_MARCA;
        oQuery.Parameters.ParamByName('DESCRICAO').Value := SEGMENTO;
        oQuery.Parameters.ParamByName('COD_ERPMARCA').Value := CODMARCA;
        oQuery.ExecSQL;
      end;

      if not clientePossuiMarca then
      begin
        oQuery.SQL.Text :=
                        'INSERT INTO CLIENTES_MARCAS(CLIENTE, MARCA, OPERADOR) ' +
                        'VALUES (:CLIENTE, :MARCA, :OPERADOR) ';
        oQuery.Parameters.ParamByName('CLIENTE').Value := CLIENTE;
        oQuery.Parameters.ParamByName('MARCA').Value := getCodigoMarca;
        oQuery.Parameters.ParamByName('OPERADOR').Value := RetornaOperador(OPERADORM);
        oQuery.ExecSQL;
      end;

      oQuery.SQL.Text := 'UPDATE ' +
                         '   CLIENTES_MARCAS ' +
                         ' SET DATA_ULTIMA_COMPRA = :DATA_ULTIMA_COMPRA, OPERADOR = :OPERADOR ' +
                         ' WHERE CLIENTE = :CLIENTE AND MARCA = :MARCA';

      if TryStrToDate(DATAULTIMACOMPRAM, odata) then
      begin
        oQuery.Parameters.ParamByName('DATA_ULTIMA_COMPRA').Value := FormatDateTime('YYYY-MM-DD', odata)
      end else
      begin
        oQuery.Parameters.ParamByName('DATA_ULTIMA_COMPRA').Value := '0000-00-00';
      end;

      oQuery.Parameters.ParamByName('OPERADOR').Value := RetornaOperador(OPERADORM);
      oQuery.Parameters.ParamByName('CLIENTE').Value := CLIENTE;
      oQuery.Parameters.ParamByName('MARCA').Value := getCodigoMarca;
      oQuery.ExecSQL;

    finally
      FreeAndNil(oQuery);
    end;
  except
    on E: Exception do
      InsereLog('Problema ao inserir marcas - ' + e.ClassName + ' - ' + e.Message + ' - Cliente: ' + CLIENTE  , mtError);
  end;
end;

procedure TfrmImportarDireto.SobrescreverCompra(const pCliente, pDescricaoCompra: string);
var
  lQueryExisteCompra: TADOQuery;
begin
  if cdsConfigSOBRESCREVER_COMPRAS.AsString <> 'S' then
  begin
    Exit;
  end;

  try
    lQueryExisteCompra := TADOQuery.Create(nil);
    lQueryExisteCompra.Connection := conexaoSGR;
    try
      lQueryExisteCompra.Close;
      lQueryExisteCompra.SQL.Clear;
      lQueryExisteCompra.SQL.Add('SELECT CODIGO ' +
        'FROM COMPRAS WHERE CLIENTE = ' + pCliente + ' AND DESCRICAO = ' +
        QuotedStr(pDescricaoCompra));
      lQueryExisteCompra.Open;

      if not lQueryExisteCompra.Eof then
      begin
        DeletaCompra(lQueryExisteCompra.FieldByName('CODIGO').AsInteger);
      end;
    finally
      FreeAndNil(lQueryExisteCompra);
    end;
  except
    on e:exception do
    begin
      InsereLog('Problemas ao buscar compras no impulse: ' + e.classname + ' - '
                + e.Message, mtError);
      InsereLog('Cliente: ' + pCliente + '; Desc.Compra: ' + pDescricaoCompra);
    end;
  end;
end;

procedure TfrmImportarDireto.DeletaCompra(pCompra: Integer);
var
  qrDelCompra : TADOQuery;
begin
  Try
    try
      qrDelCompra := TADOQuery.Create(nil);
      qrDelCompra.Connection := conexaoSGR;

      qrDelCompra.Close;
      qrDelCompra.SQL.Clear;
      qrDelCompra.SQL.Add('DELETE FROM itens_compra WHERE NOTA = ' + IntTostr(pCompra));
      qrDelCompra.ExecSql;

      qrDelCompra.Close;
      qrDelcompra.SQL.Clear;
      qrDelCompra.SQL.Add('DELETE FROM compras WHERE CODIGO = ' + IntToStr(pCompra));
      qrDelCompra.ExecSQL;
    except
      on e:exception do
      begin
        InsereLog('Erro ao excluir compra (' + IntToStr(pCompra) + ') - ' + e.ClassName + ' - ' + e.Message, mtError);
        InsereLog('Comando: ' + qrDelCompra.SQL.Text);
      end;
    end;
  Finally
    FreeAndNil(qrDelCompra);
  End;
end;

procedure TfrmImportarDireto.PopularcdsConfig(const psArquivo: string);
var
  lCdsAux: TClientDataSet;
  lIndice: Integer;
begin
  lCdsAux:= TClientDataSet.Create(nil);
  try
    lCdsAux.LoadFromFile(psArquivo);
    lCdsAux.Open;

    for lIndice := 0 to Pred(lCdsAux.Fields.Count) do
    begin
      if Assigned(cdsConfig.FindField(lCdsAux.Fields[lIndice].FieldName)) then
      begin
        cdsConfig.FindField(lCdsAux.Fields[lIndice].FieldName).Value :=
          lCdsAux.Fields[lIndice].Value;
      end;
    end;
  finally
    lCdsAux.Free;
  end;
end;

function TfrmImportarDireto.ObterValorDataSet(oDataSet: TDataSet;
  const pNome: String): string;
begin
  try
    if Assigned(oDataSet.FindField(pNome)) then
      Result := oDataSet.FieldByName(pNome).AsString
    else
      Result := EmptyStr;
  except
    Result := EmptyStr;
    log.Lines.Add('Erro: ObterValorDataSet: ' + pNome);
  end;
end;

procedure TfrmImportarDireto.DefinirMidia(poQuery: TADOQuery);
const
  sNOME_CAMPO = 'COD_MIDIA';
  sOBS_ADMIN = 'OBS_ADMIN';
var
  sMidia: string;
  qry: TADOQuery;
  ncod_unidade: integer;
begin
  if not Assigned(poQuery.Parameters.FindParam(sNOME_CAMPO)) then
    Exit;

  sMidia := ObterValorDataSet(FDatataModuloConexao.GetDataSouceClientes.DataSet, sNOME_CAMPO);
  ncod_unidade := 0;

  if sMidia <> EmptyStr then
  begin
    qry := ExecSql(' select CODIGO, cod_unidade from midias where Upper(NOME) = ' + QuotedStr(AnsiUpperCase(Trim(sMidia))));
    try
      if not qry.IsEmpty then
      begin
        poQuery.Parameters.ParamByName(sNOME_CAMPO).Value := qry.FieldByName('CODIGO').AsInteger;
        ncod_unidade := qry.FieldByName('cod_unidade').AsInteger;
      end;
    finally
      qry.free;
    end;
  end;

  if Assigned(poQuery.Parameters.FindParam(sOBS_ADMIN)) then
    poQuery.Parameters.ParamByName(sOBS_ADMIN).Value := 'Cliente cadastrado pelo site via: ' + sMidia;

  if (ncod_unidade > 0) and (Assigned(poQuery.Parameters.FindParam('COD_UNIDADE'))) then
  begin
    qry := ExecSql(' select max(CODIGO) as codigo from campanhas where unidade = ' + QuotedStr(ncod_unidade.ToString));
    try
      if (not qry.IsEmpty) and (qry.FieldByName('codigo').AsInteger > 0) then
      begin
        poQuery.Parameters.ParamByName('COD_UNIDADE').Value := ncod_unidade;
      end;
    finally
      qry.Free;
    end;
  end;
end;

function TfrmImportarDireto.GetPrioridadeAgenda(poQuery: TADOQuery): string;
const
  sNOME_CAMPO = 'COD_MIDIA';
var
  sMidia: string;
begin
  Result := '0';

  if not Assigned(poQuery.Parameters.FindParam(sNOME_CAMPO)) then
    Exit;

  if VarIsNull(poQuery.Parameters.ParamByName(sNOME_CAMPO).Value) then
    Exit;

  sMidia := poQuery.Parameters.ParamByName(sNOME_CAMPO).Value;

  with ExecSql(' select PRIORIDADE from midias where CODIGO = ' + sMidia) do
  try
    if FieldByName('PRIORIDADE').IsNull then
      Exit;

    Result := FieldByName('PRIORIDADE').AsString;
  finally
    free;
  end;
end;

function TfrmImportarDireto.GetCampanha(poQuery: TADOQuery): string;
begin
  Result := '0';

  if (Assigned(poQuery.Parameters.FindParam('COD_UNIDADE'))) then
    Result := IntToStr(poQuery.Parameters.ParamByName('COD_UNIDADE').Value);
end;

function TfrmImportarDireto.GetOPERADOR(poQuery: TADOQuery; const psOperador: string): string;
const
  sNOME_CAMPO = 'COD_MIDIA';
begin
  Result := psOperador;

  if not Assigned(poQuery.Parameters.FindParam(sNOME_CAMPO)) then
    Exit;

  with ExecSql('select OPERADOR_QUALIFICADOR from parametros') do
  try
    if not FieldByName('OPERADOR_QUALIFICADOR').IsNull then
      Result := FieldByName('OPERADOR_QUALIFICADOR').AsString;
  finally
    free;
  end;

end;

function TfrmImportarDireto.getUtilizaRegua: Boolean;
begin
  Result := tblParametros.FieldByName('UTILIZA_REGUA').AsString = 'S';
end;

end.

