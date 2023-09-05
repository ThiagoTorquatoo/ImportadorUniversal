unit uAjuda;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TfrmAjuda = class(TForm)
    pnlButton: TPanel;
    pnlDados: TPanel;
    mmoAjuda: TMemo;
    btnFechar: TButton;
    procedure btnFecharClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure showDlg(texto: TStrings);
  end;

var
  frmAjuda: TfrmAjuda;

implementation

{$R *.dfm}

procedure TfrmAjuda.btnFecharClick(Sender: TObject);
begin
   ModalResult := mrok;
end;

procedure TfrmAjuda.showDlg(texto: TStrings);
begin
  mmoAjuda.Lines := texto;
  ShowModal;
end;

end.
