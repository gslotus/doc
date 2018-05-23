unit mail;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Psock, NMsmtp, ComCtrls, Adodb,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdMessageClient, IdSMTP;

type
  TfmMail = class(TForm)
    Label4: TLabel;
    Edit4: TEdit;
    Label5: TLabel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Smtp: TNMSMTP;
    RichEdit1: TRichEdit;
    RichEdit2: TRichEdit;
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SmtpFailure(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmMail: TfmMail;
  froms:String;
  flag:boolean;
  tmpque:TAdoquery;
implementation
uses docdm;
{$R *.dfm}

procedure TfmMail.BitBtn2Click(Sender: TObject);
begin
  close;
end;

procedure TfmMail.BitBtn1Click(Sender: TObject);
begin
  richedit1.Lines.Append('郵件傳遞中請稍候......');
  smtp.PostMessage.FromName := '鉅祥企業訊息中心';
  smtp.PostMessage.fromAddress :=  'b2b@mail.gs.com.tw';
  smtp.PostMessage.Body.AddStrings(Richedit2.Lines);
  smtp.PostMessage.Subject := Trim(Edit4.Text);
  if not smtp.Connected then
    smtp.Connect;
  if not smtp.Connected then
    showmessage('無法連接郵件伺服器,請稍候再試!')
  else begin
    smtp.SendMail;
    flag:=true;
    showmessage('郵件已傳送!');
    BitBtn1.Enabled := False;
    close;
  end;
  richedit1.Clear;
end;

procedure TfmMail.FormShow(Sender: TObject);
begin
  flag:=false;
  richedit1.Clear;
  fmMail.Caption := fmMail.Caption+'--'+dm.getlogin;
  tmpque := Tadoquery.Create(self);
  tmpque.Connection := dm.glm;
  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.SQL.Add('select type_detail from sys_codedef'+docdm.dblinks+' where cd_type=''10''');
  tmpque.Open;
  if tmpque.RecordCount > 0 then
    smtp.Host := tmpque.fieldbyname('type_detail').AsString;


end;

procedure TfmMail.SmtpFailure(Sender: TObject);
begin
  close;
end;

end.
