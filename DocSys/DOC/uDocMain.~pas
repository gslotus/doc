unit uDocMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, Adodb, Menus;

type
  TfmDocMain = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    lab_fac: TLabel;
    lab_users: TLabel;
    lab_yearmonth: TLabel;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    IFRS1: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure IFRS1Click(Sender: TObject);
  private
    { Private declarations }
  public

    { Public declarations }
  end;

var
  fmDocMain: TfmDocMain;
  tmpaqu1 : Tadoquery;
implementation
uses docdm,uDoc_type, uDoc_Setup, docque, uCheDoc;
{$R *.dfm}

procedure TfmDocMain.FormShow(Sender: TObject);
var
  username : string;
  year,month,day :word;
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOCSYS'' and userid =:userid and opcode =''RUN''');
  dm.dbque.Parameters.ParamByName('userid').Value := DM.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if dm.dbque.Eof then begin
    showmessage('您無權限執行文件電子報系統(DOCSYS)');
    CLOSE;
  end
  else begin
    decodedate(now,year,month,day);
    username := dm.getlogin;
    tmpaqu1 := Tadoquery.Create(self);
    tmpaqu1.Connection := dm.glm;
    tmpaqu1.Close;
    tmpaqu1.sql.Clear;
    tmpaqu1.SQL.Add('select contact,fac,fac01 from wp_users,fac_file where fac=fac00 and user_id='''+username+'''');
    tmpaqu1.Open;
    if tmpaqu1.RecordCount > 0 then begin
      docdm.dblinks := '@'+tmpaqu1.fieldbyname('fac').AsString+'erp';
      docdm.myfac := tmpaqu1.fieldbyname('fac').AsString;
      lab_fac.Caption := tmpaqu1.fieldbyname('fac').AsString+' '+tmpaqu1.fieldbyname('fac01').AsString;
      lab_users.Caption := dm.getlogin;
      lab_yearmonth.Caption := inttostr(year)+formatfloat('00',month);
      tmpaqu1.Close;
      tmpaqu1.sql.Clear;
      tmpaqu1.SQL.Add('select GEN03 from GEN_file'+docdm.dblinks+' where gen01 ='''+username+'''');
      tmpaqu1.Open;
      docdm.mygem := copy(tmpaqu1.fieldbyname('GEN03').asstring,1,3);
    end
    else begin
      showmessage('您無權限使用文件電子報系統(DOCSYS)!');
      close;
    end;
  end;
end;

procedure TfmDocMain.N1Click(Sender: TObject);
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOC01'' and userid =:userid and opcode =''RUN''');
  dm.dbque.Parameters.ParamByName('userid').Value := DM.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if dm.dbque.Eof then begin
    showmessage('您無權限執行文件類別設定(DOC01)');
  end
  else begin
    fmDoc_type := TfmDoc_type.Create(self);
    fmDoc_type.ShowModal;
    fmDoc_type.Free;
  end;
end;

procedure TfmDocMain.N2Click(Sender: TObject);
begin
  close;
end;

procedure TfmDocMain.N3Click(Sender: TObject);
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOC02'' and userid =:userid and opcode =''RUN''');
  dm.dbque.Parameters.ParamByName('userid').Value := DM.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if dm.dbque.Eof then begin
    showmessage('您無權限執行文件基本資料定義(DOC02)');
  end
  else begin
    fmDoc_setup := TfmDoc_setup.Create(self);
    fmDoc_setup.ShowModal;
    fmDoc_setup.Free;
  end;
end;

procedure TfmDocMain.N4Click(Sender: TObject);
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOC03'' and userid =:userid and opcode =''RUN''');
  dm.dbque.Parameters.ParamByName('userid').Value := DM.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if dm.dbque.Eof then begin
    showmessage('您無權限執行設定文件讀取權限(DOC03)');
  end
  else begin
    fmDoc_setup := TfmDoc_setup.Create(self);
    fmDoc_setup.ShowModal;
    fmDoc_setup.Free;
  end;
end;

procedure TfmDocMain.N5Click(Sender: TObject);
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOC04'' and userid =:userid and opcode =''RUN''');
  dm.dbque.Parameters.ParamByName('userid').Value := DM.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if dm.dbque.Eof then begin
    showmessage('您無權限執行文件維護作業(DOC04)');
  end
  else begin
    fmquery := TfmQuery.create(self);
    docque.status := '';
    fmQuery.Panel1.Visible := True;
    fmQuery.btn_sign.Visible := False;
    fmQuery.showmodal;
    fmQuery.free;
  end;
end;

procedure TfmDocMain.N6Click(Sender: TObject);
begin
  tmpaqu1.Close;
  tmpaqu1.SQL.Clear;
  tmpaqu1.SQL.Add('select * from doc_file where doc_acti=''Y'' and docsign_np ='''+dm.getlogin+'''');
  tmpaqu1.Open;
  if tmpaqu1.RecordCount = 0 then begin
    showmessage('無須簽核的記錄!');
    Exit;
  end
  else begin
    fmquery := TfmQuery.create(self);
    docque.status := 'SIGN';
    docque.stypes := 'DOC';
    fmQuery.Panel1.Visible := False;
    fmQuery.btn_sign.Visible := True;
    fmQuery.showmodal;
    fmQuery.free;
  end;

end;

procedure TfmDocMain.N7Click(Sender: TObject);
begin
  tmpaqu1.Close;
  tmpaqu1.SQL.Clear;
  tmpaqu1.SQL.Add('select * from doc_file where doc_acti=''Y'' and websign_np ='''+dm.getlogin+'''');
  tmpaqu1.Open;
  if tmpaqu1.RecordCount = 0 then begin
    showmessage('無須簽核的記錄!');
    Exit;
  end
  else begin
    fmquery := TfmQuery.create(self);
    docque.status := 'SIGN';
    docque.stypes := 'WEB';
    fmQuery.Panel1.Visible := False;
    fmQuery.btn_sign.Visible := True;
    fmQuery.showmodal;
    fmQuery.free;
  end;

end;

procedure TfmDocMain.N8Click(Sender: TObject);
begin
  fmCheDep := TfmCheDep.create(self);
  fmCheDep.ShowModal;
  fmCheDep.free;
end;

procedure TfmDocMain.IFRS1Click(Sender: TObject);
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOCIFRS'' and userid =:userid and opcode =''RUN''');
  dm.dbque.Parameters.ParamByName('userid').Value := DM.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if dm.dbque.Eof then
    showmessage('您無權限執行IFRS文件維護作業(DOCIFRS)')
  ELSE begin
    {fmifrs := Tfmifrs.Create(self);
    fmifrs.ShowModal;
    fmifrs.Free;}
  end;

end;

end.
