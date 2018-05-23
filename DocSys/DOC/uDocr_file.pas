unit uDocr_file;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, ExtCtrls, DB, ADODB, Buttons;

type
  TfmDocr_file = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    cb_dep: TComboBox;
    DBGrid1: TDBGrid;
    Panel3: TPanel;
    DBGrid2: TDBGrid;
    Panel4: TPanel;
    sb_remove: TSpeedButton;
    sb_add: TSpeedButton;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    lab_dst01: TLabel;
    lab_dst02: TLabel;
    Label7: TLabel;
    lab_dst03: TLabel;
    btn_all: TButton;
    DOCR_FILE: TADOQuery;
    DOCR_FILEDOCR01: TWideStringField;
    DOCR_FILEDOCR02: TWideStringField;
    DOCR_FILEDOCR03: TWideStringField;
    DOCR_FILEDOCR04: TWideStringField;
    DOCR_FILEDOCR05: TWideStringField;
    DOCR_FILEDOCR_USER: TWideStringField;
    DOCR_FILEDOCR_DATE: TDateTimeField;
    DOCR_FILEDOCR_ACTI: TWideStringField;
    DOCR_FILEDOCR_FAC: TWideStringField;
    DataSource1: TDataSource;
    docr1_file: TADOQuery;
    Button1: TButton;
    DataSource2: TDataSource;
    docr1_fileGEN01: TStringField;
    docr1_fileGEN02: TStringField;
    Label8: TLabel;
    edt_name: TEdit;
    Label9: TLabel;
    lab_repdesc: TLabel;
    DOCR_FILEGEN02: TStringField;
    Label10: TLabel;
    cb_fac: TComboBox;
    cb_fac1: TComboBox;
    Label11: TLabel;
    DOCR_FILEDOCR06: TWideStringField;
    procedure sb_removeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dbinit;
    procedure sb_addClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btn_allClick(Sender: TObject);
    procedure cb_facChange(Sender: TObject);
    procedure cb_fac1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmDocr_file: TfmDocr_file;
  tmpque1 : TAdoquery;
implementation
uses docdm, uDocMain;
{$R *.dfm}

procedure TfmDocr_file.dbinit;
begin
  screen.Cursor:=crsqlwait;
  docr1_file.Close;
  docr1_file.SQL.Clear;
  //docr1_file.sql.Add('Select gen01,gen02 from gen_file'+dblinks+' where gen03 like :gen03 and gen02 like :gen02 ');
  docr1_file.sql.Add('Select gen01,gen02 from gen_file@'+copy(cb_fac.Text,1,pos(' ',cb_fac.text)-1)+'erp where gen03 like :gen03 and gen02 like :gen02 ');
  docr1_file.sql.add(' and gen01 not in (select docr04 from docr_file where docr_acti=''Y'' and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03) and genacti =''Y'' order by gen05');
  if trim(cb_dep.Text) = '' then
    docr1_file.Parameters.ParamByName('gen03').Value := '%'
  else
    docr1_file.Parameters.ParamByName('gen03').Value := copy(cb_dep.Text,1,pos(' ',cb_dep.Text)-1)+'%';

  if trim(edt_name.Text) = '' then begin                                                                            ;
    docr1_file.Parameters.ParamByName('gen02').Value := '%'
  end
  else begin
    docr1_file.Parameters.ParamByName('gen02').Value := trim(edt_name.text)+'%';
  end;
  docr1_file.Parameters.ParamByName('docr01').Value := lab_dst01.Caption;
  docr1_file.Parameters.ParamByName('docr02').Value := lab_dst02.Caption;
  docr1_file.Parameters.ParamByName('docr03').Value := lab_dst03.Caption;
  docr1_file.Open;
  docr1_file.First;

  docr_file.Close;
  docr_file.SQL.Clear;
  docr_file.sql.Add('Select docr_file.*,(select contact from wp_users where fac=docr06 and user_id=docr04) gen02 from docr_file where docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_acti=''Y'' and docr_fac =:docr_fac');
  docr_file.parameters.parambyname('docr01').value := lab_dst01.Caption;
  docr_file.parameters.parambyname('docr02').value := lab_dst02.Caption;
  docr_file.parameters.parambyname('docr03').value := lab_dst03.caption;
  docr_file.parameters.parambyname('docr_fac').value := docdm.myfac;
  if cb_fac1.Text<>'全部' then  begin
    docr_file.Filtered:=true;
    docr_file.Filter:='docr06 = ' + QuotedStr(copy(cb_fac1.Text,1,pos(' ',cb_fac1.text)-1));
  end
  else begin
    docr_file.Filtered:=false;
    docr_file.Filter:='';
  end;
  docr_file.open;
  docr_file.first;

  screen.Cursor:=crdefault;

  {docr_file.Close;
  docr_file.SQL.Clear;
  docr_file.sql.Add('Select docr_file.*, gen02 from docr_file,gen_file'+docdm.dblinks+' where gen01 = docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_acti=''Y'' and docr_fac =:docr_fac');
  docr_file.parameters.parambyname('docr01').value := lab_dst01.Caption;
  docr_file.parameters.parambyname('docr02').value := lab_dst02.Caption;
  docr_file.parameters.parambyname('docr03').value := lab_dst03.caption;
  docr_file.parameters.parambyname('docr_fac').value := docdm.myfac;
  docr_file.open;
  docr_file.first;     }
end;

procedure TfmDocr_file.sb_removeClick(Sender: TObject);
var
  trows: integer;
begin
  if (not docr_file.Active) or (docr_file.RecordCount = 0) then begin
    showmessage('請先選擇要取消權限的人員!');
    Exit;
  end;


  try
    if not dm.glm.InTransaction then
      dm.glm.BeginTrans;

    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('update docr_file set docr_acti =''N'' where docr04 =:docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_fac =:docr_fac' );
    tmpque1.Parameters.ParamByName('docr01').Value := lab_dst01.Caption;
    tmpque1.Parameters.ParamByName('docr02').Value := lab_dst02.Caption;
    tmpque1.Parameters.ParamByName('docr03').Value := lab_dst03.Caption;
    tmpque1.Parameters.ParamByName('docr_fac').Value := docdm.myfac;
    tmpque1.Parameters.ParamByName('docr04').Value := docr_file.fieldbyname('docr04').AsString;
    trows := tmpque1.ExecSQL;

    if trows = 1 then begin
      if dm.glm.InTransaction then
        dm.glm.CommitTrans;
      showmessage('異動成功!')
    end
    else begin
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('異動失敗!')
    end;
  except
    if dm.glm.InTransaction then
      dm.glm.RollbackTrans;
    showmessage('異動失敗!')
  end;
  dbinit;
end;

procedure TfmDocr_file.FormShow(Sender: TObject);
begin
  tmpque1 := TAdoquery.Create(self);
  tmpque1.Connection := dm.glm;

  cb_fac.clear;
  cb_fac1.clear;
  cb_fac1.Items.Append('全部');
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select fac00,fac01 From fac_file order by fac00');
  tmpque1.Open;
  while not tmpque1.eof do begin
    cb_fac.Items.Append(tmpque1.fieldbyname('fac00').AsString+' '+tmpque1.fieldbyname('fac01').AsString);
    cb_fac1.Items.Append(tmpque1.fieldbyname('fac00').AsString+' '+tmpque1.fieldbyname('fac01').AsString);
    tmpque1.Next;
  end;
  cb_fac.Text:=fmDocMain.lab_fac.Caption;
  cb_fac1.Text:='全部';

  cb_dep.clear;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select gem01,gem02 From gem_file@'+copy(cb_fac.Text,1,pos(' ',cb_fac.text)-1)+'erp where (length(gem01) = 3  or length(gem01) = 2) and gem07 is not null and gemacti =''Y'' order by gem01');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dep.Items.Append(tmpque1.fieldbyname('gem01').AsString+' '+tmpque1.fieldbyname('gem02').AsString);
    tmpque1.Next;
  end;

  {cb_dep.clear;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select gem01,gem02 From gem_file'+docdm.dblinks+' where (length(gem01) = 3  or length(gem01) = 2) and gem07 is not null and gemacti =''Y'' order by gem01');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dep.Items.Append(tmpque1.fieldbyname('gem01').AsString+' '+tmpque1.fieldbyname('gem02').AsString);
    tmpque1.Next;
  end;}

  docr_file.Close;
  docr_file.SQL.Clear;
  docr_file.sql.Add('Select docr_file.*,(select contact from wp_users where fac=docr06 and user_id=docr04) gen02 from docr_file,gen_file'+docdm.dblinks+' where gen01 = docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_acti=''Y'' and docr_fac =:docr_fac');
  docr_file.parameters.parambyname('docr01').value := lab_dst01.Caption;
  docr_file.parameters.parambyname('docr02').value := lab_dst02.Caption;
  docr_file.parameters.parambyname('docr03').value := lab_dst03.caption;
  docr_file.parameters.parambyname('docr_fac').value := docdm.myfac;
  if cb_fac1.Text<>'全部' then  begin
    docr_file.Filter:='docr06 = ' + QuotedStr(copy(cb_fac1.Text,1,pos(' ',cb_fac1.text)-1));
    docr_file.Filtered:=true;
  end
  else begin
    docr_file.Filtered:=false;
    docr_file.Filter:='';
  end;
  docr_file.open;
  docr_file.first;

  {docr_file.Close;
  docr_file.SQL.Clear;
  docr_file.sql.Add('Select docr_file.*,gen02 from docr_file,gen_file'+docdm.dblinks+' where gen01 = docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_acti=''Y'' and docr_fac =:docr_fac');
  docr_file.parameters.parambyname('docr01').value := lab_dst01.Caption;
  docr_file.parameters.parambyname('docr02').value := lab_dst02.Caption;
  docr_file.parameters.parambyname('docr03').value := lab_dst03.caption;
  docr_file.parameters.parambyname('docr_fac').value := docdm.myfac;
  docr_file.open;
  docr_file.first; }


end;

procedure TfmDocr_file.sb_addClick(Sender: TObject);
var
  mymail : string;
begin
  if (not docr1_file.Active) or (docr1_file.RecordCount = 0) then begin
    showmessage('請先查出要設定權限的人員!');
    Exit;
  end;
  
  tmpque1.Close;
  tmpque1.sql.Clear;
  tmpque1.sql.Add('Select * from docr_file where docr04 =:docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_fac =:docr_fac');
  tmpque1.parameters.parambyname('docr01').value := lab_dst01.Caption;
  tmpque1.parameters.parambyname('docr02').value := lab_dst02.Caption;
  tmpque1.parameters.parambyname('docr03').value := lab_dst03.caption;
  tmpque1.Parameters.ParamByName('docr04').Value := docr1_file.fieldbyname('gen01').AsString;
  tmpque1.parameters.parambyname('docr_fac').value := docdm.myfac;
  tmpque1.Open;
  tmpque1.First;
  if tmpque1.RecordCount > 0 then begin
    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('update docr_file set docr_acti =''Y'' where docr04 =:docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_fac =:docr_fac' );
    tmpque1.Parameters.ParamByName('docr01').Value := lab_dst01.Caption;
    tmpque1.Parameters.ParamByName('docr02').Value := lab_dst02.Caption;
    tmpque1.Parameters.ParamByName('docr03').Value := lab_dst03.Caption;
    tmpque1.Parameters.ParamByName('docr_fac').Value := docdm.myfac;
    tmpque1.Parameters.ParamByName('docr04').Value := docr1_file.fieldbyname('gen01').AsString;
    tmpque1.ExecSQL;
  end
  else begin
    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('select e_mail from wp_users where user_id ='''+docr1_file.fieldbyname('gen01').AsString+''' and fac ='''+copy(cb_fac.Text,1,pos(' ',cb_fac.text)-1)+'''');
    tmpque1.Open;
    tmpque1.First;
    if tmpque1.RecordCount > 0 then
      mymail := tmpque1.fieldbyname('e_mail').AsString
    else
      mymail := '';

    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('insert into docr_file (docr01,docr02,docr03,docr04,docr05,docr06,docr_user,docr_date,docr_acti,docr_fac) values (:docr01,:docr02,:docr03,:docr04,:docr05,:docr06,'''+dm.getlogin+''',sysdate,''Y'',:docr_fac) ');
    tmpque1.Parameters.ParamByName('docr01').Value := lab_dst01.Caption;
    tmpque1.Parameters.ParamByName('docr02').Value := lab_dst02.Caption;
    tmpque1.Parameters.ParamByName('docr03').Value := lab_dst03.Caption;
    tmpque1.Parameters.ParamByName('docr_fac').Value := docdm.myfac;
    tmpque1.Parameters.ParamByName('docr04').Value := docr1_file.fieldbyname('gen01').AsString;
    tmpque1.Parameters.ParamByName('docr05').Value := mymail;
    tmpque1.Parameters.ParamByName('docr06').Value := copy(cb_fac.Text,1,pos(' ',cb_fac.text)-1);
    tmpque1.ExecSQL;
  end;
  dbinit;
end;

procedure TfmDocr_file.Button1Click(Sender: TObject);
begin
  dbinit;
end;

procedure TfmDocr_file.btn_allClick(Sender: TObject);
var
  mymail : string;
begin
  if MessageDlg('請注意設定全體人員將會設定此部門所有的員工有此查詢報表權限,是否執行?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then begin
    Exit;
  end;

  dbinit;

  docr1_file.first;
  while not docr1_file.eof do begin
    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('Select * from docr_file where docr04 =:docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_fac =:docr_fac');
    tmpque1.parameters.parambyname('docr01').value := lab_dst01.Caption;
    tmpque1.parameters.parambyname('docr02').value := lab_dst02.Caption;
    tmpque1.parameters.parambyname('docr03').value := lab_dst03.caption;
    tmpque1.parameters.parambyname('docr_fac').value := docdm.myfac;
    tmpque1.Parameters.ParamByName('docr04').Value := docr1_file.fieldbyname('gen01').AsString;
    tmpque1.Open;
    tmpque1.First;
    if tmpque1.RecordCount > 0 then begin
      tmpque1.Close;
      tmpque1.sql.Clear;
      tmpque1.sql.Add('update docr_file set docr_acti =''Y'' where docr04 =:docr04 and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_fac =:docr_fac' );
      tmpque1.Parameters.ParamByName('docr01').Value := lab_dst01.Caption;
      tmpque1.Parameters.ParamByName('docr02').Value := lab_dst02.Caption;
      tmpque1.Parameters.ParamByName('docr03').Value := lab_dst03.Caption;
      tmpque1.Parameters.ParamByName('docr_fac').Value := docdm.myfac;
      tmpque1.Parameters.ParamByName('docr04').Value := docr1_file.fieldbyname('gen01').AsString;
      tmpque1.ExecSQL;
    end
    else begin
      tmpque1.Close;
      tmpque1.sql.Clear;
      tmpque1.sql.Add('select e_mail from wp_users where user_id ='''+docr1_file.fieldbyname('gen01').AsString+''' and fac ='''+docdm.myfac+'''');
      tmpque1.Open;
      tmpque1.First;
      if tmpque1.RecordCount > 0 then
        mymail := tmpque1.fieldbyname('e_mail').AsString
      else
        mymail := '';

      tmpque1.Close;
      tmpque1.sql.Clear;
      tmpque1.sql.Add('insert into docr_file (docr01,docr02,docr03,docr04,docr05,docr06,docr_user,docr_date,docr_acti,docr_fac) values (:docr01,:docr02,:docr03,:docr04,:docr05,:docr06,'''+dm.getlogin+''',sysdate,''Y'',:docr_fac) ');
      tmpque1.Parameters.ParamByName('docr01').Value := lab_dst01.Caption;
      tmpque1.Parameters.ParamByName('docr02').Value := lab_dst02.Caption;
      tmpque1.Parameters.ParamByName('docr03').Value := lab_dst03.Caption;
      tmpque1.Parameters.ParamByName('docr_fac').Value := docdm.myfac;
      tmpque1.Parameters.ParamByName('docr04').Value := docr1_file.fieldbyname('gen01').AsString;
      tmpque1.Parameters.ParamByName('docr05').Value := mymail;
      tmpque1.Parameters.ParamByName('docr06').Value := copy(cb_fac.Text,1,pos(' ',cb_fac.text)-1);
      tmpque1.ExecSQL;
    end;
    docr1_file.Next;
  end;
  dbinit;
end;

procedure TfmDocr_file.cb_facChange(Sender: TObject);
begin
  screen.Cursor:=crsqlwait;
  cb_dep.clear;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select gem01,gem02 From gem_file@'+copy(cb_fac.Text,1,pos(' ',cb_fac.text)-1)+'erp where (length(gem01) = 3  or length(gem01) = 2) and gem07 is not null and gemacti =''Y'' order by gem01');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dep.Items.Append(tmpque1.fieldbyname('gem01').AsString+' '+tmpque1.fieldbyname('gem02').AsString);
    tmpque1.Next;
  end;
  screen.Cursor:=crDefault;
end;

procedure TfmDocr_file.cb_fac1Change(Sender: TObject);
begin
  screen.Cursor:=crsqlwait;
  docr_file.Close;
  docr_file.SQL.Clear;
  docr_file.sql.Add('Select docr_file.*,(select contact from wp_users where fac=docr06 and user_id=docr04) gen02 from docr_file where docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_acti=''Y'' and docr_fac =:docr_fac');
  docr_file.parameters.parambyname('docr01').value := lab_dst01.Caption;
  docr_file.parameters.parambyname('docr02').value := lab_dst02.Caption;
  docr_file.parameters.parambyname('docr03').value := lab_dst03.caption;
  docr_file.parameters.parambyname('docr_fac').value := docdm.myfac;
  if cb_fac1.Text<>'全部' then  begin
    docr_file.Filter:='docr06 = ' + QuotedStr(copy(cb_fac1.Text,1,pos(' ',cb_fac1.text)-1));
    docr_file.Filtered:=true;
  end
  else begin
    docr_file.Filtered:=false;
    docr_file.Filter:='';
  end;
  docr_file.open;
  screen.Cursor:=crDefault;
end;

end.
