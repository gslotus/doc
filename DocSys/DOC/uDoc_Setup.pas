unit uDoc_Setup;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, ADODB, ExtCtrls;

type
  TfmDoc_setup = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Panel2: TPanel;
    Panel5: TPanel;
    Panel3: TPanel;
    Label7: TLabel;
    cb_type011: TComboBox;
    btn_que: TButton;
    doc_setup: TADOQuery;
    DataSource1: TDataSource;
    Label6: TLabel;
    cb_type021: TComboBox;
    Label8: TLabel;
    edt_dst03: TEdit;
    Label9: TLabel;
    edt_dst04: TEdit;
    Label10: TLabel;
    Label11: TLabel;
    DBGrid1: TDBGrid;
    cb_dst05: TComboBox;
    cb_dst06: TComboBox;
    Label13: TLabel;
    Label14: TLabel;
    ed_dst08: TEdit;
    Label15: TLabel;
    cb_dst071: TComboBox;
    cb_dst01: TComboBox;
    cb_dst02: TComboBox;
    cb_dst07: TComboBox;
    doc_setupDST01: TWideStringField;
    doc_setupDST02: TWideStringField;
    doc_setupDST_ACTI: TWideStringField;
    doc_setupDST03: TWideStringField;
    doc_setupDST04: TWideStringField;
    doc_setupDST05: TWideStringField;
    doc_setupDST06: TWideStringField;
    doc_setupDST07: TWideStringField;
    doc_setupDST071: TWideStringField;
    doc_setupDST08: TBCDField;
    doc_setupDST_USER: TWideStringField;
    doc_setupDST_DATE: TDateTimeField;
    doc_setupDST_FAC: TWideStringField;
    Label3: TLabel;
    Panel4: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    doc_setupdst01desc: TStringField;
    doc_setupdst02desc: TStringField;
    doc_setupdst05desc: TStringField;
    doc_setupdst06desc: TStringField;
    Label4: TLabel;
    cb_sign: TComboBox;
    doc_setupDST09: TWideStringField;
    btn_cp: TButton;
    Button6: TButton;
    procedure FormShow(Sender: TObject);
    procedure cb_type011Exit(Sender: TObject);
    procedure cb_dst01Exit(Sender: TObject);
    procedure cb_dst02Exit(Sender: TObject);
    procedure btn_queClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure cb_dst07Exit(Sender: TObject);
    procedure doc_setupAfterScroll(DataSet: TDataSet);
    procedure doc_setupCalcFields(DataSet: TDataSet);
    procedure btn_cpClick(Sender: TObject);
    procedure Button6Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmDoc_setup: TfmDoc_setup;
  tmpque1 : Tadoquery;
  status,strall : string;
implementation
uses docdm, uDocr_file, uCp;
{$R *.dfm}

procedure TfmDoc_setup.FormShow(Sender: TObject);
var
  mydep,dtype03 : string;
begin
  tmpque1 := TAdoquery.Create(self);
  tmpque1.Connection := dm.glm;

  cb_type011.Clear;
  cb_dst01.Clear;

  tmpque1.Close;
  tmpque1.sql.Clear;
  tmpque1.SQL.Add('select * from appdeta'+docdm.dblinks+' where appname =''DOC05'' and userid ='''+dm.getlogin+'''');
  tmpque1.Open;
  if tmpque1.RecordCount > 0 then
    strall := 'ALL'
  else
    strall := '';

  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('select gen03 from gen_file'+docdm.dblinks+' where gen01 ='''+dm.getlogin+'''');
  tmpque1.Open;
  mydep := copy(tmpque1.fieldbyname('gen03').AsString,1,3);

  tmpque1.Close;
  tmpque1.SQL.Clear;
  if strall = '' then
    tmpque1.SQL.Add('Select unique dtype01,dtype011 from doc_type where dtype_fac ='''+docdm.myfac+''' and dtype03 ='''+mydep+''' order by dtype01')
  else
    tmpque1.SQL.Add('Select unique dtype01,dtype011 from doc_type where dtype_fac ='''+docdm.myfac+''' and dtype03 like '''+copy(mydep,1,2)+'%'+''' order by dtype01');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dst01.Items.Append(tmpque1.fieldbyname('dtype01').AsString+' '+tmpque1.fieldbyname('dtype011').AsString);
    cb_type011.Items.Append(tmpque1.fieldbyname('dtype01').AsString+' '+tmpque1.fieldbyname('dtype011').AsString);
    tmpque1.Next;
  end;

  cb_dst05.Clear;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select subject,remark From sysparameter Where isActive = ''Y'' And sysname =''DOC'' And item =''DST05''');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dst05.Items.Append(tmpque1.fieldbyname('subject').AsString+' '+tmpque1.fieldbyname('remark').AsString);
    tmpque1.Next;
  end;

  cb_dst06.Clear;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select subject,remark From sysparameter Where isActive = ''Y'' And sysname =''DOC'' And item =''DST06''');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dst06.Items.Append(tmpque1.fieldbyname('subject').AsString+' '+tmpque1.fieldbyname('remark').AsString);
    tmpque1.Next;
  end;

  cb_dst07.Clear;
  cb_dst071.Clear;
  tmpque1.Close;
  tmpque1.sql.Clear;
  tmpque1.sql.Add('select dtype03 from doc_type where dtype01 ='''+copy(cb_dst01.Text,1,pos(' ',cb_dst01.Text)-1)+''' and dtype02 ='''+copy(cb_dst02.Text,1,pos(' ',cb_dst02.Text)-1)+''' and dtype_fac='''+docdm.myfac+'''');
  tmpque1.Open;
  if tmpque1.RecordCount > 0 then
    dtype03 := tmpque1.fieldbyname('dtype03').asstring
  else
    dtype03 := '';

  tmpque1.close;
  tmpque1.sql.Clear;
  if strall = '' then
    tmpque1.sql.Add('select gen02,gen01 from gen_file'+docdm.dblinks+' where genacti =''Y'' and gen03 like'''+mydep+'%'+''' order by gen05')
  else
    tmpque1.sql.Add('select gen02,gen01 from gen_file'+docdm.dblinks+' where genacti =''Y'' and gen03 like'''+copy(mydep,1,2)+'%'+''' order by gen05');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dst07.Items.Append(tmpque1.fieldbyname('gen01').AsString+' '+tmpque1.fieldbyname('gen02').AsString);
    tmpque1.Next;
  end;

end;

procedure TfmDoc_setup.cb_type011Exit(Sender: TObject);
begin
  cb_type021.Clear;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select dtype02,dtype021 from doc_type where dtype_fac ='''+docdm.myfac+''' and dtype01 ='''+copy(cb_type011.Text,1,pos(' ',cb_type011.Text)-1)+''' order by dtype02');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_type021.Items.Append(tmpque1.fieldbyname('dtype02').AsString+' '+tmpque1.fieldbyname('dtype021').AsString);
    tmpque1.Next;
  end;

end;

procedure TfmDoc_setup.cb_dst01Exit(Sender: TObject);
begin
  cb_dst02.Clear;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select dtype02,dtype021 from doc_type where dtype_fac ='''+docdm.myfac+''' and dtype01 ='''+copy(cb_dst01.Text,1,pos(' ',cb_dst01.Text)-1)+''' order by dtype02');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dst02.Items.Append(tmpque1.fieldbyname('dtype02').AsString+' '+tmpque1.fieldbyname('dtype021').AsString);
    tmpque1.Next;
  end;

end;

procedure TfmDoc_setup.cb_dst02Exit(Sender: TObject);
begin
  if status ='INS' then begin
    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('select count(*) from doc_setup where dst01 = '''+copy(cb_dst01.Text,1,pos(' ',cb_dst01.Text)-1)+''' and dst02 = '''+copy(cb_dst02.Text,1,pos(' ',cb_dst02.Text)-1)+'''');
    tmpque1.Open;
    edt_dst03.Text := formatfloat('00',tmpque1.Fields.Fields[0].AsInteger+1);
  end;

end;

procedure TfmDoc_setup.btn_queClick(Sender: TObject);
begin
  if trim(cb_type011.Text) = '' then begin
    showmessage('請輸入代碼大類!');
    Exit;
  end;
  
  status := '';
  doc_setup.Close;
  doc_setup.sql.Clear;
  doc_setup.SQL.Add('select * from doc_setup where dst01 like :dst01 and dst02 like :dst02 ');
  if trim(cb_type011.Text) = '' then
    doc_setup.Parameters.ParamByName('dst01').Value := '%'
  else
    doc_setup.Parameters.ParamByName('dst01').Value := copy(cb_type011.Text,1,pos(' ',cb_type011.Text)-1);

  if trim(cb_type021.Text) = '' then
    doc_setup.Parameters.ParamByName('dst02').Value := '%'
  else
    doc_setup.Parameters.ParamByName('dst02').Value := copy(cb_type021.Text,1,pos(' ',cb_type021.Text)-1);
  doc_setup.SQL.Add(' order by dst01,dst02,dst03 ');
  doc_setup.Open;
  doc_setup.First;

  panel1.Enabled := False;
   ed_dst08.Enabled := False;
  edt_dst03.ReadOnly := True;
  cb_dst01.Enabled := False;
  cb_dst02.Enabled := False;


end;

procedure TfmDoc_setup.Button1Click(Sender: TObject);
begin
  status := 'INS';
    panel1.Enabled := True;

  cb_dst01.ItemIndex := -1;
  cb_dst02.ItemIndex := -1;
  edt_dst03.Text := '';
  edt_dst04.Text := '';
  cb_dst05.ItemIndex := -1;
  cb_dst06.ItemIndex := -1;
  cb_dst07.ItemIndex := -1;
  cb_dst071.ItemIndex := -1;
  cb_sign.ItemIndex := -1;
  ed_dst08.Clear;
  ed_dst08.Enabled := True;
  edt_dst03.ReadOnly := False;
  cb_dst01.Enabled := True;
  cb_dst02.Enabled := True;
end;

procedure TfmDoc_setup.Button4Click(Sender: TObject);
var
  acts : string;
begin
  if (not doc_setup.Active) or (doc_setup.RecordCount = 0) then begin
    showmessage('請查詢要作廢的資料!');
    EXit;
  end;

  status := '';
  if doc_setup.FieldByName('dst_acti').AsString ='Y' then
    acts := ''
  else if doc_setup.FieldByName('dst_acti').AsString ='N' then
    acts := '取消';

  if MessageDlg('是否'+acts+'作廢?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    if trim(acts) = '' then begin
      tmpque1.Close;
      tmpque1.SQL.Clear;
      tmpque1.sql.Add('update doc_setup set dst_acti =''N'' where dst01 =:dst01 and dst02=:dst02 and dst03=:dst03 and dst_fac =:dst_fac');
    end
    else begin
      tmpque1.Close;
      tmpque1.SQL.Clear;
      tmpque1.sql.Add('update doc_setup set dst_acti =''Y'' where dst01 =:dst01 and dst02=:dst02 and dst03=:dst03 and dst_fac =:dst_fac');
    end;
    tmpque1.Parameters.ParamByName('dst01').Value := doc_setup.fieldbyname('dst01').AsString;
    tmpque1.Parameters.ParamByName('dst02').Value := doc_setup.fieldbyname('dst02').AsString;
    tmpque1.Parameters.ParamByName('dst03').Value := doc_setup.fieldbyname('dst03').AsString;
    tmpque1.Parameters.ParamByName('dst_fac').Value := doc_setup.fieldbyname('dst_fac').AsString;
    tmpque1.ExecSQL;
    showmessage('異動完成!');
  end;

  btn_que.Click;

end;

procedure TfmDoc_setup.Button2Click(Sender: TObject);
var
  i : integer;
begin
  status := 'EDT';
    panel1.Enabled := True;

  edt_dst03.Text := doc_setup.fieldbyname('dst03').AsString;
  edt_dst04.Text := doc_setup.fieldbyname('dst04').AsString;
  ed_dst08.Text := doc_setup.fieldbyname('dst08').AsString;

  for i := 0 to cb_dst01.Items.Count -1 do begin
    if copy(cb_dst01.Items[i],1,pos(' ',cb_dst01.Items[i])-1) = doc_setup.FieldByName('dst01').AsString then
      cb_dst01.ItemIndex := i;
  end;
  cb_dst01.OnExit(self);

  for i := 0 to cb_dst02.Items.Count -1 do begin
    if copy(cb_dst02.Items[i],1,pos(' ',cb_dst02.Items[i])-1) = doc_setup.FieldByName('dst02').AsString then
      cb_dst02.ItemIndex := i;
  end;

  for i := 0 to cb_dst05.Items.Count -1 do begin
    if copy(cb_dst05.Items[i],1,pos(' ',cb_dst05.Items[i])-1) = doc_setup.FieldByName('dst05').AsString then
      cb_dst05.ItemIndex := i;
  end;

  for i := 0 to cb_dst06.Items.Count -1 do begin
    if copy(cb_dst06.Items[i],1,pos(' ',cb_dst06.Items[i])-1) = doc_setup.FieldByName('dst06').AsString then
      cb_dst06.ItemIndex := i;
  end;

  for i := 0 to cb_dst07.Items.Count -1 do begin
    if copy(cb_dst07.Items[i],1,pos(' ',cb_dst07.Items[i])-1) = doc_setup.FieldByName('dst07').AsString then
      cb_dst07.ItemIndex := i;
  end;
  cb_dst07.OnExit(self);

  for i := 0 to cb_dst071.Items.Count -1 do begin
    if copy(cb_dst071.Items[i],1,pos(' ',cb_dst071.Items[i])-1) = doc_setup.FieldByName('dst071').AsString then
      cb_dst071.ItemIndex := i;
  end;

  cb_sign.ItemIndex := cb_sign.Items.IndexOf(trim(doc_setup.FieldByName('dst09').AsString));

  ed_dst08.Enabled := False;
  edt_dst03.ReadOnly := True;
  cb_dst01.Enabled := False;
  cb_dst02.Enabled := False;
  cb_sign.Enabled := True;
end;

procedure TfmDoc_setup.Button3Click(Sender: TObject);
var
  trows:integer;
begin
  if status = '' then begin
    showmessage('請按下新增或修改!');
    Exit;
  end;

  if trim(cb_dst01.Text) = '' then begin
    showmessage('請輸入代碼大類!');
    Exit;
  end;


  if trim(cb_dst02.Text) = '' then begin
    showmessage('請輸入代碼小類!');
    Exit;
  end;

  if trim(edt_dst03.Text) = '' then begin
    showmessage('請輸入報表序號!');
    Exit;
  end;

  if trim(edt_dst04.Text) = '' then begin
    showmessage('請輸入報表名稱!');
    Exit;
  end;

  if trim(cb_dst05.Text) = '' then begin
    showmessage('請輸入報表檔案格式');
    Exit;
  end;

  if trim(cb_dst06.Text) = '' then begin
    showmessage('請輸入報表檔案對應電子報系統');
    Exit;
  end;

  if trim(cb_dst07.Text) = '' then begin
    showmessage('請輸入承辦人');
    Exit;
  end;

  if trim(cb_dst071.Text) = '' then begin
    showmessage('請輸入代理人');
    Exit;
  end;

  if status ='INS' then begin
    tmpque1.Close;
    tmpque1.SQL.Clear;
    tmpque1.sql.Add('select * FROM doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac='''+docdm.myfac+'''');
    tmpque1.Parameters.ParamByName('dst01').Value := copy(cb_dst01.Text,1,pos(' ',cb_dst01.Text)-1);
    tmpque1.Parameters.ParamByName('dst02').Value := copy(cb_dst02.Text,1,pos(' ',cb_dst02.Text)-1);
    tmpque1.Parameters.ParamByName('dst03').Value := trim(edt_dst03.Text);
    tmpque1.Open;
    tmpque1.First;
    if tmpque1.RecordCount > 0 then begin
      showmessage('該大類與小類,報表序號碼已存在,無法儲存!');
      status := '';
      panel1.Enabled := FAlse;
      Exit;
    end;
  end;

  

  if MessageDlg('是否儲存?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    try
      if not dm.glm.InTransaction then
        dm.glm.BeginTrans;

      tmpque1.Close;
      tmpque1.SQL.Clear;
      if status ='INS' then begin
        tmpque1.SQL.Add('insert into doc_setup (dst01,dst02,dst03,dst04,dst05,dst06,dst07,dst071,dst08,dst_acti,dst_user,dst_date,dst_fac,dst09) ');
        tmpque1.sql.Add(' values (:dst01,:dst02,:dst03,:dst04,:dst05,:dst06,:dst07,:dst071,:dst08,''Y'',:dst_user,sysdate,:dst_fac,:dst09)');
      end
      else if status ='EDT' then begin
        tmpque1.SQL.Add('update doc_setup set dst04 =:dst04,dst05 =:dst05,dst06 =:dst06,dst07 =:dst07,dst071 =:dst071,dst08 =:dst08,dst_user =:dst_user,dst_date=sysdate,dst09 =:dst09 ');
        tmpque1.sql.Add(' where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac =:dst_fac');
      end;
      tmpque1.Parameters.ParamByName('dst01').Value := copy(cb_dst01.Text,1,pos(' ',cb_dst01.Text)-1);
      tmpque1.Parameters.ParamByName('dst02').Value := copy(cb_dst02.Text,1,pos(' ',cb_dst02.Text)-1);
      tmpque1.Parameters.ParamByName('dst03').Value := edt_dst03.Text;
      tmpque1.Parameters.ParamByName('dst04').Value := edt_dst04.Text;
      tmpque1.Parameters.ParamByName('dst05').Value := copy(cb_dst05.Text,1,pos(' ',cb_dst05.Text)-1);
      tmpque1.Parameters.ParamByName('dst06').Value := copy(cb_dst06.Text,1,pos(' ',cb_dst06.Text)-1);
      tmpque1.Parameters.ParamByName('dst07').Value := copy(cb_dst07.Text,1,pos(' ',cb_dst07.Text)-1);
      tmpque1.Parameters.ParamByName('dst071').Value := copy(cb_dst071.Text,1,pos(' ',cb_dst071.Text)-1);
      tmpque1.Parameters.ParamByName('dst08').Value := ed_dst08.Text;
      tmpque1.Parameters.ParamByName('dst09').Value := trim(cb_sign.Text);
      tmpque1.Parameters.ParamByName('dst_user').Value := dm.getlogin;
      tmpque1.Parameters.ParamByName('dst_fac').Value := docdm.myfac;;
      trows := tmpque1.ExecSQL;

      if trows =1 then begin
        if dm.glm.InTransaction then
          dm.glm.CommitTrans;

        showmessage('儲存成功!');
      end
      else begin
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('儲存失敗!');
      end;
      status := '';
      panel1.Enabled := False;
    except
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('儲存失敗!');
    end;
    btn_que.Click;
  end;

end;

procedure TfmDoc_setup.Button5Click(Sender: TObject);
begin
  if (not doc_setup.Active) or (doc_setup.RecordCount = 0) then begin
    showmessage('請先查詢要設定權限的報表!');
    Exit;
  end;

  fmDocr_file := TfmDocr_file.create(Self);
  fmdocr_file.lab_dst01.Caption := doc_setup.fieldbyname('dst01').AsString;
  fmdocr_file.lab_dst02.Caption := doc_setup.fieldbyname('dst02').AsString;
  fmdocr_file.lab_dst03.Caption := doc_setup.fieldbyname('dst03').AsString;
  fmdocr_file.lab_repdesc.Caption := doc_setup.fieldbyname('dst04').AsString;
  fmDocr_file.ShowModal;
  fmDocr_file.free;
end;

procedure TfmDoc_setup.cb_dst07Exit(Sender: TObject);
begin
  cb_dst071.Clear;
  tmpque1.close;
  tmpque1.sql.Clear;
  if strall = '' then
    tmpque1.sql.Add('select gen02,gen01 from gen_file'+docdm.dblinks+' where genacti =''Y'' and gen03 like'''+docdm.mygem+'%'+''' order by gen05')
  else
    tmpque1.sql.Add('select gen02,gen01 from gen_file'+docdm.dblinks+' where genacti =''Y'' and gen03 like'''+copy(docdm.mygem,1,2)+'%'+''' order by gen05');

  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dst071.Items.Append(tmpque1.fieldbyname('gen01').AsString+' '+tmpque1.fieldbyname('gen02').AsString);
    tmpque1.Next;
  end;
end;

procedure TfmDoc_setup.doc_setupAfterScroll(DataSet: TDataSet);
var
  i : integer;
begin
  edt_dst03.Text := doc_setup.fieldbyname('dst03').AsString;
  edt_dst04.Text := doc_setup.fieldbyname('dst04').AsString;
  ed_dst08.Text := doc_setup.fieldbyname('dst08').AsString;

  for i := 0 to cb_dst01.Items.Count -1 do begin
    if copy(cb_dst01.Items[i],1,pos(' ',cb_dst01.Items[i])-1) = doc_setup.FieldByName('dst01').AsString then
      cb_dst01.ItemIndex := i;
  end;
  cb_dst01.OnExit(self);

  for i := 0 to cb_dst02.Items.Count -1 do begin
    if copy(cb_dst02.Items[i],1,pos(' ',cb_dst02.Items[i])-1) = doc_setup.FieldByName('dst02').AsString then
      cb_dst02.ItemIndex := i;
  end;

  for i := 0 to cb_dst05.Items.Count -1 do begin
    if copy(cb_dst05.Items[i],1,pos(' ',cb_dst05.Items[i])-1) = doc_setup.FieldByName('dst05').AsString then
      cb_dst05.ItemIndex := i;
  end;

  for i := 0 to cb_dst06.Items.Count -1 do begin
    if copy(cb_dst06.Items[i],1,pos(' ',cb_dst06.Items[i])-1) = doc_setup.FieldByName('dst06').AsString then
      cb_dst06.ItemIndex := i;
  end;

  for i := 0 to cb_dst07.Items.Count -1 do begin
    if copy(cb_dst07.Items[i],1,pos(' ',cb_dst07.Items[i])-1) = doc_setup.FieldByName('dst07').AsString then
      cb_dst07.ItemIndex := i;
  end;
  cb_dst07.OnExit(self);

  for i := 0 to cb_dst071.Items.Count -1 do begin
    if copy(cb_dst071.Items[i],1,pos(' ',cb_dst071.Items[i])-1) = doc_setup.FieldByName('dst071').AsString then
      cb_dst071.ItemIndex := i;
  end;

  cb_sign.ItemIndex := cb_sign.Items.IndexOf(doc_setup.FieldByName('dst09').AsString);
  edt_dst03.ReadOnly := True;
  cb_dst01.Enabled := False;
  cb_dst02.Enabled := False;
  status := '';
  panel1.Enabled := False;
end;

procedure TfmDoc_setup.doc_setupCalcFields(DataSet: TDataSet);
begin
  tmpque1.Close;
  tmpque1.sql.Clear;
  tmpque1.sql.Add('select dtype011,dtype021 from DOC_TYPE where dtype01 =:dtype01 and dtype02 =:dtype02 and dtype_fac =:dtype_fac');
  tmpque1.Parameters.ParamByName('dtype01').Value := dataset.fieldbyname('dst01').AsString;
  tmpque1.Parameters.ParamByName('dtype02').Value := dataset.fieldbyname('dst02').AsString;
  tmpque1.Parameters.ParamByName('dtype_fac').Value := dataset.fieldbyname('dst_fac').AsString;
  tmpque1.Open;
  if tmpque1.RecordCount > 0 then begin
    dataset.FieldByName('dst01desc').AsString := tmpque1.fieldbyname('dtype011').AsString;
    dataset.FieldByName('dst02desc').AsString := tmpque1.fieldbyname('dtype021').AsString;
  end
  else begin
    dataset.FieldByName('dst01desc').AsString := '';
    dataset.FieldByName('dst02desc').AsString := '';
  end;

  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select subject,remark From sysparameter Where isActive = ''Y'' And sysname =''DOC'' And item =''DST05'' and subject ='''+dataset.fieldbyname('dst05').asstring+'''');
  tmpque1.Open;
  tmpque1.First;
  if tmpque1.RecordCount > 0 then
    dataset.FieldByName('dst05desc').AsString := tmpque1.fieldbyname('remark').AsString
  else
    dataset.FieldByName('dst05desc').AsString := '';

  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select subject,remark From sysparameter Where isActive = ''Y'' And sysname =''DOC'' And item =''DST06'' and subject ='''+dataset.fieldbyname('dst06').asstring+'''');
  tmpque1.Open;
  tmpque1.First;
  if tmpque1.RecordCount > 0 then
    dataset.FieldByName('dst06desc').AsString := tmpque1.fieldbyname('remark').AsString
  else
    dataset.FieldByName('dst06desc').AsString := '';

end;

procedure TfmDoc_setup.btn_cpClick(Sender: TObject);
begin
  fmcp := Tfmcp.create(self);
  fmcp.showmodal;
  fmcp.free;
end;

procedure TfmDoc_setup.Button6Click(Sender: TObject);
begin
  if MessageDlg('若設定此權限,所有人均可查看報此報表,是否設定?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    tmpque1.Close;
    tmpque1.SQL.Clear;
    tmpque1.sql.Add('select * from docr_file where docr04 =''all'' and docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03');
    tmpque1.Parameters.ParamByName('docr01').Value := doc_setup.fieldbyname('dst01').AsString;
    tmpque1.Parameters.ParamByName('docr02').Value := doc_setup.fieldbyname('dst02').AsString;
    tmpque1.Parameters.ParamByName('docr03').Value := doc_setup.fieldbyname('dst03').AsString;
    tmpque1.Open;
    if tmpque1.RecordCount > 0 then begin
      showmessage('已設定全部人員均可查看!');
      Exit;
    end
    else begin
      tmpque1.Close;
      tmpque1.SQL.Clear;
      tmpque1.SQL.Add('insert into docr_file(docr01,docr02,docr03,docr04,docr_user,docr_date,docr_acti,docr_fac) values (:docr01,:docr02,:docr03,''all'','''+dm.getlogin+''',sysdate,''Y'','''+docdm.myfac+''')');
      tmpque1.Parameters.ParamByName('docr01').Value := doc_setup.fieldbyname('dst01').AsString;
      tmpque1.Parameters.ParamByName('docr02').Value := doc_setup.fieldbyname('dst02').AsString;
      tmpque1.Parameters.ParamByName('docr03').Value := doc_setup.fieldbyname('dst03').AsString;
      tmpque1.ExecSQL;
      showmessage('設定完成!');
    end;
  end;

end;

end.
