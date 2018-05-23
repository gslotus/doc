unit uDoc_Type;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB, Menus;

type
  Tfmdoc_type = class(TForm)
    Panel2: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    DBGrid1: TDBGrid;
    Label7: TLabel;
    cb_dep1: TComboBox;
    btn_que: TButton;
    doc_type: TADOQuery;
    doc_typeDTYPE01: TWideStringField;
    doc_typeDTYPE02: TWideStringField;
    doc_typeDTYPE_ACTI: TWideStringField;
    doc_typeDTYPE011: TWideStringField;
    doc_typeDTYPE021: TWideStringField;
    doc_typeDTYPE03: TWideStringField;
    doc_typeDTYPE_USER: TWideStringField;
    doc_typeDTYPE_DATE: TDateTimeField;
    DataSource1: TDataSource;
    doc_typeDTYPE_FAC: TWideStringField;
    PopupMenu1: TPopupMenu;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    edt_type01: TEdit;
    cb_dep: TComboBox;
    edt_type02: TEdit;
    edt_type01desc: TEdit;
    edt_type02desc: TEdit;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    procedure FormShow(Sender: TObject);
    procedure btn_queClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure cb_depExit(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure edt_type01Exit(Sender: TObject);
    procedure doc_typeAfterScroll(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmdoc_type: Tfmdoc_type;
  tmpque1 : Tadoquery;
  status : string;
implementation
uses docdm;
{$R *.dfm}

procedure Tfmdoc_type.FormShow(Sender: TObject);
begin
  cb_dep1.Clear;
  cb_dep.Clear;
  tmpque1 := TAdoquery.Create(self);
  tmpque1.Connection := dm.glm;
  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('SELECT SUBJECT,GEM02 FROM SYSPARAMETER'+docdm.dblinks+' ,GEM_FILE'+docdm.dblinks+' WHERE SYSNAME =''DOCSYS'' AND SUBJECT = GEM01 ');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dep1.Items.Append(tmpque1.fieldbyname('subject').AsString+' '+tmpque1.fieldbyname('gem02').AsString);
    tmpque1.Next;
  end;

  tmpque1.Close;
  tmpque1.SQL.Clear;
  tmpque1.SQL.Add('Select gem01,gem02 From gem_file'+docdm.dblinks+',sysparameter'+docdm.dblinks+' where length(gem01) = 3 and SYSNAME =''DOCSYS'' AND SUBSTR(SUBJECT,1,2) = SUBSTR(GEM01,1,2) and gem07 is not null and gemacti =''Y'' order by gem01');
  tmpque1.Open;
  tmpque1.First;
  while not tmpque1.eof do begin
    cb_dep.Items.Append(tmpque1.fieldbyname('gem01').AsString+' '+tmpque1.fieldbyname('gem02').AsString);
    tmpque1.Next;
  end;

end;

procedure Tfmdoc_type.btn_queClick(Sender: TObject);
begin
  status := '';
  doc_type.Close;
  doc_type.sql.Clear;
  doc_type.SQL.Add('select * from doc_type where dtype03 LIKE '''+'%'+copy(cb_dep1.Text,2,1)+'%'+''' and dtype_fac ='''+docdm.myfac+''' order by dtype01,dtype02');
  doc_type.Open;
  doc_type.First;
  edt_type01.ReadOnly := True;
  edt_type02.ReadOnly := True;
  edt_type01desc.ReadOnly := True;
  edt_type02desc.ReadOnly := True;
  cb_dep.Enabled := FAlse;
end;

procedure Tfmdoc_type.Button1Click(Sender: TObject);
begin
  status := 'INS';
  cb_dep.Enabled := True;
  cb_dep.ItemIndex := -1;
  edt_type01.Clear;
  edt_type01desc.Clear;
  edt_type02.Clear;
  edt_type02desc.clear;
  edt_type01.SetFocus;

  edt_type01.ReadOnly := false;
  edt_type02.ReadOnly := false;
  edt_type01desc.ReadOnly := false;
  edt_type02desc.ReadOnly := false;
    cb_dep.Enabled := True;

end;

procedure Tfmdoc_type.Button2Click(Sender: TObject);
var
  i : integer;
begin
  status := 'EDT';
  edt_type01.Text := doc_type.fieldbyname('dtype01').AsString;
  edt_type01desc.Text := doc_type.fieldbyname('dtype011').AsString;
  edt_type02.Text := doc_type.fieldbyname('dtype02').AsString;
  edt_type02desc.Text := doc_type.fieldbyname('dtype021').AsString;
  for i := 0 to cb_dep.Items.Count -1 do begin
    if copy(cb_dep.Items[i],1,pos(' ',cb_dep.Items[i])-1) = doc_type.FieldByName('dtype03').AsString then
      cb_dep.ItemIndex := i;
  end;

 { for i := 0 to cb_emp.Items.Count -1 do begin
    if copy(cb_emp.Items[i],1,pos(' ',cb_emp.Items[i])-1) = doc_type.FieldByName('dtype_user').AsString then
      cb_emp.ItemIndex := i;
  end;    }

  cb_dep.Enabled := false;
  edt_type01.ReadOnly := True;
  edt_type02.ReadOnly := True;
  edt_type01desc.ReadOnly := False;
  edt_type02desc.ReadOnly := False;
    cb_dep.Enabled := True;

end;

procedure Tfmdoc_type.cb_depExit(Sender: TObject);
var
  s1 : string;
begin
{  if status ='INS' then begin
    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('select max(dtype01) from doc_type where dtype01 like '''+copy(cb_dep.Text,2,1)+'%'+'''');
    tmpque1.Open;
    s1 := copy(tmpque1.Fields.Fields[0].asstring,2,1);
    if s1 = '' then
      edt_type01.Text := copy(cb_dep.Text,2,1)+'A'
    else
      edt_type01.Text := copy(cb_dep.Text,2,1)+chr(ord(s1[1])+1);

    tmpque1.Close;
    tmpque1.sql.Clear;
    tmpque1.sql.Add('select count(*) from doc_type where dtype01 = '''+edt_type01.Text+'''');
    tmpque1.Open;
    edt_type02.Text := formatfloat('00',tmpque1.Fields.Fields[0].AsInteger+1);
  end;     }
end;

procedure Tfmdoc_type.Button4Click(Sender: TObject);
var
  acts : string;
begin
  if (not doc_type.Active) or (doc_type.RecordCount = 0) then begin
    showmessage('請查收要作廢的資料!');
    EXit;
  end;

   status := '';
  if doc_type.FieldByName('dtype_acti').AsString ='Y' then
    acts := ''
  else if doc_type.FieldByName('dtype_acti').AsString ='N' then
    acts := '取消';

  if MessageDlg('是否'+acts+'作廢?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    if trim(acts) = '' then begin
      tmpque1.Close;
      tmpque1.SQL.Clear;
      tmpque1.sql.Add('update doc_type set dtype_acti =''N'' where dtype01 =:dtype01 and dtype02=:dtype02 and dtype03=:dtype03 and dtype_fac =:dtype_fac');
    end
    else begin
      tmpque1.Close;
      tmpque1.SQL.Clear;
      tmpque1.sql.Add('update doc_type set dtype_acti =''Y'' where dtype01 =:dtype01 and dtype02=:dtype02 and dtype03=:dtype03 and dtype_fac =:dtype_fac');
    end;
    tmpque1.Parameters.ParamByName('dtype01').Value := doc_type.fieldbyname('dtype01').AsString;
    tmpque1.Parameters.ParamByName('dtype02').Value := doc_type.fieldbyname('dtype02').AsString;
    tmpque1.Parameters.ParamByName('dtype03').Value := doc_type.fieldbyname('dtype03').AsString;
    tmpque1.Parameters.ParamByName('dtype_fac').Value := doc_type.fieldbyname('dtype_fac').AsString;
    tmpque1.ExecSQL;
    showmessage('異動完成!');
  end;

  btn_que.Click;

end;

procedure Tfmdoc_type.Button3Click(Sender: TObject);
var
  trows : integer;
begin
  if status = '' then begin
    showmessage('請按下新增或修改!');
    Exit;
  end;

  if trim(cb_dep.Text) = '' then begin
    showmessage('請輸入承辦部門!');
    Exit;
  end;


  if trim(edt_type01.Text) = '' then begin
    showmessage('請輸入大類代碼!');
    Exit;
  end;

  if trim(edt_type02.Text) = '' then begin
    showmessage('請輸入小類代碼!');
    Exit;
  end;

  if trim(edt_type01desc.Text) = '' then begin
    showmessage('請輸入大類代碼說明!');
    Exit;
  end;

  if trim(edt_type02desc.Text) = '' then begin
    showmessage('請輸入小類代碼說明!');
    Exit;
  end;

  if status ='INS' then begin
    tmpque1.Close;
    tmpque1.SQL.Clear;
    tmpque1.sql.Add('select * FROM doc_type where dtype01 ='''+edt_type01.Text+''' and dtype02 ='''+edt_type02.Text+''' and dtype_fac='''+docdm.myfac+'''');
    tmpque1.Open;
    tmpque1.First;
    if tmpque1.RecordCount > 0 then begin
      showmessage('該大類與小類代碼已存在,無法儲存!');
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
        tmpque1.SQL.Add('insert into doc_type (dtype01,dtype02,dtype011,dtype021,dtype_acti,dtype03,dtype_user,dtype_date,dtype_fac) ');
        tmpque1.sql.Add(' values (:dtype01,:dtype02,:dtype011,:dtype021,''Y'',:dtype03,:dtype_user,sysdate,:dtypefac)');
      end
      else if status ='EDT' then begin
        tmpque1.SQL.Add('update doc_type set dtype011 =:dtype011,dtype021 =:dtype021,dtype_user =:dtype_user,dtype_date =sysdate  ');
        tmpque1.sql.Add(' where dtype01 =:dtype01 and dtype02 =:dtype02 and dtype03 =:dtype03 and dtype_fac =:dtypefac');
      end;
      tmpque1.Parameters.ParamByName('dtype01').Value := edt_type01.Text;
      tmpque1.Parameters.ParamByName('dtype02').Value := edt_type02.Text;
      tmpque1.Parameters.ParamByName('dtype011').Value := edt_type01desc.Text;
      tmpque1.Parameters.ParamByName('dtype021').Value := edt_type02desc.Text;
      tmpque1.Parameters.ParamByName('dtype03').Value := copy(cb_dep.Text,1,pos(' ',cb_dep.Text)-1);
      tmpque1.Parameters.ParamByName('dtype_user').Value := dm.getlogin;
      tmpque1.Parameters.ParamByName('dtypefac').Value := docdm.myfac;
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
    except
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('儲存失敗!');
    end;
    btn_que.Click;
  end;
end;

procedure Tfmdoc_type.edt_type01Exit(Sender: TObject);
begin
  if status = 'INS' then begin
    tmpque1.close;
    tmpque1.SQL.Clear;
    tmpque1.sql.Add('select count(*) from doc_type where dtype01 ='''+trim(edt_type01.text)+''' and dtype_fac='''+docdm.myfac+'''');
    tmpque1.Open;
    tmpque1.First;
    if tmpque1.RecordCount > 0 then begin
      edt_type02.Text := formatfloat('00',tmpque1.Fields.Fields[0].AsInteger+1);
    end;

    tmpque1.close;
    tmpque1.SQL.Clear;
    tmpque1.sql.Add('select dtype011 from doc_type where dtype01 ='''+trim(edt_type01.text)+''' and dtype_fac='''+docdm.myfac+'''');
    tmpque1.Open;
    tmpque1.First;
    if tmpque1.RecordCount > 0 then begin
      edt_type01desc.Text := tmpque1.Fields.Fields[0].AsString;
    end;

  end;
end;

procedure Tfmdoc_type.doc_typeAfterScroll(DataSet: TDataSet);
var
  i : integer;
begin
    edt_type01.Text := doc_type.fieldbyname('dtype01').AsString;
    edt_type01desc.Text := doc_type.fieldbyname('dtype011').AsString;
    edt_type02.Text := doc_type.fieldbyname('dtype02').AsString;
    edt_type02desc.Text := doc_type.fieldbyname('dtype021').AsString;
    for i := 0 to cb_dep.Items.Count -1 do begin
      if copy(cb_dep.Items[i],1,pos(' ',cb_dep.Items[i])-1) = doc_type.FieldByName('dtype03').AsString then
        cb_dep.ItemIndex := i;
    end;
    status := '';

end;

end.
