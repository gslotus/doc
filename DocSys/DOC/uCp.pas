unit uCp;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, ADODB;

type
  Tfmcp = class(TForm)
    cb_source: TComboBox;
    Label1: TLabel;
    cb_dest: TComboBox;
    btn_ok: TButton;
    ADOQuery1: TADOQuery;
    procedure FormShow(Sender: TObject);
    procedure btn_okClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmcp: Tfmcp;
  tmpque : Tadoquery;
implementation
uses docdm;
{$R *.dfm}

procedure Tfmcp.FormShow(Sender: TObject);
var
  strall : string;
begin
  tmpque := TADoquery.Create(self);
  tmpque.Connection := dm.glm;

  adoquery1.Close;
  adoquery1.sql.Clear;
  adoquery1.SQL.Add('select * from appdeta'+docdm.dblinks+' where appname =''DOC05'' and userid ='''+dm.getlogin+'''');
  adoquery1.Open;
  if adoquery1.RecordCount > 0 then
    strall := 'ALL'
  else
    strall := '';

  cb_dest.Items.Clear;
  cb_source.Items.Clear;
  adoquery1.Close;
  adoquery1.sql.Clear;
  if  strall = 'ALL' then
    adoquery1.sql.Add('select dst01,dst02,dst03,dst04,dst05,dst06  from doc_setup,doc_type where dst01 = dtype01 and dst02 = dtype02 and dtype_fac=dst_fac and dtype03 like '''+copy(docdm.mygem,1,2)+'%'+''' and dst_fac ='''+docdm.myfac+''' and dst_acti=''Y'' order by dst01,dst02,dst03')
  else
    adoquery1.sql.Add('select dst01,dst02,dst03,dst04,dst05,dst06  from doc_setup where (dst07 ='''+dm.getlogin+''' or dst01 ='''+dm.getlogin+''') and dst_fac ='''+docdm.myfac+''' and dst_acti=''Y'' order by dst01,dst02,dst03');
  adoquery1.Open;
  adoquery1.First;
  while not adoquery1.Eof do begin
    cb_source.Items.Append(adoquery1.fieldbyname('dst04').asstring+'>'+adoquery1.fieldbyname('dst01').asstring+'_'+adoquery1.fieldbyname('dst02').asstring+'_'+adoquery1.fieldbyname('dst03').asstring);
    cb_dest.Items.Append(adoquery1.fieldbyname('dst04').asstring+'>'+adoquery1.fieldbyname('dst01').asstring+'_'+adoquery1.fieldbyname('dst02').asstring+'_'+adoquery1.fieldbyname('dst03').asstring);
    adoquery1.Next;
  end;
end;

procedure Tfmcp.btn_okClick(Sender: TObject);
var
  dst01,dst02,dst03,dst011,dst021,dst031 : string;
  myres,mydst:string;
begin
  if cb_source.Text = cb_dest.Text then begin
    showmessage('複製的來源與目地文件不可為同一個!');
    Exit;
  end;

  if MessageDlg('請注意會將目地文件的權限先刪除,再進行複製功能,是否複製?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    myres := copy(cb_source.Text,pos('>',cb_source.Text)+1,length(cb_source.Text)-(pos(' ',cb_source.Text)));
    dst01 := copy(myres,1,2);
    dst02 := copy(myres,4,2);
    dst03 := copy(myres,7,2);

    mydst := copy(cb_dest.text,pos('>',cb_dest.Text)+1,length(cb_dest.Text)-(pos(' ',cb_dest.Text)+1));
    dst011 := copy(mydst,1,2);
    dst021 := copy(mydst,4,2);
    dst031 := copy(mydst,7,2);

    try
      if not dm.glm.InTransaction then
        dm.glm.BeginTrans;

      adoquery1.Close;
      adoquery1.sql.Clear;
      adoquery1.sql.Add('delete from docr_file where docr01 ='''+dst011+''' and docr02 ='''+dst021+''' and docr03 ='''+dst031+''' and docr_fac ='''+docdm.myfac+'''');
      adoquery1.ExecSQL;

      adoquery1.close;
      adoquery1.sql.Clear;
      adoquery1.sql.Add('select * from docr_file where docr01 ='''+dst01+''' and docr02 ='''+dst02+''' and docr03 ='''+dst03+''' and docr_fac ='''+docdm.myfac+'''');
      adoquery1.Open;
      adoquery1.First;
      while not adoquery1.Eof do begin
        tmpque.Close;
        tmpque.sql.Clear;
        tmpque.sql.Add('insert into docr_file (docr01,docr02,docr03,docr04,docr05,docr_user,docr_date,docr_acti,docr_fac) values ');
        tmpque.sql.Add(' (:docr01,:docr02,:docr03,:docr04,:docr05,'''+dm.getlogin+''',sysdate,''Y'','''+docdm.myfac+''') ');
        tmpque.Parameters.ParamByName('docr01').Value := dst011;
        tmpque.Parameters.ParamByName('docr02').Value := dst021;
        tmpque.Parameters.ParamByName('docr03').Value := dst031;
        tmpque.Parameters.ParamByName('docr04').Value := adoquery1.fieldbyname('docr04').AsString;
        tmpque.Parameters.ParamByName('docr05').Value := adoquery1.fieldbyname('docr05').AsString;
        tmpque.ExecSQL;
        adoquery1.Next;
      end;

      if dm.glm.InTransaction then
        dm.glm.CommitTrans;

      showmessage('複製完成!');
    except
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('複製失敗!');
    end;


  end;
end;

end.
