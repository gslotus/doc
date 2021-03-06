unit docque;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Buttons, Grids, DBGrids, Menus, ExtCtrls, Shellapi,
  OleCtrls, SHDocVw, OleCtnrs, AxCtrls, VCF1;

type
  TfmQuery = class(TForm)
    queMain: TADOQuery;
    queMainDOC01: TWideStringField;
    queMainDOC02: TWideStringField;
    queMainDOC03: TWideStringField;
    queMainDOC04: TWideStringField;
    queMainDOC05: TWideStringField;
    queMainDOC06: TWideStringField;
    queMainDOC07: TWideStringField;
    queMainDOC08: TWideStringField;
    queMainCREATEDATE: TDateTimeField;
    queMainCREATEUSER: TWideStringField;
    queMainUPDATEDATE: TDateTimeField;
    queMainUPDATEUSER: TWideStringField;
    queMainDOC09: TWideStringField;
    DataSource1: TDataSource;
    ADOQuery1: TADOQuery;
    queMaintypedesc: TStringField;
    queMainchilddesc: TStringField;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    queMainsigndesc: TStringField;
    queMainDOCMEMO: TWideStringField;
    IFRS1: TMenuItem;
    queMainDOC10: TWideStringField;
    queMainDOC_ACTI: TWideStringField;
    queMainDOC_FAC: TWideStringField;
    queMainDOCSIGN_NP: TWideStringField;
    queMainWEBSIGN_NP: TWideStringField;
    queMaindetaildesc: TStringField;
    queMainDOCSIGN: TWideStringField;
    queMainWEBSIGN: TWideStringField;
    queMainDOC11: TWideStringField;
    queMainDOC12: TWideStringField;
    Panel2: TPanel;
    Label4: TLabel;
    btn_sign: TBitBtn;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Panel1: TPanel;
    gbyear: TGroupBox;
    edyear: TEdit;
    gbmon: TGroupBox;
    edMonth: TEdit;
    gbsea: TGroupBox;
    cbSeason: TComboBox;
    Panel3: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    ls_type01: TListBox;
    ls_type02: TListBox;
    ls_type03: TListBox;
    Panel4: TPanel;
    Panel5: TPanel;
    DBGrid1: TDBGrid;
    queMainDOC081: TWideStringField;
    wb1: TWebBrowser;
    BitBtn3: TBitBtn;
    Label5: TLabel;
    Button1: TButton;
    procedure BitBtn2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure edMonthExit(Sender: TObject);
    procedure cbSeasonExit(Sender: TObject);
    procedure queMainCalcFields(DataSet: TDataSet);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure IFRS1Click(Sender: TObject);
    procedure btn_signClick(Sender: TObject);
    procedure btn_sign1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure ls_type01Click(Sender: TObject);
    procedure ls_type02Click(Sender: TObject);
    procedure queMainAfterScroll(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmQuery: TfmQuery;
  status,stypes : string;
  tmpque : TAdoquery;
  deps:string;
implementation
uses docdm,docbrowse, docins, docmaintain, docmantain1, docsign,
  mail;
{$R *.dfm}

procedure TfmQuery.BitBtn2Click(Sender: TObject);
begin
  close;
end;

procedure TfmQuery.FormShow(Sender: TObject);
var
  s1:string;
begin

  tmpque := TAdoquery.Create(self);
  tmpque.Connection := dm.glm;
  if status ='SIGN' then begin
    n1.Visible := FAlse;
    n2.Visible := FAlse;
    btn_sign.Visible := True;
    panel4.Visible := False;
  end
  else begin
    n1.Visible := True;
    n2.Visible := True;
    btn_sign.Visible := False;
    panel4.Visible := False;
    panel5.Align := alClient;
  end;

  adoquery1.Close;
  adoquery1.SQL.Clear;
  adoquery1.SQL.Add('select * from azc_file'+docdm.dblinks+' where azc03 ='''+dm.getlogin+''' order by azc01 asc');
  adoquery1.Open;
  s1 := '0';
  if adoquery1.RecordCount > 0 then begin
    if adoquery1.RecordCount > 0 then
      s1 := '1';

    deps := copy(adoquery1.fieldbyname('azc01').AsString,1,3);
  end
  else
    deps := '';

  ls_type01.Items.Clear;
  adoquery1.Close;
  adoquery1.SQL.Clear;
  if status ='SIGN' then
    adoquery1.SQL.Add('select unique dtype01,dtype011 from doc_type where dtype_acti=''Y'' and dtype03 like '''+copy(docdm.mygem,1,2)+'%'+''' and dtype_fac='''+docdm.myfac+''' order by dtype01')
  else if (copy(deps,3,1)='0') or (s1 ='1') then
    adoquery1.SQL.Add('select unique dtype01,dtype011 from doc_type where dtype_acti=''Y'' and dtype03 like '''+copy(deps,1,2)+'%'+''' and dtype_fac='''+docdm.myfac+''' order by dtype01')
  else if deps <> '' then
    adoquery1.SQL.Add('select unique dtype01,dtype011 from doc_type where dtype_acti=''Y'' and dtype03 like '''+deps+'%'''' and dtype_fac='''+docdm.myfac+''' order by dtype01')
  else
    adoquery1.SQL.Add('select unique dtype01,dtype011 from doc_type,doc_setup where dst01 = dtype01 and dst02 = dtype02 and ( dst07 ='''+dm.getlogin+''' OR dst071 = '''+dm.getlogin+''' ) and dtype_acti=''Y'' and dtype03 like '''+copy(docdm.mygem,1,2)+'%'+''' and dtype_fac='''+docdm.myfac+''' order by dtype01');

  adoquery1.Open;
  adoquery1.First;
  //showmessage(adoquery1.SQL.Text);
  while not adoquery1.Eof do begin
    ls_type01.Items.Append(adoquery1.Fields[0].asstring+' '+adoquery1.Fields[1].asstring);
    adoquery1.Next;
  end;
  {if adoquery1.FieldByName('dtype01').AsString='VA' then begin
    button1.Visible:=true;
  end
  else begin
    button1.Visible:=false;
  end;           }

  ls_type01.ItemIndex := 0;
  ls_type01.OnClick(self);
  //ls_tyoe02.ItemIndex := -1;
  edyear.Clear;
  edmonth.Clear;
  cbseason.ItemIndex := -1;

  if status ='SIGN' then
    bitbtn1.Click;

end;

procedure TfmQuery.BitBtn1Click(Sender: TObject);
var
  i : integer;
  type01,type02,type03:integer;
begin
  if status <> 'SIGN' then begin
    type01 := -1;
    type02 := -1;
    type03 := -1;
    for i := 0 to ls_type01.Items.Count -1 do begin
      if ls_type01.Selected[i] = True then
        type01 := i
    end;

    for i := 0 to ls_type02.Items.Count -1 do begin
      if ls_type02.Selected[i] = True then
        type02 := i
    end;

    for i := 0 to ls_type03.Items.Count -1 do begin
      if ls_type03.Selected[i] = True then
        type03 := i
    end;

    if type01 = -1 then begin
      showmessage('請務必選擇資料大類!');
      exit;
    end;

    if type02 = -1 then begin
      showmessage('請務必選擇資料小類!');
      exit;
    end;

    //if trim(cb_detail.text) = '' then begin
    //  showmessage('請務必選擇報表序號!');
    //  exit;
   // end;
  end;

  if trim(edmonth.Text) <> '' then
    edmonth.Text :=formatfloat('00',strtoint(edmonth.Text));
  queMain.Close;
  queMain.SQL.Clear;
  if status ='SIGN' then begin
    if stypeS = 'DOC' then
      queMain.SQL.Add('select * from doc_file where doc05 like :doc05 and doc_acti=''Y'' and docsign_np ='''+dm.getlogin+'''')
    else
      queMain.SQL.Add('select * from doc_file where doc05 like :doc05 and doc_acti=''Y'' and websign_np ='''+dm.getlogin+'''');
  end
  else
    queMain.SQL.Add('select * from doc_file where doc05 =:doc05 and doc_acti =''Y''  ');

  if status ='SIGN' then
    quemain.Parameters.ParamByName('doc05').Value := '%'
  else
    quemain.Parameters.ParamByName('doc05').Value := copy(ls_type01.Items[type01],1,2);

  if status <> 'SIGN' then begin
    if type02 <> -1 then begin
      quemain.SQL.Add(' and doc06 =:doc06');
      quemain.Parameters.ParamByName('doc06').Value := copy(ls_type02.Items[type02],1,2);
    end;
  end;

  if status <> 'SIGN' then begin
    if type03 <> -1 then begin
      quemain.SQL.Add(' and doc11 =:doc11');
      quemain.Parameters.ParamByName('doc11').Value := copy(ls_type03.Items[type03],1,2);
    end;
  end;

  if trim(edyear.Text) <> '' then begin
    quemain.SQL.Add(' and doc02 =:doc02');
    quemain.Parameters.ParamByName('doc02').Value := trim(edyear.Text);
  end;
  if trim(edmonth.Text) <> '' then begin
    quemain.SQL.Add(' and doc03 =:doc03');
    quemain.Parameters.ParamByName('doc03').Value := trim(edmonth.Text);
  end;
  if trim(cbseason.Text) <> '' then begin
    quemain.SQL.Add(' and doc04 =:doc04');
    quemain.Parameters.ParamByName('doc04').Value := formatfloat('00',cbseason.ItemIndex+1);
  end;
  quemain.SQL.Add(' order by doc05,doc06,doc02,doc03,doc04');
  quemain.Open;
  quemain.First;
end;

procedure TfmQuery.edMonthExit(Sender: TObject);
begin
  if trim(edmonth.Text) <> '' then begin
    try
      strtoint(edmonth.Text);
    except
      edmonth.Clear;
      showmessage('請輸入正確的月份 1-12');
     Exit;
    end ;

    if (strtoint(edmonth.Text) > 12)  or (strtoint(edmonth.Text) < 1) then begin
      showmessage('請輸入正確的月份 1-12');
      edmonth.Clear;
    end;

    if trim(edmonth.Text) <> '' then
      cbSeason.ItemIndex := -1;
  end;
end;

procedure TfmQuery.cbSeasonExit(Sender: TObject);
begin
  if trim(cbSeason.Text) <> '' then
    edmonth.Clear;
end;

procedure TfmQuery.queMainCalcFields(DataSet: TDataSet);
begin
  adoquery1.Close;
  adoquery1.SQL.Clear;
  adoquery1.sql.Add('select dtype011 from doc_type where dtype01 =:dtype01 and dtype_fac='''+docdm.myfac+'''');
  adoquery1.Parameters.ParamByName('dtype01').Value := dataset.fieldbyname('doc05').AsString;

  adoquery1.Open;
  adoquery1.First;
  if not adoquery1.Eof then
    DataSet.FieldByName('typedesc').AsString := adoquery1.Fields[0].AsString;

  adoquery1.Close;
  adoquery1.SQL.Clear;
  adoquery1.sql.Add('select dtype021 from doc_type where dtype01 =:dtype01 and dtype02 =:dtype02');
  adoquery1.Parameters.ParamByName('dtype01').Value := dataset.fieldbyname('doc05').AsString;
  adoquery1.Parameters.ParamByName('dtype02').Value := dataset.fieldbyname('doc06').AsString;
  adoquery1.Open;
  adoquery1.First;
  if not adoquery1.Eof then
    DataSet.FieldByName('childdesc').AsString := adoquery1.Fields[0].AsString;


  adoquery1.Close;
  adoquery1.SQL.Clear;
  adoquery1.sql.Add('select DST04,dst06 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac=:dst_fac');
  adoquery1.Parameters.ParamByName('dst01').Value := dataset.fieldbyname('doc05').AsString;
  adoquery1.Parameters.ParamByName('dst02').Value := dataset.fieldbyname('doc06').AsString;
  adoquery1.Parameters.ParamByName('dst03').Value := dataset.fieldbyname('doc11').AsString;
  adoquery1.Parameters.ParamByName('dst_fac').Value := dataset.fieldbyname('doc_fac').AsString;
  adoquery1.Open;
  adoquery1.First;
  if not adoquery1.Eof then
    DataSet.FieldByName('detaildesc').AsString := adoquery1.Fields[0].AsString;


  if (adoquery1.fieldbyname('dst06').asstring = 'D') or ((stypes ='DOC') and (adoquery1.fieldbyname('dst06').asstring = 'A')) then begin
    if dataset.FieldByName('docsign').AsString = '0' then
      dataset.FieldByName('signdesc').AsString := '待課級簽核'
    else if dataset.FieldByName('docsign').AsString = '1' then
      dataset.FieldByName('signdesc').AsString := '待理級簽核'
    else if dataset.FieldByName('docsign').AsString = '2' then
      dataset.FieldByName('signdesc').AsString := '簽核完成'
    else
      dataset.FieldByName('signdesc').AsString := '不需簽核';
  end
  else begin
    if dataset.FieldByName('websign').AsString  = '0' then
      dataset.FieldByName('signdesc').AsString := '待課級簽核'
    else if dataset.FieldByName('websign').AsString = '1' then
      dataset.FieldByName('signdesc').AsString := '待理級簽核'
    else if dataset.FieldByName('websign').AsString = '2' then
      dataset.FieldByName('signdesc').AsString := '簽核完成'
    else
      dataset.FieldByName('signdesc').AsString := '不需簽核';
  end;

end;

procedure TfmQuery.DBGrid1DblClick(Sender: TObject);
var
  myext : string;
begin
  if (quemain.Active = False) or (quemain.RecordCount = 0) then begin
    showmessage('請查出要修改的資料!');
    Exit;
  end;

  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.sql.Add('select docr04 from docr_file where docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_acti=''Y'' and docr04 ='''+dm.getlogin+''' and docr_fac='''+docdm.myfac+'''');
  tmpque.Parameters.ParamByName('docr01').Value := quemain.fieldbyname('doc05').AsString;
  tmpque.Parameters.ParamByName('docr02').Value := quemain.fieldbyname('doc06').AsString;
  tmpque.Parameters.ParamByName('docr03').Value := quemain.fieldbyname('doc11').AsString;
  tmpque.Open;
  if tmpque.RecordCount  = 0 then begin
    if (dm.getlogin <> quemain.FieldByName('createuser').AsString) and (dm.getlogin <> quemain.FieldByName('docsign_np').AsString) and (dm.getlogin <> quemain.FieldByName('websign_np').AsString) then begin
      showmessage('您無權限可查看明細資料!');
      Exit;
    end
  end;

  if status ='SIGN' then begin
    tmpque.Close;
    tmpque.sql.Clear;
    if stypes ='DOC' then
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCFTP'' and isactive =''Y''')
    else
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCWEBFTP'' and isactive =''Y''');
    tmpque.Open;
    if tmpque.RecordCount > 0 then begin
      panel4.Visible := True;
      myext := uppercase(copy(ExtractFileExt(trim(quemain.FieldByName('doc07').AsString)),2,3));
      if myext ='XLS' then
        wb1.Navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+copy(quemain.FieldByName('doc07').AsString,1,pos('.',quemain.FieldByName('doc07').AsString)-1)+'.mht')
      else
      //ShellExecute(0, PChar('open'), PChar('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString), nil , nil, SW_SHOWNA);
        wb1.Navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString);
    end;
  end
  else begin
    fmIns := TfmIns.create(Self);
    docIns.status :='BROWSE';
    fmIns.ShowModal;
    fmIns.free;
  end;
end;

procedure TfmQuery.N1Click(Sender: TObject);
begin
  fmIns := TfmIns.create(Self);
  docIns.status :='INS';
  fmIns.ShowModal;
  fmIns.free;
end;

procedure TfmQuery.N2Click(Sender: TObject);
begin
  if (quemain.Active = False) or (quemain.RecordCount = 0) then begin
    showmessage('請查出要修改的資料!');
    Exit;
  end;

 { if (quemain.FieldByName('docsign').AsString > '0') and (quemain.FieldByName('docsign').AsString <> 'X') then begin
    showmessage('該文件已簽核無法修改!');
    Exit;
  end;

  if (quemain.FieldByName('websign').AsString > '0') and (quemain.FieldByName('websign').AsString <> 'X') then begin
    showmessage('該文件(WEB)已簽核無法修改!');
    Exit;
  end;

  if dm.getlogin <> quemain.FieldByName('createuser').AsString then begin
    showmessage('您非建立人員不可修改該筆資料!');
    Exit;
  end;      }

  fmIns := TfmIns.create(Self);
  docIns.status :='UPD';
  fmIns.ShowModal;
  fmIns.free;

end;

procedure TfmQuery.N3Click(Sender: TObject);
begin
  fmMaintain := TfmMaintain.create(self);
  fmMaintain.adoquery1.close;
  fmMaintain.adoquery1.Open;
  fmMaintain.showmodal;
  fmMaintain.free;
end;

procedure TfmQuery.N4Click(Sender: TObject);
begin
  fmbigtype := tfmbigtype.create(self);
  fmbigtype.adoquery1.close;
  fmbigtype.adoquery1.Open;
  fmbigtype.showmodal;
  fmbigtype.free;

end;

procedure TfmQuery.N6Click(Sender: TObject);
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOCBOSS'' and userid =:id');
  dm.dbque.Parameters.ParamByName('id').Value := dm.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if not dm.dbque.Eof then begin
    fmsign := Tfmsign.create(self);
    docsign.status := 'BOSS';
    fmsign.showmodal;
    fmsign.free;
  end
  else
    showmessage('您無文件放行權限(DOCBOSS)');
end;

procedure TfmQuery.N5Click(Sender: TObject);
begin
  dm.dbque.Close;
  dm.dbque.SQL.Clear;
  dm.dbque.SQL.Add('select * from appdeta where appname =''DOCMAN'' and userid =:id');
  dm.dbque.Parameters.ParamByName('id').Value := dm.getlogin;
  dm.dbque.Open;
  dm.dbque.First;
  if not dm.dbque.Eof then begin
    fmsign := Tfmsign.create(self);
    docsign.status := 'MANAGER';
    fmsign.showmodal;
    fmsign.free;
  END
  else
    showmessage('您無文件放行權限(DOCMAN)');
end;

procedure TfmQuery.IFRS1Click(Sender: TObject);
begin
  {fmifrs := Tfmifrs.create(Self);
  fmifrs.ShowModal;
  fmifrs.free;}
end;

procedure TfmQuery.btn_signClick(Sender: TObject);
var
  signdep,np,signno, mailstr : string;
  docsign : integer;
begin
  if (quemain.Active = False) or (quemain.RecordCount = 0) then begin
    showmessage('請查出要簽核的資料!');
    Exit;
  end;

  if (quemain.FieldByName('docsign_np').AsString <> dm.getlogin) and (quemain.FieldByName('websign_np').AsString <> dm.getlogin) then begin
    showmessage('您非待簽人員,無法簽核!');
    Exit;
  end;

  if MessageDlg('簽核OK?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    if stypes ='DOC' then begin
      tmpque.Close;
      tmpque.sql.Clear;
      tmpque.SQL.Add('select doc12,docsign from doc_file where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''');
      tmpque.Open;
      tmpque.First;
      signdep := tmpque.fieldbyname('doc12').AsString;
      docsign := tmpque.fieldbyname('docsign').AsInteger+2;
    end
    else begin
      tmpque.Close;
      tmpque.sql.Clear;
      tmpque.SQL.Add('select doc12,websign from doc_file where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''');
      tmpque.Open;
      tmpque.First;
      signdep := 'WEB';
      docsign := tmpque.fieldbyname('websign').AsInteger+2;
    end;

    tmpque.Close;
    tmpque.sql.Clear;
    tmpque.sql.Add('select * from azc_file'+docdm.dblinks+' where azc01 ='''+signdep+''' and azc02 = '+inttostr(docsign));
    tmpque.Open;
    tmpque.First;
    if tmpque.RecordCount > 0 then begin
      signno := inttostr(docsign -1);
      np := tmpque.fieldbyname('azc03').AsString;
    end
    else begin
      if docsign >= 2 then
        signno := '2'
      else
        signno := '';
      np := '';
    end;

    try
      if not dm.glm.InTransaction then
        dm.glm.BeginTrans;

      tmpque.Close;
      tmpque.sql.Clear;
      if stypes ='DOC' then begin
        if np <> '' then
          tmpque.SQL.Add('update doc_file set docsign_np ='''+np+''',docsign ='+signno+' where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''')
        else
         tmpque.SQL.Add('update doc_file set docsign_np ='''',docsign =2 where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''');
      end
      else begin
        if np <> '' then
          tmpque.SQL.Add('update doc_file set websign_np ='''+np+''',websign ='+signno+' where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''')
        else
          tmpque.SQL.Add('update doc_file set websign_np ='''',websign =2 where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''');
      end;
      tmpque.ExecSQL;

      tmpque.Close;
      tmpque.sql.Clear;
      tmpque.sql.Add('insert into doc_sign_log (dslog01,dslog02,dslog03,dslog04,dslog05,dslog06,dslog07,dslog_fac,dslog011) values ');
      tmpque.sql.Add(' (:dslog01,:dslog02,:dslog03,sysdate,''OK'','''',:dslog07,'''+docdm.myfac+''',''DOCSIGN'') ');
      tmpque.Parameters.ParamByName('dslog01').Value := quemain.fieldbyname('doc01').AsString;
      tmpque.Parameters.ParamByName('dslog02').Value := signno;
      tmpque.Parameters.ParamByName('dslog03').Value := dm.getlogin;
      tmpque.Parameters.ParamByName('dslog07').Value := signdep;
      tmpque.ExecSQL;

      if dm.glm.InTransaction then
        dm.glm.CommitTrans;

      if np <> '' then begin
        tmpque.Close;
        tmpque.sql.Clear;
        tmpque.SQL.Add('select type_detail from sys_codedef'+docdm.dblinks+' where cd_type=''09''');
        tmpque.Open;
        if tmpque.RecordCount > 0 then
          mailstr := tmpque.fieldbyname('type_detail').AsString
        else
          mailstr := '';

        fmMail := TfmMail.Create(self);
        if np <> '' then
          fmMail.Smtp.PostMessage.ToAddress.Append(np+mailstr);
        fmMail.smtp.postmessage.toaddress.append(dm.getlogin+mailstr);
        fmMail.Edit4.Text := '文件電子報系統 : '+quemain.fieldbyname('doc09').AsString+' '+quemain.fieldbyname('doc081').AsString+' 等待簽核';
        fmMail.RichEdit2.Lines.Append(np+' 您好 : ');
        fmMail.RichEdit2.Lines.Append(quemain.fieldbyname('doc09').AsString+' '+quemain.fieldbyname('doc081').AsString+' 等待您的簽核放行');
        fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心');
        fmMail.ShowModal;
        fmMail.Free;
      end
      else if np = '' then begin
        tmpque.Close;
        tmpque.sql.Clear;
        tmpque.SQL.Add('select type_detail from sys_codedef'+docdm.dblinks+' where cd_type=''09''');
        tmpque.Open;
        if tmpque.RecordCount > 0 then
          mailstr := tmpque.fieldbyname('type_detail').AsString
        else
          mailstr := '';

        tmpque.Close;
        tmpque.sql.Clear;
        tmpque.SQL.Add('select docr05 from docr_file where docr01 =:docr01 and docr02 =:docr02 and docr03 =:docr03 and docr_fac ='''+docdm.myfac+''' and docr_acti =''Y''');
        tmpque.Parameters.ParamByName('docr01').Value := queMain.fieldbyname('doc05').AsString;
        tmpque.Parameters.ParamByName('docr02').Value := queMain.fieldbyname('doc06').AsString;
        tmpque.Parameters.ParamByName('docr03').Value := queMain.fieldbyname('doc11').AsString;
        tmpque.Open;

        fmMail := TfmMail.Create(self);
        tmpque.First;
        while not tmpque.Eof do begin
          fmMail.Smtp.PostMessage.ToAddress.Append(tmpque.fieldbyname('docr05').AsString);
          tmpque.Next;
        end;
        fmMail.Smtp.PostMessage.ToAddress.Append(quemain.fieldbyname('createuser').AsString+mailstr);
        fmMail.Edit4.Text := '文件電子報系統 : '+quemain.fieldbyname('doc09').AsString+' '+quemain.fieldbyname('doc081').AsString+' 發佈完成';
          
        fmMail.RichEdit2.Lines.Append('各位同仁您好 : ');
        fmMail.RichEdit2.Lines.Append('  ');
        fmMail.RichEdit2.Lines.Append(quemain.fieldbyname('doc09').AsString+' '+quemain.fieldbyname('doc081').AsString+' 已發佈完成');
        if signdep ='WEB' then
          fmMail.RichEdit2.Lines.Append('請點選 http://www.gs.com.tw/chinese/CompanyPreview6.html  查看文件 ')
        else
          fmMail.RichEdit2.Lines.Append('請點選 http://i.gshank.com  查看您的文件 ');
        fmMail.RichEdit2.Lines.Append('若有任何使用上的問題請與資訊室連絡 分機:126');
        fmMail.RichEdit2.lines.Append('   ');
        fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心 b2b@mail.gs.com.tw');
        fmMail.ShowModal;
        fmMail.Free;
      end;               
    except
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('簽核失敗!')
    end;
    bitbtn1.Click;
  end;
end;

procedure TfmQuery.btn_sign1Click(Sender: TObject);
var
  mailstr: string;
var
  signdep,np,signno : string;
  websign : integer;
begin
  if (quemain.Active = False) or (quemain.RecordCount = 0) then begin
    showmessage('請查出要簽核的資料!');
    Exit;
  end;

  if MessageDlg('網頁文件簽核OK?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
    tmpque.Close;
    tmpque.sql.Clear;
    tmpque.SQL.Add('select websign from doc_file where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''');
    tmpque.Open;
    tmpque.First;
    signdep := 'WEB';
    websign := tmpque.fieldbyname('websign').AsInteger+2;

    tmpque.Close;
    tmpque.sql.Clear;
    tmpque.sql.Add('select * from azc_file'+docdm.dblinks+' where azc01 ='''+signdep+''' and azc02 = '+inttostr(websign));
    tmpque.Open;
    tmpque.First;
    if tmpque.RecordCount > 0 then begin
      signno := inttostr(websign-1);
      np := tmpque.fieldbyname('azc03').AsString;
    end
    else begin
      if websign >= 2 then
        signno := '2'
      else
        signno := '';
      np := '';
    end;

    try
      if not dm.glm.InTransaction then
        dm.glm.BeginTrans;

      tmpque.Close;
      tmpque.sql.Clear;
      if np <> '' then
        tmpque.SQL.Add('update doc_file set websign_np ='''+np+''',websign ='+signno+' where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''')
      else
        tmpque.SQL.Add('update doc_file set websign_np ='''',websign =2 where doc01 ='''+quemain.fieldbyname('doc01').AsString+'''');
      tmpque.ExecSQL;

      tmpque.Close;
      tmpque.sql.Clear;
      tmpque.sql.Add('insert into doc_sign_log (dslog01,dslog02,dslog03,dslog04,dslog05,dslog06,dslog07,dslog_fac,dslog011) values ');
      tmpque.sql.Add(' (:dslog01,:dslog02,:dslog03,sysdate,''OK'','''',:dslog07,'''+docdm.myfac+''',''WEBSIGN'') ');
      tmpque.Parameters.ParamByName('dslog01').Value := quemain.fieldbyname('doc01').AsString;
      tmpque.Parameters.ParamByName('dslog02').Value := signno;
      tmpque.Parameters.ParamByName('dslog03').Value := dm.getlogin;
      tmpque.Parameters.ParamByName('dslog07').Value := signdep;
      tmpque.ExecSQL;

      if dm.glm.InTransaction then
        dm.glm.CommitTrans;

      showmessage('簽核完成!');

      if np <> '' then begin
        tmpque.Close;
        tmpque.sql.Clear;
        tmpque.SQL.Add('select type_detail from sys_codedef'+docdm.dblinks+' where cd_type=''09''');
        tmpque.Open;
        if tmpque.RecordCount > 0 then
          mailstr := tmpque.fieldbyname('type_detail').AsString
        else
          mailstr := '';

        fmMail := TfmMail.Create(self);
        //if docsignnp <> '' then
        //  fmMail.Smtp.PostMessage.ToAddress.Append(np+mailstr);
        fmMail.smtp.postmessage.toaddress.append('lotus@mail.gs.com.tw');
        fmMail.Edit4.Text := '文件電子報系統 : '+quemain.fieldbyname('doc02').AsString+'年'+quemain.fieldbyname('doc03').AsString+'月'+quemain.fieldbyname('doc08').AsString+' 等待簽核';
        fmMail.RichEdit2.Lines.Append('鉅祥同仁您好 : ');
        fmMail.RichEdit2.Lines.Append(quemain.fieldbyname('doc02').AsString+'年'+quemain.fieldbyname('doc03').AsString+'月'+quemain.fieldbyname('doc08').AsString+' 等待您的簽核放行');
        fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心');
        fmMail.ShowModal;
        fmMail.Free;
      end
      else if np = '' then begin
        tmpque.Close;
        tmpque.sql.Clear;
        tmpque.SQL.Add('select type_detail from sys_codedef'+docdm.dblinks+' where cd_type=''09''');
        tmpque.Open;
        if tmpque.RecordCount > 0 then
          mailstr := tmpque.fieldbyname('type_detail').AsString
        else
          mailstr := '';

        fmMail := TfmMail.Create(self);
        //  fmMail.Smtp.PostMessage.ToAddress.Append(quemain.fieldbyname('createuser').AsString+mailstr);
        fmMail.smtp.postmessage.toaddress.append('lotus@mail.gs.com.tw');
        fmMail.Edit4.Text := '文件電子報系統 : '+quemain.fieldbyname('doc02').AsString+'年'+quemain.fieldbyname('doc03').AsString+'月'+quemain.fieldbyname('doc08').AsString+' 簽核完成';
        fmMail.RichEdit2.Lines.Append('鉅祥同仁您好 : ');
        fmMail.RichEdit2.Lines.Append(quemain.fieldbyname('doc02').AsString+'年'+quemain.fieldbyname('doc03').AsString+'月'+quemain.fieldbyname('doc08').AsString+' 已簽核完成');
        fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心');
        fmMail.ShowModal;
        fmMail.Free;
      end;
      
    except
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('簽核失敗!')
    end;
    bitbtn1.Click;
  end;

end;

procedure TfmQuery.Button2Click(Sender: TObject);
begin
 if quemain.FieldByName('doc07').AsString <> '' then begin
    fmBrowse := TfmBrowse.create(Self);
    fmbrowse.WebBrowser1.Navigate('http://1.1.166.6/doc/'+docdm.myfac+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString);
    fmBrowse.ShowModal;
    fmBrowse.free;
  end
  else
    showmessage('尚未上載檔案');
end;

procedure TfmQuery.Button4Click(Sender: TObject);
begin
 if quemain.FieldByName('doc07').AsString <> '' then begin
    tmpque.Close;
    tmpque.sql.Clear;
    tmpque.SQL.Add('select dst06 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03');
    tmpque.Parameters.ParamByName('dst01').Value := quemain.FieldByName('doc05').AsString;
    tmpque.Parameters.ParamByName('dst02').Value := quemain.FieldByName('doc06').AsString;
    tmpque.Parameters.ParamByName('dst03').Value := quemain.FieldByName('doc11').AsString;
    tmpque.Open;
    if tmpque.FieldByName('dst06').AsString = 'W' then begin
      fmBrowse := TfmBrowse.create(Self);
      fmbrowse.WebBrowser1.Navigate('http://1.1.130.9/doc/'+docdm.myfac+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString);
      fmBrowse.ShowModal;
      fmBrowse.free;
    end
    else
      showmessage('未有網頁文件可瀏覽!');
  end
  else
    showmessage('尚未上載檔案');

end;

procedure TfmQuery.ls_type01Click(Sender: TObject);
var
  i : integer;
  type01 : integer;
begin
  ls_type02.Clear;
  ls_type03.Clear;
  type01 := -1;
  for i := 0 to ls_type01.Count -1 do begin
    if ls_type01.Selected[i] = True then
      type01 := i;
  end;
  if type01 <> -1 then begin
    ls_type02.Clear;
    adoquery1.Close;
    adoquery1.SQL.Clear;
    if copy(deps,3,1) ='0' then
      adoquery1.SQL.Add('select unique dtype02,dtype021 from doc_type,doc_setup where dst01 = dtype01 and dst02 = dtype02 and dtype03 like '''+copy(deps,1,2)+'%'+''' and dtype01 =:dtype01 and dtype_acti =''Y'' and dtype_fac='''+docdm.myfac+''' order by dtype02')
    else if deps <> '' then
      adoquery1.SQL.Add('select unique dtype02,dtype021 from doc_type,doc_setup where dst01 = dtype01 and dst02 = dtype02 and dtype03 like '''+copy(deps,1,3)+'%'+''' and dtype01 =:dtype01 and dtype_acti =''Y'' and dtype_fac='''+docdm.myfac+''' order by dtype02')
    else
      adoquery1.SQL.Add('select unique dtype02,dtype021 from doc_type,doc_setup where dst01 = dtype01 and dst02 = dtype02 and ( dst07 ='''+dm.getlogin+''' OR dst071 = '''+dm.getlogin+''' ) and dtype01 =:dtype01 and dtype_acti =''Y'' and dtype_fac='''+docdm.myfac+''' order by dtype02');
    adoquery1.Parameters.ParamByName('dtype01').Value := copy(ls_type01.Items[type01],1,2);
    adoquery1.Open;
    adoquery1.First;
    while not adoquery1.Eof do begin
      ls_type02.Items.Append(adoquery1.Fields[0].asstring+' '+adoquery1.Fields[1].asstring);
      adoquery1.Next;
    end;
  end;
end;

procedure TfmQuery.ls_type02Click(Sender: TObject);
var
  i : integer;
  type01,type02 : integer;
begin
  ls_type03.Clear;
  type01 := -1;
  type02 := -1;
  for i := 0 to ls_type01.Count -1 do begin
    if ls_type01.Selected[i] = True then
      type01 := i;
  end;

  for i := 0 to ls_type02.Count -1 do begin
    if ls_type02.Selected[i] = True then
      type02 := i;
  end;

  ls_type03.Clear;
  adoquery1.Close;
  adoquery1.SQL.Clear;
  if copy(deps,3,1) = '0' then
    adoquery1.SQL.Add('select unique dst03,dst04 from doc_setup,doc_type where dst01 = dtype01 and dst02 = dtype02 and dst01 =:dst01 and dtype03 like '''+copy(deps,1,2)+'%'+'''  and dst02 =:dst02 and dst_acti =''Y'' and dst_fac='''+docdm.myfac+''' order by dst03')
  else if deps <> '' then
    adoquery1.SQL.Add('select unique dst03,dst04 from doc_setup,doc_type where dst01 = dtype01 and dst02 = dtype02 and dst01 =:dst01 and dtype03 like '''+deps+'%'+''' and  dst02 =:dst02 and dst_acti =''Y'' and dst_fac='''+docdm.myfac+''' order by dst03')
  else
    adoquery1.SQL.Add('select unique dst03,dst04 from doc_setup where dst01 =:dst01 and ( dst07 ='''+dm.getlogin+''' OR dst071 = '''+dm.getlogin+''' ) and dst02 =:dst02 and dst_acti =''Y'' and dst_fac='''+docdm.myfac+''' order by dst03');
  adoquery1.Parameters.ParamByName('dst01').Value := copy(ls_type01.Items[type01],1,2);
  adoquery1.Parameters.ParamByName('dst02').Value := copy(ls_type02.Items[type02],1,2);
  adoquery1.Open;
  adoquery1.First;
  while not adoquery1.Eof do begin
    ls_type03.Items.Append(adoquery1.Fields[0].asstring+' '+adoquery1.Fields[1].asstring);
    adoquery1.Next;
  end;
  //li_type03.ItemIndex := -1;
end;

procedure TfmQuery.queMainAfterScroll(DataSet: TDataSet);
var
  fe,f,myext:string;
begin
   if status ='SIGN' then begin
    tmpque.Close;
    tmpque.sql.Clear;
    if stypes ='DOC' then
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCFTP'' and isactive =''Y''')
    else
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCWEBFTP'' and isactive =''Y''');
    tmpque.Open;
    if tmpque.RecordCount > 0 then begin
      if quemain.RecordCount > 0 then begin
        panel4.Visible := True;
        if stypes ='DOC' then begin
          myext := uppercase(copy(ExtractFileExt(trim(quemain.FieldByName('doc07').AsString)),2,3));
          if myext ='XLS' then
            wb1.Navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString)
          else
          //ShellExecute(0, PChar('open'), PChar('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString), nil , nil, SW_SHOWNA);
            wb1.Navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString);
        end
        else begin
            wb1.navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc07').AsString);
        end;
      end;
    end;

  end
  else
    panel4.Visible := False;

end;

procedure TfmQuery.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  wb1.Navigate('#');
end;

procedure TfmQuery.BitBtn3Click(Sender: TObject);
var
  fe,f,myext:string;
begin
   if status ='SIGN' then begin
    tmpque.Close;
    tmpque.sql.Clear;
    if stypes ='DOC' then
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCFTP'' and isactive =''Y''')
    else
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCWEBFTP'' and isactive =''Y''');
    tmpque.Open;
    if tmpque.RecordCount > 0 then begin
      if quemain.RecordCount > 0 then begin
        panel4.Visible := True;
        if stypes ='DOC' then begin
          myext := uppercase(copy(ExtractFileExt(trim(quemain.FieldByName('doc07').AsString)),2,3));
          if myext ='XLS' then
            //wb1.Navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString)
            ShellExecute(0, PChar('open'), PChar('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString), nil , nil, SW_SHOWNA)
          else
            ShellExecute(0, PChar('open'), PChar('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString), nil , nil, SW_SHOWNA);
            //wb1.Navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc_fac').AsString+'/'+quemain.FieldByName('doc02').AsString+'/'+copy(quemain.FieldByName('doc05').AsString,1,1)+'/'+quemain.FieldByName('doc07').AsString);
        end
        else begin
            ShellExecute(0, PChar('open'), PChar('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc07').AsString), nil , nil, SW_SHOWNA);
            //wb1.navigate('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+quemain.FieldByName('doc07').AsString);
        end;
      end;
    end;

  end
  else
    panel4.Visible := False;

end;

procedure TfmQuery.Button1Click(Sender: TObject);
var query1,query2,query3 : TAdoquery;
begin


  if quemain.Active and (quemain.RecordCount>0) then begin
    if MessageDlg('請問是否確定執行作廢?',  mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      Exit;
    try
      if not dm.glm.InTransaction then
        dm.glm.BeginTrans;
      query1 := TAdoquery.Create(self);
      query1.Connection := dm.glm;
      query2 := TAdoquery.Create(self);
      query2.Connection := dm.glm;
      query3 := TAdoquery.Create(self);
      query3.Connection := dm.glm;
      query1.Close;
      query1.SQL.Clear;
      query1.SQL.Add('update doc_file set doc_acti=''N'' where doc01=:doc01');
      query1.Parameters.ParamByName('doc01').Value:=quemain.fieldbyname('doc01').asstring;
      query1.ExecSQL;
      if (ls_type01.Count>0) and (copy(ls_type01.Items[0],1,2)='VA') then begin  //訓練中心
        query2.Close;
        query2.SQL.Clear;
        query2.SQL.Add('update trll_doc set tdoc03=''N'' where tdoc07=:tdoc07');
        query2.Parameters.ParamByName('tdoc07').Value:=quemain.fieldbyname('doc01').asstring;
        query2.ExecSQL;
        query3.Close;
        query3.SQL.Clear;
        query3.SQL.Add('update trll_file set trllacti=''N'' where trll03=:trll03');
        query3.Parameters.ParamByName('trll03').Value:= quemain.fieldbyname('doc01').asstring;
        query3.ExecSQL;
      end;
      //button1.Visible:=false;

      showmessage('作廢成功!');
    except
      if dm.glm.InTransaction then
        dm.glm.RollbackTrans;
      showmessage('作廢失敗!');
    end;
    if dm.glm.InTransaction then
      dm.glm.CommitTrans;
  end
  else begin
    showmessage('請先查詢出需作廢資料');
    exit;
  end;
end;

end.
