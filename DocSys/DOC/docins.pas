unit docins;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, DB, ADODB, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdFTP, Shellapi, Grids, DBGrids, Comobj;

type
  ech = exception;
  TfmIns = class(TForm)
    gbmon: TGroupBox;
    edMonth: TEdit;
    gbyear: TGroupBox;
    edyear: TEdit;
    Label4: TLabel;
    lbno: TLabel;
    BitBtn1: TBitBtn;
    Label5: TLabel;
    edfile: TEdit;
    Button1: TButton;
    btn_docview1: TButton;
    ADOQuery1: TADOQuery;
    OpenDialog1: TOpenDialog;
    lbfiles: TLabel;
    adoexec: TADOQuery;
    Button3: TButton;
    Label6: TLabel;
    Memo2: TMemo;
    gbseason: TGroupBox;
    cbseason: TComboBox;
    Label8: TLabel;
    lab_filetype: TLabel;
    IdFTPDOC: TIdFTP;
    IdFTPWEB: TIdFTP;
    btn_webview1: TButton;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    mydoc: TADOQuery;
    mydocDST01: TWideStringField;
    mydocDST02: TWideStringField;
    mydocDST03: TWideStringField;
    mydocDST04: TWideStringField;
    mydocDST05: TWideStringField;
    mydoclastdate: TStringField;
    mydocdst05desc: TStringField;
    mydocdst06desc: TStringField;
    Label1: TLabel;
    edt_doc081: TEdit;
    mydocDST06: TStringField;
    mydocDST08: TBCDField;
    btn_ins: TBitBtn;
    btn_view1: TButton;
    lbfiles1: TLabel;
    ADOQuery2: TADOQuery;
    IdFTP1: TIdFTP;
    Label2: TLabel;
    edfile1: TEdit;
    Button2: TButton;
    Button4: TButton;
    lab_filetype1: TLabel;
    lbfiles2: TLabel;
    Memo1: TMemo;
    IdFTP2: TIdFTP;
    Label3: TLabel;
    Label7: TLabel;
    OpenDialog2: TOpenDialog;
    Label9: TLabel;
    Edit1: TEdit;
    procedure FormShow(Sender: TObject);
    procedure cbTypeExit(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btn_docview1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure cb_detailExit(Sender: TObject);
    function ftpdocsys(filenames:string):boolean;
    function ftpclasssys(filenames,filenames1:string):boolean;
    function ftpwebdocsys(filenames:string):boolean;
    procedure cbChildExit(Sender: TObject);
    procedure btn_webview1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure mydocAfterScroll(DataSet: TDataSet);
    procedure mydocCalcFields(DataSet: TDataSet);
    procedure btn_insClick(Sender: TObject);
    procedure btn_view1Click(Sender: TObject);
    procedure btn_viewpdfClick(Sender: TObject);
    procedure btn_etopClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Edit1Exit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmIns: TfmIns;
  status,fileadd,fileno : string;
  year,month,day : word;
  e :ech;
  yy,mm,ss : string;
  list1 : Tstringlist;
  tmpque : Tadoquery;
implementation
uses docdm,docbrowse, docque, mail;
{$R *.dfm}

procedure TfmIns.FormShow(Sender: TObject);
var
  i : integer;
begin
  tmpque := Tadoquery.Create(self);
  tmpque.Connection := dm.glm;
  list1 := TStringlist.Create;
  decodedate(date,year,month,day);
  mydoc.Close;
  mydoc.SQL.Clear;
  if status ='INS' then
    mydoc.sql.Add('select dst01,dst02,dst03,dst04,dst05,dst06,dst08  from doc_setup where (dst07 ='''+dm.getlogin+''' or dst071 ='''+dm.getlogin+''') and dst_fac ='''+docdm.myfac+''' and dst_acti=''Y'' order by dst01,dst02,dst03')
  else begin
    mydoc.sql.Add('select dst01,dst02,dst03,dst04,dst05,dst06,dst08  from doc_setup where (dst07 ='''+dm.getlogin+''' or dst071 ='''+dm.getlogin+''') ');
    mydoc.sql.add(' and dst_fac ='''+docdm.myfac+''' and dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_acti=''Y'' order by dst01,dst02,dst03');
    mydoc.Parameters.ParamByName('dst01').Value := fmquery.quemain.fieldbyname('doc05').asstring;
    mydoc.Parameters.ParamByName('dst02').Value := fmquery.quemain.fieldbyname('doc06').asstring;
    mydoc.Parameters.ParamByName('dst03').Value := fmquery.quemain.fieldbyname('doc11').asstring;
  end;
  mydoc.Open;
  if Status = 'INS' THEN BEGIN
    edyear.Enabled := True;
    gbmon.Enabled := True;
    gbseason.Enabled := True;
    lbno.Caption := '';
    edyear.Text := inttostr(year);
    edmonth.Clear;
    cbseason.ItemIndex := -1;
    edt_doc081.clear;
    edt_doc081.Text := mydoc.fieldbyname('dst04').AsString;
    //ed_repdesc.clear;
    edfile.Clear;
  END
  else begin
    gbyear.Enabled := false;
    gbmon.Enabled := false;
    gbseason.Enabled := False;
    //ed_repdesc.Enabled := false;
    edfile.Enabled := True;

    lbno.Caption := fmquery.quemain.fieldbyname('doc01').AsString;
    lbfiles.Caption := fmquery.quemain.fieldbyname('doc07').AsString;
    if uppercase(copy(fmquery.quemain.fieldbyname('doc07').AsString,pos('.',fmquery.quemain.fieldbyname('doc07').AsString)+1,3)) ='XLS' then begin
      lbfiles1.Caption := copy(fmquery.quemain.fieldbyname('doc07').AsString,1,pos('.',fmquery.quemain.fieldbyname('doc07').AsString)-1)+'.pdf';
      //edt_excelpdf.Enabled := True;
      //btn_etop.Enabled := True;
      //btn_view2.Enabled := True;
      //btn_viewpdf.Enabled := True;
    end
    else begin
      lbfiles1.caption := '';
      //edt_excelpdf.Enabled := False;
      //btn_etop.Enabled := False;
      //btn_view2.Enabled := False;
      //btn_viewpdf.Enabled := False;
    end;
      
    edyear.text :=fmquery.quemain.fieldbyname('doc02').AsString;
    //ed_repdesc.Text :=fmquery.quemain.fieldbyname('doc08').AsString;
    memo2.Text :=fmquery.quemain.fieldbyname('docmemo').AsString;
    if fmquery.quemain.fieldbyname('doc03').AsString = '00' then begin
      gbmon.Visible := False;
      edMonth.Clear;
    end
    else if fmquery.quemain.fieldbyname('doc03').AsString <> '00' then begin
      gbmon.Visible := True;
      edmonth.Text := fmquery.quemain.fieldbyname('doc03').AsString;
    end;

    if fmquery.quemain.fieldbyname('doc04').AsString = '00' then begin
      gbseason.Visible := False;
      cbseason.ItemIndex := -1;
    end
    else begin
      gbseason.Visible := True;
      cbSeason.ItemIndex := fmquery.quemain.fieldbyname('doc04').AsInteger -1;
    end;
    edt_doc081.Text := fmquery.quemain.fieldbyname('doc081').AsString;
    {for i := 0 to cbtype.Items.Count -1 do begin
      if copy(cbtype.Items.Strings[i],1,2) = fmquery.quemain.fieldbyname('doc05').AsString then
        cbtype.ItemIndex := i;
    end;
    cbtype.OnClick(self);

    for i := 0 to cbchild.Items.Count -1 do begin
      if copy(cbchild.Items.Strings[i],1,2) = fmquery.quemain.fieldbyname('doc06').AsString then
        cbchild.ItemIndex := i;
    end;
    cbChild.OnExit(self);

    for i := 0 to cb_detail.Items.Count -1 do begin
      if copy(cb_detail.Items.Strings[i],1,2) = fmquery.quemain.fieldbyname('doc11').AsString then
        cb_detail.itemIndex := i;
    end;
    cb_detail.OnExit(self); }
  end;

  if status = 'BROWSE' then begin
    bitbtn1.Visible := false;
    btn_ins.Visible := false;
  end
  else begin
    bitbtn1.Visible := true;
    btn_ins.Visible := True;
  end;
  if mydoc.FieldByName('dst05').AsString='V' then begin
    Label3.Visible:=true;
    Label7.Visible:=true;
    btn_docview1.Visible:=false;
    btn_webview1.Visible:=false;
    btn_ins.Visible:=false;
    Label2.Visible:=true;
    edfile1.Visible:=true;
    button2.Visible:=true;
    button4.Visible:=true;
    Label1.Caption:='課程名稱';
    Label9.Visible:=true;
    edit1.Visible:=true;
  end
  else begin
    Label3.Visible:=false;
    Label7.Visible:=false;
    btn_docview1.Visible:=true;
    btn_webview1.Visible:=true;
    Label2.Visible:=false;
    edfile1.Visible:=false;
    button2.Visible:=false;
    button4.Visible:=false;
    Label1.Caption:='報表名稱';
    Label9.Visible:=false;
    edit1.Visible:=false;
  end;
end;

procedure TfmIns.cbTypeExit(Sender: TObject);
begin
  {cbchild.Clear;
  adoquery1.Close;
  adoquery1.SQL.Clear;
  adoquery1.SQL.Add('select dtype02,dtype021 from doc_type where dtype01 =:dtype01 and dtype_acti =''Y'' and dtype_fac='''+docdm.myfac+'''');
  adoquery1.Parameters.ParamByName('dtype01').Value := copy(cbtype.Text,1,2);
  adoquery1.Open;
  adoquery1.First;
  while not adoquery1.Eof do begin
    cbchild.Items.Append(adoquery1.Fields[0].asstring+' '+adoquery1.Fields[1].asstring);
    adoquery1.Next;
  end;
  cbchild.ItemIndex := -1; }
  //gbseason.Visible := FAlse;
  //gbmon.Visible := False;
end;

procedure TfmIns.Button1Click(Sender: TObject);
var
  myext : string;
begin
  if opendialog1.Execute then begin
    edfile.Text := opendialog1.Files.Text;
    myext := ExtractFileExt(edfile.Text);
    if trim(myext) <> '.'+copy(lab_filetype.Caption,pos(' ',lab_filetype.Caption)+1,length(lab_filetype.Caption)-2) then begin
      showmessage('您輸入的檔案類型錯誤,檔案類型必需為:'+copy(lab_filetype.Caption,pos(' ',lab_filetype.Caption)+1,length(lab_filetype.Caption)-2));
      edfile.Clear;
    end;
  end;


end;

procedure TfmIns.btn_docview1Click(Sender: TObject);
begin
  if lbfiles.Caption <> '' then begin
    tmpque.Close;
    tmpque.sql.Clear;
    tmpque.sql.Add('select subject from sysparameter where sysname=''DOCFTP'' and isactive =''Y''');
    tmpque.Open;
    if tmpque.RecordCount > 0 then
      ShellExecute(0, PChar('open'), PChar('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+docdm.myfac+'/'+edyear.Text+'/'+copy(mydoc.fieldbyname('dst01').AsString,1,1)+'/'+lbfiles.Caption), nil , nil, SW_SHOWNA);
  end
  else
    showmessage('尚未上載檔案');
end;

procedure TfmIns.BitBtn1Click(Sender: TObject);
var
  docno : integer;
  no,mon,sea,dst09,no_str,s : string;
  copyed,puted : boolean;
  webs,doctypes,docsignnp,websignnp,mailstr : string;
  query1,query2,query3,query4 : TAdoquery;
  W: WideString;
begin
  BitBtn1.Cursor:= crappstart;
  query1 := TAdoquery.Create(self);
  query1.Connection := dm.glm;
  query2 := TAdoquery.Create(self);
  query2.Connection := dm.glm;
  query3 := TAdoquery.Create(self);
  query3.Connection := dm.glm;
  query4 := TAdoquery.Create(self);
  query4.Connection := dm.glm;
  //edmonth.Text:=formatfloat('00',strtoint(edmonth.Text));
  adoquery2.Close;
  adoquery2.SQL.Clear;
  adoquery2.SQL.Add('select dst01,dst02,dst03,dst04,dst05 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
  adoquery2.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
  adoquery2.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
  adoquery2.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
  adoquery2.Open;
  query4.Close;
  query4.SQL.Clear;
  query4.SQL.Add('select gem03,gen05 from gem_file,gen_file where gem01=gen03 and gen02=:gen02 and gem00=:gem00 and gemacti=''Y''');
  query4.Parameters.ParamByName('gen02').Value:=trim(edit1.Text);
  query4.Parameters.ParamByName('gem00').Value:=docdm.myfac;
  query4.Open;
  if adoquery2.FieldByName('dst05').AsString='V' then begin
    if status = 'INS' then begin
      if edfile.Text='' then begin
        Showmessage('請選擇上傳文件!');
        Exit;
      end;
      if edfile1.Text='' then begin 
        Showmessage('請選擇上傳文件!');
        Exit;
      end;
      adoquery1.Close;
      adoquery1.sql.Clear;
      if  ss = '00' then begin
        adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno06 =:docno06 and docno_fac='''+docdm.myfac+'''');
        adoquery1.Parameters.ParamByName('docno06').Value := edmonth.Text;
      end
      else if ss ='01' then begin
        adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno05 =:docno05 and docno_fac='''+docdm.myfac+'''');
        adoquery1.Parameters.ParamByName('docno05').Value := formatfloat('00',cbseason.itemindex+1);
      end
      else if ss = '02' then begin
        adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno05 =''06'' and docno_fac='''+docdm.myfac+'''');
      end
      else if ss = '03' then begin
        adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno05 =''99'' and docno_fac='''+docdm.myfac+'''');
      end;
      adoquery1.Parameters.ParamByName('docno01').Value := mydoc.fieldbyname('dst01').AsString;
      adoquery1.Parameters.ParamByName('docno02').Value := mydoc.fieldbyname('dst02').AsString;
      adoquery1.Parameters.ParamByName('docno03').Value := mydoc.fieldbyname('dst03').AsString;
      adoquery1.Parameters.ParamByName('docno04').Value := edyear.Text;
      adoquery1.Open;
      adoquery1.First;
      if not adoquery1.Eof then begin
        docno := adoquery1.fieldbyname('docno07').AsInteger+1;
        adoexec.Close;
        adoexec.sql.Clear;
        if ss = '00' then begin
          adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno06 =:docno06 and docno_fac='''+docdm.myfac+'''');
          adoexec.Parameters.ParamByName('docno06').Value := edmonth.Text;
        end
        else if ss = '01' then begin
          adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno05 =:docno05 and docno_fac='''+docdm.myfac+'''');
          adoexec.Parameters.ParamByName('docno05').Value := formatfloat('00',cbseason.itemindex+1);
        end
        else if ss = '02' then begin
          adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno05 =''06'' and docno_fac='''+docdm.myfac+'''');
        end
        else if ss = '03' then begin
          adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno05 =''99'' and docno_fac='''+docdm.myfac+'''');
        end;
        adoexec.Parameters.ParamByName('docno01').Value := mydoc.fieldbyname('dst01').AsString;
        adoexec.Parameters.ParamByName('docno02').Value := mydoc.fieldbyname('dst02').AsString;
        adoexec.Parameters.ParamByName('docno03').Value := mydoc.fieldbyname('dst03').AsString;
        adoexec.Parameters.ParamByName('docno07').Value := formatfloat('0000',docno);
        adoexec.ExecSQL;
      end
      else begin
        docno := 1;
        adoexec.Close;
        adoexec.sql.Clear;
        adoexec.SQL.Add('insert into doc_no (docno01,docno02,docno03,docno04,docno05,docno06,docno07,docno_date,docno_fac) values ');
        adoexec.SQL.Add(' (:docno01,:docno02,:docno03,:docno04,:docno05,:docno06,:docno07,sysdate,'''+docdm.myfac+''') ');
        adoexec.Parameters.ParamByName('docno01').Value := mydoc.fieldbyname('dst01').AsString;
        adoexec.Parameters.ParamByName('docno02').Value := mydoc.fieldbyname('dst02').AsString;
        adoexec.Parameters.ParamByName('docno03').Value := mydoc.fieldbyname('dst03').AsString;
        adoexec.Parameters.ParamByName('docno04').Value := edyear.Text;
        adoexec.Parameters.ParamByName('docno07').Value := formatfloat('0000',docno);
        if ss = '00' then begin
          adoexec.Parameters.ParamByName('docno06').Value := formatfloat('00',strtoint(edmonth.Text));
          adoexec.Parameters.ParamByName('docno05').Value := '00';
        end
        else if ss = '01' then begin
          adoexec.Parameters.ParamByName('docno06').Value := '00';
          adoexec.Parameters.ParamByName('docno05').Value := formatfloat('00',cbseason.ItemIndex+1);;
        end
        else if ss = '02' then begin
          adoexec.Parameters.ParamByName('docno06').Value := '00';
          adoexec.Parameters.ParamByName('docno05').Value := '00';
        end
        else if ss = '03' then begin
          adoexec.Parameters.ParamByName('docno06').Value := '00';
          adoexec.Parameters.ParamByName('docno05').Value := '99';
        end;
        adoexec.ExecSQL;
      end;
      if ss = '00' then
        no := docdm.myfac+trim(edyear.Text)+formatfloat('00',strtoint(edmonth.Text))+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno)
      else if ss = '01' then
        no := docdm.myfac+trim(edyear.Text)+formatfloat('00',cbseason.ItemIndex+1)+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno)
      else if ss = '02' then
        no := docdm.myfac+trim(edyear.Text)+'06'+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno)
      else if ss = '03' then
        no := docdm.myfac+trim(edyear.Text)+'99'+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno);
      try
        if not dm.glm.InTransaction then
          dm.glm.BeginTrans;
      with TMemoryStream.Create do try
        memo1.Lines[71]:='              <h2>' +edt_doc081.Text+ '</h2>';
        memo1.Lines[76]:='              <source src="media/'+edyear.Text+'/'+mydoc.fieldbyname('dst01').AsString+'/'+no+'.mp4" type=''video/mp4; ';
        memo1.Lines[77]:='              codecs="avc1.42E01E, mp4a.40.2"''></source>';
        memo1.Lines[78]:='              <source src="media/'+edyear.Text+'/'+mydoc.fieldbyname('dst01').AsString+'/'+no+'.ogv" type="video/ogg"></source>';
        S := #$FF#$FE;
        Write(S[1], Length(S));
        W := Memo1.Text;
        Write(W[1], Length(W) * SizeOf(WideChar));
        Position := 0;
        SaveToFile('D:\'+no+'.html');
        fileadd:='D:\'+no+'.html';
        fileno:= no;
        finally
          Free;
      end;
      adoexec.Close;
      adoexec.SQL.Clear;
      adoexec.SQL.Add('insert into trll_doc (tdoc01,tdoc02,tdoc03,tdoc04,tdoc05,tdoc06,tdoc07) values(:tdoc01,:tdoc02,:tdoc03,:tdoc04,:tdoc05,:tdoc06,:tdoc07)');
      adoexec.Parameters.ParamByName('tdoc01').Value:= trim(edt_doc081.Text);
      adoexec.Parameters.ParamByName('tdoc02').Value:= 'html5/'+no+'.html';
      adoexec.Parameters.ParamByName('tdoc03').Value:='Y';
      adoexec.Parameters.ParamByName('tdoc04').Value:='E';
      adoexec.Parameters.ParamByName('tdoc05').Value:=adoquery2.fieldbyname('dst01').AsString;
      adoexec.Parameters.ParamByName('tdoc06').Value:=adoquery2.fieldbyname('dst02').AsString;
      adoexec.Parameters.ParamByName('tdoc07').Value:=no;
      adoexec.ExecSQL;
      query2.Close;
      query2.SQL.Add('insert into doc_file (doc01,doc02,doc03,doc11,doc04,doc05,doc06,doc07,doc08,doc081,createdate,createuser,doc09,docsign,websign,docmemo,doc_acti,doc_fac,docsign_np,websign_np,doc12) ');
      query2.SQL.Add(' values (:doc01,:doc02,:doc03,:doc11,:doc04,:doc05,:doc06,:doc07,:doc08,:doc081,sysdate,'''+dm.getlogin+''',:doc09,:docsign,:websign,:docmemo,''Y'','''+docdm.myfac+''',:docsign_np,:websign_np,:doc12)');
      query2.Parameters.ParamByName('doc01').Value:=no;
      query2.Parameters.ParamByName('doc02').Value := trim(edyear.Text);
      if ss = '00' then begin
        query2.Parameters.ParamByName('doc03').Value := edmonth.text;
        query2.Parameters.ParamByName('doc04').Value := '00';
      end
      else if ss = '01' then begin
        query2.Parameters.ParamByName('doc03').Value := '00';
        query2.Parameters.ParamByName('doc04').Value := cbseason.itemindex+1;
      end
      else if ss = '02' then begin
        query2.Parameters.ParamByName('doc03').Value := '00';
        query2.Parameters.ParamByName('doc04').Value := '06';
      end
      else if ss = '03' then begin
        query2.Parameters.ParamByName('doc03').Value := '00';
        query2.Parameters.ParamByName('doc04').Value := '99';
      end;
      query2.Parameters.ParamByName('doc05').Value := adoquery2.fieldbyname('dst01').AsString;
      query2.Parameters.ParamByName('doc06').Value := adoquery2.fieldbyname('dst02').AsString;
      query2.Parameters.ParamByName('doc11').Value := adoquery2.fieldbyname('dst03').AsString;
      query2.Parameters.ParamByName('doc07').Value:=no+'.mp4';
      query2.Parameters.ParamByName('doc08').Value:=adoquery2.fieldbyname('dst04').AsString;
      query2.Parameters.ParamByName('doc081').Value:=trim(edt_doc081.Text);
      if ss = '00' then
        query2.Parameters.ParamByName('doc09').Value := edyear.Text+'年'+edmonth.text+'月'
      else if ss = '01' then
        query2.Parameters.ParamByName('doc09').Value := edyear.Text+'年'+cbseason.Text
      else if ss = '02' then
        query2.Parameters.ParamByName('doc09').Value := edyear.Text+'年半年報'
      else if ss = '03' then
        query2.Parameters.ParamByName('doc09').Value := edyear.Text+'年';
      query2.Parameters.ParamByName('docsign').Value := 'X';
      query2.Parameters.ParamByName('WEBSIGN').Value := 'X';
      query2.Parameters.ParamByName('docmemo').Value:=trim(Memo2.Text);
      query2.Parameters.ParamByName('docsign_np').Value:='';
      query2.Parameters.ParamByName('websign_np').Value:='';
      query2.Parameters.ParamByName('doc12').Value:=docdm.mygem;
      query2.ExecSQL;

      query3.Close;
      query3.SQL.Clear;
      query3.SQL.Add('insert into trll_file (trllfac,trll01,trll02,trll03,trll04,trll061,trll062,trll071,trll072,trll15,trllacti,trllcla1) ');
      query3.SQL.Add(' values(:trllfac,:trll01,:trll02,:trll03,:trll04,:trll061,:trll062,:trll071,:trll072,:trll15,:trllacti,:trllcla1)');
      query3.Parameters.ParamByName('trllfac').Value:=docdm.myfac;
      query3.Parameters.ParamByName('trll01').Value:=trim(edyear.Text);
      query3.Parameters.ParamByName('trll02').Value:=formatdatetime('yyyy/mm/dd',now);
      query3.Parameters.ParamByName('trll03').Value:= no;
      query3.Parameters.ParamByName('trll04').Value:=query4.fieldbyname('gem03').AsString;
      query3.Parameters.ParamByName('trll061').Value:=formatdatetime('yyyy/mm/dd',now);
      query3.Parameters.ParamByName('trll062').Value:=formatdatetime('yyyy/mm/dd',now);
      query3.Parameters.ParamByName('trll071').Value:='00:00';
      query3.Parameters.ParamByName('trll072').Value:='00:00';
      query3.Parameters.ParamByName('trll15').Value:=query4.fieldbyname('gen05').AsString;
      query3.Parameters.ParamByName('trllacti').Value:='Y';
      query3.Parameters.ParamByName('trllcla1').Value:=trim(edt_doc081.Text);
      query3.ExecSQL;
      no_str:=no;
      puted :=  ftpclasssys(no,no_str);
      if puted = False then begin
        showmessage('檔案上傳失敗，請稍後再試!');
        if dm.topdb.InTransaction then
          dm.topdb.RollbackTrans;
        exit;
      end;
      if fileExists('D:\'+no+'.html') then
        Deletefile('D:\'+no+'.html');
      showmessage('儲存成功!');
      except
        if dm.glm.InTransaction then
          dm.glm.RollbackTrans;
        showmessage('儲存失敗!');
      end;
      if dm.glm.InTransaction then
        dm.glm.CommitTrans;
    end
    else begin
      if edfile.Text='' then begin
        Showmessage('請選擇上傳文件!');
        Exit;
      end;
      if edfile1.Text='' then begin  
        Showmessage('請選擇上傳文件!');
        Exit;
      end;
      try
        if not dm.glm.InTransaction then
          dm.glm.BeginTrans;
      adoexec.Close;
      adoexec.SQL.Clear;
      adoexec.SQL.Add('update doc_file set doc081=:doc081 where doc01=:doc01');
      adoexec.Parameters.ParamByName('doc01').Value:=lbno.Caption;
      adoexec.Parameters.ParamByName('doc081').Value:=trim(edt_doc081.Text);
      adoexec.ExecSQL;
      query1.Close;
      query1.SQL.Clear;
      query1.SQL.Add('update trll_doc set tdoc01=:tdoc01 where tdoc07=:tdoc07');
      query1.Parameters.ParamByName('tdoc01').Value:=trim(edt_doc081.Text);
      query1.Parameters.ParamByName('tdoc07').Value:=lbno.Caption;
      query1.ExecSQL;
      query2.Close;
      query2.SQL.Clear;
      query2.SQL.Add('update trll_file set trll04=:trll04,trll15=:trll15,trllcla1=:trllcla1 where trll03=:trll03');
      query2.Parameters.ParamByName('trll04').Value:=query4.fieldbyname('gem03').AsString;
      query2.Parameters.ParamByName('trll15').Value:=query4.fieldbyname('gen05').AsString;
      query2.Parameters.ParamByName('trll03').Value:=lbno.Caption;
      query3.Parameters.ParamByName('trllcla1').Value:=trim(edt_doc081.Text);
      query2.ExecSQL;
      no:=lbno.Caption;
      no_str:=no;
      with TMemoryStream.Create do try
        memo1.Lines[71]:='              <h2>' +edt_doc081.Text+ '</h2>';
        memo1.Lines[76]:='              <source src="media/'+edyear.Text+'/'+mydoc.fieldbyname('dst01').AsString+'/'+no+'.mp4" type=''video/mp4; ';
        memo1.Lines[77]:='              codecs="avc1.42E01E, mp4a.40.2"''></source>';
        memo1.Lines[78]:='              <source src="media/'+edyear.Text+'/'+mydoc.fieldbyname('dst01').AsString+'/'+no+'.ogv" type="video/ogg"></source>';
        S := #$FF#$FE;
        Write(S[1], Length(S));
        W := Memo1.Text;
        Write(W[1], Length(W) * SizeOf(WideChar));
        Position := 0;
        SaveToFile('D:\'+no+'.html');
        fileadd:='D:\'+no+'.html';
        fileno:= no;
        finally
          Free;
      end;
      puted :=  ftpclasssys(no,no_str);
      if puted = False then begin
        showmessage('檔案上傳失敗，請稍後再試!');
        if dm.topdb.InTransaction then
          dm.topdb.RollbackTrans;
        exit;
      end;
      if fileExists('D:\'+no+'.html') then
        Deletefile('D:\'+no+'.html');
      showmessage('儲存成功!');
      except
        if dm.glm.InTransaction then
          dm.glm.RollbackTrans;
        showmessage('儲存失敗!');
      end;
      if dm.glm.InTransaction then
        dm.glm.CommitTrans;
    end;
  end
  else begin
    adoquery1.Close;
    adoquery1.sql.Clear;
    adoquery1.SQL.Add('select dst09 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
    adoquery1.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
    adoquery1.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
    adoquery1.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
    adoquery1.Open;
    dst09 := adoquery1.fieldbyname('dst09').AsString;

    if status = 'INS' then begin
      if mydoc.FieldByName('dst08').AsInteger <> 999 then begin
        adoquery1.Close;
        adoquery1.sql.Clear;
        adoquery1.sql.Add('select * from doc_file where doc05 =:doc05 and doc06 =:doc06 and doc11 =:doc11 and doc02 =:doc02 and doc03 =:doc03 and doc04 =:doc04 and doc_acti=''Y'' and doc_fac ='''+docdm.myfac+'''');
        adoquery1.Parameters.ParamByName('doc05').Value := mydoc.fieldbyname('dst01').AsString;
        adoquery1.Parameters.ParamByName('doc06').Value := mydoc.fieldbyname('dst02').AsString;
        adoquery1.Parameters.ParamByName('doc11').Value := mydoc.fieldbyname('dst03').AsString;
        adoquery1.Parameters.ParamByName('doc02').Value := edyear.Text;
        if gbmon.Visible = True then
          adoquery1.Parameters.ParamByName('doc03').Value := edmonth.Text
        else
          adoquery1.Parameters.ParamByName('doc03').Value := '00';
        if gbseason.Visible = True then
          adoquery1.Parameters.ParamByName('doc04').Value := formatfloat('00',cbseason.itemindex+1)
        else
          adoquery1.Parameters.ParamByName('doc04').Value := '00';
        adoquery1.open;
        if adoquery1.recordcount > 0 then begin
          Showmessage('該文件已存在,無法新增!');
          Exit;
        end;
      end;

      try
        if not dm.glm.InTransaction then
          dm.glm.BeginTrans;

        //找出文件編號
        adoquery1.Close;
        adoquery1.sql.Clear;
        if  ss = '00' then begin
          adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno06 =:docno06 and docno_fac='''+docdm.myfac+'''');
          adoquery1.Parameters.ParamByName('docno06').Value := edmonth.Text;
        end
        else if ss ='01' then begin
          adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno05 =:docno05 and docno_fac='''+docdm.myfac+'''');
          adoquery1.Parameters.ParamByName('docno05').Value := formatfloat('00',cbseason.itemindex+1);
        end
        else if ss = '02' then begin
          adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno05 =''06'' and docno_fac='''+docdm.myfac+'''');
        end
        else if ss = '03' then begin
          adoquery1.sql.Add('select docno07 from doc_no where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno04 =:docno04 and docno05 =''99'' and docno_fac='''+docdm.myfac+'''');
        end;
        adoquery1.Parameters.ParamByName('docno01').Value := mydoc.fieldbyname('dst01').AsString;
        adoquery1.Parameters.ParamByName('docno02').Value := mydoc.fieldbyname('dst02').AsString;
        adoquery1.Parameters.ParamByName('docno03').Value := mydoc.fieldbyname('dst03').AsString;
        adoquery1.Parameters.ParamByName('docno04').Value := edyear.Text;
        adoquery1.Open;
        adoquery1.First;
        if not adoquery1.Eof then begin
          docno := adoquery1.fieldbyname('docno07').AsInteger+1;
          adoexec.Close;
          adoexec.sql.Clear;
          if ss = '00' then begin
            adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno06 =:docno06 and docno_fac='''+docdm.myfac+'''');
            adoexec.Parameters.ParamByName('docno06').Value := edmonth.Text;
          end
          else if ss = '01' then begin
            adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno05 =:docno05 and docno_fac='''+docdm.myfac+'''');
            adoexec.Parameters.ParamByName('docno05').Value := formatfloat('00',cbseason.itemindex+1);
          end
          else if ss = '02' then begin
            adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno05 =''06'' and docno_fac='''+docdm.myfac+'''');
          end
          else if ss = '03' then begin
            adoexec.SQL.Add('update doc_no set docno07 =:docno07 where docno01 =:docno01 and docno02 =:docno02 and docno03 =:docno03 and docno05 =''99'' and docno_fac='''+docdm.myfac+'''');
          end;
          adoexec.Parameters.ParamByName('docno01').Value := mydoc.fieldbyname('dst01').AsString;
          adoexec.Parameters.ParamByName('docno02').Value := mydoc.fieldbyname('dst02').AsString;
          adoexec.Parameters.ParamByName('docno03').Value := mydoc.fieldbyname('dst03').AsString;
          adoexec.Parameters.ParamByName('docno07').Value := formatfloat('0000',docno);
          adoexec.ExecSQL;
        end
        else begin
          docno := 1;
          adoexec.Close;
          adoexec.sql.Clear;
          adoexec.SQL.Add('insert into doc_no (docno01,docno02,docno03,docno04,docno05,docno06,docno07,docno_date,docno_fac) values ');
          adoexec.SQL.Add(' (:docno01,:docno02,:docno03,:docno04,:docno05,:docno06,:docno07,sysdate,'''+docdm.myfac+''') ');
          adoexec.Parameters.ParamByName('docno01').Value := mydoc.fieldbyname('dst01').AsString;
          adoexec.Parameters.ParamByName('docno02').Value := mydoc.fieldbyname('dst02').AsString;
          adoexec.Parameters.ParamByName('docno03').Value := mydoc.fieldbyname('dst03').AsString;
          adoexec.Parameters.ParamByName('docno04').Value := edyear.Text;
          adoexec.Parameters.ParamByName('docno07').Value := formatfloat('0000',docno);
          if ss = '00' then begin
            adoexec.Parameters.ParamByName('docno06').Value := formatfloat('00',strtoint(edmonth.Text));
            adoexec.Parameters.ParamByName('docno05').Value := '00';
          end
          else if ss = '01' then begin
            adoexec.Parameters.ParamByName('docno06').Value := '00';
            adoexec.Parameters.ParamByName('docno05').Value := formatfloat('00',cbseason.ItemIndex+1);;
          end
          else if ss = '02' then begin
            adoexec.Parameters.ParamByName('docno06').Value := '00';
            adoexec.Parameters.ParamByName('docno05').Value := '00';
          end
          else if ss = '03' then begin
            adoexec.Parameters.ParamByName('docno06').Value := '00';
            adoexec.Parameters.ParamByName('docno05').Value := '99';
          end;
          adoexec.ExecSQL;
        end;

        if ss = '00' then
          no := docdm.myfac+trim(edyear.Text)+formatfloat('00',strtoint(edmonth.Text))+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno)
        else if ss = '01' then
          no := docdm.myfac+trim(edyear.Text)+formatfloat('00',cbseason.ItemIndex+1)+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno)
        else if ss = '02' then
          no := docdm.myfac+trim(edyear.Text)+'06'+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno)
        else if ss = '03' then
          no := docdm.myfac+trim(edyear.Text)+'99'+mydoc.fieldbyname('dst01').AsString+mydoc.fieldbyname('dst02').AsString+mydoc.fieldbyname('dst03').AsString+formatfloat('0000',docno);

        //新增記錄
        adoexec.Close;
        adoexec.sql.Clear;
        adoexec.SQL.Add('insert into doc_file(doc01,doc02,doc03,doc11,doc04,doc05,doc06,doc07,doc08,doc081,createdate,createuser,doc09,docsign,websign,docmemo,doc_acti,doc_fac,docsign_np,websign_np,doc12) ');
        adoexec.sql.Add(' values (:doc01,:doc02,:doc03,:doc11,:doc04,:doc05,:doc06,:doc07,:doc08,:doc081,sysdate,'''+dm.getlogin+''',:doc09,:docsign,:websign,:docmemo,''Y'','''+docdm.myfac+''',:docsign_np,:websign_np,:doc12)');
        adoexec.Parameters.ParamByName('doc01').Value := no;
        adoexec.Parameters.ParamByName('doc02').Value := trim(edyear.Text);
        if ss = '00' then begin
          adoexec.Parameters.ParamByName('doc03').Value := edmonth.text;
          adoexec.Parameters.ParamByName('doc04').Value := '00';
        end
        else if ss = '01' then begin
          adoexec.Parameters.ParamByName('doc03').Value := '00';
          adoexec.Parameters.ParamByName('doc04').Value := cbseason.itemindex+1;
        end
        else if ss = '02' then begin
          adoexec.Parameters.ParamByName('doc03').Value := '00';
          adoexec.Parameters.ParamByName('doc04').Value := '06';
        end
        else if ss = '03' then begin
          adoexec.Parameters.ParamByName('doc03').Value := '00';
          adoexec.Parameters.ParamByName('doc04').Value := '99';
        end;

        adoexec.Parameters.ParamByName('doc05').Value := mydoc.fieldbyname('dst01').AsString;
        adoexec.Parameters.ParamByName('doc06').Value := mydoc.fieldbyname('dst02').AsString;
        adoexec.Parameters.ParamByName('doc11').Value := mydoc.fieldbyname('dst03').AsString;
        adoexec.Parameters.ParamByName('doc07').Value := trim(no)+'.'+copy(lab_filetype.caption,3,3);
        adoexec.Parameters.ParamByName('doc08').Value := mydoc.fieldbyname('dst04').AsString;
        adoexec.Parameters.ParamByName('doc081').Value := trim(edt_doc081.Text);
        if ss = '00' then
          adoexec.Parameters.ParamByName('doc09').Value := edyear.Text+'年'+edmonth.text+'月'
        else if ss = '01' then
          adoexec.Parameters.ParamByName('doc09').Value := edyear.Text+'年'+cbseason.Text+'季'
        else if ss = '02' then
          adoexec.Parameters.ParamByName('doc09').Value := edyear.Text+'年半年報'
        else if ss = '03' then
          adoexec.Parameters.ParamByName('doc09').Value := edyear.Text+'年';

        adoexec.Parameters.ParamByName('docsign').Value := '0';

        adoquery1.Close;
        adoquery1.sql.Clear;
        adoquery1.SQL.Add('select dst09 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
        adoquery1.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
        adoquery1.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
        adoquery1.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
        adoquery1.Open;
        dst09 := adoquery1.fieldbyname('dst09').AsString;


        adoquery1.close;
        adoquery1.sql.clear;
        adoquery1.sql.add('Select subject From sysparameter,doc_setup Where isActive = ''Y'' And sysname = ''DOC'' And item = ''DST06'' And subject = dst06 and dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
        adoquery1.Parameters.parambyname('dst01').value := mydoc.fieldbyname('dst01').AsString;
        adoquery1.Parameters.parambyname('dst02').value := mydoc.fieldbyname('dst02').AsString;
        adoquery1.Parameters.parambyname('dst03').value := mydoc.fieldbyname('dst03').AsString;
        adoquery1.open;
        if dst09 = 'Y' then begin
          if adoquery1.fieldbyname('subject').asstring ='D' then begin
            adoexec.Parameters.ParamByName('DOCSIGN').Value := '0';
            adoexec.Parameters.ParamByName('WEBSIGN').Value := 'X';
          end
          else if adoquery1.fieldbyname('subject').asstring ='W' then begin
            adoexec.Parameters.ParamByName('DOCSIGN').Value := 'X';
            adoexec.Parameters.ParamByName('WEBSIGN').Value := '0';
          end
          else if adoquery1.fieldbyname('subject').asstring ='A' then begin
            adoexec.Parameters.ParamByName('DOCSIGN').Value := '0';
            adoexec.Parameters.ParamByName('WEBSIGN').Value := '0';
          end;
        end
        else begin
          adoexec.Parameters.ParamByName('DOCSIGN').Value := 'X';
          adoexec.Parameters.ParamByName('WEBSIGN').Value := 'X';
        end;
        webs :=  adoquery1.fieldbyname('subject').asstring;
        doctypes := adoquery1.fieldbyname('subject').asstring;
        adoexec.Parameters.ParamByName('docmemo').Value := trim(memo2.Text);

        if ((webs ='D') or (webs ='A')) and (dst09 <> 'N') then begin
          adoquery1.close;
          adoquery1.sql.clear;
          adoquery1.sql.add('select azc03 from azc_file'+docdm.dblinks+' where azc01 ='''+docdm.mygem+''' and azc02 =1 ');
          adoquery1.open;
          if adoquery1.recordcount > 0 then begin
            adoexec.Parameters.ParamByName('docsign_np').Value := adoquery1.fieldbyname('azc03').asstring;
            docsignnp := adoquery1.fieldbyname('azc03').asstring;
          end
          else begin
            adoexec.Parameters.ParamByName('docsign_np').Value := '';
            docsignnp := '';
          end;
        end
        else begin
          adoexec.Parameters.ParamByName('docsign_np').Value := '';
          docsignnp := '';
        end;

        if ((webs ='W') or (webs ='A')) and (dst09 <> 'N') then begin
          adoquery1.close;
          adoquery1.sql.clear;
          adoquery1.sql.add('select azc03 from azc_file'+docdm.dblinks+' where azc01 =''WEB'' and azc02 =1 ');
          adoquery1.open;
          if adoquery1.recordcount > 0 then begin
            adoexec.Parameters.ParamByName('websign_np').Value := adoquery1.fieldbyname('azc03').asstring;
            websignnp := adoquery1.fieldbyname('azc03').asstring;
          end
          else begin
            adoexec.Parameters.ParamByName('websign_np').Value := '';
            websignnp := '';
          end;
        end
        else begin
          adoexec.Parameters.ParamByName('websign_np').Value := '';
          websignnp := '';
        end;

        adoexec.Parameters.ParamByName('doc12').Value := docdm.mygem;
        adoexec.ExecSQL;
        lbfiles.Caption := no+ExtractFileExt(edfile.text);

        //ftp檔案
        if (doctypes ='D') or (doctypes ='A') then begin
          puted :=  ftpdocsys(no);
          if puted = False then begin
            showmessage('檔案上傳失敗，請稍後再試!');
            if dm.topdb.InTransaction then
              dm.topdb.RollbackTrans;
            exit;
          end;
        end;

        //ftp檔案
        if (doctypes ='W') or (doctypes ='A') then begin
          puted :=  ftpwebdocsys(no);
          if puted = False then begin
            showmessage('檔案上傳失敗，請稍後再試!');
            if dm.topdb.InTransaction then
              dm.topdb.RollbackTrans;
            exit;
          end;
        end;

        if dm.glm.InTransaction then
          dm.glm.CommitTrans;
        showmessage('儲存成功!');
        lbno.caption := no;

        if dst09 ='Y'  then begin
          tmpque.Close;
          tmpque.sql.Clear;
          tmpque.SQL.Add('select type_detail from sys_codedef'+docdm.dblinks+' where cd_type=''09''');
          tmpque.Open;
          if tmpque.RecordCount > 0 then
            mailstr := tmpque.fieldbyname('type_detail').AsString
          else
            mailstr := '';

          fmMail := TfmMail.Create(self);
          if docsignnp <> '' then
            fmMail.Smtp.PostMessage.ToAddress.Append(docsignnp+mailstr);
          if websignnp <> '' then
            fmMail.Smtp.PostMessage.ToAddress.Append(websignnp+mailstr);
          fmMail.smtp.postmessage.ToCarbonCopy.Append(dm.getlogin+mailstr);
          fmMail.Edit4.Text := '文件電子報系統 : '+edyear.Text+'年'+edmonth.text+'月'+mydoc.fieldbyname('dst04').AsString+' 等待簽核';
          fmMail.RichEdit2.Lines.Append(docsignnp+' '+websignnp+' '+'您好 : ');
          fmMail.RichEdit2.Lines.Append(edyear.Text+'年'+edmonth.text+'月'+mydoc.fieldbyname('dst04').AsString+' 等待您的簽核放行');
          if trim(memo2.Text) <> '' then begin
            fmMail.RichEdit2.Lines.Append('備註說明:'+trim(memo2.Text));
          end;
          fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心');
          fmMail.ShowModal;
          fmMail.Free;
        end
        else begin
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
          tmpque.Parameters.ParamByName('docr01').Value := mydoc.fieldbyname('dst01').AsString;
          tmpque.Parameters.ParamByName('docr02').Value := mydoc.fieldbyname('dst02').AsString;
          tmpque.Parameters.ParamByName('docr03').Value := mydoc.fieldbyname('dst03').AsString;
          tmpque.Open;

          fmMail := TfmMail.Create(self);
          tmpque.First;
          while not tmpque.Eof do begin
            if tmpque.fieldbyname('docr05').AsString <> 'all' then
              fmMail.Smtp.PostMessage.ToAddress.Append(tmpque.fieldbyname('docr05').AsString);
            tmpque.Next;
          end;
          fmMail.Smtp.PostMessage.ToAddress.Append(dm.getlogin+mailstr);
          fmMail.Edit4.Text := '文件電子報系統 : '+edyear.text+'年'+edmonth.text+'月'+edt_doc081.Text+' 發佈完成';

          fmMail.RichEdit2.Lines.Append('各位同仁您好 : ');
          fmMail.RichEdit2.Lines.Append('  ');
          fmMail.RichEdit2.Lines.Append(edyear.text+'年'+edmonth.text+'月'+edt_doc081.Text+' 發佈完成');
          fmMail.RichEdit2.Lines.Append('請點選 http://i.gshank.com  查看您的文件 ');
          fmMail.RichEdit2.Lines.Append('若有任何使用上的問題請與資訊室連絡 分機:126');
          fmMail.RichEdit2.lines.Append('   ');
          fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心 b2b@mail.gs.com.tw');
          fmMail.ShowModal;
          fmMail.Free;

        end;

        if MessageDlg('是否新增文件?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          btn_ins.Click
        else
          status := 'UPD';
      except
        if dm.glm.InTransaction then
          dm.glm.RollbackTrans;
        showmessage('儲存失敗!');
      end;
    end
    else begin   //update


      try
        if not dm.glm.InTransaction then
          dm.glm.BeginTrans;

        //修改記錄
        adoexec.Close;
        adoexec.sql.Clear;

        adoquery1.close;
        adoquery1.sql.clear;
        adoquery1.sql.add('Select subject From sysparameter,doc_setup Where isActive = ''Y'' And sysname = ''DOC'' And item = ''DST06'' And subject = dst06 and dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
        adoquery1.parameters.parambyname('dst01').value := mydoc.fieldbyname('dst01').AsString;
        adoquery1.parameters.parambyname('dst02').value := mydoc.fieldbyname('dst02').AsString;
        adoquery1.parameters.parambyname('dst03').value := mydoc.fieldbyname('dst03').AsString;
        adoquery1.open;

        adoexec.SQL.Add('update doc_file set docsign =:docsign,websign =:websign,docsign_np =:docsign_np,websign_np =:websign_np,doc081 =:doc081,docmemo=:docmemo,updatedate =to_date(:updatedate,''yyyy/MM/dd''),updateuser =:updateuser where doc01 ='''+lbno.Caption+'''');

        adoquery1.Close;
        adoquery1.sql.Clear;
        adoquery1.SQL.Add('select dst09 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
        adoquery1.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
        adoquery1.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
        adoquery1.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
        adoquery1.Open;
        dst09 := adoquery1.fieldbyname('dst09').AsString;


        adoquery1.close;
        adoquery1.sql.clear;
        adoquery1.sql.add('Select subject From sysparameter,doc_setup Where isActive = ''Y'' And sysname = ''DOC'' And item = ''DST06'' And subject = dst06 and dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
        adoquery1.Parameters.parambyname('dst01').value := mydoc.fieldbyname('dst01').AsString;
        adoquery1.Parameters.parambyname('dst02').value := mydoc.fieldbyname('dst02').AsString;
        adoquery1.Parameters.parambyname('dst03').value := mydoc.fieldbyname('dst03').AsString;
        adoquery1.open;
        if dst09 = 'Y' then begin
          if adoquery1.fieldbyname('subject').asstring ='D' then begin
            adoexec.Parameters.ParamByName('DOCSIGN').Value := '0';
            adoexec.Parameters.ParamByName('WEBSIGN').Value := 'X';
          end
          else if adoquery1.fieldbyname('subject').asstring ='W' then begin
            adoexec.Parameters.ParamByName('DOCSIGN').Value := 'X';
            adoexec.Parameters.ParamByName('WEBSIGN').Value := '0';
          end
          else if adoquery1.fieldbyname('subject').asstring ='A' then begin
            adoexec.Parameters.ParamByName('DOCSIGN').Value := '0';
            adoexec.Parameters.ParamByName('WEBSIGN').Value := '0';
          end;
        end
        else begin
          adoexec.Parameters.ParamByName('DOCSIGN').Value := 'X';
          adoexec.Parameters.ParamByName('WEBSIGN').Value := 'X';
        end;
        webs :=  adoquery1.fieldbyname('subject').asstring;
        doctypes := adoquery1.fieldbyname('subject').asstring;
        adoexec.Parameters.ParamByName('docmemo').Value := trim(memo2.Text);

        if ((webs ='D') or (webs ='A')) and (dst09 <> 'N') then begin
          adoquery1.close;
          adoquery1.sql.clear;
          adoquery1.sql.add('select azc03 from azc_file'+docdm.dblinks+' where azc01 ='''+docdm.mygem+''' and azc02 =1 ');
          adoquery1.open;
          if adoquery1.recordcount > 0 then begin
            adoexec.Parameters.ParamByName('docsign_np').Value := adoquery1.fieldbyname('azc03').asstring;
            docsignnp := adoquery1.fieldbyname('azc03').asstring;
          end
          else begin
            adoexec.Parameters.ParamByName('docsign_np').Value := '';
            docsignnp := '';
          end;
        end
        else begin
          adoexec.Parameters.ParamByName('docsign_np').Value := '';
          docsignnp := '';
        end;

        if ((webs ='W') or (webs ='A')) and (dst09 <> 'N') then begin
          adoquery1.close;
          adoquery1.sql.clear;
          adoquery1.sql.add('select azc03 from azc_file'+docdm.dblinks+' where azc01 =''WEB'' and azc02 =1 ');
          adoquery1.open;
          if adoquery1.recordcount > 0 then begin
            adoexec.Parameters.ParamByName('websign_np').Value := adoquery1.fieldbyname('azc03').asstring;
            websignnp := adoquery1.fieldbyname('azc03').asstring;
          end
          else begin
            adoexec.Parameters.ParamByName('websign_np').Value := '';
            websignnp := '';
          end;
        end
        else begin
          adoexec.Parameters.ParamByName('websign_np').Value := '';
          websignnp := '';
        end;

        adoexec.Parameters.ParamByName('updatedate').Value := datetostr(date);
        adoexec.Parameters.ParamByName('updateuser').Value := dm.getlogin;
        adoexec.Parameters.ParamByName('docmemo').Value := trim(memo2.Text);
        adoexec.Parameters.ParamByName('doc081').Value := trim(edt_doc081.Text);
        adoexec.ExecSQL;

        //ftp檔案
        if (webs ='D') or (webs ='A') then begin
          puted :=  ftpdocsys(lbno.Caption);
          if puted = False then begin
            showmessage('檔案上傳失敗，請稍後再試!');
            if dm.topdb.InTransaction then
              dm.topdb.RollbackTrans;
            exit;
          end;
        end;

        //ftp檔案
        if (webs ='W') or (webs ='A') then begin
          puted :=  ftpwebdocsys(lbno.Caption);
          if puted = False then begin
            showmessage('檔案上傳失敗，請稍後再試!');
            if dm.topdb.InTransaction then
              dm.topdb.RollbackTrans;
            exit;
          end;
        end;


        if dst09 ='Y'  then begin
          tmpque.Close;
          tmpque.sql.Clear;
          tmpque.SQL.Add('select type_detail from sys_codedef'+docdm.dblinks+' where cd_type=''09''');
          tmpque.Open;
          if tmpque.RecordCount > 0 then
            mailstr := tmpque.fieldbyname('type_detail').AsString
          else
            mailstr := '';

          fmMail := TfmMail.Create(self);
          if docsignnp <> '' then
            fmMail.Smtp.PostMessage.ToAddress.Append(docsignnp+mailstr);
          if websignnp <> '' then
            fmMail.Smtp.PostMessage.ToAddress.Append(websignnp+mailstr);
          fmMail.smtp.postmessage.ToCarbonCopy.Append(dm.getlogin+mailstr);
          fmMail.Edit4.Text := '文件電子報系統 : '+edyear.Text+'年'+edmonth.text+'月'+mydoc.fieldbyname('dst04').AsString+' 等待簽核';
          fmMail.RichEdit2.Lines.Append(docsignnp+' '+websignnp+' '+'您好 : ');
          fmMail.RichEdit2.Lines.Append(edyear.Text+'年'+edmonth.text+'月'+mydoc.fieldbyname('dst04').AsString+' 等待您的簽核放行');
          if trim(memo2.Text) <> '' then begin
            fmMail.RichEdit2.Lines.Append('備註說明:'+trim(memo2.Text));
          end;
          fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心');
          fmMail.ShowModal;
          fmMail.Free;
        end
        else begin
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
          tmpque.Parameters.ParamByName('docr01').Value := mydoc.fieldbyname('dst01').AsString;
          tmpque.Parameters.ParamByName('docr02').Value := mydoc.fieldbyname('dst02').AsString;
          tmpque.Parameters.ParamByName('docr03').Value := mydoc.fieldbyname('dst03').AsString;
          tmpque.Open;

          fmMail := TfmMail.Create(self);
          tmpque.First;
          while not tmpque.Eof do begin
            if tmpque.fieldbyname('docr05').AsString <> 'all' then
              fmMail.Smtp.PostMessage.ToAddress.Append(tmpque.fieldbyname('docr05').AsString);
            tmpque.Next;
          end;
          fmMail.Smtp.PostMessage.ToAddress.Append(dm.getlogin+mailstr);
          fmMail.Edit4.Text := '文件電子報系統 : '+edyear.text+'年'+edmonth.text+'月'+edt_doc081.Text+' 發佈完成';

          fmMail.RichEdit2.Lines.Append('各位同仁您好 : ');
          fmMail.RichEdit2.Lines.Append('  ');
          fmMail.RichEdit2.Lines.Append(edyear.text+'年'+edmonth.text+'月'+edt_doc081.Text+' 發佈完成');
          fmMail.RichEdit2.Lines.Append('請點選 http://i.gshank.com 查看您的文件 ');
          fmMail.RichEdit2.Lines.Append('若有任何使用上的問題請與資訊室連絡 分機:126');
          fmMail.RichEdit2.lines.Append('   ');
          fmMail.RichEdit2.Lines.Append('鉅祥企業訊息中心 b2b@mail.gs.com.tw');
          fmMail.ShowModal;
          fmMail.Free;

        end;

        if dm.glm.InTransaction then
          dm.glm.CommitTrans;
        showmessage('儲存成功!');
      except
        if dm.glm.InTransaction then
          dm.glm.RollbackTrans;
        showmessage('儲存失敗!');
      end;
    end;
  end;
  query1.Free;
  query2.Free;
  query3.Free;
  query4.Free;
  BitBtn1.Cursor:= crdefault;
end;

function TfmIns.ftpclasssys(filenames,filenames1:string):boolean;
var
  i : integer;
  myfile,myfile1,myfile2 : string;
  str_ext,str_ext1:string;
  Excel :variant;
  query1 : TAdoquery;
begin
  query1 := TAdoquery.Create(self);
  query1.Connection := dm.glm;
  query1.Close;
  query1.SQL.Clear;
  query1.SQL.Add('select subject from sysparameter where item=''doc05'' and isactive =''Y'' and subject=''V''');
  query1.Open;
  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.sql.Add('select subject from sysparameter where sysname=''DOCVFTP'' and isactive =''Y''');
  tmpque.Open;
  if tmpque.RecordCount > 0 then begin
    IdFTP1.Host := tmpque.fieldbyname('subject').AsString;
    IdFTP2.Host := tmpque.fieldbyname('subject').AsString;
  end
  else begin
    result := False;
    Exit;
  end;

  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.sql.Add('select item,subject from sysparameter where sysname=''DOCVFTPU'' and isactive =''Y''');
  tmpque.Open;
  if tmpque.RecordCount > 0 then begin
    IdFTP1.Password := tmpque.fieldbyname('subject').AsString;
    IdFTP1.User := tmpque.fieldbyname('item').AsString;
    IdFTP2.Password := tmpque.fieldbyname('subject').AsString;
    IdFTP2.User := tmpque.fieldbyname('item').AsString;
  end
  else begin
    result := False;
    Exit;
  end;

   str_ext := uppercase(copy(ExtractFileExt(trim(edfile.Text)),2,3));
   if str_ext = 'MP4' then  begin
     myfile1 := filenames+'.MP4';
     lbfiles1.Caption := myfile1;
   end
   else
     myfile1 := '';

   str_ext1 := uppercase(copy(ExtractFileExt(trim(edfile1.Text)),2,3));
   if str_ext1 = 'OGV' then  begin
     myfile2 := filenames1+'.OGV';
     lbfiles2.Caption := myfile2;
   end
   else
     myfile2 := '';
  list1.Clear;
  if IdFTP1.Connected then
    IdFTP1.Disconnect;
  IdFTP1.Connect;
  IdFTP1.ChangeDir('media');
  try
    IdFTP1.ChangeDir(edyear.Text);
  except
    IdFTP1.MakeDir(edyear.Text);
    IdFTP1.ChangeDir(edyear.Text);
  end;
  try
    IdFTP1.ChangeDir(mydoc.fieldbyname('dst01').AsString);
  except
    IdFTP1.MakeDir(mydoc.fieldbyname('dst01').AsString);
    IdFTP1.ChangeDir(mydoc.fieldbyname('dst01').AsString);
  end;
  try
    IdFTP1.List(list1, '*.*', false);
  except
  end;
  try
    myfile := filenames+'.'+copy(lab_filetype.Caption,pos(' ',lab_filetype.Caption)+1,length(lab_filetype.Caption)-2);
    for i := 0 to list1.Count -1 do begin
      if list1.Strings[i] = myfile  then
        IdFTP1.Delete(trim(myfile));
    end;
  except
  end;

  if trim(myfile1) <> '' then begin
    try
      for i := 0 to list1.Count -1 do begin
          if list1.Strings[i] = myfile1  then
            IdFTP1.Delete(trim(myfile));
      end;
    except
    end;
  end;

try
    IdFTP1.List(list1, '*.*', false);
  except
  end;
  try
    myfile2 := filenames1+'.ogv';
    for i := 0 to list1.Count -1 do begin
      if list1.Strings[i] = myfile2  then
        IdFTP1.Delete(trim(myfile2));
    end;
  except
  end;

  if IdFTP2.Connected then
    IdFTP2.Disconnect;
  IdFTP2.Connect;
  try
    IdFTP2.Put(trim(fileadd),fileno+'.html');
    result := True;
    IdFTP2.Disconnect;
  except
    if IdFTP2.Connected then
      IdFTP2.Disconnect;
    result := false;
  end;
  try
    IdFTP1.Put(trim(edfile.Text),myfile);
    IdFTP1.Put(trim(edfile1.Text),myfile2);
    result := True;
    IdFTP1.Disconnect;
  except
    if IdFTP1.Connected then
      IdFTP1.Disconnect;
    result := false;
  end;

end;

function TfmIns.ftpdocsys(filenames:string):boolean;
var
  i : integer;
  myfile,myfile1 : string;
  str_ext:string;
  Excel :variant;
begin
  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.sql.Add('select subject from sysparameter where sysname=''DOCFTP'' and isactive =''Y''');
  tmpque.Open;
  if tmpque.RecordCount > 0 then
    idftpdoc.Host := tmpque.fieldbyname('subject').AsString
  else begin
    result := False;
    Exit;
  end;

  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.sql.Add('select item,subject from sysparameter where sysname=''DOCFTPU'' and isactive =''Y''');
  tmpque.Open;
  if tmpque.RecordCount > 0 then begin
    idftpdoc.Password := tmpque.fieldbyname('subject').AsString;
    idftpdoc.User := tmpque.fieldbyname('item').AsString;
  end
  else begin
    result := False;
    Exit;
  end;

   str_ext := uppercase(copy(ExtractFileExt(trim(edfile.Text)),2,3));
   if str_ext = 'XLS' then  begin
     myfile1 := filenames+'.pdf';
     lbfiles1.Caption := myfile1;
   end
   else
     myfile1 := '';


  list1.Clear;
  if idftpdoc.Connected then
    idftpdoc.Disconnect;
  idftpdoc.Passive := true;
  idftpdoc.Connect;
  try
    idftpdoc.ChangeDir(docdm.myfac);
  except
    idftpdoc.MakeDir(docdm.myfac);
    idftpdoc.ChangeDir(docdm.myfac);
  end;

  try
    idftpdoc.ChangeDir(edyear.Text);
  except
    idftpdoc.MakeDir(edyear.Text);
    idftpdoc.ChangeDir(edyear.Text);
  end;

  try
    idftpdoc.ChangeDir(copy(mydoc.fieldbyname('dst01').AsString,1,1));
  except
    idftpdoc.MakeDir(copy(mydoc.fieldbyname('dst01').AsString,1,1));
    idftpdoc.ChangeDir(copy(mydoc.fieldbyname('dst01').AsString,1,1));
  end;

  try

    //idftpdoc.List(list1, '*.*', false);
  except
  end;

  try
    myfile := filenames+'.'+copy(lab_filetype.Caption,pos(' ',lab_filetype.Caption)+1,length(lab_filetype.Caption)-2);
    if idftpdoc.Size(myfile) > 0 then
       idftpdoc.Delete(trim(myfile));

  except
  end;

  if trim(myfile1) <> '' then begin
    if idftpdoc.Size(myfile) > 0 then
       idftpdoc.Delete(trim(myfile));
  end;

  try
    idftpdoc.Put(trim(edfile.Text),myfile);
    result := True;
    idftpdoc.Disconnect;
  except
    if idftpdoc.Connected then
      idftpdoc.Disconnect;
    result := false;
  end;

end;

function TfmIns.ftpwebdocsys(filenames:string):boolean;
var
  myfile,myfile1 : string;
  i: integer;
  str_ext:string;
  Excel :variant;
begin
  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.sql.Add('select subject from sysparameter where sysname=''DOCWEBFTP'' and isactive =''Y''');
  tmpque.Open;
  if tmpque.RecordCount > 0 then
    idftpweb.Host := tmpque.fieldbyname('subject').AsString
  else begin
    result := False;
    Exit;
  end ;

  tmpque.Close;
  tmpque.sql.Clear;
  tmpque.sql.Add('select item,subject from sysparameter where sysname=''DOCWEBFTPU'' and isactive =''Y''');
  tmpque.Open;
  if tmpque.RecordCount > 0 then begin
    idftpweb.Password := tmpque.fieldbyname('subject').AsString;
    idftpweb.User := tmpque.fieldbyname('item').AsString;
  end
  else begin
    result := False;
    Exit;
  end;

   str_ext := uppercase(copy(ExtractFileExt(trim(edfile.Text)),2,3));
   if str_ext = 'XLS' then  begin
     myfile1 := filenames+'.pdf';
     lbfiles1.Caption := myfile1;
   end
   else
     myfile1 := '';


  list1.Clear;
  if idftpweb.Connected then
    idftpweb.Disconnect;
  idftpweb.Connect;

  idftpweb.List(list1, filenames+'.'+copy(lab_filetype.Caption,pos(' ',lab_filetype.Caption)+1,length(lab_filetype.Caption)-2), true);

  try
    myfile := filenames+'.'+copy(lab_filetype.Caption,pos(' ',lab_filetype.Caption)+1,length(lab_filetype.Caption)-2);
    for i := 0 to list1.Count -1 do begin
      if list1.Strings[i] = myfile  then
        idftpdoc.Delete(trim(myfile));
      
    end;
  except
  end;

  if trim(myfile1) <> '' then begin
    try
      for i := 0 to list1.Count -1 do begin
          if list1.Strings[i] = myfile1  then
            idftpweb.Delete(trim(myfile));
      end;
    except
    end;
  end;

  try
    idftpweb.Put(trim(edfile.Text),myfile);
    result := True;
    idftpweb.Disconnect;
  except
    if idftpweb.Connected then
      idftpweb.Disconnect;
    result := false;
  end;

end;

procedure TfmIns.BitBtn2Click(Sender: TObject);
var
  copyed : boolean;
begin
  if fileexists('T:\FinXml\WEB\'+lbfiles.caption) then begin
    if MessageDlg('請問是否刪除已上載檔案?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
      deletefile('T:\FinXml\WEB\'+lbfiles.caption);
      copyfile(pchar(edfile.text),pchar('T:\FinXml\WEB\'+lbno.caption+ExtractFileExt(edfile.text)),copyed);
      if copyed then begin
        adoexec.close;
        adoexec.sql.clear;
        adoexec.sql.add('update doc_file set doc07 =:doc07 where doc01 =:doc01');
        adoexec.parameters.parambyname('doc01').value := lbno.caption;
        adoexec.parameters.parambyname('doc07').value := lbno.caption+ExtractFileExt(edfile.text);
        adoexec.execsql;
        showmessage('上載成功!');
      end
      else
        showmessage('上載失敗!');
    end;
  end;
end;

procedure TfmIns.Button3Click(Sender: TObject);
begin
  close;
end;

procedure TfmIns.cb_detailExit(Sender: TObject);
var
  dst05 : string;
begin
{  adoquery1.Close;
  adoquery1.sql.Clear;
  adoquery1.sql.Add('select * from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
  adoquery1.Parameters.ParamByName('dst01').Value := copy(cbtype.Text,1,2);
  adoquery1.Parameters.ParamByName('dst02').Value := copy(cbchild.Text,1,2);
  adoquery1.Parameters.ParamByName('dst03').Value := copy(cb_detail.Text,1,2);
  adoquery1.Open;
  gbyear.Visible := True;
  if not adoquery1.Eof then begin
    //小於90代表需輸入月
    if adoquery1.FieldByName('dst08').AsInteger < 90 then begin
      gbmon.Visible := True;
      gbseason.Visible := false;
      ss := '00';
    end //90輸入季
    else if (adoquery1.FieldByName('dst08').AsInteger = 90) then begin
      gbseason.Visible := True;
      gbmon.Visible := False;
      ss := '01';
    end // 180是半年一次輸入資料
    else if (adoquery1.FieldByName('dst08').AsInteger = 180) then begin
      gbseason.Visible := false;
      gbmon.Visible := false;
      ss := '02';
    end   //999隨時可輸入資料
    else if (adoquery1.FieldByName('dst08').AsInteger = 999) then begin
      gbseason.Visible := false;
      gbmon.Visible := false;
      ss := '03';
    end;

    dst05 := adoquery1.fieldbyname('dst05').AsString;
  end;

  ed_repdesc.Text := adoquery1.fieldbyname('dst04').AsString;
  adoquery1.Close;
  adoquery1.sql.Clear;
  adoquery1.SQL.Add('Select substr(remark,1,3) From sysparameter Where isActive = ''Y'' And sysname = ''DOC'' And item = ''DST05'' and subject ='''+dst05+'''');
  adoquery1.Open;
  adoquery1.First;
  if adoquery1.RecordCount > 0 then
    lab_filetype.Caption := dst05+' '+adoquery1.Fields.Fields[0].AsString
  else
    lab_filetype.Caption := '';     }
  
end;

procedure TfmIns.cbChildExit(Sender: TObject);
begin
{  cb_detail.Clear;
  adoquery1.Close;
  adoquery1.SQL.Clear;
  if docque.status = 'SIGN' then
    adoquery1.SQL.Add('select dst03,dst04 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst_acti =''Y'' and dst_fac='''+docdm.myfac+'''')
  else
    adoquery1.SQL.Add('select dst03,dst04 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst_acti =''Y'' and dst_fac='''+docdm.myfac+''' and (dst07 ='''+dm.getlogin+''' or dst071 ='''+dm.getlogin+''')');
  adoquery1.Parameters.ParamByName('dst01').Value := copy(cbtype.Text,1,2);
  adoquery1.Parameters.ParamByName('dst02').Value := copy(cbchild.Text,1,2);
  adoquery1.Open;
  adoquery1.First;
  while not adoquery1.Eof do begin
    cb_detail.Items.Append(adoquery1.Fields[0].asstring+' '+adoquery1.Fields[1].asstring);
    adoquery1.Next;
  end;
  cb_detail.ItemIndex := 0;  }

end;

procedure TfmIns.btn_webview1Click(Sender: TObject);
begin
  if lbfiles.Caption <> '' then begin
    tmpque.Close;
    tmpque.sql.Clear;
    tmpque.SQL.Add('select dst06 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03');
    tmpque.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
    tmpque.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
    tmpque.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
    tmpque.Open;
    if (tmpque.FieldByName('dst06').AsString = 'W') or (tmpque.FieldByName('dst06').AsString = 'A') then begin
      tmpque.Close;
      tmpque.sql.Clear;
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCWEBFTP'' and isactive =''Y''');
      tmpque.Open;
      if tmpque.RecordCount > 0 then
        ShellExecute(0, PChar('open'), PChar('http://'+tmpque.FieldByName('subject').AsString+'/doc/'+lbfiles.Caption), nil , nil, SW_SHOWNA);
    end
    else
      showmessage('未有網頁文件可瀏覽!');
  end
  else
    showmessage('尚未上載檔案');

end;

procedure TfmIns.FormClose(Sender: TObject; var Action: TCloseAction);
var  query1 : TAdoquery;
begin
  query1 := TAdoquery.Create(self);
  query1.Connection := dm.glm;
  query1.Close;
  query1.SQL.Clear;
  query1.SQL.Add('select dst01,dst02,dst05 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
  query1.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
  query1.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
  query1.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
  query1.Open;
  if query1.FieldByName('dst05').AsString='V' then begin
    Label2.Visible:=true;
    edfile1.Visible:=true;
    button3.Visible:=true;
    button4.Visible:=true;
  end
  else begin
    Label2.Visible:=false;
    edfile1.Visible:=false;
    button3.Visible:=false;
    button4.Visible:=false;
  end;

  if idftpdoc.Connected then
    idftpdoc.Disconnect;

  if idftpweb.Connected then
    idftpweb.Disconnect;

  list1.Free;
end;

procedure TfmIns.mydocAfterScroll(DataSet: TDataSet);
var
  year,month,day : word;
begin
  if status = 'INS' then
    edt_doc081.Text := mydoc.fieldbyname('dst04').AsString;
  if mydoc.fieldbyname('dst05').AsString='V' then
    edt_doc081.Text :='';

  decodedate(now,year,month,day);
  adoquery1.Close;
  adoquery1.sql.Clear;
  adoquery1.sql.Add('select * from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03 and dst_fac ='''+docdm.myfac+'''');
  adoquery1.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
  adoquery1.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
  adoquery1.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
  adoquery1.Open;
  gbyear.Visible := True;
  if not adoquery1.Eof then begin
    //小於90代表需輸入月
    if adoquery1.FieldByName('dst08').AsInteger < 90 then begin
      gbmon.Visible := True;
      edmonth.Text := formatfloat('00',month);
      gbseason.Visible := false;
      ss := '00';
    end //90輸入季
    else if (adoquery1.FieldByName('dst08').AsInteger = 90) then begin
      gbseason.Visible := True;
      gbmon.Visible := False;
      ss := '01';
    end // 180是半年一次輸入資料
    else if (adoquery1.FieldByName('dst08').AsInteger = 180) then begin
      gbseason.Visible := false;
      gbmon.Visible := false;
      ss := '02';
    end   //999隨時可輸入資料
    else if (adoquery1.FieldByName('dst08').AsInteger = 999) then begin
      gbseason.Visible := false;
      gbmon.Visible := false;
      ss := '03';
    end;

  end;

  adoquery1.Close;
  adoquery1.sql.Clear;
  adoquery1.SQL.Add('Select substr(remark,1,3) From sysparameter Where isActive = ''Y'' And sysname = ''DOC'' And item = ''DST05'' and subject ='''+dataset.fieldbyname('dst05').AsString+'''');
  adoquery1.Open;
  adoquery1.First;
  if adoquery1.RecordCount > 0 then
    lab_filetype.Caption := dataset.fieldbyname('dst05').AsString+' '+adoquery1.Fields.Fields[0].AsString
  else
    lab_filetype.Caption := '';

end;

procedure TfmIns.mydocCalcFields(DataSet: TDataSet);
begin
  tmpque.Close;
  tmpque.SQL.Clear;
  tmpque.sql.Add('select * from doc_file where doc05 =:doc05 and doc06 =:doc06 and doc11 =:doc11 and doc_fac =:doc_fac order by createdate desc' );
  tmpque.Parameters.ParamByName('doc05').Value := dataset.fieldbyname('dst01').AsString;
  tmpque.Parameters.ParamByName('doc06').Value := dataset.fieldbyname('dst02').AsString;
  tmpque.Parameters.ParamByName('doc11').Value := dataset.fieldbyname('dst03').AsString;
  tmpque.Parameters.ParamByName('doc_fac').Value := docdm.myfac;
  tmpque.Open;
  if tmpque.RecordCount > 0 then begin
    if tmpque.FieldByName('doc02').AsString <> '' then
      dataset.FieldByName('lastdate').AsString := tmpque.FieldByName('doc02').AsString+'年';
    if tmpque.FieldByName('doc03').AsString <> '' then
      dataset.FieldByName('lastdate').AsString := dataset.FieldByName('lastdate').AsString+tmpque.FieldByName('doc03').AsString+'月';
    if tmpque.FieldByName('doc04').AsString <> '' then
      dataset.FieldByName('lastdate').AsString := dataset.FieldByName('lastdate').AsString+tmpque.FieldByName('doc04').AsString+'季';
  end
  else
    dataset.FieldByName('lastdate').AsString := '';

  tmpque.Close;
  tmpque.SQL.Clear;
  tmpque.SQL.Add('Select subject,remark From sysparameter Where isActive = ''Y'' And sysname =''DOC'' And item =''DST05'' and subject ='''+dataset.fieldbyname('dst05').asstring+'''');
  tmpque.Open;
  tmpque.First;
  if tmpque.RecordCount > 0 then
    dataset.FieldByName('dst05desc').AsString := tmpque.fieldbyname('remark').AsString
  else
    dataset.FieldByName('dst05desc').AsString := '';

  tmpque.Close;
  tmpque.SQL.Clear;
  tmpque.SQL.Add('Select subject,remark From sysparameter Where isActive = ''Y'' And sysname =''DOC'' And item =''DST06'' and subject ='''+dataset.fieldbyname('dst06').asstring+'''');
  tmpque.Open;
  tmpque.First;
  if tmpque.RecordCount > 0 then
    dataset.FieldByName('dst06desc').AsString := tmpque.fieldbyname('remark').AsString
  else
    dataset.FieldByName('dst06desc').AsString := '';


end;

procedure TfmIns.btn_insClick(Sender: TObject);
begin
  status := 'INS';
  lbno.Caption := '';
  lbfiles.Caption := '';
  lbfiles1.Caption := '';
  edfile.Clear;
  memo2.Clear;
end;

procedure TfmIns.btn_view1Click(Sender: TObject);
begin
  if trim(edfile.Text) <> '' then
    ShellExecute(0, PChar('open'), PChar(trim(edfile.text)), nil , nil, SW_SHOWNA);
end;

procedure TfmIns.btn_viewpdfClick(Sender: TObject);
begin
  if lbfiles1.Caption <> '' then begin
    tmpque.Close;
    tmpque.sql.Clear;
    tmpque.SQL.Add('select dst06 from doc_setup where dst01 =:dst01 and dst02 =:dst02 and dst03 =:dst03');
    tmpque.Parameters.ParamByName('dst01').Value := mydoc.fieldbyname('dst01').AsString;
    tmpque.Parameters.ParamByName('dst02').Value := mydoc.fieldbyname('dst02').AsString;
    tmpque.Parameters.ParamByName('dst03').Value := mydoc.fieldbyname('dst03').AsString;
    tmpque.Open;
    if (tmpque.FieldByName('dst06').AsString = 'W') or (tmpque.FieldByName('dst06').AsString = 'A') then begin
      tmpque.Close;
      tmpque.sql.Clear;
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCWEBFTP'' and isactive =''Y''');
      tmpque.Open;
      if tmpque.RecordCount > 0 then
        ShellExecute(0, PChar('open'), PChar('http://'+tmpque.FieldByName('subject').AsString+'/doc/'+lbfiles1.Caption), nil , nil, SW_SHOWNA);
    end
    else if (tmpque.FieldByName('dst06').AsString = 'D') then begin
      tmpque.Close;
      tmpque.sql.Clear;
      tmpque.sql.Add('select subject from sysparameter where sysname=''DOCFTP'' and isactive =''Y''');
      tmpque.Open;
      if tmpque.RecordCount > 0 then
        ShellExecute(0, PChar('open'), PChar('http://'+tmpque.fieldbyname('subject').AsString+'/doc/'+docdm.myfac+'/'+edyear.Text+'/'+copy(mydoc.fieldbyname('dst01').AsString,1,1)+'/'+lbfiles1.Caption), nil , nil, SW_SHOWNA);
    end
    else
      showmessage('未有文件可瀏覽!');
  end
  else
    showmessage('尚未上載檔案');

end;

procedure TfmIns.btn_etopClick(Sender: TObject);
var
  myext : string;
begin
  if opendialog1.Execute then begin
    if uppercase(trim(myext)) <> '.'+'PDF' then begin
      showmessage('您輸入的檔案類型錯誤,檔案類型必需為PDF');
    end;
  end;
end;

procedure TfmIns.Button2Click(Sender: TObject);
var
  myext : string;
begin
  if opendialog2.Execute then begin
    edfile1.Text := opendialog2.Files.Text;
    myext := ExtractFileExt(edfile1.Text);
    if trim(myext) <> '.ogv' then begin
      showmessage('您輸入的檔案類型錯誤,檔案類型必需為:ogv');
      edfile1.Clear;
    end;
  end;


end;

procedure TfmIns.Button4Click(Sender: TObject);
begin
  if trim(edfile1.Text) <> '' then
    ShellExecute(0, PChar('open'), PChar(trim(edfile1.text)), nil , nil, SW_SHOWNA);
end;

procedure TfmIns.Edit1Exit(Sender: TObject);
var query1 : TAdoquery;
begin
  query1 := TAdoquery.Create(self);
  query1.Connection := dm.glm;
  query1.Close;
  query1.SQL.Clear;
  query1.SQL.Add('select gen01 from gen_file where gen02=:gen02');
  query1.Parameters.ParamByName('gen02').Value:=trim(edit1.Text);
  query1.Open;
  if query1.Eof then begin
    showmessage('無此講師，煩請確認');
    exit;
  end;
end;

end.
