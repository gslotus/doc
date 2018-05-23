unit docsign;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, DB, ADODB, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdFTP;

type
  Tfmsign = class(TForm)
    Button1: TButton;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    queMain: TADOQuery;
    queMainDOC01: TWideStringField;
    queMainDOC09: TWideStringField;
    queMaintypedesc: TStringField;
    queMainchilddesc: TStringField;
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
    ADOQuery1: TADOQuery;
    IdFTP1: TIdFTP;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure queMainCalcFields(DataSet: TDataSet);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure Refresh;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmsign: Tfmsign;
  status : string;
implementation

uses docdm, docbrowse;

{$R *.dfm}

procedure Tfmsign.Refresh;
begin
  if status ='MANAGER' THEN begin
    quemain.close;
    quemain.SQL.Clear;
    quemain.sql.add('select * from doc_file where docsign =''0'' order by doc02,doc03,doc04,doc05,doc06 ');
    quemain.Open;
    quemain.First;
    if quemain.RecordCount = 0 then
      showmessage('無資料簽核');
  end
  else begin
    quemain.close;
    quemain.SQL.Clear;
    quemain.sql.add('select * from doc_file where docsign =''1'' order by doc02,doc03,doc04,doc05,doc06 ');
    quemain.Open;
    quemain.First;
    if quemain.RecordCount = 0 then
      showmessage('無資料簽核');
  end;
end;

procedure Tfmsign.FormShow(Sender: TObject);
begin
  Refresh;
end;

procedure Tfmsign.Button1Click(Sender: TObject);
begin
  if status ='MANAGER' THEN begin     //經理簽核
    try
      if not dm.gshank.InTransaction then
        dm.gshank.BeginTrans;
      adoquery1.Close;
      adoquery1.sql.Clear;
      adoquery1.SQL.Add('update doc_file set docsign =''2'' where doc01 =:doc01');
      adoquery1.Parameters.ParamByName('doc01').Value := quemain.fieldbyname('doc01').AsString ;
      adoquery1.ExecSQL;
      if dm.gshank.InTransaction then
        dm.gshank.CommitTrans;
      showmessage('放行成功!');
    except
      if dm.gshank.InTransaction then
        dm.gshank.RollbackTrans;
      showmessage('放行失敗!');
    end;
  end;
  //因處長休養暫時取消處長簽核
  {
  else begin                               //處長簽核
    try
      if not dm.gshank.InTransaction then
        dm.gshank.BeginTrans;
      adoquery1.Close;
      adoquery1.sql.Clear;
      adoquery1.SQL.Add('update doc_file set docsign =''2'' where doc01 =:doc01');
      adoquery1.Parameters.ParamByName('doc01').Value := quemain.fieldbyname('doc01').AsString ;
      adoquery1.ExecSQL;
      if not idftp1.Connected then
        idftp1.Connect(True);
      idftp1.Put('T:\FinXml\Web\'+quemain.fieldbyname('doc07').asstring,quemain.fieldbyname('doc07').asstring,False);
      if dm.gshank.InTransaction then
        dm.gshank.CommitTrans;
      showmessage('放行成功!');
    except
      if dm.gshank.InTransaction then
        dm.gshank.RollbackTrans;
      showmessage('放行失敗!');
    end;
  end;
  }
  Refresh
end;

procedure Tfmsign.queMainCalcFields(DataSet: TDataSet);
begin
  adoquery1.Close;
  adoquery1.SQL.Clear;
  adoquery1.sql.Add('select big02 from doc_big_type where big01 =:big01');
  adoquery1.Parameters.ParamByName('big01').Value := dataset.fieldbyname('doc05').AsString;
  adoquery1.Open;
  adoquery1.First;
  if not adoquery1.Eof then
    DataSet.FieldByName('typedesc').AsString := adoquery1.Fields[0].AsString;

  adoquery1.Close;
  adoquery1.SQL.Clear;
  adoquery1.sql.Add('select dtype03 from doc_type where dtype01 =:dtype01 and dtype02 =:dtype02');
  adoquery1.Parameters.ParamByName('dtype01').Value := dataset.fieldbyname('doc05').AsString;
  adoquery1.Parameters.ParamByName('dtype02').Value := dataset.fieldbyname('doc06').AsString;
  adoquery1.Open;
  adoquery1.First;
  if not adoquery1.Eof then
    DataSet.FieldByName('childdesc').AsString := adoquery1.Fields[0].AsString;


end;

procedure Tfmsign.DBGrid1DblClick(Sender: TObject);
begin
  //fmBrowse.WebBrowser1.Navigate('T:\Finxml\web\'+quemain.fieldbyname('doc07').AsString);
  //fmBrowse.Show;
  fmBrowse := TfmBrowse.create(Self);
  fmBrowse.WebBrowser1.Navigate('T:\Finxml\web\'+quemain.fieldbyname('doc07').AsString);
  fmBrowse.ShowModal;
  fmBrowse.free;
end;

end.
