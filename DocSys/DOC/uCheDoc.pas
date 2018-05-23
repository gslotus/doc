unit uCheDoc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, DB, ADODB, StdCtrls, ComCtrls, ExtCtrls, DateUtils;

type
  TfmCheDep = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    dtp_date1: TDateTimePicker;
    dtp_date2: TDateTimePicker;
    Button1: TButton;
    ds_new: TDataSource;
    ado_new: TADOQuery;
    ado_newCPF02: TWideStringField;
    ado_newCQA03: TWideStringField;
    ado_newCQA04: TWideStringField;
    ado_newBDEP: TStringField;
    ado_newCQA05: TWideStringField;
    ado_newADEP: TStringField;
    dep_new: TDBGrid;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmCheDep: TfmCheDep;

implementation
uses docdm;
{$R *.dfm}

procedure TfmCheDep.Button1Click(Sender: TObject);
begin
  ado_new.Close;   //部門員工異動
  ado_new.sql.Clear;
  ado_new.SQL.Add('select cpf02,CQA03,CQA04, ');
  ado_new.SQL.Add(' (select gem03 from gem_file'+docdm.dblinks+' where gem01=cqa04) as bdep, ');
  ado_new.SQL.Add(' CQA05,(select gem03 from gem_file'+docdm.dblinks+' where gem01=cqa05) as adep ');
  ado_new.SQL.Add(' from cqa_file'+docdm.dblinks+',cpf_file'+docdm.dblinks+',cqaa_file'+docdm.dblinks+' where cqa01 = cqaa01 and cpf01 = cqa03  ');
  ado_new.SQL.Add(' and cqaa03 between :Date1 and :Date2 and cqa06 =''2'' ');
  ado_new.SQL.Add(' and cpf35 is null and cpf36<>''Y'' and cpf37<>''Y'' order by cqa04,cqa09');
  ado_new.Parameters.ParamByName('date1').Value := dateof(dtp_date1.Date);
  ado_new.Parameters.ParamByName('date2').Value := dateof(dtp_date2.Date);
  ado_new.Open;
  ado_new.First;
end;

procedure TfmCheDep.FormShow(Sender: TObject);
begin
  dtp_date1.Date := now-7;
  dtp_date2.Date := now;
end;

end.
