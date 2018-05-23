unit docmaintain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ExtCtrls, DBCtrls, Grids, DBGrids, StdCtrls;

type
  TfmMaintain = class(TForm)
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    ADOQuery1: TADOQuery;
    ADOQuery1DTYPE01: TWideStringField;
    ADOQuery1DTYPE02: TWideStringField;
    ADOQuery1DTYPE03: TWideStringField;
    ADOQuery1DTYPE04: TWideStringField;
    ADOQuery1DTYPE05: TWideStringField;
    ADOQuery1DTYPE06: TWideStringField;
    Label2: TLabel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmMaintain: TfmMaintain;

implementation
uses docdm;
{$R *.dfm}

end.
