unit docmantain1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, ExtCtrls, DBCtrls;

type
  Tfmbigtype = class(TForm)
    DBNavigator1: TDBNavigator;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    ADOQuery1: TADOQuery;
    ADOQuery1BIG01: TWideStringField;
    ADOQuery1BIG02: TWideStringField;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmbigtype: Tfmbigtype;

implementation
uses docdm;
{$R *.dfm}

end.
