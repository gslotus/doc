unit docdm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DB, ADODB, StdCtrls, DBCtrls, Mask, Grids, DBGrids,
  ComCtrls, Buttons, CheckLst, DateUtils, mATH;

type
  Tdm = class(TDataModule)
    gshank: TADOConnection;
    topdb: TADOConnection;
    dbque: TADOQuery;
    glm: TADOConnection;
  private
    { Private declarations }
  public
    function getlogin:string;
    { Public declarations }
  end;

var
  dm: Tdm;
  dblinks :string;
  myfac : string;
  mygem : string;
implementation

{$R *.dfm}

function Tdm.getlogin:string;
var
  lpBuffer: PChar;
  nSize: DWord;
begin
  {¿À¨d≈v≠≠}
  nSize := 255;
  GetMem(lpBuffer, nSize);
  GetUserName(lpBuffer, nSize);
  result := lpBuffer;
  //result := 'winy';
end;

end.
