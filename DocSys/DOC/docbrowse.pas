unit docbrowse;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, SHDocVw;

type
  TfmBrowse = class(TForm)
    WebBrowser1: TWebBrowser;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmBrowse: TfmBrowse;

implementation

{$R *.dfm}

end.
