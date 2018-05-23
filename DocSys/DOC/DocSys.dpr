program DocSys;

uses
  Forms,
  docins in 'docins.pas' {fmIns},
  docque in 'docque.pas' {fmQuery},
  docdm in 'docdm.pas' {dm: TDataModule},
  docbrowse in 'docbrowse.pas' {fmBrowse},
  docmaintain in 'docmaintain.pas' {fmMaintain},
  docmantain1 in 'docmantain1.pas' {fmbigtype},
  docsign in 'docsign.pas' {fmsign},
  uDoc_Type in 'uDoc_Type.pas' {fmdoc_type},
  uDocMain in 'uDocMain.pas' {fmDocMain},
  uDoc_Setup in 'uDoc_Setup.pas' {fmDoc_setup},
  uDocr_file in 'uDocr_file.pas' {fmDocr_file},
  uCheDoc in 'uCheDoc.pas' {fmCheDep},
  uCp in 'uCp.pas' {fmcp},
  mail in 'mail.pas' {fmMail};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(Tdm, dm);
  Application.CreateForm(TfmDocMain, fmDocMain);
  Application.CreateForm(TfmMail, fmMail);
  Application.Run;
end.
