library VBACodeHelper;

uses
  ComServ,
  VBACH_TLB in 'VBACH_TLB.pas',
  VBACHAddin in 'VBACHAddin.pas' {VBAHelperAddin: CoClass},
  VBACHConst in 'VBACHConst.pas',
  VBACHButton in 'VBACHButton.pas',
  VBACHIndent in 'VBACHIndent.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  DllInstall;

{$R *.TLB}

{$R *.RES}

begin
end.
