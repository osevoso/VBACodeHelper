unit VBACHConst;

interface

const
  sCHBT_Indent           = 'CHBT_Indent';
  { CodeHelper button tags }
  sCHBT_IndentModule     = 'CHBT_IndentModule';
  sCHBT_IndentSubProgram = 'CHBT_IndentSubProgram';
  sCHBT_IndentLines      = 'CHBT_IndentLines';

resourcestring
{$IFDEF WIN32}
  rs_RegBaseKey          = 'Software\Microsoft\VBA\VBE\6.0\Addins\';
{$ENDIF}
{$IFDEF WIN64}
  rs_RegBaseKey          = 'Software\Microsoft\VBA\VBE\6.0\Addins64\';
{$ENDIF}
  rs_RegLoadBehavior     = 'LoadBehavior';
  rs_RegCommandLineSafe  = 'CommandLineSafe';
  rs_RegFriendlyName     = 'FriendlyName';
  rs_RegDescription      = 'Description';

  rs_RegAddinFriendlyName = 'VBA CodeHelper Add-In';
  rs_RegAddinDescription  = 'CodeHelper Add-In';

resourcestring
  rsErr_CanNotRegisterAddin     = 'Can''t register Add-In %s';
  rsErr_CanNotDetermineSubProg  = 'Can''t determine subprogram';

implementation

end.
