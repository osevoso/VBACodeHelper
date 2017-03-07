unit VBACH_TLB;

{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers.
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL,
     Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:
//   Type Libraries     : LIBID_xxxx
//   CoClasses          : CLASS_xxxx
//   DISPInterfaces     : DIID_xxxx
//   Non-DISP interfaces: IID_xxxx
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  VBACHMajorVersion = 1;
  VBACHMinorVersion = 0;

  LIBID_VBACH: TGUID = '{DC3D6BBF-5978-468C-ABF9-67AC5B0CFD41}';

  IID_IVBACodeHelper: TGUID = '{0398855D-F151-464E-8495-79AA4530C983}';
  CLASS_VBACodeHelper: TGUID = '{A305C2EE-3B28-497D-820D-E5A4571FE52C}';
  IID_IDTExtensibility2: TGUID = '{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library
// *********************************************************************//
// Constants for enum ext_ConnectMode
type
  ext_ConnectMode = TOleEnum;
const
  ext_cm_AfterStartup = $00000000;
  ext_cm_Startup = $00000001;
  ext_cm_External = $00000002;
  ext_cm_CommandLine = $00000003;

// Constants for enum ext_DisconnectMode
type
  ext_DisconnectMode = TOleEnum;
const
  ext_dm_HostShutdown = $00000000;
  ext_dm_UserClosed = $00000001;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//
  IVBACodeHelper = interface;
  IVBACodeHelperDisp = dispinterface;
  IDTExtensibility2 = interface;
  IDTExtensibility2Disp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library
// (NOTE: Here we map each CoClass to its Default Interface)
// *********************************************************************//
  VBACodeHelper = IVBACodeHelper;


// *********************************************************************//
// Interface: IVBACodeHelper
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0398855D-F151-464E-8495-79AA4530C983}
// *********************************************************************//
  IVBACodeHelper = interface(IDispatch)
    ['{0398855D-F151-464E-8495-79AA4530C983}']
  end;

// *********************************************************************//
// DispIntf:  IVBACodeHelperDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {0398855D-F151-464E-8495-79AA4530C983}
// *********************************************************************//
  IVBACodeHelperDisp = dispinterface
    ['{0398855D-F151-464E-8495-79AA4530C983}']
  end;

// *********************************************************************//
// Interface: IDTExtensibility2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B65AD801-ABAF-11D0-BB8B-00A0C90F2744}
// *********************************************************************//
  IDTExtensibility2 = interface(IDispatch)
    ['{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}']
    procedure OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode;
                           const AddInInst: IDispatch; var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
  end;

// *********************************************************************//
// DispIntf:  IDTExtensibility2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B65AD801-ABAF-11D0-BB8B-00A0C90F2744}
// *********************************************************************//
  IDTExtensibility2Disp = dispinterface
    ['{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}']
    procedure OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode;
                           const AddInInst: IDispatch;
                           var custom: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 201;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode;
                              var custom: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 202;
    procedure OnAddInsUpdate(var custom: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 203;
    procedure OnStartupComplete(var custom: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 204;
    procedure OnBeginShutdown(var custom: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 205;
  end;

// *********************************************************************//
// The Class CoVBACodeHelper provides a Create and CreateRemote method to
// create instances of the default interface IVBACodeHelper exposed by
// the CoClass VBACodeHelper. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoVBACodeHelper = class
    class function Create: IVBACodeHelper;
    class function CreateRemote(const MachineName: string): IVBACodeHelper;
  end;

implementation

uses System.Win.ComObj;

class function CoVBACodeHelper.Create: IVBACodeHelper;
begin
  Result := CreateComObject(CLASS_VBACodeHelper) as IVBACodeHelper;
end;

class function CoVBACodeHelper.CreateRemote(const MachineName: string): IVBACodeHelper;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_VBACodeHelper) as IVBACodeHelper;
end;

end.

