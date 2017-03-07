unit VBACHButton;

interface

uses
  ActiveX, OleServer, Office2010;

type
  TCommandBarButtonClick = procedure(const Ctrl: CommandBarButton;
    var CancelDefault: WordBool) of object;

  TCommandBarButton = class(TOleServer)
  private
    FIntf: _CommandBarButton;
    FOnClick: TCommandBarButtonClick;
    function GetDefaultInterface: _CommandBarButton;
  protected
    procedure InitServerData; override;
    procedure InvokeEvent(DispID: TDispID; var Params: TVariantArray); override;
  public
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _CommandBarButton);
    procedure Disconnect; override;
    property DefaultInterface: _CommandBarButton read GetDefaultInterface;
  published
    property OnClick: TCommandBarButtonClick read FOnClick write FOnClick;
  end;

implementation

{ TCommandBarButton }

procedure TCommandBarButton.Connect;
var pUnk: IUnknown;
begin
  if FIntf = nil then
  begin
    pUnk := GetServer;
    ConnectEvents(pUnk);
    FIntf := pUnk as _CommandBarButton;
  end;
end;

procedure TCommandBarButton.ConnectTo(svrIntf: _CommandBarButton);
begin
  Disconnect;
  FIntf := svrIntf;
  ConnectEvents(FIntf);
end;

procedure TCommandBarButton.Disconnect;
begin
  if Fintf <> nil then
  begin
    DisconnectEvents(FIntf);
    FIntf:= nil;
  end;
end;

function TCommandBarButton.GetDefaultInterface: _CommandBarButton;
begin
  if FIntf = nil then
    Connect;
  Result := FIntf;
end;

procedure TCommandBarButton.InitServerData;
const
  CServerData: TServerData = (
    ClassID:    '{55F88891-7708-11D1-ACEB-006008961DA5}'; // CLASS_CommandBarButton
    IntfIID:    '{000C030E-0000-0000-C000-000000000046}'; // IID__CommandBarButton;
    EventIID:   '{000C0351-0000-0000-C000-000000000046}'; // DIID__CommandBarButtonEvents;
    LicenseKey: nil;
    Version:    500);
begin
  ServerData:= @CServerData;
end;

procedure TCommandBarButton.InvokeEvent(DispID: TDispID;
  var Params: TVariantArray);
begin
  case DispID of
    1: if Assigned(FOnClick) then
         FOnClick(IUnknown(TVarData(Params[0]).VPointer) as _CommandBarButton,
           WordBool((TVarData(Params[1]).VPointer)^));
  end;
end;

end.
