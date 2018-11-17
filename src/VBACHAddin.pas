unit VBACHAddin;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Contnrs, Variants,
  ComObj, ActiveX, OleServer, IniFiles,
  Office2010, VBACH_TLB;

type
  TVBACHFactory = class(TAutoObjectFactory)
    procedure UpdateRegistry(Register: Boolean); override;
  end;

  TVBACHAddin = class(TAutoObject, IVBACodeHelper, IDTExtensibility2)
  private
    FVBEApp: IDispatch;
    FBList: TObjectList;
    FVBEWindowHandle: HWND;
    FHotKeyWindowHandle: HWND;
    FHotKeyCommentLine: Cardinal;
    FHotKeyUnCommentLine: Cardinal;
    FHotKeyToggleBookmark: Cardinal;
    FHotKeyNextBookmark: Cardinal;
    FHotKeyPrevBookmark: Cardinal;
    procedure InitButtons;
    procedure DestroyButtons;
    procedure WndProc(var Msg: TMessage);
    procedure BtnClick(const Ctrl: CommandBarButton; var CancelDefault: WordBool);
  protected
    { IDTExtensibility2 }
    procedure OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode;
                           const AddInInst: IDispatch; var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
  public
    procedure Initialize; override;
    destructor Destroy; override;
    procedure InitHotKeys;
    procedure DestroyHotKeys;
  end;

implementation

uses
  ComServ, Registry, VBIDE2010,
  VBACHConst, VBACHButton, VBACHIndent;

{ TVBACHFactory }

procedure TVBACHFactory.UpdateRegistry(Register: Boolean);
var RootKey: HKEY;
    AddInKey: String;
    Reg: TRegistry;
begin
  { Add and remove registration VBE Add-in }
  Rootkey := HKEY_CURRENT_USER;
  AddInKey := rs_RegBaseKey + ProgID;
  Reg := TRegistry.Create;
  Reg.RootKey := RootKey;
  try
    if Register then
    begin
      if Reg.OpenKey(AddInKey, True) then
      begin
        Reg.WriteInteger(rs_RegLoadBehavior, 3);
        Reg.WriteInteger(rs_RegCommandLineSafe, 0);
        Reg.WriteString(rs_RegFriendlyName, rs_RegAddinFriendlyName);
        Reg.WriteString(rs_RegDescription, rs_RegAddinDescription);
        Reg.CloseKey;
      end
      else
        raise EOleError.CreateFmt(rsErr_CanNotRegisterAddin, [ProgID]);
    end
    else
      if Reg.KeyExists(AddInKey) then
        Reg.DeleteKey(AddInKey);
  finally
    FreeAndNil(Reg);
  end;
  inherited;
end;

{ TVBACHAddin }

procedure TVBACHAddin.Initialize;
begin
  inherited;
  FBList := TObjectList.Create;
  FHotKeyWindowHandle := 0;
end;

destructor TVBACHAddin.Destroy;
begin
  FreeAndNil(FBList);
  inherited;
end;

procedure TVBACHAddin.InitHotKeys;
var AtomText: array[0..31] of Char;
begin
  if Assigned(FVBEApp) then
    if (FHotKeyWindowHandle = 0) then
    begin
      FVBEWindowHandle := (FVBEApp as VBE).MainWindow.HWnd;
      FHotKeyWindowHandle := AllocateHWnd(WndProc);
      { Comment lines }
      FHotKeyCommentLine := GlobalAddAtom(StrFmt(AtomText, 'CommentHotVBE%.8X%.8X', [HInstance, GetCurrentThreadID]));
      RegisterHotKey(FHotKeyWindowHandle, FHotKeyCommentLine, MOD_ALT + MOD_CONTROL, VK_ADD);
      { UnComment lines }
      FHotKeyUnCommentLine := GlobalAddAtom(StrFmt(AtomText, 'UnCommentHotVBE%.8X%.8X', [HInstance, GetCurrentThreadID]));
      RegisterHotKey(FHotKeyWindowHandle, FHotKeyUnCommentLine, MOD_ALT + MOD_CONTROL, VK_SUBTRACT);
      { Toggle bookmark }
      FHotKeyToggleBookmark := GlobalAddAtom(StrFmt(AtomText, 'TgBmarkHotVBE%.8X%.8X', [HInstance, GetCurrentThreadID]));
      RegisterHotKey(FHotKeyWindowHandle, FHotKeyToggleBookmark, MOD_ALT + MOD_CONTROL, VK_MULTIPLY);
      { Next bookmark }
      FHotKeyNextBookmark := GlobalAddAtom(StrFmt(AtomText, 'NxBkmarkVBE%.8X%.8X', [HInstance, GetCurrentThreadID]));
      RegisterHotKey(FHotKeyWindowHandle, FHotKeyNextBookmark, MOD_CONTROL, VK_OEM_3);
      { Previous bookmark }
      FHotKeyPrevBookmark := GlobalAddAtom(StrFmt(AtomText, 'PvBkmarkVBE%.8X%.8X', [HInstance, GetCurrentThreadID]));
      RegisterHotKey(FHotKeyWindowHandle, FHotKeyPrevBookmark, MOD_SHIFT + MOD_CONTROL, VK_OEM_3);
    end;
end;

procedure TVBACHAddin.InitButtons;
var i: Integer;
    cmBar: CommandBar;
    cmBars: CommandBars;
    cmPopup: CommandBarPopup;
    ToggleIndex: Integer;
    Button: TCommandBarButton;
begin
  if Assigned(FVBEApp) then
  begin
    cmBars := (FVBEApp as VBE).CommandBars;
    cmBar := cmBars.Item['Code Window'];

    { Check indent items }
    for i := 1 to cmBar.Controls.Count do
      if cmBar.Controls.Item[i].Tag = sCHBT_Indent then
        Exit;

    ToggleIndex := 1;
    { Get index "Toggle" item }
    for i := 1 to cmBar.Controls.Count do
      if cmBar.Controls.Item[i].Id = 32805 then
      begin
        ToggleIndex := cmBar.Controls.Item[i].Index;
        Break;
      end;

   { Add "Code Indent" popup before "Toggle" }
    cmPopup := cmBar.Controls.Add(msoControlPopup,
      Variants.EmptyParam, Variants.EmptyParam,
      ToggleIndex, True) as CommandBarPopup;
    cmPopup.Set_Tag(sCHBT_Indent);
    cmPopup.Set_BeginGroup(True);
    cmPopup.Set_Caption('Code Indent');
    cmPopup.Set_Visible(True);
    cmPopup.Set_TooltipText('Code Indent');

    { Adding Indent buttons }
    Button := TCommandBarButton.Create(nil);
    Button.OnClick := Self.BtnClick;
    FBList.Add(Button);
    Button.ConnectTo(cmPopup.CommandBar.Controls.Add(msoControlButton,
      Variants.EmptyParam, Variants.EmptyParam, Variants.EmptyParam,
      True) as CommandBarButton);
    with Button.DefaultInterface do
    begin
      Set_Style(msoButtonIconAndCaption);
      Set_BeginGroup(False);
      Set_Width(80);
      Set_Caption('Entire Module');
      Set_FaceId(15);
      Set_DescriptionText('Indent all lines in module');
      Set_Enabled(True);
      Set_Tag(sCHBT_IndentModule);
      Set_Visible(True);
    end;

    Button := TCommandBarButton.Create(nil);
    Button.OnClick := Self.BtnClick;
    FBList.Add(Button);
    Button.ConnectTo(cmPopup.CommandBar.Controls.Add(msoControlButton,
      Variants.EmptyParam, Variants.EmptyParam, Variants.EmptyParam,
      True) as CommandBarButton);
    with Button.DefaultInterface do
    begin
      Set_Style(msoButtonIconAndCaption);
      Set_BeginGroup(False);
      Set_Width(80);
      Set_Caption('SubProgram');
      Set_FaceId(15);
      Set_DescriptionText('Indent current subprogram lines');
      Set_Enabled(True);
      Set_Tag(sCHBT_IndentSubProgram);
      Set_Visible(True);
    end;

    Button := TCommandBarButton.Create(nil);
    Button.OnClick := Self.BtnClick;
    FBList.Add(Button);
    Button.ConnectTo(cmPopup.CommandBar.Controls.Add(msoControlButton,
      Variants.EmptyParam, Variants.EmptyParam, Variants.EmptyParam,
      True) as CommandBarButton);
    with Button.DefaultInterface do
    begin
      Set_Style(msoButtonIconAndCaption);
      Set_BeginGroup(False);
      Set_Width(80);
      Set_Caption('Selected Lines');
      Set_FaceId(15);
      Set_DescriptionText('Indent selected lines');
      Set_Enabled(True);
      Set_Tag(sCHBT_IndentLines);
      Set_Visible(True);
    end;
  end;
end;

procedure TVBACHAddin.DestroyButtons;
var i, j: Integer;
    cmBar: CommandBar;
    cmBars: CommandBars;
    cmPopup: CommandBarPopup;
begin
  FBList.Clear;
  if Assigned(FVBEApp) then
  begin
    cmBars := (FVBEApp as VBE).CommandBars;
    cmBar := cmBars.Item['Code Window'];
    for i := cmBar.Controls.Count downto 1 do
    begin
      if cmBar.Controls.Item[i].Tag = sCHBT_Indent then
      begin
        cmPopup := cmBar.Controls.Item[i] as CommandBarPopup;
        for j := cmPopup.Controls.Count downto 1 do
          cmPopup.Controls.Item[j].Delete(False);
        cmPopup.Delete(False);
        Break;
      end;
    end;
  end;
end;

procedure TVBACHAddin.DestroyHotKeys;
begin
  if FHotKeyWindowHandle <> 0 then
  begin
    UnRegisterHotKey(FHotKeyWindowHandle, FHotKeyCommentLine);
    GlobalDeleteAtom(FHotKeyCommentLine);
    UnRegisterHotKey(FHotKeyWindowHandle, FHotKeyUnCommentLine);
    GlobalDeleteAtom(FHotKeyUnCommentLine);
    UnRegisterHotKey(FHotKeyWindowHandle, FHotKeyToggleBookmark);
    GlobalDeleteAtom(FHotKeyToggleBookmark);
    UnRegisterHotKey(FHotKeyWindowHandle, FHotKeyNextBookmark);
    GlobalDeleteAtom(FHotKeyNextBookmark);
    UnRegisterHotKey(FHotKeyWindowHandle, FHotKeyPrevBookmark);
    GlobalDeleteAtom(FHotKeyPrevBookmark);
    DeallocateHWnd(FHotKeyWindowHandle);
    FHotKeyWindowHandle := 0;
  end;
end;

procedure TVBACHAddin.BtnClick(const Ctrl: CommandBarButton;
  var CancelDefault: WordBool);
var CM: CodeModule;
    CP: CodePane;
begin
  if (Ctrl.Tag = sCHBT_IndentModule) or
    (Ctrl.Tag = sCHBT_IndentSubProgram) or
    (Ctrl.Tag = sCHBT_IndentLines) then
  begin
    CP := (FVBEApp as VBE).ActiveCodePane;
    if Assigned(CP) then
    begin
      CM := CP.CodeModule;
      if Assigned(CM) then
      begin
        if (Ctrl.Tag = sCHBT_IndentModule) then
          IndentCodeModule(CM)
        else if (Ctrl.Tag = sCHBT_IndentSubProgram) then
          IndentSubrogramm(CM)
        else if (Ctrl.Tag = sCHBT_IndentLines) then
          IndentSelectedLines(CM);
      end;
    end;
  end;
end;

procedure TVBACHAddin.WndProc(var Msg: TMessage);
var fgWnd: HWND;
    cmBars: CommandBars;
begin
  if Msg.Msg = WM_HOTKEY then
  begin
    fgWnd := GetForegroundWindow;
    if (fgWnd = FVBEWindowHandle) and
       ((Msg.WParam = FHotKeyCommentLine) or
        (Msg.WParam = FHotKeyUnCommentLine) or
        (Msg.WParam = FHotKeyToggleBookmark) or
        (Msg.WParam = FHotKeyNextBookmark) or
        (Msg.WParam = FHotKeyPrevBookmark)) then
    begin
      if Assigned(FVBEApp) then
      begin
        cmBars := (FVBEApp as VBE).CommandBars;
        if Msg.WParam = FHotKeyCommentLine then
          cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_CommentLines].Execute
        else if Msg.WParam = FHotKeyUnCommentLine then
          cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_UnCommentLines].Execute
        else if Msg.WParam = FHotKeyToggleBookmark then
        begin
          if cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_ToggleBookmark].Enabled then
            cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_ToggleBookmark].Execute;
        end
        else if Msg.WParam = FHotKeyNextBookmark then
        begin
          if cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_NextBookmark].Enabled then
            cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_NextBookmark].Execute
        end
        else if Msg.WParam = FHotKeyPrevBookmark then
        begin
          if cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_PreviousBookmark].Enabled then
            cmBars.Item[sVBE_EditBar].Controls.Item[sVBE_PreviousBookmark].Execute;
        end;
      end;
    end;
  end
  else
    Msg.Result := DefWindowProc(FHotKeyWindowHandle, Msg.Msg, Msg.wParam, Msg.lParam);
end;

{$REGION 'IDTExtensibility2 Impl'}
procedure TVBACHAddin.OnConnection(const Application: IDispatch;
  ConnectMode: ext_ConnectMode; const AddInInst: IDispatch;
  var custom: PSafeArray);
begin
  FVBEApp := Application;
  if ConnectMode = ext_cm_AfterStartup then
  begin
    InitHotKeys;
    InitButtons;
  end;
end;

procedure TVBACHAddin.OnDisconnection(RemoveMode: ext_DisconnectMode;
  var custom: PSafeArray);
begin
  DestroyHotKeys;
  DestroyButtons;
  FVBEApp := nil;
end;

procedure TVBACHAddin.OnStartupComplete(var custom: PSafeArray);
begin
  InitHotKeys;
  InitButtons;
end;

procedure TVBACHAddin.OnAddInsUpdate(var custom: PSafeArray);
begin
end;

procedure TVBACHAddin.OnBeginShutdown(var custom: PSafeArray);
begin
end;
{$ENDREGION}

initialization
  TVBACHFactory.Create(ComServer, TVBACHAddin, Class_VBACodeHelper,
    ciMultiInstance, tmApartment);

end.
