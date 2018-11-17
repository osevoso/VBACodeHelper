{ **********************************************************************************
  Originals algorithm author: Michael Ciurescu (CVMichael)
  http://www.vbforums.com/showthread.php?479449-VBTools-AddIn-Auto-indent-VB-code-!
  ********************************************************************************** }

unit VBACHIndent;

interface

uses
  Windows, SysUtils, StrUtils, Types, RegularExpressionsCore,
  VBIDE2010;

  procedure IndentCodeModule(ACodeModule: CodeModule);
  procedure IndentSubrogramm(ACodeModule: CodeModule);
  procedure IndentSelectedLines(ACodeModule: CodeModule);

implementation

uses
  VBACHConst;

const
  TABIndent: Char = #9;
  BlockingIndent: Char = #1;

var
  arrBlockStart: array [0..27] of string = ('^If\s{1,}.*\s{1,}Then((\s*)$|(\s*)''(.*)$|$)', '^#If\s{1,}.*\s{1,}Then.*',
    '^For\s{1,}.*', '^Do\s{1,}.*', '^Do', '^Select\s{1,}Case\s{1,}.*', '^While\s{1,}.*', '^With\s{1,}.*',
    '^Private\s{1,}Function\s{1,}.*', '^Public\s{1,}Function\s{1,}.*', '^Friend\s{1,}Function\s{1,}.*',
    '^Function\s{1,}.*', '^Private\s{1,}Sub\s{1,}.*', '^Public\s{1,}Sub\s{1,}.*', '^Friend\s{1,}Sub\s{1,}.*',
    '^Sub\s{1,}.*', '^Private\s{1,}Property\s{1,}.*', '^Public\s{1,}Property\s{1,}.*', '^Friend\s{1,}Property\s{1,}.*',
    '^Property\s{1,}.*', '^Private\s{1,}Enum\s{1,}.*', '^Public\s{1,}Enum\s{1,}.*', '^Friend\s{1,}Enum\s{1,}.*',
    '^Enum\s{1,}.*', '^Private\s{1,}Type\s{1,}.*', '^Public\s{1,}Type\s{1,}.*', '^Friend\s{1,}Type\s{1,}.*', '^Type\s{1,}.*');

  arrBlockEnd: array [0..27] of string = ('^End\s{1,}If', '^#End\s{1,}If', '^Next(\s{1,}.*)*', '^Loop',
    '^Loop\s{1,}.*', '^End\s{1,}Select', '^Wend', '^End\s{1,}With', '^End\s{1,}Function', '^End\s{1,}Function',
    '^End\s{1,}Function', '^End\s{1,}Function', '^End\s{1,}Sub', '^End\s{1,}Sub', '^End\s{1,}Sub',
    '^End\s{1,}Sub', '^End\s{1,}Property', '^End\s{1,}Property', '^End\s{1,}Property', '^End\s{1,}Property',
    '^End\s{1,}Enum', '^End\s{1,}Enum', '^End\s{1,}Enum', '^End\s{1,}Enum', '^End\s{1,}Type', '^End\s{1,}Type',
    '^End\s{1,}Type', '^End\s{1,}Type');

  arrBlockMiddle: array [0..4] of string = ('^ElseIf\s{1,}.*\s{1,}Then', '^#ElseIf\s{1,}.*\s{1,}Then', '^Else', '^#Else', '^Case\s{1,}.*');

  BlockStartList: TPerlRegExList;
  BlockEndList: TPerlRegExList;
  BlockMiddleList: TPerlRegExList;


procedure IndentInitialize;

  procedure FillBlockList(var AList: TPerlRegExList; const Arr: array of string);
  var i: Integer;
      ex: TPerlRegEx;
  begin
    AList := TPerlRegExList.Create;
    for i := Low(Arr) to High(Arr) do
    begin
      ex := TPerlRegEx.Create;
      ex.RegEx := Arr[i];
      ex.Options := [preCaseLess];
      AList.Add(ex);
    end;
  end;

begin
  FillBlockList(BlockStartList, arrBlockStart);
  FillBlockList(BlockEndList, arrBlockEnd);
  FillBlockList(BlockMiddleList, arrBlockMiddle);
end;

procedure IndentFinalize;

  procedure FreeBlockList(var AList: TPerlRegExList);
  var i: Integer;
      ex: TPerlRegEx;
  begin
    for i := AList.Count - 1 downto 0 do
    begin
      ex := AList.RegEx[i];
      AList.Delete(i);
      ex.Free;
    end;
    FreeAndNil(AList);
  end;

begin
  FreeBlockList(BlockStartList);
  FreeBlockList(BlockEndList);
  FreeBlockList(BlockMiddleList);
end;

function StrArrayJoin(const StringArray: array of string): string;
var i : Integer;
begin
  Result := '';
  for i := Low(StringArray) to High(StringArray) do
    Result := Result + StringArray[i] + sLineBreak;
  Result := Result.TrimRight;
end;

function GetCurrentVBELine(ACodeModule: CodeModule): Integer;
var CurrentVBELine, EndLine, StartColumn, EndColumn: Integer;
begin
  ACodeModule.CodePane.GetSelection(CurrentVBELine, StartColumn, EndLine, EndColumn);
  Result := CurrentVBELine;
end;

procedure IndentLineBlock(AVBLines: TArray<string>; StartLine: Integer; EndLine: Integer);
var i: Integer;
begin
  if StartLine > 0 then
    AVBLines[StartLine - 1] := BlockingIndent + AVBLines[StartLine - 1];
  if EndLine < High(AVBLines) then
    AVBLines[EndLine + 1] := BlockingIndent + AVBLines[EndLine + 1];
  for i := StartLine to EndLine do
    AVBLines[i] := TABIndent + AVBLines[i];
end;

function RemoveLineComments(AVBLine: string): string;
var i: Integer;
begin
  i := Pos('''', AVBLine, 1);
  if i > 0 then
    AVBLine := Copy(AVBLine, 1, i - 1);
  Result := AVBLine.TrimLeft([BlockingIndent]);
  Result := AVBLine.TrimRight([BlockingIndent]);
end;

procedure IndentBlock(AVBLines: TArray<string>);
var i: Integer;
    VBLine: String;
    FoundStartEnd: Boolean;
    StartPos, EndPos: Integer;
begin
  repeat
    StartPos := 0;
    EndPos := High(AVBLines);
    repeat
      FoundStartEnd := False;
      for i := EndPos downto StartPos do
      begin
        VBLine := RemoveLineComments(AVBLines[i]);
        if Length(VBLine) > 0 then
        begin
          BlockStartList.Subject := VBLine;
          if BlockStartList.Match then
          begin
            StartPos := i + 1;
            FoundStartEnd := True;
            Break;
          end;
        end;
      end;
      if FoundStartEnd then
      begin
        for i := StartPos to EndPos do
        begin
          VBLine := RemoveLineComments(AVBLines[i]);
          if Length(VBLine) > 0 then
          begin
            BlockEndList.Subject := VBLine;
            if BlockEndList.Match then
            begin
              EndPos := i - 1;
              FoundStartEnd := True;
              Break;
            end;
          end;
        end;
      end;
    until not FoundStartEnd;
    if not (not FoundStartEnd and (StartPos = 0) and (EndPos = High(AVBLines))) then
       IndentLineBlock(AVBLines, StartPos, EndPos);
  until not FoundStartEnd and (StartPos = 0) and (EndPos = High(AVBLines));
end;

procedure IndentCode(ACodeModule: CodeModule; AStartInsPos: Integer;
  ACountOrigLines: Integer; ATopLine: Integer; AAllCode: string);
var i: Integer;
    VBLine: String;
    AllLines: TArray<string>;
begin
  if AAllCode.Trim = EmptyStr then
    Exit;
  IndentInitialize;
  try
    AAllCode := StringReplace(AAllCode, ' _' + sLineBreak, ' ', [rfReplaceAll]);
    AllLines := AAllCode.Split([sLineBreak]);
    for i := 0 to High(AllLines) do
      AllLines[i] := Trim(AllLines[i]);
    IndentBlock(AllLines);
    for i := 0 to High(AllLines) do
    begin
      VBLine := RemoveLineComments(AllLines[i]);
      BlockMiddleList.Subject := VBLine.Replace(TABIndent, '', [rfReplaceAll]);
      if BlockMiddleList.Match then
      begin
        if VBLine.Chars[0] = TABIndent then
          AllLines[i] := Copy(AllLines[i], 2);
      end;
    end;
    AAllCode := StrArrayJoin(AllLines).Replace(BlockingIndent, '', [rfReplaceAll]);
    ACodeModule.DeleteLines(AStartInsPos, ACountOrigLines);
    ACodeModule.InsertLines(AStartInsPos, AAllCode);
    ACodeModule.CodePane.SetSelection(AStartInsPos, 1, AStartInsPos, 1);
    ACodeModule.CodePane.TopLine := ATopLine;
  finally
    IndentFinalize;
  end;
end;

procedure IndentCodeModule(ACodeModule: CodeModule);
var AllCode: String;
    TopLine: Integer;
begin
  TopLine := ACodeModule.CodePane.TopLine;
  AllCode := ACodeModule.Lines[1, ACodeModule.CountOfLines];
  IndentCode(ACodeModule, 1, ACodeModule.CountOfLines, TopLine, AllCode);
end;

procedure IndentSubrogramm(ACodeModule: CodeModule);
var ProcKind: vbext_ProcKind;
    AllCode, ProcName: String;
    StartProcLine, CountProcLines, CurrLine: Integer;
begin
  CurrLine := GetCurrentVBELine(ACodeModule);
  ProcName := ACodeModule.ProcOfLine[CurrLine, ProcKind];
  if ProcName <> '' then
  begin
    StartProcLine := ACodeModule.ProcStartLine[ProcName, ProcKind];
    CountProcLines := ACodeModule.ProcCountLines[ProcName, ProcKind];
    AllCode := ACodeModule.Lines[StartProcLine, CountProcLines];
    IndentCode(ACodeModule, StartProcLine, CountProcLines, StartProcLine, AllCode);
  end
  else
    MessageBoxW(0, PWideChar(rsErr_CanNotDetermineSubProg),
      PWideChar(rs_RegAddinFriendlyName), MB_OK);
end;

procedure IndentSelectedLines(ACodeModule: CodeModule);
var AllCode: String;
    StartCursorPos, StartColumn, EndLine, EndColumn, CountLines: Integer;
begin
  ACodeModule.CodePane.GetSelection(StartCursorPos, StartColumn, EndLine, EndColumn);
  CountLines := EndLine - StartCursorPos + 1;
  if CountLines >= 1 then
  begin
    AllCode := ACodeModule.Lines[StartCursorPos, CountLines];
    IndentCode(ACodeModule, StartCursorPos, CountLines, StartCursorPos, AllCode);
  end;
end;

end.
