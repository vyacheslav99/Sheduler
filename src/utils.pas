unit utils;

interface

uses
  Windows, Messages, ShellAPI, SysUtils, Classes, IniFiles, Registry, FileCtrl;

const
  WinClass = 'TShedulerAgentClass';
  AppTitle = 'Планировщик заданий';
  ICOSTOPPED = 'MAINICONSTOP';
  ICOSTARTED = 'MAINICON';
  AUTORUNKEY = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run';
  AUTORUNPARAM = 'Sheduler';
  TASK_SECTION_PREFIX = 'TASK_';
  PARAM_SECTION = 'CONFIG';
  idTaskTimer = 1;
  Ico_Message = WM_USER + $110;
  U_WM_RELOAD_PARAM = Ico_Message + 1;
  U_WM_BREAK = Ico_Message + 2;
  pmExit = 1;
  pmStart = 2;
  pmStop = 3;
  pmSettings = 4;
  pmOpenLog = 5;
  pmBreak = 6;
  pmBreakAndExit = 7;
  pmEditTasks = 8;
  pmReloadTasks = 9;
  // команды командной строки
  CMD_PAUSE = 'pause';
  CMD_RESUME = 'resume';
  CMD_RELOAD = 'reload';
  CMD_STOP = 'stop';  // c флагом -b - убить все процессы и выйти
  CMD_RESTART = 'restart';
  CMD_BRAEK = 'break';

type
  TAutorunSection = (asNone, asAllUsers, asCurrentUser);
  TWorkLogType = (ltInfo, ltWarning, ltError, ltDebug);

var
  Instance, FHandle: HWND;
  WindowClass: TWndClass;
  Msg: TMsg;
  MPos: TPoint;
  noIconData: TNotifyIconData;
  HIcon1: HICON;
  PopupMenu: HMENU;
  MonStatus: integer;
  WorkLog: TStringList;
  LastError: string;
  ParamFolder: string;
  ParamFile: string;
  LogFile: string;
  StartTime: TDateTime;
  // настройки
  MaxTerminatedThread: integer;
  LastTaskId: integer;
  CWLeft: integer;
  CWTop: integer;
  CWWidth: integer;
  CWHeight: integer;
  CWMaximized: boolean;

function LoadParams(Startup: boolean = false): boolean;
function SaveParams: boolean;
procedure SetDefaultParams;
procedure RestartTimer(IdTimer: integer);
procedure AddToWorkLog(LogType: TWorkLogType; str: string; NewLine: boolean = true; CanAddTime: boolean = true);
procedure WriteWorkLog;
procedure ChangeStatus(Code: integer);
procedure Close;
procedure CreateMyIcon;
procedure ShowShedulerWnd;
function ReadRegValue(RootKey: HKEY; Key, Param, default: string): string;
function WriteRegValue(RootKey: HKEY; Key, Param, Value: string; CanCreate: boolean): boolean;
function DelRegValue(RootKey: HKEY; Key, Param: string): boolean;
function CheckAutorun: TAutorunSection;
function InstallAutorun: TAutorunSection;
function UninstallAutorun: TAutorunSection;
function ExtractWordEx(n: integer; s: string; WordDelims: TSysCharSet; IgnoreBlockChars: TSysCharSet): string;
function WordCountEx(s: string; WordDelims: TSysCharSet; IgnoreBlockChars: TSysCharSet): integer;
function GenRandString(genrule, vlength: byte): string;
function BytesToStr(b: int64): string;
function GetTempDir: string;
function KillProcess(PID: Cardinal; var ErrMsg: string): boolean;
function CreateEmptyFile(FileName: string): boolean;
function SelectDir(cap: string; var dir: string): boolean;

implementation

uses tasks, ctrlWnd;

function LoadParams(Startup: boolean): boolean;
var
  f: TIniFile;
  canSave: boolean;

begin
  result := true;
  SetDefaultParams;
  canSave := not FileExists(ParamFolder + ParamFile);
  f := TIniFile.Create(ParamFolder + ParamFile);

  try
    MaxTerminatedThread := f.ReadInteger(PARAM_SECTION, 'MaxTerminatedThread', 100);
    if Startup then LastTaskId := f.ReadInteger(PARAM_SECTION, 'LastTaskId', 0);
  except
    on e: Exception do
    begin
      result := false;
      AddToWorkLog(ltError, 'Не удалось загрузить настройки! Ошибка: ' + e.Message);
    end;
  end;

  f.Free;
  AddToWorkLog(ltInfo, 'Параметры загружены');
  if canSave then SaveParams;
  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

function SaveParams: boolean;
var
  f: TIniFile;

begin
  result := true;
  f := TIniFile.Create(ParamFolder + ParamFile);

  try
    f.WriteInteger(PARAM_SECTION, 'MaxTerminatedThread', MaxTerminatedThread);
    f.WriteInteger(PARAM_SECTION, 'LastTaskId', LastTaskId);
  except
    on e: Exception do
    begin
      result := false;
      AddToWorkLog(ltError, 'Не удалось сохранить настройки! Ошибка: ' + e.Message);
    end;
  end;
  
  f.Free;
end;

procedure SetDefaultParams;
begin
  MaxTerminatedThread := 100;
  LastTaskId := 0;
end;

procedure RestartTimer(IdTimer: integer);
begin
  KillTimer(FHandle, IdTimer);
  if (MonStatus = pmStart) then
  begin
    case IdTimer of
      idTaskTimer: SetTimer(FHandle, IdTimer, 60 * 1000, nil);
    end;
  end;
end;

procedure AddToWorkLog(LogType: TWorkLogType; str: string; NewLine: boolean; CanAddTime: boolean);
var
  sType: string;
  
begin
  //формат лога:
  //04.04.2010  10:15:00  <ТИП СООБЩЕНИЯ>  <current_user>  сообщение

  case LogType of
    ltInfo: sType := '<INFO>   ';
    ltWarning: sType := '<WARNING>';
    ltError: sType := '<ERROR>  ';
    ltDebug: sType := '<DEBUG>  ';
  end;

  if CanAddTime then str := DateToStr(Now) + #9 + TimeToStr(Now) + #9 + sType + #9 + str;

  if NewLine or (WorkLog.Count = 0) then WorkLog.Add(str)
  else
    if (Length(WorkLog.Strings[WorkLog.Count - 1]) >= 128) then WorkLog.Add(str)
    else WorkLog.Strings[WorkLog.Count - 1] := WorkLog.Strings[WorkLog.Count - 1] + str;

  WriteWorkLog;
end;

procedure WriteWorkLog;
var
  fs: TFileStream;
  ss: TStringStream;
  fn: string;

begin
  if WorkLog.Count = 0 then exit;
  fn := ParamFolder + LogFile;

  try
    try
      ss := TStringStream.Create('');
      if FileExists(fn) then
      begin
        fs := TFileStream.Create(fn, fmOpenWrite);
        fs.Seek(fs.Size, soFromBeginning);
      end else
        fs := TFileStream.Create(fn, fmCreate);

      WorkLog.SaveToStream(ss);
      ss.Seek(0, soFromBeginning);
      fs.CopyFrom(ss, ss.Size);
      WorkLog.Clear;
    finally
      if Assigned(ss) then ss.Free;
      if Assigned(fs) then fs.Free;
    end;
  except
  end;
end;

procedure ChangeStatus(Code: integer);
var
  s: string;
  i: integer;

begin
  MonStatus := Code;
  case Code of
    pmStart:
    begin
      ModifyMenu(PopupMenu, pmStart, MF_BYCOMMAND or MF_STRING or MF_CHANGE, pmStop, 'Приостановить');
      HIcon1 := LoadIcon(Instance, ICOSTARTED);
      s := AppTitle + ' - Активен';
    end;
    //pmStop:
    else begin
      ModifyMenu(PopupMenu, pmStop, MF_BYCOMMAND or MF_STRING or MF_CHANGE, pmStart, 'Возобновить');
      HIcon1 := LoadIcon(Instance, ICOSTOPPED);
      s := AppTitle + ' - Приостановлен';
    end;
  end;

  i := 0;
  while (i < Length(s)) and (i < High(noIconData.szTip)) do
  begin
    noIconData.szTip[i] := s[i + 1];
    Inc(i);
  end;
  noIconData.szTip[i] := #0;
  noIconData.hIcon := HIcon1;
  Shell_NotifyIcon(NIM_MODIFY, @noIconData);
  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

procedure Close;
begin
  try
    FCtrlWnd.Close;
    FCtrlWnd.Free;
    SaveParams;
    TaskList.Free;
    AddToWorkLog(ltInfo, 'Завершение работы');
    AddToWorkLog(ltInfo, '*************************************************************************');
    AddToWorkLog(ltInfo, '');
    WriteWorkLog;
    Shell_NotifyIcon(NIM_DELETE, @noIconData);
  finally
    Halt;
  end;
end;

procedure CreateMyIcon;
begin
  HIcon1 := LoadIcon(Instance, ICOSTOPPED);

  with noIconData do
  begin
    cbSize := SizeOf(TNotifyIconData);
    wnd := FHandle;
    uID := 0;
    uFlags := NIF_MESSAGE or NIF_ICON or NIF_TIP;
    FillChar(szTip, SizeOf(szTip), #0);
    Move(AppTitle[1], szTip, Length(AppTitle));
    //szTip := AppTitle;
    hIcon := HIcon1;
    uCallbackMessage := Ico_Message;
  end;
  Shell_NotifyIcon(NIM_ADD, @noIconData);
end;

procedure ShowShedulerWnd;
begin
  if Assigned(FCtrlWnd) then
  begin
    if (not FCtrlWnd.Visible) then FCtrlWnd.Show;
    FCtrlWnd.BringToFront;
  end;
end;

function ReadRegValue(RootKey: HKEY; Key, Param, default: string): string;
var
  reg: TRegistry;

begin
  result := default;
  reg := TRegistry.Create(KEY_READ);
  reg.RootKey := RootKey;

  try
    if reg.OpenKey(Key, False) then result := reg.ReadString(Param);
  except
  end;
  
  reg.CloseKey;
  reg.Free;
end;

function WriteRegValue(RootKey: HKEY; Key, Param, Value: string; CanCreate: boolean): boolean;
var
  reg: TRegistry;

begin
  result := false;
  reg := TRegistry.Create(KEY_ALL_ACCESS);
  try
    reg.RootKey := RootKey;
    reg.LazyWrite := false;
    if reg.OpenKey(Key, CanCreate) then
    begin
      reg.WriteString(Param, Value);
      result := true;
    end;
  except
  end;
  reg.CloseKey;
  reg.Free;
end;

function DelRegValue(RootKey: HKEY; Key, Param: string): boolean;
var
  reg: TRegistry;

begin
  result := false;
  reg := TRegistry.Create(KEY_ALL_ACCESS);
  try
    reg.RootKey := RootKey;
    reg.LazyWrite := False;
    if reg.OpenKey(Key, false) then
      result := reg.DeleteValue(Param);
  except
    result := false;
  end;
  reg.Free;
end;

function CheckAutorun: TAutorunSection;
begin
  result := asNone;
  if (ReadRegValue(HKEY_LOCAL_MACHINE, AUTORUNKEY, AUTORUNPARAM, '') <> '') then
    result := asAllUsers;

  if (result = asNone) then
    if (ReadRegValue(HKEY_CURRENT_USER, AUTORUNKEY, AUTORUNPARAM, '') <> '') then
      result := asCurrentUser;
end;

function InstallAutorun: TAutorunSection;
var
  s: string;

begin
  result := CheckAutorun;
  if result in [asNone, asCurrentUser] then
  begin
    if WriteRegValue(HKEY_LOCAL_MACHINE, AUTORUNKEY, AUTORUNPARAM, '"' + ParamStr(0) + '"', true) then
    begin
      result := asAllUsers;
      s := 'Автозапуск: ДОБАВЛЕН ВСЕМ пользователям';
    end else
      s := 'Автозапуск: неудачная попытка записи в HKLM';
  end;

  if result = asNone then
  begin
    if WriteRegValue(HKEY_CURRENT_USER, AUTORUNKEY, AUTORUNPARAM, '"' + ParamStr(0) + '"', true) then
    begin
      result := asCurrentUser;
      s := 'Автозапуск: ДОБАВЛЕН ТЕКУЩЕМУ пользователю';
    end else
      s := 'Автозапуск: неудачная попытка записи HKCU';
  end;

  if s <> '' then AddToWorkLog(ltWarning, s);
  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

function UninstallAutorun: TAutorunSection;
var
  s: string;

begin
  result := CheckAutorun;
  if result = asNone then exit;
  if not DelRegValue(HKEY_CURRENT_USER, AUTORUNKEY, AUTORUNPARAM) then s := 'Автозапуск: НЕ УДАЛЕН у ТЕКУЩЕГО пользователя';
  if not DelRegValue(HKEY_LOCAL_MACHINE, AUTORUNKEY, AUTORUNPARAM) then
  begin
    if s = '' then s := 'Автозапуск: НЕ УДАЛЕН у ВСЕХ пользователей'
    else s := s + ', НЕ УДАЛЕН у ВСЕХ пользователей';
  end;
  result := CheckAutorun;

  if result = asNone then s := 'Автозапуск: УДАЛЕН';
  if s <> '' then AddToWorkLog(ltWarning, s);
  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

function ExtractWordEx(n: integer; s: string; WordDelims: TSysCharSet; IgnoreBlockChars: TSysCharSet): string;
var
  CurrBlChar: char;
  iblock: boolean;
  i: integer;
  wn: integer;

begin
  result := '';
  iblock := false;
  CurrBlChar := #0;
  wn := 1;

  for i := 1 to Length(s) do
  begin
    if (iblock) then
    begin
      if (s[i] = CurrBlChar) then
      begin
        iblock := false;
        CurrBlChar := #0;
      end;
      if (wn = n) then result := result + s[i];
      continue;
    end;
    if (s[i] in IgnoreBlockChars) then
    begin
      iblock := true;
      CurrBlChar := s[i];
      if (wn = n) then result := result + s[i];
      continue;
    end;
    if (s[i] in WordDelims) then
    begin
      Inc(wn);
      if (wn > n) then exit;
    end else
      if (wn = n) then result := result + s[i];
  end;
end;

function WordCountEx(s: string; WordDelims: TSysCharSet; IgnoreBlockChars: TSysCharSet): integer;
var
  CurrBlChar: char;
  iblock: boolean;
  i: integer;

begin
  if (s = '') then result := 0
  else result := 1;
  iblock := false;
  CurrBlChar := #0;

  for i := 1 to Length(s) do
  begin
    if (iblock) then
    begin
      if (s[i] = CurrBlChar) then
      begin
        iblock := false;
        CurrBlChar := #0;
      end;
      continue;
    end;
    if (s[i] in IgnoreBlockChars) then
    begin
      iblock := true;
      CurrBlChar := s[i];
      continue;
    end;
    if ((s[i] in WordDelims) and (i < Length(s))) then Inc(result);
  end;
end;

function GenRandString(genrule, vlength: byte): string;
var
  i: integer;
  c: byte;
  symbs: set of byte;

begin
  result := '';
  symbs := [];
  if genrule > 6 then genrule := 6;

  case genrule of
    0: symbs := symbs + [48..57];                           //цифры 0..9
    1: symbs := symbs + [65..90, 97..122];                  //буквы A..Z, a..z
    2: symbs := symbs + [65..90];                           //буквы A..Z
    3: symbs := symbs + [97..122];                          //буквы a..z
    4: symbs := symbs + [48..57, 65..90];                   //цифры + буквы A..Z
    5: symbs := symbs + [48..57, 97..122];                  //цифры + буквы a..z
    6: symbs := symbs + [48..57, 65..90, 97..122];          //цифры + буквы A..Z, a..z
  end;
  if symbs = [] then exit;

  Randomize;
  for i := 1 to vlength do
  begin
    c := Random(123);
    while not (c in symbs) do c := Random(123);
    result := result + chr(c);
  end;
end;

function BytesToStr(b: int64): string;
const
  GB = 1024 * 1024 * 1024;
  MB = 1024 * 1024;
  KB = 1024;

begin
  if b div GB > 0 then
    result := Format('%.2f Gb', [b / GB])
  else if b div MB > 0 then
    Result := Format('%.2f Mb', [b / MB])
  else if b div KB > 0 then
    Result := Format('%.2f kb', [b / KB])
  else
    Result := IntToStr(b) + ' b';
end;

function GetTempDir: string;
var
  buff: array [0..255] of char;

begin
  GetEnvironmentVariable(pchar('TEMP'), buff, SizeOf(buff));
  result := string(buff);
  if (Length(result) > 0) and (result[Length(result)] = '\') then Delete(result, Length(result), 1);
end;

function KillProcess(PID: Cardinal; var ErrMsg: string): boolean;
const
  PROCESS_TERMINATE = $0001;

var
  h: THandle;
  
begin
  result := false;
  ErrMsg := '';

  h := OpenProcess(PROCESS_TERMINATE, false, PID);
  if h <> 0 then
  begin
    result := TerminateProcess(h, 0);
    CloseHandle(h);
  end;

 if not result then ErrMsg := SysErrorMessage(GetLastError);
end;

function CreateEmptyFile(FileName: string): boolean;
var
  fs: TFileStream;

begin
  result := false;
  if FileExists(FileName) then exit;

  try
    fs := TFileStream.Create(FileName, fmCreate);
    result := true;
  finally
    fs.Free;
  end;
end;

function SelectDir(cap: string; var dir: string): boolean;
begin
  result := SelectDirectory(cap, '', dir, []);
end;

initialization
  ParamFolder := ExtractFilePath(ParamStr(0));
  ParamFile := ChangeFileExt(ExtractFileName(ParamStr(0)), '.conf');
  LogFile := ChangeFileExt(ExtractFileName(ParamStr(0)), '.log');
  WorkLog := TStringList.Create;

finalization
  WorkLog.Free;

end.
