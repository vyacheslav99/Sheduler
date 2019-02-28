program sheduler;

uses
  Windows,
  Messages,
  ShellAPI,
  SysUtils,
  utils in 'src\utils.pas',
  tasks in 'src\tasks.pas',
  ctrlWnd in 'src\ctrlWnd.pas' {FCtrlWnd};

{$R *.res}
{$R icon.res}

function WindowProc(hWnd, Msg, wParam, lParam: Longint): Longint; stdcall;
begin
  result := DefWindowProc(hWnd, Msg, wParam, lParam);

  case Msg of
    WM_DESTROY, WM_CLOSE, WM_QUIT, WM_ENDSESSION:
    begin
      if Msg = WM_ENDSESSION then
      begin
        AddToWorkLog(ltInfo, 'Получен сигнал завершения работы системы');
        // тут имеет смысл завершить все выполняющиеся задачи. Они всеравно будут обрублены, так хоть в лог будет соотв. запись
        TaskList.BreakAll;        
      end;
      Close;
    end;
    Ico_Message:
    begin
      case lParam of
        WM_RBUTTONUP:
        begin
          GetCursorPos(MPos);
          SetForegroundWindow(FHandle);        //так рекомендует Microsoft
          TrackPopupMenu(PopupMenu, TPM_RIGHTALIGN + TPM_RIGHTBUTTON, MPos.X, MPos.Y, 0, FHandle, nil);
          PostMessage(FHandle, WM_NULL, 0, 0); //так рекомендует Microsoft
        end;
        WM_LBUTTONUP:
        begin
          SetForegroundWindow(FHandle);
          PostMessage(FHandle, WM_NULL, 0, 0);
          ShowShedulerWnd;
        end;
      end;
    end;
    WM_COMMAND:
    begin
      case wParam of
        pmExit: Close;
        pmStart: TaskList.Start;
        pmStop: TaskList.Stop;
        pmSettings: ShowShedulerWnd;
        pmEditTasks:
        begin
          if not FileExists(ParamFolder + ParamFile) then CreateEmptyFile(ParamFolder + ParamFile);
          MessageBox(FHandle, pchar('Не забудь, что после ручного редактирования конфигурационного файла, чтобы изменения ' +
            'вступили в силу, нужно нажать кнопку "Перечитать параметры" в окне управления или в меню значка в трее!'),
            pchar(AppTitle), MB_OK + MB_ICONINFORMATION);

          ShellExecute(FHandle, 'open', pchar(ParamFolder + ParamFile), nil, pchar(ParamFolder), SW_SHOWNORMAL);
        end;
        pmReloadTasks:
        begin
          LoadParams;
          TaskList.LoadTasks;
        end;
        pmOpenLog:
          if FileExists(ParamFolder + LogFile) then
            ShellExecute(FHandle, 'open', pchar(ParamFolder + LogFile), nil, pchar(ParamFolder), SW_SHOWNORMAL)
          else
            MessageBox(FHandle, 'Лог отсутствует', pchar(AppTitle), MB_OK + MB_ICONINFORMATION);
        pmBreak:
        begin
          if MessageBox(FHandle, 'Завершить все выполняющиеся процессы?', pchar(AppTitle), MB_YESNO + MB_ICONQUESTION) = ID_YES then
            TaskList.BreakAll;
        end;
        pmBreakAndExit:
        begin
          if MessageBox(FHandle, 'Выйти и завершить все выполняющиеся процессы?', pchar(AppTitle), MB_YESNO + MB_ICONQUESTION) = ID_YES then
          begin
            TaskList.BreakAll;
            Close;
          end;
        end;
      end;
    end;
    WM_TIMER:
    begin
      case wParam of
        idTaskTimer: TaskList.ExecCurrent;
      end;
    end;
    U_WM_RELOAD_PARAM:
    begin
      LoadParams;
      TaskList.LoadTasks;
    end;
    U_WM_BREAK: TaskList.BreakAll;
  end;
end;

var
  cmd: string;
  sh: HWND;
  noexit: boolean;
  
begin
  if ParamCount > 0 then
  begin
    // посылаем команды из консоли выполняющемуся экземпляру - возможность управления из консоли
    sh := FindWindow(WinClass, nil);
    if sh = 0 then exit;

    noexit := false;
    cmd := LowerCase(ParamStr(1));
    AddToWorkLog(ltInfo, 'Посылаю команду: ' + cmd);

    if cmd = CMD_PAUSE then SendMessage(sh, WM_COMMAND, pmStop, 0)
    else if cmd = CMD_RESUME then SendMessage(sh, WM_COMMAND, pmStart, 0)
    else if cmd = CMD_RELOAD then SendMessage(sh, U_WM_RELOAD_PARAM, 0, 0)
    else if cmd = CMD_BRAEK then SendMessage(sh, U_WM_BREAK, 0, 0)
    else if cmd = CMD_STOP then
    begin
      if LowerCase(ParamStr(2)) = '-b' then SendMessage(sh, U_WM_BREAK, 0, 0);
      SendMessage(sh, WM_CLOSE, 0, 0);
    end else if cmd = CMD_RESTART then
    begin
      SendMessage(sh, WM_CLOSE, 0, 0);
      noexit := true;
    end;

    if not noexit then exit;
  end;

  if FindWindow(WinClass, nil) <> 0 then exit;
  StartTime := Now;
  AddToWorkLog(ltInfo, 'Начинаю работу');
  Instance := GetModuleHandle(nil);
  with WindowClass do
  begin
    style := CS_HREDRAW or CS_VREDRAW;
    lpfnWndProc := @WindowProc;
    hInstance := Instance;
    hbrBackground := COLOR_BTNFACE;
    lpszClassName := pchar(WinClass);
    hCursor := LoadCursor(0, IDC_ARROW);
    hIcon := LoadIcon(Instance, pchar(ICOSTOPPED));
  end;

  RegisterClass(WindowClass);
  FHandle := CreateWindowEx(0, WinClass, AppTitle, WS_POPUP, 5, 5, 200, 200, 0, 0, Instance, nil);

  FCtrlWnd := TFCtrlWnd.Create(nil);

  CreateMyIcon;
  PopupMenu := CreatePopupMenu;
  AppendMenu(PopupMenu, MF_STRING, pmStart, 'Возобновить');
  AppendMenu(PopupMenu, MF_SEPARATOR, 0, '');
  AppendMenu(PopupMenu, MF_STRING, pmSettings, 'Управление');
  AppendMenu(PopupMenu, MF_STRING, pmEditTasks, 'Редактировать файл заданий');
  AppendMenu(PopupMenu, MF_STRING, pmReloadTasks, 'Перечитать параметры');
  AppendMenu(PopupMenu, MF_STRING, pmOpenLog, 'Просмотр лога');
  AppendMenu(PopupMenu, MF_SEPARATOR, 0, '');
  AppendMenu(PopupMenu, MF_STRING, pmBreak, 'Завершить все задачи');
  AppendMenu(PopupMenu, MF_SEPARATOR, 0, '');
  AppendMenu(PopupMenu, MF_STRING, pmBreakAndExit, 'Выйти и завершить все задачи');
  AppendMenu(PopupMenu, MF_STRING, pmExit, 'Вйти');
  LoadParams(true);

  TaskList := TTaskList.Create;
  TaskList.Start;

  while (GetMessage(msg, 0, 0, 0)) do
  begin
    TranslateMessage(msg);       
    DispatchMessage(msg);
  end;
end.
