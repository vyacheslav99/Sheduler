unit ctrlWnd;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, utils, tasks, ExtCtrls, StdCtrls,
  ComCtrls, XpMan, Buttons, ShellAPI, MemTableDataEh, Db, MemTableEh, DBGridEhGrouping, GridsEh, DBGridEh, EhLibMTE, IniFiles,
  Spin, DateUtils, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, EhLibVCL, DBAxisGridsEh;

type
  TFCtrlWnd = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    lbTaskCount: TLabel;
    Label2: TLabel;
    lbRunningCount: TLabel;
    Label3: TLabel;
    lbTerminatedThreads: TLabel;
    Label4: TLabel;
    lbStartTime: TLabel;
    Label5: TLabel;
    lbWorkTime: TLabel;
    Label6: TLabel;
    lbStatus: TLabel;
    pcMain: TPageControl;
    tsControl: TTabSheet;
    tsTask: TTabSheet;
    tsThread: TTabSheet;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    btnStart: TBitBtn;
    btnStop: TBitBtn;
    btnOpenLog: TBitBtn;
    btnDeleteLog: TBitBtn;
    btnBreakAll: TBitBtn;
    btnAddToAutorun: TButton;
    btnSaveParams: TBitBtn;
    btnOpenConfig: TBitBtn;
    btnReloadParams: TButton;
    lbAutorun: TLabel;
    mtTask: TMemTableEh;
    dbgTask: TDBGridEh;
    dsoTask: TDataSource;
    mtTaskTaskId: TIntegerField;
    mtTaskTaskName: TStringField;
    mtTaskExecCommand: TStringField;
    mtTaskParametres: TStringField;
    mtTaskWorkDir: TStringField;
    mtTaskShowWindow: TIntegerField;
    mtTaskMinute: TStringField;
    mtTaskHour: TStringField;
    mtTaskDay: TStringField;
    mtTaskMonth: TStringField;
    mtTaskWeekday: TStringField;
    mtTaskAllowMultiple: TBooleanField;
    mtTaskWaitProcess: TBooleanField;
    btnSaveTask: TBitBtn;
    mtTaskChanged: TBooleanField;
    btnAddTask: TSpeedButton;
    btnDelTask: TSpeedButton;
    GroupBox3: TGroupBox;
    Label8: TLabel;
    lbTaskStatus: TLabel;
    Label9: TLabel;
    lbTaskCopyCnt: TLabel;
    Label11: TLabel;
    lbLastTaskStart: TLabel;
    Label10: TLabel;
    lbLastTaskStop: TLabel;
    btnStartTask: TSpeedButton;
    Label12: TLabel;
    lbLastMessage: TLabel;
    btnBreakTask: TSpeedButton;
    mtThread: TMemTableEh;
    dsoThread: TDataSource;
    mtThreadProcessId: TIntegerField;
    mtThreadStartTime: TDateTimeField;
    mtThreadStopTime: TDateTimeField;
    mtThreadTerminated: TBooleanField;
    mtThreadLastError: TStringField;
    dbgThread: TDBGridEh;
    mtThreadTaskId: TIntegerField;
    mtThreadTaskName: TStringField;
    mtThreadExecCommand: TStringField;
    mtThreadParametres: TStringField;
    Label7: TLabel;
    edMaxThreadCount: TSpinEdit;
    mtThreadIndex: TIntegerField;
    btnBreakThread: TSpeedButton;
    OpenDialog: TOpenDialog;
    btnRefresh: TSpeedButton;
    mtTaskActive: TBooleanField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnStartClick(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure btnBreakAllClick(Sender: TObject);
    procedure btnOpenConfigClick(Sender: TObject);
    procedure btnOpenLogClick(Sender: TObject);
    procedure btnDeleteLogClick(Sender: TObject);
    procedure btnReloadParamsClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnAddToAutorunClick(Sender: TObject);
    procedure btnSaveParamsClick(Sender: TObject);
    procedure btnSaveTaskClick(Sender: TObject);
    procedure mtTaskBeforePost(DataSet: TDataSet);
    procedure dbgTaskGetCellParams(Sender: TObject; Column: TColumnEh; AFont: TFont; var Background: TColor;
      State: TGridDrawState);
    procedure dbgTaskKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnAddTaskClick(Sender: TObject);
    procedure btnDelTaskClick(Sender: TObject);
    procedure mtTaskAfterScroll(DataSet: TDataSet);
    procedure mtTaskAfterOpen(DataSet: TDataSet);
    procedure btnStartTaskClick(Sender: TObject);
    procedure btnBreakTaskClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnBreakThreadClick(Sender: TObject);
    procedure dbgThreadGetCellParams(Sender: TObject; Column: TColumnEh; AFont: TFont; var Background: TColor;
      State: TGridDrawState);
    procedure mtThreadAfterScroll(DataSet: TDataSet);
    procedure mtThreadAfterOpen(DataSet: TDataSet);
    procedure dbgTaskColumns2EditButtonClick(Sender: TObject; var Handled: Boolean);
    procedure dbgTaskColumns4EditButtonClick(Sender: TObject; var Handled: Boolean);
    procedure btnRefreshClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    TaskListChanged: boolean;
    TaskListLoading: boolean;
    procedure LoadWndState;
    procedure SaveWndState;
  public
    procedure Refresh;
    procedure LoadTaskList;
    procedure SaveTaskList;
    procedure RefreshThreadList;
  end;

var
  FCtrlWnd: TFCtrlWnd;

implementation

{$R *.dfm}

procedure TFCtrlWnd.btnAddTaskClick(Sender: TObject);
begin
  if not mtTask.Active then exit;
  mtTask.Append;
  mtTaskShowWindow.AsInteger := Ord(swDefault);
  mtTaskMinute.AsString := '*';
  mtTaskHour.AsString := '*';
  mtTaskDay.AsString := '*';
  mtTaskMonth.AsString := '*';
  mtTaskWeekday.AsString := '*';
  mtTaskAllowMultiple.AsBoolean := false;
  mtTaskWaitProcess.AsBoolean := true;
  mtTaskActive.AsBoolean := true;
end;

procedure TFCtrlWnd.btnAddToAutorunClick(Sender: TObject);
begin
  case CheckAutorun of
    asNone: InstallAutorun;
    asAllUsers, asCurrentUser: UninstallAutorun;
  end;
end;

procedure TFCtrlWnd.btnBreakAllClick(Sender: TObject);
begin
  if MessageBox(FHandle, 'Завершить все выполняющиеся процессы?', pchar(AppTitle), MB_YESNO + MB_ICONQUESTION) = ID_YES then
    TaskList.BreakAll;
end;

procedure TFCtrlWnd.btnBreakTaskClick(Sender: TObject);
var
  err: string;

begin
  if (not mtTask.Active) or mtTask.IsEmpty then exit;

  if mtTaskTaskId.IsNull or TaskListChanged then
  begin
    MessageBox(FHandle, pchar('Сначала сохрани список задач!'), pchar(AppTitle), MB_OK + MB_ICONINFORMATION);
    exit;
  end;

  if MessageBox(FHandle, pchar('Завершить задачу ' + mtTaskTaskId.AsString + ': ' + mtTaskTaskName.AsString + '? ' +
    'Будут завершены все выполняющиеся экземпляры задачи!'), pchar(AppTitle), MB_YESNO + MB_ICONQUESTION) <> ID_YES then exit;

  TaskList.BreakTask(mtTaskTaskId.AsInteger, err);
  if err <> '' then
    MessageBox(FHandle, pchar('Ошибка! ' + err), pchar(AppTitle), MB_OK + MB_ICONERROR);

  mtTaskAfterScroll(mtTask);
end;

procedure TFCtrlWnd.btnBreakThreadClick(Sender: TObject);
begin
  if (not mtThread.Active) or mtThread.IsEmpty then exit;

  if MessageBox(FHandle, pchar('Завершить процесс ' + mtThreadProcessId.AsString + ', задача: ' + mtTaskTaskId.AsString + ': ' +
    mtTaskTaskName.AsString + '?'), pchar(AppTitle), MB_YESNO + MB_ICONQUESTION) <> ID_YES then exit;

  TaskList.BreakThread(mtThreadIndex.AsInteger);
  RefreshThreadList;
  mtThreadAfterScroll(mtThread);
end;

procedure TFCtrlWnd.btnDeleteLogClick(Sender: TObject);
begin
  if MessageBox(FHandle, 'Уверен, что хочешь очистить лог?', pchar(AppTitle), MB_YESNO + MB_ICONQUESTION) = ID_YES then
    if FileExists(ParamFolder + LogFile) then DeleteFile(ParamFolder + LogFile);
end;

procedure TFCtrlWnd.btnDelTaskClick(Sender: TObject);
begin
  if (not mtTask.Active) or mtTask.IsEmpty then exit;
  
  if MessageBox(FHandle, pchar('Удалить задачу "' + mtTaskTaskId.AsString + ': ' + mtTaskTaskName.AsString + '"?'),
    pchar(AppTitle), MB_YESNO + MB_ICONQUESTION) <> ID_YES then exit;

  mtTask.Delete;
end;

procedure TFCtrlWnd.btnOpenConfigClick(Sender: TObject);
begin
  if not FileExists(ParamFolder + ParamFile) then CreateEmptyFile(ParamFolder + ParamFile);

  MessageBox(FHandle, pchar('Не забудь, что после ручного редактирования конфигурационного файла нужно нажать кнопку ' +
    '"Перечитать параметры" тут или в меню значка в трее!'), pchar(AppTitle), MB_OK + MB_ICONINFORMATION);
    
  ShellExecute(FHandle, 'open', pchar(ParamFolder + ParamFile), nil, pchar(ParamFolder), SW_SHOWNORMAL);
end;

procedure TFCtrlWnd.btnOpenLogClick(Sender: TObject);
begin
  if FileExists(ParamFolder + LogFile) then
    ShellExecute(FHandle, 'open', pchar(ParamFolder + LogFile), nil, pchar(ParamFolder), SW_SHOWNORMAL)
  else
    MessageBox(FHandle, pchar('Лог пока отсутствует'), pchar(AppTitle), MB_OK + MB_ICONINFORMATION);
end;

procedure TFCtrlWnd.btnRefreshClick(Sender: TObject);
begin
  RefreshThreadList;
end;

procedure TFCtrlWnd.btnReloadParamsClick(Sender: TObject);
begin
  LoadParams;
  TaskList.LoadTasks;
end;

procedure TFCtrlWnd.btnSaveParamsClick(Sender: TObject);
begin
  SaveParams;
end;

procedure TFCtrlWnd.btnSaveTaskClick(Sender: TObject);
begin
  SaveTaskList;
  TaskList.LoadTasks;
  LoadTaskList;
end;

procedure TFCtrlWnd.btnStartClick(Sender: TObject);
begin
  TaskList.Start;
end;

procedure TFCtrlWnd.btnStartTaskClick(Sender: TObject);
var
  err: string;

begin
  if (not mtTask.Active) or mtTask.IsEmpty then exit;

  if mtTaskTaskId.IsNull or TaskListChanged then
  begin
    MessageBox(FHandle, pchar('Сначала сохрани список задач!'), pchar(AppTitle), MB_OK + MB_ICONINFORMATION);
    exit;
  end;

  if TaskList.StartTaskById(mtTaskTaskId.AsInteger, err) < 0 then
    MessageBox(FHandle, pchar('Ошибка! ' + err), pchar(AppTitle), MB_OK + MB_ICONERROR);

  mtTaskAfterScroll(mtTask);
end;

procedure TFCtrlWnd.btnStopClick(Sender: TObject);
begin
  TaskList.Stop;
end;

procedure TFCtrlWnd.dbgTaskColumns2EditButtonClick(Sender: TObject; var Handled: Boolean);
begin
  if (not mtTask.Active) or (mtTask.IsEmpty) then exit;
  
  OpenDialog.FileName := mtTaskExecCommand.AsString;
  if OpenDialog.Execute then
  begin
    if not (mtTask.State in [dsEdit, dsInsert]) then mtTask.Edit;
    mtTaskExecCommand.AsString := OpenDialog.FileName;
  end;
end;

procedure TFCtrlWnd.dbgTaskColumns4EditButtonClick(Sender: TObject; var Handled: Boolean);
var
  s: string;

begin
  if (not mtTask.Active) or (mtTask.IsEmpty) then exit;

  s := mtTaskWorkDir.AsString;
  if SelectDir('Выбери папку', s) then
  begin
    if not (mtTask.State in [dsEdit, dsInsert]) then mtTask.Edit;
    mtTaskWorkDir.AsString := s;
  end;
end;

procedure TFCtrlWnd.dbgTaskGetCellParams(Sender: TObject; Column: TColumnEh; AFont: TFont; var Background: TColor;
  State: TGridDrawState);
begin
  if (not mtTask.Active) or (mtTask.IsEmpty) then exit;

  if mtTaskChanged.AsBoolean then
  begin
    if mtTaskTaskId.IsNull then Background := RGB(255, 255, 190)
    else Background := RGB(255, 210, 255);
  end else
    if not mtTaskActive.AsBoolean then AFont.Color := clGray;
end;

procedure TFCtrlWnd.dbgTaskKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_INSERT: if Shift = [] then btnAddTaskClick(btnAddTask);
    VK_DELETE: if ssCtrl in Shift then btnDelTaskClick(btnDelTask);
  end;
end;

procedure TFCtrlWnd.dbgThreadGetCellParams(Sender: TObject; Column: TColumnEh; AFont: TFont; var Background: TColor;
  State: TGridDrawState);
begin
  if (not mtThread.Active) or (mtThread.IsEmpty) then exit;
  if mtThreadTerminated.AsBoolean then AFont.Color := clGray;
end;

procedure TFCtrlWnd.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caHide;
end;

procedure TFCtrlWnd.FormCreate(Sender: TObject);
begin
  pcMain.ActivePage := tsControl;
end;

procedure TFCtrlWnd.FormDestroy(Sender: TObject);
begin
 // SaveWndState;
end;

procedure TFCtrlWnd.FormShow(Sender: TObject);
begin
  Refresh;
  LoadTaskList;
  //LoadWndState;
end;

procedure TFCtrlWnd.LoadTaskList;
var
  i, id: integer;

begin
  TaskListLoading := true;
  id := -1;

  try
    mtTask.DisableControls;
    if mtTask.Active then
    begin
      if not mtTask.IsEmpty then id := mtTaskTaskId.AsInteger;
      mtTask.EmptyTable;
      //mtTask.Close;
    end;

    if not mtTask.Active then mtTask.CreateDataSet;

    for i := 0 to TaskList.Count - 1 do
    begin
      mtTask.Append;
      mtTaskTaskId.AsInteger := TaskList.Item[i].TaskId;
      mtTaskTaskName.AsString := Copy(TaskList.Item[i].TaskName, 1, 256);
      mtTaskExecCommand.AsString := Copy(TaskList.Item[i].ExecCommand, 1, 256);
      mtTaskParametres.AsString := Copy(TaskList.Item[i].Parametres, 1, 4097);
      mtTaskWorkDir.AsString := Copy(TaskList.Item[i].WorkDir, 1, 256);
      mtTaskShowWindow.AsInteger := Ord(TaskList.Item[i].ShowWindow);
      mtTaskMinute.AsString := Copy(TaskList.Item[i].ExecTime.Minute, 1, 256);
      mtTaskHour.AsString := Copy(TaskList.Item[i].ExecTime.Hour, 1, 256);
      mtTaskDay.AsString := Copy(TaskList.Item[i].ExecTime.Day, 1, 256);
      mtTaskMonth.AsString := Copy(TaskList.Item[i].ExecTime.Month, 1, 256);
      mtTaskWeekday.AsString := Copy(TaskList.Item[i].ExecTime.Weekday, 1, 256);
      mtTaskAllowMultiple.AsBoolean := TaskList.Item[i].AllowMultiple;
      mtTaskWaitProcess.AsBoolean := TaskList.Item[i].WaitProcess;
      mtTaskActive.AsBoolean := TaskList.Item[i].Active;
      mtTaskChanged.AsBoolean := false;
      mtTask.Post;
    end;
  finally
    if not mtTask.Locate('TaskId', id, []) then mtTask.First;

    TaskListLoading := false;
    TaskListChanged := false;
    mtTask.EnableControls;
    mtTaskAfterScroll(mtTask);
  end;
end;

procedure TFCtrlWnd.LoadWndState;
var
  f: TIniFile;
  mx: boolean;

begin
  try
    f := TIniFile.Create(ParamFolder + ParamFile);
    try
      Width := f.ReadInteger(PARAM_SECTION, 'Width', 1100);
      Height := f.ReadInteger(PARAM_SECTION, 'Height', 700);
      Left := f.ReadInteger(PARAM_SECTION, 'Left', Round(Screen.Width / 2) - Round(Width / 2));
      Top := f.ReadInteger(PARAM_SECTION, 'Top', Round(Screen.Height / 2) - Round(Height / 2));
      mx := f.ReadBool(PARAM_SECTION, 'Maximized', false);
      if mx then
        WindowState := wsMaximized
      else
        WindowState := wsNormal;
    finally
      f.Free;
    end;
  except
    Width := 1100;
    Height := 700;
    Left := Round(Screen.Width / 2) - Round(Width / 2);
    Top := Round(Screen.Height / 2) - Round(Height / 2);
    WindowState := wsNormal;
  end;
end;

procedure TFCtrlWnd.mtTaskAfterOpen(DataSet: TDataSet);
begin
  lbTaskStatus.Caption := '';
  lbTaskCopyCnt.Caption := '';
  lbLastTaskStart.Caption := '';
  lbLastTaskStop.Caption := '';
  lbLastMessage.Caption := '';
  btnAddTask.Enabled := mtTask.Active;
  btnDelTask.Enabled := mtTask.Active and (not mtTask.IsEmpty);
  btnStartTask.Enabled := mtTask.Active and (not mtTask.IsEmpty);
  btnBreakTask.Enabled := mtTask.Active and (not mtTask.IsEmpty);
end;

procedure TFCtrlWnd.mtTaskAfterScroll(DataSet: TDataSet);
var
  st, msg: string;
  cnt: integer;
  rdt, edt: TDateTime;

begin
  btnDelTask.Enabled := mtTask.Active and (not mtTask.IsEmpty);
  btnStartTask.Enabled := mtTask.Active and (not mtTask.IsEmpty);
  btnBreakTask.Enabled := mtTask.Active and (not mtTask.IsEmpty);

  TaskList.GetTaskInfo(mtTaskTaskId.AsInteger, st, msg, cnt, rdt, edt);

  lbTaskStatus.Caption := st;
  lbTaskCopyCnt.Caption := IntToStr(cnt);

  if rdt <= StrToDateTime('01.01.1990') then
    lbLastTaskStart.Caption := 'Неизвестно'
  else
    lbLastTaskStart.Caption := DateTimeToStr(rdt);

  if edt <= StrToDateTime('01.01.1990') then
    lbLastTaskStop.Caption := 'Неизвестно'
  else
    lbLastTaskStop.Caption := DateTimeToStr(edt);

  lbLastMessage.Caption := msg;
end;

procedure TFCtrlWnd.mtTaskBeforePost(DataSet: TDataSet);
begin
  if not TaskListLoading then
  begin
    mtTaskChanged.AsBoolean := true;
    TaskListChanged := true;
  end;
end;

procedure TFCtrlWnd.mtThreadAfterOpen(DataSet: TDataSet);
begin
  btnRefresh.Enabled := mtThread.Active;
  btnBreakThread.Enabled := mtThread.Active and (not mtThread.IsEmpty) and (not mtThreadTerminated.AsBoolean);
end;

procedure TFCtrlWnd.mtThreadAfterScroll(DataSet: TDataSet);
begin
  btnBreakThread.Enabled := mtThread.Active and (not mtThread.IsEmpty) and (not mtThreadTerminated.AsBoolean);
end;

procedure TFCtrlWnd.Refresh;
var
  rtCnt, rthCnt, sCnt: integer;
  
begin
  // основная инфа
  lbStartTime.Caption := DateTimeToStr(StartTime);
  lbWorkTime.Caption := IntToStr(DaysBetween(Now, StartTime)) + ' дн ' + FormatDateTime('hh ч nn мин', Now - StartTime);

  if TaskList.Started then
  begin
    lbStatus.Caption := 'Активен';
    lbStatus.Font.Color := clGreen;
    btnStart.Enabled := false;
    btnStop.Enabled := true;
  end else
  begin
    lbStatus.Caption := 'Приостановлен';
    lbStatus.Font.Color := clGray;
    btnStart.Enabled := true;
    btnStop.Enabled := false;
  end;

  lbTaskCount.Caption := IntToStr(TaskList.Count);
  TaskList.ExecutionStatus(rtCnt, rthCnt, sCnt);
  lbRunningCount.Caption := IntToStr(rtCnt) + ' (' + IntToStr(rthCnt) + ')';
  lbTerminatedThreads.Caption := IntToStr(sCnt);

  // автозапуск
  case CheckAutorun of
    asNone:
    begin
      btnAddToAutorun.Caption := 'Добавить в автозапуск';
      lbAutorun.Caption := 'Автозапуск отсутствует';
    end;
    asAllUsers:
    begin
      btnAddToAutorun.Caption := 'Убрать из автозапуска';
      lbAutorun.Caption := 'Автозапуск у всех пользователей';
    end;
    asCurrentUser:
    begin
      btnAddToAutorun.Caption := 'Убрать из автозапуска';
      lbAutorun.Caption := 'Автозапуск у текущего пользователя';
    end;
  end;

  // настройки
  edMaxThreadCount.Value := MaxTerminatedThread;
  
  // выполняющиеся потоки
  RefreshThreadList;
end;

procedure TFCtrlWnd.RefreshThreadList;
var
  i, idx: integer;
  
begin
  mtThread.DisableControls;
  idx := -1;

  try
    if mtThread.Active then
    begin
      if not mtThread.IsEmpty then idx := mtThreadIndex.AsInteger;
      mtThread.EmptyTable;
      //mtThread.Close;
    end;

    if not mtThread.Active then mtThread.CreateDataSet;

    for i := 0 to TaskList.ThreadCount - 1 do
    begin
      if Assigned(TaskList.Thread[i]) then
      begin
        mtThread.Append;
        mtThreadIndex.AsInteger := TaskList.Thread[i].Index;
        mtThreadTaskId.AsInteger := TaskList.Thread[i].Task.TaskId;
        mtThreadTaskName.AsString := Copy(TaskList.Thread[i].Task.TaskName, 1, 256);
        mtThreadExecCommand.AsString := Copy(TaskList.Thread[i].Task.ExecCommand, 1, 256);
        mtThreadParametres.AsString := Copy(TaskList.Thread[i].Task.Parametres, 1, 4097);
        mtThreadProcessId.AsInteger := TaskList.Thread[i].PrID;
        if TaskList.Thread[i].StartTime <= StrToDateTime('01.01.1990') then
          mtThreadStartTime.Clear
        else
          mtThreadStartTime.AsDateTime := TaskList.Thread[i].StartTime;
        if TaskList.Thread[i].StopTime <= StrToDateTime('01.01.1990') then
          mtThreadStopTime.Clear
        else
          mtThreadStopTime.AsDateTime := TaskList.Thread[i].StopTime;
        mtThreadTerminated.AsBoolean := TaskList.Thread[i].Terminated;
        mtThreadLastError.AsString := Copy(TaskList.Thread[i].LastError, 1, 1025);
        mtThread.Post;
      end;
    end;
  finally
    if not mtThread.Locate('Index', idx, []) then mtThread.First;
    mtThread.EnableControls;
    mtThreadAfterScroll(mtThread);
  end;
end;

procedure TFCtrlWnd.SaveTaskList;
var
  f: TIniFile;
  sl: TStringList;
  i: integer;
  
begin
  sl := TStringList.Create;
  f := TIniFile.Create(ParamFolder + ParamFile);
  try
    if mtTask.State in [dsEdit, dsInsert] then mtTask.Post;
    mtTask.DisableControls;

    f.WriteInteger(PARAM_SECTION, 'LastTaskId', LastTaskId);

    f.ReadSections(sl);
    for i := 0 to sl.Count - 1 do
      if Pos(TASK_SECTION_PREFIX, sl.Strings[i]) > 0 then f.EraseSection(sl.Strings[i]);

    mtTask.First;
    try
      while not mtTask.Eof do
      begin
        if mtTaskTaskId.IsNull then
          f.WriteInteger(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'TaskId', TaskList.GetNextTaskId)
        else
          f.WriteInteger(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'TaskId', mtTaskTaskId.AsInteger);
        f.WriteString(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'TaskName', mtTaskTaskName.AsString);
        f.WriteString(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'ExecCommand', mtTaskExecCommand.AsString);
        f.WriteString(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'Parametres', mtTaskParametres.AsString);
        f.WriteString(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'WorkDir', mtTaskWorkDir.AsString);
        f.WriteString(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'ShowWindow', TaskList.SWtoString(TShowWindow(mtTaskShowWindow.AsInteger)));
        f.WriteBool(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'AllowMultiple', mtTaskAllowMultiple.AsBoolean);
        f.WriteBool(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'WaitProcess', mtTaskWaitProcess.AsBoolean);
        f.WriteBool(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'Active', mtTaskActive.AsBoolean);
        f.WriteString(TASK_SECTION_PREFIX + IntToStr(mtTask.RecNo-1), 'ExecTime', mtTaskMinute.AsString + '|' + mtTaskHour.AsString + '|' +
          mtTaskDay.AsString + '|' + mtTaskMonth.AsString + '|' + mtTaskWeekday.AsString);
        mtTask.Next;
      end;
    except
      on e: Exception do
      begin
        MessageBox(FHandle, pchar('Не удалось сохранить список задач! Ошибка: ' + e.Message), pchar(AppTitle), MB_OK + MB_ICONERROR);
      end;
    end;
  finally
    mtTask.First;
    mtTask.EnableControls;
    sl.Free;
    f.Free;
  end;
end;

procedure TFCtrlWnd.SaveWndState;
var
  f: TIniFile;

begin
  f := TIniFile.Create(ParamFolder + ParamFile);

  try
    f.WriteBool(PARAM_SECTION, 'Maximized', WindowState = wsMaximized);
    if WindowState <> wsMaximized then
    begin
      f.WriteInteger(PARAM_SECTION, 'Left', Left);
      f.WriteInteger(PARAM_SECTION, 'Top', Top);
      f.WriteInteger(PARAM_SECTION, 'Width', Width);
      f.WriteInteger(PARAM_SECTION, 'Height', Height);
    end;
  finally
    f.Free;
  end;
end;

end.
