unit tasks;

interface

uses
  Windows, Messages, SysUtils, Classes, utils, IniFiles, DateUtils;

type
  TPartType = (ptMin, ptHour, ptDay, ptMonth, ptWeekday);
  TShowWindow = (swDefault, swShow, swHidden, swNormal, swMinimized, swMaximized);

  TExecTime = record
    Minute: string;
    Hour: string;
    Day: string;
    Month: string;
    Weekday: string;
  end;

  TTaskItem = class
  private
  public
    TaskId: integer;
    TaskName: string;
    ExecCommand: string;
    Parametres: string;
    WorkDir: string;
    //LogFile: string;
    ShowWindow: TShowWindow;
    ExecTime: TExecTime;
    AllowMultiple: boolean;
    WaitProcess: boolean;
    Active: boolean; 
    constructor Create;
  end;

  TTaskThread = class(TThread)
  private
    ProcessId: Cardinal;
    FTask: TTaskItem;
    FStartTime: TDateTime;
    FStopTime: TDateTime;
    Finalized: boolean;
    ForcedTerm: boolean;
    procedure SetTask(const Value: TTaskItem);
  protected
    procedure Execute; override;
  public
    Index: integer;
    LastError: string;
    WaitProcess: boolean;
    Success: boolean;
    destructor Destroy; override;
    function BreakAndTerminate: boolean;
    procedure OnFinalExecution;
    procedure SafeRefreshCtrlWnd;
    procedure ForcedTerminate;
    property Terminated;
    property Task: TTaskItem read FTask write SetTask;
    property StartTime: TDateTime read FStartTime;
    property StopTime: TDateTime read FStopTime;
    property PrID: Cardinal read ProcessId;
  end;
  
  TTaskList = class
  private
    Items: array of TTaskItem;
    FStarted: boolean;
    ThreadList: array of TTaskThread;
    FNow: TDateTime;
    procedure ClearTaskList;
    procedure ClearThreadList;
    function GetThreadIndex: integer;
    function IsRunning(TaskId: integer): boolean;
    function CheckTaskId(TaskId: integer): boolean;
    // проверка расписания
    function unpackRange(s: string): string;
    function isNumeric(s: string; var d: integer): boolean;
    function itIsTime(part: string; partType: TPartType): boolean;
    function MustRun(et: TExecTime): boolean;
    //
    function GetTaskItem(Index: integer): TTaskItem;
    function GetThread(Index: integer): TTaskThread;
    function GetTaskById(TaskId: integer): integer;
  public
    constructor Create;
    destructor Destroy; override;
    function GetNextTaskId: integer;
    function SWtoString(sw: TShowWindow): string;
    function SWfromString(sw: string): TShowWindow;
    procedure Start;
    procedure Stop;
    procedure LoadTasks;
    procedure SaveTasks;
    procedure SuspendThread(ThreadIdx: integer);
    procedure ResumeThread(ThreadIdx: integer);
    procedure BreakThread(ThreadIdx: integer);
    function StartTaskById(TaskId: integer; var err: string): integer;
    function StartTask(Idx: integer; var err: string): integer;
    procedure BreakTask(TaskId: integer; var err: string);
    procedure BreakAll;
    procedure ExecCurrent;
    function Count: integer;
    function ThreadCount: integer;
    procedure ExecutionStatus(var CntRunTask, CntRunThread, CntEndThread: integer);
    procedure GetTaskInfo(TaskId: integer; var tStatus, msg: string; var cntThread: integer; var LastRun, LastStop: TDateTime);
    property Started: boolean read FStarted;
    property Item[Index: integer]: TTaskItem read GetTaskItem;
    property Thread[Index: integer]: TTaskThread read GetThread;
  end;

var
  TaskList: TTaskList;
  
implementation

uses ctrlWnd;

{ TTaskItem }

constructor TTaskItem.Create;
begin
  inherited Create;
  TaskId := -1;
  AllowMultiple := false;
  WaitProcess := true;
  Active := true;
end;

{ TTaskList }

procedure TTaskList.BreakTask(TaskId: integer; var err: string);
var
  i: integer;
  
begin
  // остановка всех выполняющихся экземпляров задачи
  for i := 0 to Length(ThreadList) - 1 do
    if Assigned(ThreadList[i]) and (not ThreadList[i].Terminated) and (ThreadList[i].Task.TaskId = TaskId) then
    begin
      if not ThreadList[i].BreakAndTerminate then
      begin
        AddToWorkLog(ltError, 'Принудительное завершение задачи № ' + IntToStr(TaskId) + ': ' + ThreadList[i].Task.TaskName +
          '. Ошибка: ' + ThreadList[i].LastError);
        if err = '' then err := 'При завершении работы потока возникла ошибка: ' + ThreadList[i].LastError
        else err := err + #13#10 + 'При завершении работы потока возникла ошибка: ' + ThreadList[i].LastError;
      end else
        AddToWorkLog(ltWarning, 'Принудительное завершение задачи № ' + IntToStr(TaskId) + ': ' + ThreadList[i].Task.TaskName);

      //FreeAndNil(ThreadList[i]); - это сделает мой сборщик мусора
    end;

  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

procedure TTaskList.BreakThread(ThreadIdx: integer);
begin
  // остановка указанного птока (экземпляра задачи)
  if (ThreadIdx < 0) or (ThreadIdx >= Length(ThreadList)) then
    raise Exception.Create('Индекс потока вышел за границы списка потоков!');

  if (not Assigned(ThreadList[ThreadIdx])) or ThreadList[ThreadIdx].Terminated then
    raise Exception.Create('Поток уже завершил свое выполнение!');

  if not ThreadList[ThreadIdx].BreakAndTerminate then
  begin
    AddToWorkLog(ltError, 'Принудительное завершение потока ' + IntToStr(ThreadList[ThreadIdx].ProcessId) + ' задачи № ' +
      IntToStr(ThreadList[ThreadIdx].Task.TaskId) + ': ' + ThreadList[ThreadIdx].Task.TaskName + '. ' + ThreadList[ThreadIdx].LastError);
    raise Exception.Create('При завершении работы потока возникла ошибка: ' + ThreadList[ThreadIdx].LastError);
  end else
    AddToWorkLog(ltWarning, 'Принудительное завершение потока ' + IntToStr(ThreadList[ThreadIdx].ProcessId) + ' задачи № ' +
      IntToStr(ThreadList[ThreadIdx].Task.TaskId) + ': ' + ThreadList[ThreadIdx].Task.TaskName);

  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

procedure TTaskList.BreakAll;
var
  i: integer;

begin
  AddToWorkLog(ltWarning, 'Принудительное завершение всех задач!');

  for i := 0 to Length(ThreadList) - 1 do
    if Assigned(ThreadList[i]) and (not ThreadList[i].Terminated) then
    begin
      if not ThreadList[i].BreakAndTerminate then
        AddToWorkLog(ltError, 'Принудительное завершение задачи № ' + IntToStr(ThreadList[i].Task.TaskId) + ': ' +
          ThreadList[i].Task.TaskName + '. ' + ThreadList[i].LastError)
      else
        AddToWorkLog(ltWarning, 'Принудительное завершение задачи № ' + IntToStr(ThreadList[i].Task.TaskId) + ': ' +
          ThreadList[i].Task.TaskName);

      //FreeAndNil(ThreadList[i]); - это сделает мой сборщик мусора
    end;

  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

function TTaskList.CheckTaskId(TaskId: integer): boolean;
var
  i: integer;

begin
  // проверяет - свободен ли переданный id задачи
  result := true;

  for i := 0 to Length(Items) - 1 do
    if Assigned(Items[i]) and (Items[i].TaskId = TaskId) then
    begin
      result := false;
      break;
    end;
end;

procedure TTaskList.ClearTaskList;
var
  i: integer;

begin
  for i := 0 to Length(Items) - 1 do
    if Assigned(Items[i]) then FreeAndNil(Items[i]);

  SetLength(Items, 0);
end;

procedure TTaskList.ClearThreadList;
var
  i: integer;
  
begin
  for i := 0 to Length(ThreadList) - 1 do
    if Assigned(ThreadList[i]) then
    begin
      if not ThreadList[i].Terminated then ThreadList[i].ForcedTerminate;
      FreeAndNil(ThreadList[i]);
    end;

  SetLength(ThreadList, 0);
end;

constructor TTaskList.Create;
begin
  inherited Create;
  SetLength(Items, 0);
  SetLength(ThreadList, 0);
  LoadTasks;
end;

destructor TTaskList.Destroy;
begin
  //SaveTasks;
  ClearTaskList;
  ClearThreadList;
  inherited Destroy;
end;

procedure TTaskList.ExecCurrent;
var
  i: integer;
  s: string;
  
begin
  // процедура запуска заданий по расписанию, вызывается по таймеру раз в минуту
  if not FStarted then exit;
  FNow := Now;
  
  for i := 0 to Length(Items) - 1 do
    try
      if Items[i].Active and MustRun(Items[i].ExecTime) then StartTask(i, s);
    except
      on e: Exception do
      begin
        AddToWorkLog(ltError, 'Задача ' + IntToStr(Items[i].TaskId) + ': ' + Items[i].TaskName +
          ' - Синтаксическая ошибка в расписании или другая ошибка: ' + e.Message);
      end;
    end;
end;

function TTaskList.Count: integer;
begin
  result := Length(Items);
end;

procedure TTaskList.ExecutionStatus(var CntRunTask, CntRunThread, CntEndThread: integer);
var
  i, n: integer;
  sl: TStringList;
  
begin
  sl := TStringList.Create;

  CntRunTask := 0;
  CntRunThread := 0;
  CntEndThread := 0;
  
  try
    for i := 0 to Length(ThreadList) - 1 do
    begin
      if Assigned(ThreadList[i]) then
      begin
        if ThreadList[i].Terminated then Inc(CntEndThread)
        else begin
          Inc(CntRunThread);
          if not sl.Find(IntToStr(ThreadList[i].Task.TaskId), n) then
          begin
            Inc(CntRunTask);
            sl.Add(IntToStr(ThreadList[i].Task.TaskId));
          end;
        end;
      end;
    end;
  finally
    sl.Free;
  end;
end;

function TTaskList.GetNextTaskId: integer;
begin
  while not CheckTaskId(LastTaskId) do Inc(LastTaskId);
  result := LastTaskId;
end;

function TTaskList.GetTaskById(TaskId: integer): integer;
var
  i: integer;

begin
  result := -1;
  
  for i := 0 to Length(Items) - 1 do
    if Items[i].TaskId = TaskId then
    begin
      result := i;
      break;
    end;
end;

procedure TTaskList.GetTaskInfo(TaskId: integer; var tStatus, msg: string; var cntThread: integer; var LastRun, LastStop: TDateTime);
var
  i: integer;
  
begin
  tStatus := 'Ожидает';
  msg := '';
  cntThread := 0;
  
  for i := 0 to Length(ThreadList) - 1 do
    if Assigned(ThreadList[i]) and (ThreadList[i].Task.TaskId = TaskId) then
    begin
      msg := ThreadList[i].LastError;
      LastRun := ThreadList[i].StartTime;
      LastStop := ThreadList[i].StopTime;
      if (not ThreadList[i].Terminated) then
      begin
        Inc(cntThread);
        tStatus := 'Выполняется';
      end;
    end;
end;

function TTaskList.GetTaskItem(Index: integer): TTaskItem;
begin
  result := Items[Index];
end;

function TTaskList.GetThread(Index: integer): TTaskThread;
begin
  result := ThreadList[Index];
end;

function TTaskList.GetThreadIndex: integer;
var
  i: integer;
  
begin
  result := -1;

  // поищем свободную ячейку, и если ее нет, но есть завершенные потоки - удалим их (чистка мусора) 
  for i := 0 to Length(ThreadList) - 1 do
  begin
    if not Assigned(ThreadList[i]) then
    begin
      result := i;
      break;
    end;

    if ThreadList[i].Terminated and ((MaxTerminatedThread < 0) or
      (Length(ThreadList) > MaxTerminatedThread)) then
    begin
      FreeAndNil(ThreadList[i]);
      result := i;
      break;
    end;
  end;

  // массив или пустой или нет свободных ячеек - надо расширить
  if result = -1 then
  begin
    SetLength(ThreadList, Length(ThreadList) + 1);
    result := High(ThreadList);
  end;
end;

function TTaskList.itIsTime(part: string; partType: TPartType): boolean;

  function _check(d: integer; curr: word; min_, max_: integer; extend_: boolean; name_: string): boolean;
  begin
    if extend_ then
    begin
      if (d = 0) or (d > 31) then raise Exception.Create('День месяца вне диапазона: ' + IntToStr(d));
      if d > DaysInMonth(FNow) then d := DaysInMonth(FNow);
      if d < 0 then d := DaysInMonth(FNow) - d;
      if d <= 0 then raise Exception.Create('День месяца вне диапазона: ' + IntToStr(d));
    end else
    begin
      if (d < min_) or (d > max_) then raise Exception.Create(name_ + ' вне диапазона: ' + IntToStr(d));
    end;
    result := d = curr;
  end;

  function _checkRange(d: integer; curr: word; min_, max_: integer; extend_: boolean; name_: string): boolean;
  begin
    if extend_ then
    begin
      if (d = 0) or (d > 31) then raise Exception.Create('Интервал дней месяца вне диапазона: ' + IntToStr(d));
      if d > DaysInMonth(FNow) then d := DaysInMonth(FNow);
      if d < 0 then d := DaysInMonth(FNow) - d;
      if d <= 0 then raise Exception.Create('Интервал дней месяца вне диапазона: ' + IntToStr(d));
    end else
    begin
      if (d < min_) or (d > max_) then raise Exception.Create(name_ + ' вне диапазона (интервал): ' + IntToStr(d));
    end;
    result := (curr mod d) = 0;
  end;

var
  d: integer;
  ms, cSec, cMin, cHour, cDay, cMonth, cYear, cWD, currElem: word;
  sSec, sMin, sHour, sDay, sMonth, sYear, sWD: word;
  n: integer;
  s, s1, name_: string;
  min_, max_: integer;

begin
  // парсим расписание - проверка заданной части даты
  // текущие значения
  DecodeDateTime(FNow, cYear, cMonth, cDay, cHour, cMin, cSec, ms);
  cWD := DayOfTheWeek(FNow);
  // вычисляем время старта для замены знака ? на соотв. часть времени старта
  DecodeDateTime(StartTime, sYear, sMonth, sDay, sHour, sMin, sSec, ms);
  sWD := DayOfTheWeek(StartTime);

  case partType of
    ptMin:
    begin
      currElem := cMin;
      min_ := 0;
      max_ := 59;
      name_ := 'Минуты';
      part := StringReplace(part, '?', IntToStr(sMin), [rfReplaceAll]);
    end;
    ptHour:
    begin
      currElem := cHour;
      min_ := 0;
      max_ := 23;
      name_ := 'Часы';
      part := StringReplace(part, '?', IntToStr(sHour), [rfReplaceAll]);
    end;
    ptDay:
    begin
      currElem := cDay;
      min_ := 1;
      max_ := 31;
      name_ := 'День';
      part := StringReplace(part, '?', IntToStr(sDay), [rfReplaceAll]);
    end;
    ptMonth:
    begin
      currElem := cMonth;
      min_ := 1;
      max_ := 12;
      name_ := 'Месяц';
      part := StringReplace(part, '?', IntToStr(sMonth), [rfReplaceAll]);
    end;
    ptWeekday:
    begin
      currElem := cWD;
      min_ := 1;
      max_ := 7;
      name_ := 'День недели';
      part := StringReplace(part, '?', IntToStr(sWD), [rfReplaceAll]);
    end;
  end;

  // может каждый?
  result := (part = '') or (part = '*') or (part = '/1');
  if not result then
  begin
    // указано конкретное число?
    if isNumeric(part, d) then
    begin
      // конкретное число - сверимся с текущим
      result := _check(d, currElem, min_, max_, partType = ptDay, name_);
    end else
    begin
      // или список значеий через , или ; (в т.ч. элементом списка м.б. интервал /N)
      // разберемся с диапазонами
      for n := 1 to WordCountEx(part, [',', ';'], []) do
      begin
        s := Trim(ExtractWordEx(n, part, [',', ';'], []));
        s := unpackRange(s);
        if s1 = '' then s1 := s
        else s1 := s1 + ',' + s;
      end;
      part := s1;

      for n := 1 to WordCountEx(part, [',', ';'], []) do
      begin
        s := Trim(ExtractWordEx(n, part, [',', ';'], []));
        if isNumeric(s, d) then
        begin
          result := _check(d, currElem, min_, max_, partType = ptDay, name_);
          if result then break;
        end else
        begin
          // значит указан интервал - первый символ слэш (/), его обрезаем
          s := Copy(s, 2, Length(s));
          if not isNumeric(s, d) then raise Exception.Create('Неверное значение в поле ' + name_ + ': ' + s);
          result := _checkRange(d, currElem, min_, max_, partType = ptDay, name_);
          if result then break;
        end;
      end;
    end;
  end;
end;

procedure TTaskList.LoadTasks;
var
  f: TIniFile;
  sl: TStringList;
  i, tid: integer;
  s: string;

begin
  sl := TStringList.Create;
  f := TIniFile.Create(ParamFolder + ParamFile);

  try
    ClearTaskList;
    f.ReadSections(sl);

    for i := 0 to sl.Count - 1 do
    begin
      if Pos(TASK_SECTION_PREFIX, sl.Strings[i]) > 0 then
      begin
        tid := f.ReadInteger(sl.Strings[i], 'TaskId', GetNextTaskId);
        if not CheckTaskId(tid) then tid := GetNextTaskId;
        SetLength(Items, Length(Items) + 1);
        Items[High(Items)] := TTaskItem.Create;
        Items[High(Items)].TaskId := tid;
        Items[High(Items)].TaskName := f.ReadString(sl.Strings[i], 'TaskName', '');
        Items[High(Items)].ExecCommand := f.ReadString(sl.Strings[i], 'ExecCommand', '');
        Items[High(Items)].Parametres := f.ReadString(sl.Strings[i], 'Parametres', '');
        Items[High(Items)].WorkDir := f.ReadString(sl.Strings[i], 'WorkDir', '');
        //Items[High(Items)].LogFile := f.ReadString(sl.Strings[i], 'LogFile', '');
        Items[High(Items)].ShowWindow := SWfromString(f.ReadString(sl.Strings[i], 'ShowWindow', 'Default'));
        Items[High(Items)].AllowMultiple := f.ReadBool(sl.Strings[i], 'AllowMultiple', false);
        Items[High(Items)].WaitProcess := f.ReadBool(sl.Strings[i], 'WaitProcess', true);
        Items[High(Items)].Active := f.ReadBool(sl.Strings[i], 'Active', true);
        s := f.ReadString(sl.Strings[i], 'ExecTime', '');
        Items[High(Items)].ExecTime.Minute := Trim(ExtractWordEx(1, s, ['|'], []));
        Items[High(Items)].ExecTime.Hour := Trim(ExtractWordEx(2, s, ['|'], []));
        Items[High(Items)].ExecTime.Day := Trim(ExtractWordEx(3, s, ['|'], []));
        Items[High(Items)].ExecTime.Month := Trim(ExtractWordEx(4, s, ['|'], []));
        Items[High(Items)].ExecTime.Weekday := Trim(ExtractWordEx(5, s, ['|'], []));
      end;
    end;
  except
    on e: Exception do
    begin
      AddToWorkLog(ltError, 'Не удалось загрузить список задач! Ошибка: ' + e.Message);
    end;
  end;

  sl.Free;
  f.Free;
  AddToWorkLog(ltInfo, 'Загружен список задач');

  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

function TTaskList.MustRun(et: TExecTime): boolean;
begin
  // парсим и проверяем расписание
  result := itIsTime(et.Day, ptDay) and itIsTime(et.Month, ptMonth) and itIsTime(et.Weekday, ptWeekday) and
    itIsTime(et.Hour, ptHour) and itIsTime(et.Minute, ptMin);
end;

procedure TTaskList.ResumeThread(ThreadIdx: integer);
begin
  // на самом деле идея суспендить процессы не очень хорошая, поэтому пока я точно не буду значть - зачем это нужно, делать не буду

{  if (ThreadIdx < 0) or (ThreadIdx >= Length(ThreadList)) then
    raise Exception.Create('Индекс потока вышел за границы списка потоков!');

  if (not Assigned(ThreadList[ThreadIdx])) or ThreadList[ThreadIdx].Terminated then
    raise Exception.Create('Поток уже завершил свое выполнение!');

  // тут еще должен быть вызов ResumeProcess, но он пока не реализован, а суспендить сам контролирующий поток возможно вобще не надо
  ThreadList[ThreadIdx].Resume;}
end;

procedure TTaskList.SaveTasks;
var
  f: TIniFile;
  i: integer;

begin
  f := TIniFile.Create(ParamFolder + ParamFile);

  try
    f.WriteInteger(PARAM_SECTION, 'LastTaskId', LastTaskId);

    for i := 0 to Length(Items) - 1 do
    begin
      f.WriteInteger(TASK_SECTION_PREFIX + IntToStr(i), 'TaskId', Items[i].TaskId);
      f.WriteString(TASK_SECTION_PREFIX + IntToStr(i), 'TaskName', Items[i].TaskName);
      f.WriteString(TASK_SECTION_PREFIX + IntToStr(i), 'ExecCommand', Items[i].ExecCommand);
      f.WriteString(TASK_SECTION_PREFIX + IntToStr(i), 'Parametres', Items[i].Parametres);
      f.WriteString(TASK_SECTION_PREFIX + IntToStr(i), 'WorkDir', Items[i].WorkDir);
      //f.WriteString(TASK_SECTION_PREFIX + IntToStr(i), 'LogFile', Items[i].LogFile);
      f.WriteString(TASK_SECTION_PREFIX + IntToStr(i), 'ShowWindow', SWtoString(Items[i].ShowWindow));
      f.WriteBool(TASK_SECTION_PREFIX + IntToStr(i), 'AllowMultiple', Items[i].AllowMultiple);
      f.WriteBool(TASK_SECTION_PREFIX + IntToStr(i), 'WaitProcess', Items[i].WaitProcess);
      f.WriteBool(TASK_SECTION_PREFIX + IntToStr(i), 'Active', Items[i].Active);
      f.WriteString(TASK_SECTION_PREFIX + IntToStr(i), 'ExecTime', Items[i].ExecTime.Minute + '|' + Items[i].ExecTime.Hour + '|' +
        Items[i].ExecTime.Day + '|' + Items[i].ExecTime.Month + '|' + Items[i].ExecTime.Weekday);
    end;
  except
    on e: Exception do
    begin
      AddToWorkLog(ltError, 'Не удалось сохранить список задач! Ошибка: ' + e.Message);
    end;
  end;

  f.Free;
end;

procedure TTaskList.Start;
begin
  AddToWorkLog(ltInfo, 'Включена обработка расписания');
  FStarted := true;
  ChangeStatus(pmStart);
  RestartTimer(idTaskTimer);
  ExecCurrent;
end;

function TTaskList.StartTask(Idx: integer; var err: string): integer;
begin
  result := -1;
  if (Idx < 0) or (Idx >= Length(Items)) then exit;

  if IsRunning(Items[Idx].TaskId) and (not Items[Idx].AllowMultiple) then
  begin
    err := 'У задачи стоит запрет на запуск нескольких копий и предыдущий запуск еще не завершен';
    AddToWorkLog(ltWarning, 'Задача ' + IntToStr(Items[Idx].TaskId) + ': ' + Items[Idx].TaskName +
      ' - пропуск! Предыдущий запуск задачи еще не завершен');
    exit;
  end;

  AddToWorkLog(ltInfo, 'Задача ' + IntToStr(Items[Idx].TaskId) + ': ' + Items[Idx].TaskName + ' - запуск');
  result := GetThreadIndex;

  ThreadList[result] := TTaskThread.Create(true);
  ThreadList[result].Index := result;
  ThreadList[result].Task := Items[Idx];
  //ThreadList[result].WaitProcess := true; - это задается при привязке параметра Task
  ThreadList[result].Suspended := false;

  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

function TTaskList.StartTaskById(TaskId: integer; var err: string): integer;
var
  idx: integer;

begin
  result := -1;
  idx := GetTaskById(TaskId);

  if idx < 0 then
  begin
    err := 'Задачи с ID ' + IntToStr(TaskId) + ' не существует!';
    exit;
  end;

  result := StartTask(idx, err);
end;

procedure TTaskList.Stop;
begin
  AddToWorkLog(ltInfo, 'Приостановлен (обработка расписания не выполняется)');
  FStarted := false;
  ChangeStatus(pmStop);
  RestartTimer(idTaskTimer);
end;

procedure TTaskList.SuspendThread(ThreadIdx: integer);
begin
  // на самом деле идея суспендить процессы не очень хорошая, поэтому пока я точно не буду значть - зачем это нужно, делать не буду

{  if (ThreadIdx < 0) or (ThreadIdx >= Length(ThreadList)) then
    raise Exception.Create('Индекс потока вышел за границы списка потоков!');

  if (not Assigned(ThreadList[ThreadIdx])) or ThreadList[ThreadIdx].Terminated then
    raise Exception.Create('Поток уже завершил свое выполнение!');

  // тут еще должен быть вызов SuspendProcess, но он пока не реализован, а суспендить сам контролирующий поток возможно вобще не надо
  ThreadList[ThreadIdx].Suspend; }
end;

function TTaskList.SWfromString(sw: string): TShowWindow;
begin
  sw := LowerCase(sw);
  if sw = 'default' then result := swDefault
  else if sw = 'show' then result := swShow
  else if sw = 'hidden' then result := swHidden
  else if sw = 'normal' then result := swNormal
  else if sw = 'minimized' then result := swMinimized
  else if sw = 'maximized' then result := swMaximized
  else result := swDefault;
end;

function TTaskList.SWtoString(sw: TShowWindow): string;
begin
  case sw of
    swDefault: result := 'Default';
    swShow: result := 'Show';
    swHidden: result := 'Hidden';
    swNormal: result := 'Normal';
    swMinimized: result := 'Minimized';
    swMaximized: result := 'Maximized';
    else result := 'Default';
  end;
end;

function TTaskList.ThreadCount: integer;
begin
  result := Length(ThreadList);
end;

function TTaskList.unpackRange(s: string): string;
var
  b, e: string;
  i, m, n: integer;

begin
  if Pos('-', s) <= 1 then
  begin
    result := s;
    exit;
  end;

  b := ExtractWordEx(1, s, ['-'], []);
  e := ExtractWordEx(2, s, ['-'], []);

  if (not isNumeric(b, m)) or (not isNumeric(e, n)) then
    raise Exception.Create('Неверно задан диапазон ' + s);

  for i := m to n do
    if result = '' then result := IntToStr(i)
    else result := result + ',' + IntToStr(i);
end;

function TTaskList.isNumeric(s: string; var d: integer): boolean;
begin
  try
    d := StrToInt(s);
    result := true;
  except
    result := false;
  end;
end;

function TTaskList.IsRunning(TaskId: integer): boolean;
var
  i: integer;

begin
  result := false;

  for i := 0 to Length(ThreadList) - 1 do
  begin
    if Assigned(ThreadList[i]) and (not ThreadList[i].Terminated) and (ThreadList[i].Task.TaskId = TaskId) then
    begin
      result := true;
      exit;
    end;
  end;
end;

{ TTaskThread }

function TTaskThread.BreakAndTerminate: boolean;
begin
  Success := true; // чтоб не писал в лог, что ошибка запуска
  result := KillProcess(ProcessId, LastError);
  Terminate;
end;

destructor TTaskThread.Destroy;
begin
  FTask.Free;
  inherited Destroy;
end;

procedure TTaskThread.Execute;
var
  SI: TStartupInfo;
  PrI: TProcessInformation;
  // некоторые параметры процесса
  CurrDir: pchar;
  ShowWnd: integer;
  w: Cardinal;

begin
  LastError := '';
  Success := false;
  FStartTime := Now;
  
  try
    if FTask.WorkDir = '' then CurrDir := nil
    else CurrDir := pchar(FTask.WorkDir);

    case FTask.ShowWindow of
      swDefault: ShowWnd := SW_SHOWDEFAULT;
      swShow: ShowWnd := SW_SHOW;
      swHidden: ShowWnd := SW_HIDE;
      swNormal: ShowWnd := SW_SHOWNORMAL;
      swMinimized: ShowWnd := SW_SHOWMINIMIZED;
      swMaximized: ShowWnd := SW_SHOWMAXIMIZED;
      else ShowWnd := SW_SHOWDEFAULT;
    end;

    GetStartupInfo(SI);
    SI.wShowWindow := ShowWnd;
    SI.dwFlags := STARTF_USESHOWWINDOW;
    if CreateProcess(nil, PChar('"' + FTask.ExecCommand + '"' + ' ' + FTask.Parametres), nil, nil, false,
      NORMAL_PRIORITY_CLASS, nil, CurrDir, SI, PrI) then
    begin
      ProcessId := PrI.dwProcessId;
      Synchronize(SafeRefreshCtrlWnd);
      if WaitProcess then
        while (not ForcedTerm) and (not Terminated) do
        begin
          w := WaitForSingleObject(PrI.hProcess, 1000);
          if (w = WAIT_OBJECT_0) or (w = WAIT_ABANDONED) or (w = WAIT_FAILED) then break;
        end;
      CloseHandle(PrI.hProcess);
      CloseHandle(PrI.hThread);
      Success := true;
    end else
      LastError := SysErrorMessage(GetLastError);
  finally
    OnFinalExecution;
  end;
end;

procedure TTaskThread.ForcedTerminate;
begin
  ForcedTerm := true;
  Terminate;
end;

procedure TTaskThread.OnFinalExecution;
begin
  if ForcedTerm then exit;
  if Finalized then exit;

  Finalized := true;
  FStopTime := Now;
  Terminate;
  if not Success then
    AddToWorkLog(ltError, 'Задача ' + IntToStr(Task.TaskId) + ': ' + Task.TaskName + ' - ошбка запуска: ' + LastError)
  else
    if WaitProcess then
      AddToWorkLog(ltInfo, 'Задача ' + IntToStr(Task.TaskId) + ': ' + Task.TaskName + ' - выполнение завершено');

  Synchronize(SafeRefreshCtrlWnd);
end;

procedure TTaskThread.SafeRefreshCtrlWnd;
begin
  // вызывать только из метода Synchronize!
  if ForcedTerm then exit;
  if Assigned(FCtrlWnd) and FCtrlWnd.Visible then FCtrlWnd.Refresh;
end;

procedure TTaskThread.SetTask(const Value: TTaskItem);
begin
  if not Assigned(FTask) then FTask := TTaskItem.Create;
  FTask.TaskId := Value.TaskId;
  FTask.TaskName := Value.TaskName;
  FTask.ExecCommand := Value.ExecCommand;
  FTask.Parametres := Value.Parametres;
  FTask.WorkDir := Value.WorkDir;
  //FTask.LogFile := Value.LogFile;
  FTask.ShowWindow := Value.ShowWindow;
  FTask.ExecTime := Value.ExecTime;
  FTask.AllowMultiple := Value.AllowMultiple;
  FTask.WaitProcess := Value.WaitProcess;
  FTask.Active := Value.Active;
  WaitProcess := FTask.WaitProcess;
end;

end.
