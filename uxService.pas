unit uxService;

interface

uses
  Windows, WinSVC, SysUtils;

type
  tSvcProc = function: boolean;

var
  SvcCount: cardinal = 1;
  WaitAfterReady : boolean = True;
  fLoopDelay: integer = 1000;
  {$I ServiceNames.inc}
  StopEvent: THandle;


procedure PrepareProcessParams(fStart,fLoop,fStop: tSvcProc);

implementation

uses
  uxLogWriter;

var
  SvcStart, SvcLoop, SvcStop: tSvcProc;
  SvcThread: THandle;
  SCManager: THandle;
  Handle: THandle;
  Status: SERVICE_STATUS;
  SvcParams: String;
  ServTableEntry: SERVICE_TABLE_ENTRYW;

function ChangeSvcDescription(hService:THandle; dwInfoLevel: cardinal;lpInfo: Pointer):Bool; stdcall; external advapi32 name 'ChangeServiceConfig2W';

function ConnectSCM:boolean;
begin
  Debug('SCM connect','started', dlCommon);
  SCManager := OpenSCManager(Nil, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS);
  Result := (SCManager<>0);
  If not Result then Debug('Error',SysErrorMessage(GetLastError));
end;

function CreateSvc:boolean;
begin
  Debug('Creating service database record', 'started', dlCommon);
  Handle := CreateService(SCManager, @SvcName[1], @DispName[1], SERVICE_QUERY_CONFIG or SERVICE_CHANGE_CONFIG,
    SERVICE_WIN32_OWN_PROCESS, SERVICE_AUTO_START, SERVICE_ERROR_NORMAL, PChar(ParamStr(0)+' '+SvcParams), Nil, Nil, Nil, Nil, Nil);
  Result := (Handle<>0);
  If not Result then Debug('Error',SysErrorMessage(GetLastError));
end;

function OpenSvc:boolean;
begin
  Debug('Open service database record', 'started', dlCommon);
  Handle := OpenService(SCManager,pChar(SvcName),STANDARD_RIGHTS_REQUIRED or SERVICE_QUERY_STATUS or SERVICE_STOP);
  Result := (Handle<>0);
  If not Result then Debug('Error',SysErrorMessage(GetLastError));
end;

function QuerySvcStatus: boolean;
begin
  Debug('QueryServiceStatus', 'started', dlCommon);
  Result := QueryServiceStatus(Handle, Status);
  If not Result then Debug('Error',SysErrorMessage(GetLastError));
end;

function DeleteSvc: boolean;
begin
  Debug('Deleting Service', 'started' , dlCommon);
  Result := DeleteService(Handle);
  If not Result then Debug('Error',SysErrorMessage(GetLastError));
end;

function SetSvcDescr:boolean;
begin
  Debug('Setting service description', 'started', dlCommon);
  Result := ChangeSvcDescription(Handle, 1, @SvcDescr);
  If not Result then raise Exception.Create(SysErrorMessage(GetLastError));
end;

procedure ProcessInstall;
begin
  try
    Debug('Service install', 'start', dlCommon);
    if not ConnectSCM then Exit;
    if not CreateSvc then Exit;
    if not SetSvcDescr then Exit;
    CloseServiceHandle(Handle);
    CloseServiceHandle(SCManager);
    Debug('Service install', 'success', dlCommon);
  finally
    if WaitAfterReady then Readln;
  end;
end;

procedure ProcessUninstall;
begin
  Debug('Service uninstall', 'start', dlCommon);
  if not ConnectSCM then Exit;
  if not OpenSvc then Exit;
  if not QuerySvcStatus then Exit;
  if Status.dwCurrentState<>SERVICE_STOPPED then
  begin
    ControlService(Handle, SERVICE_CONTROL_STOP, Status);
    while Status.dwCurrentState<>SERVICE_STOPPED do
    begin
      Sleep(250);
      QuerySvcStatus;
    end;
  end;
  if not DeleteSvc then Exit;
  CloseServiceHandle(Handle);
  CloseServiceHandle(SCManager);
  Debug('Service uninstall', 'success', dlCommon);
  if WaitAfterReady then Readln;
end;

function SetState(aState: DWORD): DWORD;
begin
  Status.dwCurrentState := aState;
  If (Handle<>0) then SetServiceStatus(Handle, Status);
  Result := Status.dwCurrentState;
end;

procedure ServiceHandler(fControl: DWORD); stdcall;
  begin
    Case fControl of
      SERVICE_CONTROL_STOP:
        begin
          Debug('Stop', '', dlCommon);
          SetState(SERVICE_STOP_PENDING);
          SetEvent(StopEvent);
          ResumeThread(SvcThread);
        end;
      SERVICE_CONTROL_PAUSE:
        begin
          SetState(SERVICE_PAUSE_PENDING);
          SuspendThread(SvcThread);
          SetState(SERVICE_PAUSED);
        end;
      SERVICE_CONTROL_CONTINUE:
        begin
          SetState(SERVICE_CONTINUE_PENDING);
          ResumeThread(SvcThread);
          SetState(SERVICE_RUNNING);
        end;
      SERVICE_CONTROL_INTERROGATE:
        begin
          SetState(Status.dwCurrentState);
        end;
    else
      SuspendThread(SvcThread);
      Status.dwWin32ExitCode := ERROR_SUCCESS;
      SetState(Status.dwCurrentState);
      ResumeThread(SvcThread);
    end;
  end;

procedure MainLoop;
begin
  If SvcStart then
  repeat
    SvcLoop;
  until WaitForSingleObject(StopEvent, fLoopDelay) = WAIT_OBJECT_0;
  SvcStop;
end;

procedure MainSvcProc(dwArgc: DWORD;lpszArgv: Pointer); stdcall;
begin
  Debug('SvcName', SvcName, dlCommon);
  Handle := RegisterServiceCtrlHandler(PWideChar(SvcName),@ServiceHandler);
  If Handle=0 then
  begin
    ExitCode := GetLastError;
    Exit;
  end;
  ZeroMemory(@Status, SizeOf(Status));
  Status.dwServiceType := SERVICE_WIN32_OWN_PROCESS;
  Status.dwControlsAccepted := SERVICE_ACCEPT_STOP or SERVICE_ACCEPT_PAUSE_CONTINUE;
  SetState(SERVICE_START_PENDING);
  If not DuplicateHandle(GetCurrentProcess,GetCurrentThread,GetCurrentProcess,@SvcThread, 0, FALSE, DUPLICATE_SAME_ACCESS) then
  begin
    Status.dwWin32ExitCode := GetLastError;
    SetState(SERVICE_STOPPED);
    Exit;
  end;
  SetState(SERVICE_RUNNING);
  ResetEvent(StopEvent);
  MainLoop;
  CloseHandle(SvcThread);
  SetState(SERVICE_STOPPED);
end;

procedure RunAsConsole;
begin
  MainLoop;
end;

procedure RunAsService;
begin
  ZeroMemory(@ServTableEntry, SizeOf(ServTableEntry));
  ServTableEntry.lpServiceName := PChar(SvcName);
  ServTableEntry.lpServiceProc := @MainSvcProc;
  If StartServiceCtrlDispatcher(ServTableEntry) then
  begin
    raise Exception.Create('StartServiceCtrlDispatcher');
    Exit;
  end;
  ExitCode := GetLastError;
end;

procedure SetSvcNames(Suff: string);
begin
  SvcName := SvcShortName + '$' + Suff;
  DispName := SvcShortName + '.' + Suff;
end;

procedure PrepareProcessParams(fStart,fLoop,fStop: tSvcProc);
var I:integer;S:AnsiString;
begin
  SvcStart := fStart;
  SvcLoop := fLoop;
  SvcStop := fStop;
  SvcParams := '';
  StartMode:=0;
  WaitAfterReady := False;
  For I:=1 to ParamCount do
  begin
    S:=ParamStr(I);
    If (S[1] in ['-','/']) then
      Case S[2] of
        'i','I': StartMode:=1;
        'u','U': StartMode:=2;
        'e','E': StartMode:=3;
        'N': SetSvcNames(Copy(S,4,MaxInt));
        'T': fLoopDelay := StrToIntDef(Copy(S,4,MaxInt), 10000);
        'W': WaitAfterReady := True;
        'v','V': StartMode:=4;
      else
        Debug('Неверный параметр командной строки ', S, dlCommon);
        Exit;
      end;
    if I>1 then SvcParams := SvcParams + ' ' + S;
  end;
  Debug('SvcParams', SvcParams, dlCommon);
  Debug('StartMode', StartMode, dlCommon);
  Case StartMode of
    0:RunAsService;
    1:ProcessInstall;
    2:ProcessUninstall;
  else
    if StartMode=4 then Readln else RunAsConsole;
  end;
end;

initialization
  StopEvent := CreateEvent(nil, true, false, nil);
finalization
  CloseHandle(StopEvent);
end.



