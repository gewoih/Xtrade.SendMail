unit uxLogWriter;

interface

uses SysUtils, Windows;

type
  txDebugLevel = (dlAlways, dlError, dlValue, dlCommon, dlDebug);
  txDebugLevels = set of txDebugLevel;

var
  fDebugLevels: txDebugLevels = [dlAlways];

procedure Debug(S:string;E:Exception; Level: txDebugLevel = dlCommon); overload;
procedure Debug(S:string;V:integer; Level: txDebugLevel = dlCommon); overload;
procedure Debug(S:string;V:string; Level: txDebugLevel = dlCommon); overload;
procedure Debug(S:string;V:TDateTime; Level: txDebugLevel = dlCommon); overload;
procedure WriteLog(S:string; Level: txDebugLevel = dlCommon);

type
  tMapRec = record
    DevID: integer;
    DevName: ShortString;
    LastActive: TDateTime;
    Flag: byte;
  end;
  tMapRecs = array[0..2] of tMapRec;
  pMapRecs = ^tMapRecs;

var
  StartMode : byte;
  fMapData: pMapRecs = nil;
  FExePath: string;

  FLoopResult: boolean;

threadvar
  SocketAddr: string;

implementation

uses
  Classes,
  SyncObjs;

var
  FLog: Text;
  FLogPath: string = '';
  FOpened: boolean = False;
  fLast: String;
  fLock: TCriticalSection;

procedure Debug(S:string;E:Exception; Level: txDebugLevel = dlCommon); overload;
begin
  WriteLog(S + ': исключение <' + E.Message + '>', Level);
end;

procedure Debug(S:string;V:integer; Level: txDebugLevel = dlCommon);
begin
  if S<>'' then S := S + ' : I = ';
  WriteLog(S + IntToStr(V), Level);
end;

procedure Debug(S:string;V:string; Level: txDebugLevel = dlCommon);
begin
  if S<>'' then S := S + ' : S = ';
  WriteLog(S + V, Level);
end;

procedure Debug(S:string;V:TDateTime; Level: txDebugLevel = dlCommon);
begin
  if S<>'' then S := S + ' : D = ';
  WriteLog(S + FormatDateTime('YYYY.MM.DD HH:NN:SS.ZZZ',V), Level);
end;

procedure WriteLog(S:string; Level: txDebugLevel = dlCommon);
var Sp: string;
begin
//  if not (Level in fDebugLevels) then Exit;
  fLock.Enter;
  try
    if S = fLast then Exit;
    fLast := S;
    Sp := FExePath + 'log\' +FormatDateTime('yyyy\mm\dd', Now)+'.txt';
    if Sp<>FLogPath then
    begin
      if FLogPath<>'' then CloseFile(FLog);
      ForceDirectories(ExtractFilePath(Sp));
      FLogPath := Sp;
      AssignFile(FLog,FLogPath);
      if FileExists(FLogPath) then Append(FLog) else Rewrite(FLog);
      FOpened := True;
    end;
    Sp := FormatDateTime('YYYY.MM.DD HH:NN:SS.ZZZ',Now);
    Writeln(FLog, Sp,' | ', SocketAddr, ' | ', S);
    Flush(FLog);
    if (IsConsole) and (StartMode>2) then
    begin
      AnsiToOem(@S[1],@S[1]);
      Writeln(Sp,' | ', SocketAddr, ' | ', S);
    end;
  finally
    fLock.Release;
  end;
end;

procedure GetLogSettings;
begin
  Include(fDebugLevels, dlDebug);
  Include(fDebugLevels, dlError);
  Include(fDebugLevels, dlValue);
  Include(fDebugLevels, dlCommon);
  Include(fDebugLevels, dlAlways);
end;

initialization
   FExePath := ExtractFilePath(ParamStr(0));
   fLock := TCriticalSection.Create;
finalization
  if FOpened then CloseFile(FLog);
  fLock.Free;
end.
