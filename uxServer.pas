unit uxServer;

interface

uses
	SysUtils,
	Winsock,
	Windows,
	Messages,
	ComObj,
	ActiveX,
	Classes,
	AnsiStrings, 
	IniFiles,
	DateUtils,
	Variants,
	
	smtpsend,
	mimemess, 
	mimepart, 
	blcksock,
	synautil, 
	synachar, 
	SynaCode,
	WideStrUtils;

var
	Mail: TSMTPSend;
	Mime: TMimemess;
	Params: TIniFile;


function SvcStart: boolean;
function SvcLoop: boolean;
function SvcStop: boolean;

implementation

uses
	uxLogWriter,
	uxService;

var fCon: OleVariant;	

function GetConStr: string;
begin
    Result := 'Provider='+params.ReadString('SQL', 'Prov', '') + ';'; //SQLNCLI11;';
    Result := Result + 'Persist Security Info=False;';
    Result := Result + 'Data Source='+params.ReadString('SQL', 'Serv', '') + ';';//192.168.44.100;';
    Result := Result + 'Initial Catalog='+params.ReadString('SQL', 'Base', '') + ';';//vtk;';
    Result := Result + 'User ID='+params.ReadString('SQL', 'User', '') + ';';//sa;';
    Result := Result + 'Application Name=' + ExtractFileName(ParamStr(0))+ ';';
    Result := Result + 'MultipleActiveResultSets=True;';
    Result := Result + 'Password='+params.ReadString('SQL', 'Pass', '') + ';';//icq99802122;'
end;

function ConnectSQL(Var Con:OleVariant): boolean;
begin
    try
        Con := CreateOleObject('ADODB.Connection');
        Con.CursorLocation:= 3;
        Con.CommandTimeout := 60000;
        Con.ConnectionTimeout := 10;
        Con.Open(GetConStr);
        Con.Execute('set nocount on');
        Con.Execute(Format('select %d as userid into #tuser', [0]));
        Result := True;

        except
        on E:Exception do
        begin
            Debug('SQL connect error', E.Message);
            Debug('Connection string', GetConStr);
            Result := False;
        end;
    end;
end;

function IsNull(A,B:variant):variant;
begin
  if VarIsNull(A) then Result := B else Result := A;
end;

function SvcStart: boolean;
begin
	CoInitializeEx(nil, 0);
	Result := True;
	try
        Mail := TSMTPSend.Create;
        Mime := TMimeMess.Create;
        Mail.TargetHost:=params.ReadString('MAILBOX', 'TargetHost', '');
        Mail.UserName:=params.ReadString('MAILBOX', 'UserName', '');
        Mail.Password:=params.ReadString('MAILBOX', 'Password', '');
        Mail.AutoTLS:=False;
        Mail.FullSSL:=False;
        Mail.TargetPort:=params.ReadString('MAILBOX', 'Port', '');

        fLoopDelay:=params.ReadInteger('Common', 'LoopDelay', 20000);
	except
		on E: exception do
		begin
			Debug('Error', E.Message);
			If IsConsole then
			begin
				WriteLn('Press ENTER to exit');
				Readln;
			end;
			Result := False;
		end;
	end;
end;

function SvcLoop: boolean;
var R: OleVariant;
begin
    if not ConnectSQL(fCon) then Exit;
    R := fCon.Execute('exec mng_makepost;6');
    if R.State <> 1 then
    begin
  	  Debug('Очередь пустая', Now);
      Exit;
    end
    else Debug('Loop start', Now);

	var sl: TStringList;
	var Msg: TMimeMess;
	var Mmp: TMimePart;
	var sended: boolean;
	var S: RawByteString;
	//var is_html: boolean;
	try
		try
			sl := TStringList.Create;
			if Mail.Login then
			begin
				Msg:=TMimeMess.Create;
				try
					Msg.Header.From := 'W -- T -- C <prices@wtc.ru>';
					Msg.Header.ToList.DelimitedText := R.Fields[0].Value;
					Msg.Header.ReplyTo := R.Fields[1].Value;
                    Msg.Header.Subject := R.Fields[3].Value;
					Mmp := Msg.AddPartMultipart('mixed', nil);

                    S := R.Fields[2].Value;

					if IsUTF8String(S) then	S := UTF8ToString(S);

                    Sl.Text := S;
					//if is_html then
                    Msg.AddPartHTML(Sl, Mmp);
					//else
					//	Msg.AddPartText(sl, Mmp);

					{проверяем и добавляем вложения}

                    //for S in memFiles.Lines do
                    //begin
                    //if (S<>'') and FileExists(S) then Msg.AddPartBinaryFromFile(S, Mmp);
                    //end;

					Msg.EncodeMessage;
					if (Mail.Login) and (Mail.AuthDone) then
					begin
						if Mail.MailFrom(GetEmailAddr(Msg.Header.From), Length(Msg.Lines.Text)) then
						begin
                        	sended := Mail.MailTo(Msg.Header.ToList.DelimitedText);
							if sended then sended := Mail.MailData(Msg.Lines);
						end;
						if sended then
							Debug('Письмо успешно отправлено',Msg.Header.ToList.DelimitedText)
						else
							Debug('Во время отправки письма произошла ошибка',Msg.Header.ToList.DelimitedText);
						Mail.Logout;
					end;
				finally
					Msg.Free;
				end;
			 end;
		finally
		end;
	except
		On E: Exception do Debug('Loop error', E.Message);
	end;
end;

function SvcStop: boolean;
begin
    Mail.Free;
    Mime.Free;
    params.free;
    CoUninitialize;
end;

initialization
  Params := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
	FormatSettings.DecimalSeparator := '.';
finalization
  Params.Free;
end.

