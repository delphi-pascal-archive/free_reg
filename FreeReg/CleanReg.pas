Unit CleanReg;

interface

uses
Windows, SysUtils, Classes, Registry, dialogs, ShlObj;

Const
  PATH_DESKTOP                       = $0000;  // Рабочий стол
  PATH_INTERNET                      = $0001;
  PATH_PROGRAMS                      = $0002;  // "Все программы"
  PATH_CONTROLS                      = $0003;  // ХЗ
  PATH_PRINTERS                      = $0004;
  PATH_PERSONAL                      = $0005;  // Мои документы
  PATH_FAVORITES                     = $0006;  // Избранное
  PATH_STARTUP                       = $0007;  // Автозагрузка
  PATH_RECENT                        = $0008;  // Ярлыки на недавние документы
  PATH_SENDTO                        = $0009;
  PATH_BITBUCKET                     = $000a;
  PATH_STARTMENU                     = $000b;  // Меню "Пуск"
  PATH_DESKTOPDIRECTORY              = $0010;  // Рабочий стол
  PATH_DRIVES                        = $0011;
  PATH_NETWORK                       = $0012;
  PATH_NETHOOD                       = $0013;
  PATH_FONTS                         = $0014;
  PATH_TEMPLATES                     = $0015;
  PATH_COMMON_STARTMENU              = $0016;
  PATH_COMMON_PROGRAMS               = $0017;
  PATH_COMMON_STARTUP                = $0018;
  PATH_COMMON_DESKTOPDIRECTORY       = $0019;
  PATH_APPDATA                       = $001a;
  PATH_PRINTHOOD                     = $001b;
  PATH_ALTSTARTUP                    = $001d;
  PATH_COMMON_ALTSTARTUP             = $001e;
  PATH_COMMON_FAVORITES              = $001f;
  PATH_INTERNET_CACHE                = $0020;
  PATH_COOKIES                       = $0021;
  PATH_HISTORY                       = $0022;
  PATH_TEMP                          = $0023; // My
  PATH_MYMUSIC_XP                    = $000d;
  PATH_MYGRAPHICS_XP                 = $0027;
  PATH_WINDOWS                       = $0024; // My
  PATH_SYSTEM                        = $0025; // My
  PATH_PROGRAMFILES                  = $0026; // My
  PATH_COMMONFILES                   = $002b; // My
  PATH_RESOURCES_XP                  = $0038;
  PATH_CURRENTUSER_XP                = $0028;

  OS_Version                         = $0001;
  OS_Platform                        = $0002;
  OS_Name                            = $0003;
  OS_Organization                    = $0004;
  OS_Owner                           = $0005;
  OS_SerNumber                       = $0006;
  OS_WinPath                         = $0007;
  OS_SysPath                         = $0008;
  OS_TempPath                        = $0009;
  OS_ProgramFilesPath                = $000a;
  OS_IPName                          = $000b;

Type TSystemPath=(Desktop,StartMenu,Programs,Startup,Personal, winroot, winsys);

type
  TCleanThread = class(TThread)
  private
    Procedure GetSub(Node,NodeKey: String; Root: Cardinal);
    Procedure RegRecurseScan(ANode: String; Key, OldKey: string; Level: Integer);
    Procedure ScanRegistry(Root: Cardinal; Metod: integer{; ListView});
  protected
    procedure Execute; override;
  public
    Root: Cardinal;
    Key: String;
    scanning: boolean;
    InvExt: Boolean;
    InvFlp: Boolean;
  end;


var
FReg,reg: TRegistry;
Log : TStringList;
Ext: String;
ScanMb : Real;
FSize: integer;
P: TStringList;i:byte;

Procedure InitLog;
Procedure FreeLog(Save: Boolean ;SavePath: String);
Function ClearKey(RootKey: Cardinal; Key: String; Value: String; Param: String ;DelKey: Boolean; ReservCopy: Boolean):Boolean;
Procedure CleanWindows;
procedure ScanUNINST;
function Check:string;
implementation
uses MainFrm;
/////////////////////////////////////////////
procedure ScanUNINST;
begin
MainForm.ListBox2.Clear;

reg:=TRegistry.Create;
p:=TStringList.Create;
reg.RootKey:=HKEY_LOCAL_MACHINE;

if reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',false) then
reg.GetKeyNames(p);
reg.Free;

if p.Count>0 then
for i := 0 to p.Count-1 do
begin
reg:=TRegistry.Create;
reg.RootKey:=HKEY_LOCAL_MACHINE;
  if  reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'+p[i],false) then
     if reg.ValueExists('DisplayName') then
       if reg.ValueExists('UninstallString') then
       MainForm.ListBox2.Items.Add(reg.ReadString('DisplayName'));
reg.Free;
end;
if MainForm.ListBox2.Count>0 then
MainForm.ListBox2.Selected[0]:=true;
end;
//====================================================================
function Check:string;
begin
result:='';
for i := 0 to p.Count-1 do
  begin
  reg:=TRegistry.Create;
  reg.RootKey:=HKEY_LOCAL_MACHINE;
    if  reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'+p[i],false) then
       if reg.ValueExists('UninstallString') then
          if MainForm.ListBox2.Items.Strings[MainForm.ListBox2.ItemIndex]=reg.ReadString('DisplayName') then
          begin
          result:='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'+p[i];
          reg.CloseKey;
          reg.Free;
          exit;
          end;
  end;
end;

/////////////////////////////////////////////
Function GetSize(FileN: String): String;
var
hdc : cardinal;
Buf: integer;
begin
try
hdc := FileOpen(FileN,0);
buf := GetFileSize(hdc,0);
result := PChar(inttostr(buf));
FileClose(hdc);
except
end;
end;

function GetSystemPath(Handle: THandle; PATH_type: WORD): String;
var
  P:PItemIDList;
  C:array [0..1000] of char;
  PathArray:array [0..255] of char;
begin
  if PATH_type = PATH_TEMP then
  begin
    FillChar(PathArray,SizeOf(PathArray),0);
    ExpandEnvironmentStrings('%TEMP%', PathArray, 255);
    Result := Format('%s',[PathArray])+'\';
    Exit;
  end;

  if PATH_type = PATH_WINDOWS then
  begin
    FillChar(PathArray,SizeOf(PathArray),0);
    GetWindowsDirectory(PathArray,255);
    Result := Format('%s',[PathArray])+'\';
    Exit;
  end;

  if PATH_type = PATH_SYSTEM then
  begin
    FillChar(PathArray,SizeOf(PathArray),0);
    GetSystemDirectory(PathArray,255);
    Result := Format('%s',[PathArray])+'\';
    Exit;
  end;

  if (PATH_type = PATH_PROGRAMFILES) or(PATH_type = PATH_COMMONFILES) then
  begin
    FillChar(PathArray,SizeOf(PathArray),0);
    ExpandEnvironmentStrings('%ProgramFiles%', PathArray, 255);
    Result := Format('%s',[PathArray])+'\';
    if Result[1] = '%' then
    begin
      FillChar(PathArray,SizeOf(PathArray),0);
      GetSystemDirectory(PathArray,255);
      Result := Format('%s',[PathArray]);
      Result := Result[1]+':\Program Files\';
    end;
    if (PATH_type = PATH_COMMONFILES) then Result := Result+'Common Files\';
    Exit;
  end;

  if SHGetSpecialFolderLocation(Handle,PATH_type,p)=NOERROR then
  begin
    SHGetPathFromIDList(P,C);
    if   StrPas(C)<>'' then
      Result := StrPas(C)+'\' else  Result:='';
  end;
end;

Function FindFile(Dir:String): Boolean;
Var
  SR:TSearchRec;
  FindRes,i:Integer;
  EX,tmp : String;
  MDHash : String;
  c: cardinal;
  Four: integer;
begin
  Four := 0;
  FindRes:=FindFirst(Dir+'*.*',faAnyFile,SR);
  While FindRes=0 do
   begin

    if ((SR.Attr and faDirectory)=faDirectory) and
    ((SR.Name='.')or(SR.Name='..')) then
      begin
      FindRes:=FindNext(SR);
      Continue;
      end;

    if ((SR.Attr and faDirectory)=faDirectory) then
      begin
      FindFile(Dir+SR.Name+'\');
      FindRes:=FindNext(SR);
      Continue;
      end;
    MainForm.RefreshApp;
    Ex := ExtractFileExt(Dir+SR.Name);
    // Scan for exestension
    if Ext <> '' then
    if LowerCase(Ex) = LowerCase(Ext)then
      begin
      FSize :=  strtoint(GetSize(Dir+SR.Name));
      ScanMb := ScanMb + (FSize div 1024) / 1024;
      FSize := (FSize div 1024);
      MainForm.ListBox1.Items.Add(Dir+SR.Name);
      end;
    if Ext = '' then
      begin
      FSize :=  strtoint(GetSize(Dir+SR.Name));
      ScanMb := ScanMb + (FSize div 1024) / 1024;
      FSize := (FSize div 1024);
      MainForm.ListBox1.Items.Add(Dir+SR.Name);
      end;

    FindRes:=FindNext(SR);
  end;
  FindClose(SR);
end;


Procedure CleanWindows;
begin
// Temp Folder:
ScanMb := 0;
if DirectoryExists(GetSystemPath(0,PATH_TEMP)) then begin
Ext := '';
FindFile(GetSystemPath(0,PATH_TEMP));
end;
if DirectoryExists(GetSystemPath(0,PATH_WINDOWS)+'Prefetch\') then begin
Ext := '.pf';
FindFile(GetSystemPath(0,PATH_WINDOWS)+'Prefetch\');
end;
MainForm.ListBox1.Items.Add('');
MainForm.ListBox1.Items.Add('======================================================================');
MainForm.ListBox1.Items.Add(FormatFloat('0.00',ScanMb)+'Мб. будет удалено. (Примерный размер)');

MainForm.Button1.Enabled := True;
MainForm.Button3.Enabled := True;
MainForm.RCl.TabVisible := True;
MainForm.TabSheet2.TabVisible := True;
end;
/////////////////////////////////////////////
Procedure InitLog;
begin
Log := TStringList.Create;
Log.Add('Windows Registry Editor Version 5.00');
Log.Add('');
Log.Add('');
end;

Procedure FreeLog(Save: Boolean ;SavePath: String);
begin
try
if Save = true then begin
Log.SaveToFile(SavePath);
end;
Log.Free;
except
end;
end;

Procedure SaveReservCopy(RootKey: Cardinal; Key, Value, Param: String);
var
RootStr: String;
RG: TRegistry;
begin
case RootKey of
 HKEY_CLASSES_ROOT: RootStr := 'HKEY_CLASSES_ROOT';
 HKEY_CURRENT_USER: RootStr := 'HKEY_CURRENT_USER';
 HKEY_LOCAL_MACHINE: RootStr := 'HKEY_LOCAL_MACHINE';
end;

try
//
Rg := TRegistry.Create;
Rg.RootKey := RootKey;

if Rg.OpenKey(key,True) then begin
Log.Add('['+RootStr+Key+']');
if Value <> '' then begin
Log.Add('"'+Value+'"="'+Param+'"');
end;
end;

except
end;
RG.Free;
end;

Function ClearKey(RootKey: Cardinal; Key: String; Value: String; Param: String ;DelKey: Boolean; ReservCopy: Boolean):Boolean;
var
RG: TRegistry;
RootStr: String;
begin
case RootKey of
 HKEY_CLASSES_ROOT: RootStr := 'HKEY_CLASSES_ROOT';
 HKEY_CURRENT_USER: RootStr := 'HKEY_CURRENT_USER';
 HKEY_LOCAL_MACHINE: RootStr := 'HKEY_LOCAL_MACHINE';
end;
RG := TRegistry.Create;
RG.RootKey := RootKey;

if DelKey = True then begin
Result := RG.DeleteKey(KEY);
if ReservCopy = true then SaveReservCopy(RootKey, Key,'','');
end else begin
RG.OpenKey(Key,False);
Result := RG.DeleteValue(Value);
if ReservCopy = true then SaveReservCopy(RootKey, Key, Value, Param);
end;

RG.Free;
end;
/////////////////////////////////////////////
function FixupPath(Key: string): string;
begin
  if Key = '' then
    Result := '\'
  else
  if AnsiLastChar(Key) <> '\' then
    Result := Key + '\'
  else
    Result := Key;
  if Length(Result) > 1 then
    if (Result[1] = '\') and (Result[2] = '\') then
      Result := Copy(Result, 2, Length(Result));
end;

function GetPreviousKey(Key: string): string;
var
  I: Integer;
begin
  Result := Key;
  if (Result = '') or (Result = '\') then Exit;
  for I := Length(Result) - 1 downto 1 do
    if Result[I] = '\' then
    begin
      Result := Copy(Result,1,I - 1);
      Exit;
    end;
end;

Function DelExt(Fname: String): String;
var
TmpStr, Ext: String;
Ps: integer;
begin
Ext := ExtractFileExt(Fname);
ps := pos(ext,Fname);
tmpstr := Fname;
if ps <> 0 then begin
Delete(TmpStr,ps,length(TmpStr)-ps+1);
end;
Result := TmpStr;
end;

function GetHDDSerial(ADisk : char): dword;
var
SerialNum : dword;
a, b : dword;
VolumeName : array [0..255] of char;
begin
try
Result := 0;
if GetVolumeInformation(PChar(ADisk + ':\'), VolumeName, SizeOf(VolumeName),
@SerialNum, a, b, nil, 0) then
Result := SerialNum;
except
end;
end;

Function isPath(Param: string; var FormatPath: String): boolean;
var
Fname,Str,Dir,Ext,Name, TempStr: String;
Ps,Len,i: integer;
Bol, ValidName: Boolean;
begin
//TempStr := Param;
ValidName := true;
Bol := False;
Result := False;
Fname := Param;
Dir := '';
Name := '';
Ext := '';
Dir := ExtractFileDrive(Fname);
Name := ExtractFileName(Fname);
Ext := ExtractFileExt(Fname);

//del TRASH
try
ps := Pos(' ',Ext);
if ps <> 0 then begin
Delete(Ext,ps,length(ext)-ps+1)
end;
except
end;

try
ps := Pos('!',Ext);
if ps <> 0 then begin
Delete(Ext,ps,length(Ext)-ps+1)
end;
except
end;

try
ps := Pos(',',Ext);
if ps <> 0 then begin
Delete(Ext,ps,length(Ext)-ps+1)
end;
except
end;

try
ps := Pos('^',Name);
if ps <> 0 then begin
Delete(Name,ps,length(Name)-ps+1)
end;
except
end;

try
ps := Pos('%',Name);
if ps <> 0 then begin
Delete(Name,ps,length(Name)-ps+1)
end;
except
end;

try
tempstr := Param;
if pos(':\',tempstr) <> 0 then
Delete(tempstr,pos(':\',tempstr),2);
if pos(':\',tempstr) <> 0 then ValidName := False;
except
end;


//Valid symbols
if pos('|',Fname) <> 0 then ValidName := False;
if pos('=',Fname) <> 0 then ValidName := False;
if pos('*',Fname) <> 0 then ValidName := False;
if pos('/',Name) <> 0 then ValidName := False;
if pos('\',Name) <> 0 then ValidName := False;
if pos('/',Name) <> 0 then ValidName := False;
if pos('"',Name) <> 0 then ValidName := False;

//Gen Folders
Ps := length(Fname);
try
if Fname[ps] = '.' then Bol := true;
except
end;

if Length(Fname) > 3 then
if Fname[2] = ':' then if Fname[3] = '\' then if ValidName = True then
if (Dir <> '') and (Name <> '') and (Ext <> '') then begin
  Result := true;
  if Bol = False then
  FormatPath := ExtractFilePath(Fname)+DelExt(ExtractFileName(Name))+Ext
  else
  FormatPath := ExtractFilePath(Fname)+ExtractFileName(Name);
end;
end;
/////////////////////////////////////////////
/// Get Sub Values                        ///
/////////////////////////////////////////////
Procedure TCleanThread.GetSub(Node,NodeKey: String; Root: Cardinal);
var s,v,tmp: string;
KeyInfo : TRegKeyInfo;
ValueNames,Strtemp : TStringList;
i,sn : Integer;
DataType : TRegDataType;
reg : TRegistry;
begin
if scanning = False then Exit;
 s:= Node;
 reg := TRegistry.Create;
 reg.RootKey :=Root;
 if not reg.OpenKeyReadOnly(s) then Exit;
 reg.GetKeyInfo(KeyInfo);
 if (KeyInfo.NumValues<=0) and (Root <> HKEY_CLASSES_ROOT) then Exit;

 ValueNames := TStringList.Create;
 reg.GetValueNames(ValueNames);

 // proverke on ext
   if reg.RootKey = HKEY_CLASSES_ROOT then begin
   Strtemp := TStringList.Create;
   reg.GetKeyNames(Strtemp);

   if NodeKey[1] = '.' then begin
   if ValueNames.Count-1 = -1 then
   if Strtemp.Count-1 = -1 then
   if '\'+NodeKey = Node then
   with MainForm.ListView1.Items.Add do begin
    Caption := 'Неверное расширение';
    SubItems.Add(NodeKey);
    SubItems.Add('HKCR');
    SubItems.Add(Node);
    SubItems.Add('');
   end;
   end;
   Strtemp.Free;
   exit;
   end;
   //

 for i := 0 to ValueNames.Count-1 do
 begin
   // proverka:  ///////////////////////////////////////
   if reg.GetDataType(ValueNames[i]) = rdString then begin
   s := reg.ReadString(ValueNames[i]);
   // disk skan
   tmp := S;
   if (S <> '') and (S[1] <> 'A') and (S[1] <> 'a') then
   if GetHDDSerial(S[1]) <> 0 then
   if FileExists(S) = false then
   if IsPath(S,S) = true then begin
   V := ExtractFileExt(S);
   if V <> '' then begin
   if DirectoryExists(S) = False then
   if FileExists(S) = false then
   //
   with MainForm.ListView1.Items.Add do begin
    if (reg.RootKey =HKEY_CURRENT_USER ) or (reg.RootKey = HKEY_LOCAL_MACHINE ) then begin
    Caption := 'Неверная ссылка на файл';
    SubItems.Add(tmp);

    if Root =HKEY_CURRENT_USER then
    SubItems.Add('HKCU') else
    if Root =HKEY_LOCAL_MACHINE then
    SubItems.Add('HKLM');

    SubItems.Add(Node);
    SubItems.Add(ValueNames[i]);
    end;

   end;
   //
   end;
   
   end;
   end;
   end;

 ValueNames.Free;
 reg.Free;
end;
/////////////////////////////////////////////
/// Scan registry fo keys                 ///
/////////////////////////////////////////////
Procedure TCleanThread.RegRecurseScan(ANode: String; Key, OldKey: string; Level: Integer);
var
  AStrings: TStringList;
  I: Integer;
  //NewNode: TTreeNode;
  AKey: string;
begin
if scanning = False then Exit;
  AKey := FixupPath(OldKey);
  if FReg.OpenKeyReadOnly(Key) and FReg.HasSubKeys then
  begin
    if Level = 1 then
    begin
      AStrings := TStringList.Create;
      try
        FReg.GetKeyNames(AStrings);
        for I := 0 to AStrings.Count - 1 do
        begin
        if scanning = False then Exit;
          if AStrings[I] = '' then
            AStrings[I] := Format('%.04d', [I]);
          GetSub(ANode+'\'+AStrings[I],AStrings[I],FReg.RootKey);
          if Freg.RootKey <> HKEY_CLASSES_ROOT then
          RegRecurseScan(ANode+'\'+AStrings[I], AStrings[I], AKey + Key, Level);
        end;
      finally
        AStrings.Free;
      end;
    end;
  end;
  FReg.OpenKeyReadOnly(AKey);
end;

Procedure TCleanThread.ScanRegistry(Root: Cardinal; Metod: integer{; ListView});
begin
//
end;

Procedure TCleanThread.Execute;
begin
//MainForm.Label1.Caption := 'Scan in progress';
//    InvExt: Boolean;
//    InvFlp: Boolean;

Scanning := True;
if InvExt = True then begin
MainForm.Panel1.Caption := 'Поиск неверных расширений. Пожалуйста ждите...';
FReg := TRegistry.Create;
FReg.RootKey := HKEY_CLASSES_ROOT;
RegRecurseScan('','','',1);
FReg.Free;
end;
if InvFlp = True then begin
MainForm.Panel1.Caption := 'Поиск неверных ссылок на файлы. Пожалуйста ждите...';
FReg := TRegistry.Create;
FReg.RootKey := HKEY_LOCAL_MACHINE;
RegRecurseScan('\SOFTWARE','\SOFTWARE','',1);
FReg.Free;
end;
if InvFlp = True then begin
MainForm.Panel1.Caption := 'Поиск неверных ссылок на файлы. Пожалуйста ждите...';
FReg := TRegistry.Create;
FReg.RootKey := HKEY_CURRENT_USER;
RegRecurseScan('\SOFTWARE','\SOFTWARE','',1);
FReg.Free;
end;
MainForm.Panel1.Caption := 'Анализ ключей реестра завершен.';
Scanning := False;

MainForm.Button1.Enabled := True;
MainForm.Button2.Enabled := False;
MainForm.TabSheet1.TabVisible := True;
MainForm.TabSheet2.TabVisible := True;
if MainForm.ListView1.Items.Count > 0 then MainForm.Button3.Enabled := True;
end;

end.
