unit MainFrm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, XPMan, CleanReg,registry;

type
  TMainForm = class(TForm)
    TopPanel: TPanel;
    Image12: TImage;
    Image13: TImage;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    PageControl1: TPageControl;
    RCl: TTabSheet;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ListView1: TListView;
    Bevel1: TBevel;
    XPManifest1: TXPManifest;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Panel1: TPanel;
    SaveDialog1: TSaveDialog;
    ListBox1: TListBox;
    ListBox2: TListBox;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure RClShow(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);
    procedure TabSheet2Show(Sender: TObject);
    Procedure RefreshApp;
    procedure ListBox2Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
  public
  Scan : TCleanThread;
    { Public declarations }
  end;

var
  MainForm: TMainForm;
  UninstPath: String;
implementation

uses Math, AboutFrm;

{$R *.dfm}
Procedure TMainForm.RefreshApp;
begin
Application.ProcessMessages;
end;

procedure TMainForm.Button1Click(Sender: TObject);
begin
if Button1.Tag = 0 then begin
ListView1.Clear;
Scan := TCleanThread.Create(True);
Scan.InvExt := True;
Scan.InvFlp := True;
Scan.FreeOnTerminate := True;
Scan.Resume;

TabSheet1.TabVisible := False;
TabSheet2.TabVisible := False;
Button1.Enabled := False;
Button3.Enabled := False;
Button2.Enabled := True;
end;

if Button1.Tag = 1 then begin
RCl.TabVisible := False;
TabSheet2.TabVisible := False;
ListBox1.Clear;
CleanWindows;
Button1.Enabled := False;
end;

if Button1.Tag = 2 then ScanUNINST;
end;


procedure TMainForm.Button3Click(Sender: TObject);
var
i,P: integer;
Root: Cardinal;
SavRC: Boolean;
s:string;
begin
if Button1.Tag = 0 then begin
// - Clear reg ///////////////////////////////////
P := MessageDlg('Сохранить резервную копию исправляемых значений?',mtInformation,mbOKCancel,0);

if P = 1 then begin
SavRC := True;
SaveDialog1.Execute;
end else SavRC := False;

InitLog;
For i := 0 to ListView1.Items.Count-1 do begin
  if ListView1.Items.Item[i].Checked = true then begin
    if ListView1.Items.Item[i].SubItems[1] = 'HKCU' then Root := HKEY_CURRENT_USER;
    if ListView1.Items.Item[i].SubItems[1] = 'HKLM' then Root := HKEY_LOCAL_MACHINE;
    if ListView1.Items.Item[i].SubItems[1] = 'HKCR' then Root := HKEY_CLASSES_ROOT;
    Application.ProcessMessages;
    Sleep(100);
    if Root = HKEY_CLASSES_ROOT then begin
    if clearkey(Root,ListView1.Items.Item[i].SubItems[2],'','',True,SavRC) then
    ListView1.Items.Item[i].Caption := 'Исправлено';
    end
    else begin
    if clearkey(Root,ListView1.Items.Item[i].SubItems[2],ListView1.Items.Item[i].SubItems[3],ListView1.Items.Item[i].SubItems[2],False,SavRC) then
    ListView1.Items.Item[i].Caption := 'Исправлено';
    end;
  end;
end;
ListView1.Clear;
CleanReg.FreeLog(SavRC,SaveDialog1.FileName);
///////////////////////////////////////////////////

Button3.Enabled := False;
End;
if Button1.Tag = 1 then begin
If ListBox1.Count > 0 then if ListBox1.Count > 3 then begin

For i := 0 to ListBox1.Count - 3 do begin
try
DeleteFile(PChar(ListBox1.Items.Strings[i]));
except
end;
end;

end;
ListBox1.Clear;
Button3.Enabled := False;
Button1.Enabled := True;
end;

if Button1.Tag = 2 then begin
s:=Check;
if ListBox2.Count>0 then
if messagebox(MainForm.Handle,'Удалить ключ из реестра?','Удаление', 4)=IDYes then
begin
reg:=TRegistry.Create;
reg.RootKey:=HKEY_LOCAL_MACHINE;

if reg.DeleteKey(s) then
                 showmessage('Раздел удален')
                 else
                 showmessage('Ошибка');
reg.Free;

ScanUNINST;
ListBox2Click(Sender);
end;

end;
end;

procedure TMainForm.Button2Click(Sender: TObject);
begin
if Button1.Tag = 0 then begin
try
Scan.scanning := False;
Scan.Terminate;
except
end;
end;
if Button1.Tag = 2 then begin
windows.WinExec(PChar(UninstPath),SW_SHOW);
end;
end;

procedure TMainForm.Button4Click(Sender: TObject);
begin
Close;
end;

procedure TMainForm.RClShow(Sender: TObject);
begin
ListView1.Items.Clear;
Button3.Caption := 'Исправить';
Button2.Caption := 'Отмена';
Button1.Caption := 'Начать';
Button1.Enabled := True;
Button2.Enabled := False;
Button3.Enabled := False;
Button2.Visible := True;
Button1.Tag := 0;
Panel1.Caption := '';
end;

procedure TMainForm.TabSheet1Show(Sender: TObject);
begin
ListBox1.Clear;
Button3.Caption := 'Очистить';
//Button2.Caption := 'Отмена';
Button1.Caption := 'Начать';
Button1.Enabled := True;
Button3.Enabled := False;
Button2.Visible := False;
Button1.Tag := 1;
Panel1.Caption := '';
end;

procedure TMainForm.TabSheet2Show(Sender: TObject);
begin
Button3.Caption := 'Удалить';
Button2.Caption := 'Uninstall';
Button1.Caption := 'Обновить';
Button2.Visible := True;
Button1.Enabled := True;
Button2.Enabled := True;
Button3.Enabled := True;
Button1.Tag := 2;
Panel1.Caption := '';
ScanUNINST;
end;

procedure TMainForm.ListBox2Click(Sender: TObject);
var s:string;
begin
s:=Check;
reg:=TRegistry.Create;
reg.RootKey:=HKEY_LOCAL_MACHINE;
if not reg.OpenKey(s,false) then
       begin
       showmessage('Ошибка открытия ключа');
       exit;
       end else
       if not reg.ValueExists('UninstallString') then
       begin
       showmessage('UninstallString не найденно');
       exit;
       end else
       UninstPath :=reg.ReadString('UninstallString');
Panel1.Caption := UninstPath;
reg.Free;
end;

procedure TMainForm.Button5Click(Sender: TObject);
begin
aboutform.ShowModal;
end;

end.
