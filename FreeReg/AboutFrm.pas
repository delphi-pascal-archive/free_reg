unit AboutFrm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, ShellAPI;

type
  TAboutForm = class(TForm)
    BitBtn1: TBitBtn;
    Image1: TImage;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Label10: TLabel;
    Label19: TLabel;
    Label18: TLabel;
    Label17: TLabel;
    Label20: TLabel;
    procedure BitBtn1Click(Sender: TObject);
    procedure Label5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutForm: TAboutForm;

implementation

{$R *.dfm}

procedure TAboutForm.BitBtn1Click(Sender: TObject);
begin
Close;
end;

procedure TAboutForm.Label5Click(Sender: TObject);
Const 
  URL : String = 'http://BlackCash2006@Yandex.ru';
begin
  ShellExecute(0,'',pChar('mailto:'+URL),NIL,NIL,SW_SHOWNORMAL);
end;
end.
