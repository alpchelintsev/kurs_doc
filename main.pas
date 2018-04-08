unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, ExtCtrls, ComObj, Menus;

type
  TFormInput = class(TForm)
    LabeledEdit1: TLabeledEdit;
    Label1: TLabel;
    DateTimePicker1: TDateTimePicker;
    Label2: TLabel;
    DateTimePicker2: TDateTimePicker;
    LabeledEdit2: TLabeledEdit;
    Label3: TLabel;
    ComboBox1: TComboBox;
    LabeledEdit3: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    StatusBar1: TStatusBar;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    LabeledEdit6: TLabeledEdit;
    Edit1: TEdit;
    CheckBox1: TCheckBox;
    LabeledEdit7: TLabeledEdit;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    Label4: TLabel;
    Button4: TButton;
    OpenDialog1: TOpenDialog;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
    procedure CreteListStudents(FN: String);
    procedure StringReplace(W: OLEVariant; SearchStr,ReplaceStr: String;
                            ReplaceAll: Boolean);
  public
    { Public declarations }
  end;

  TInfoStudent = record
    Fam   : String;
    IO    : String;
    Theme : String;
    ocenka: String;
    page  : String
  end;

var
  FormInput: TFormInput;

implementation

{$R *.dfm}

var
  studs: array of TInfoStudent;
  n    : Integer = 0;

  fname_IUL       : String = '';
  fname_Perechen  : String = '';
  fname_stick     : String = '';
  fname_ved       : String = '';
  ex_file_IUL     : Boolean = false;
  ex_file_Perechen: Boolean = false;
  ex_file_stick   : Boolean = false;
  ex_file_ved     : Boolean = false;
  path            : String = '';
  config_path     : String = '';

procedure TFormInput.CreteListStudents(FN: String);
var
  XL     : OLEVariant;
  XLRun  : Boolean;
  s      : String;
  ist,i,j: Integer;
  Name   : String;
  buf    : String;
  status : String;
  flag   : Boolean;
begin
  studs:=nil;
  n:=0;
  Button1.Enabled:=false;
  Button2.Enabled:=false;
  Button3.Enabled:=false;
  Button4.Enabled:=false;
  XLRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Excel, формирование списка студентов. Ждите...';
    XL:=CreateOleObject('Excel.Application');
    XLRun:=true;
    XL.WorkBooks.Open(FN);
    ist:=1;
    while true do
    begin
      s:=XL.WorkBooks[1].WorkSheets[1].Cells[ist,1];
      if s = '' then
        break;
      inc(n);
      SetLength(studs, n);
      status:=XL.WorkBooks[1].WorkSheets[1].Cells[ist,2];
      Name:='';
      if status <> '*' then
      begin
        flag:=false;
        for i:=2 to Length(s) do
          if s[i] in ['А'..'Я'] then
          begin
            flag:=true;
            break
          end;
        if flag then
          Name:=s[i] + '.';
        flag:=false;
        for j:=i+1 to Length(s) do
          if s[j] in ['А'..'Я'] then
          begin
            flag:=true;
            break
          end;
        if flag then
          Name:=Name + s[j] + '.';
        for j:=i downto 1 do
          if s[j] in ['а'..'я'] then
            break;
        buf:=Copy(s, 1, j);
        s:=buf
      end;
      studs[n-1].Fam   :=s;
      studs[n-1].IO    :=Name;
      studs[n-1].Theme :=XL.WorkBooks[1].WorkSheets[1].Cells[ist,3];
      studs[n-1].ocenka:=XL.WorkBooks[1].WorkSheets[1].Cells[ist,4];
      studs[n-1].page  :=XL.WorkBooks[1].WorkSheets[1].Cells[ist,5];
      inc(ist)
    end;
    XL.Quit;
    XL:=Unassigned;
    Button1.Enabled:=ex_file_IUL;
    Button2.Enabled:=ex_file_Perechen;
    Button4.Enabled:=ex_file_stick;
    Button3.Enabled:=ex_file_ved;
    StatusBar1.SimpleText:='Готово'
  except
    MessageDlg('Ошибка работы с xlsx-файлом',mtError,[mbOk],0);
    try
      if XLRun then XL.Quit
    except
    end;
    XL:=Unassigned;
    StatusBar1.SimpleText:=''
  end
end;

procedure TFormInput.StringReplace(W: OLEVariant; SearchStr,ReplaceStr: String;
                                   ReplaceAll: Boolean);
begin
  try
    W.Selection.Find.ClearFormatting;
    W.Selection.Find.Text:=SearchStr;
    W.Selection.Find.Replacement.Text:=ReplaceStr;
    W.Selection.Find.Forward:=true;
    W.Selection.Find.Wrap:=1;
    W.Selection.Find.Format:=false;
    W.Selection.Find.MatchCase:=false;
    W.Selection.Find.MatchWholeWord:=false;
    W.Selection.Find.MatchWildcards:=false;
    W.Selection.Find.MatchSoundsLike:=false;
    W.Selection.Find.MatchAllWordForms:=false;
    if ReplaceAll then W.Selection.Find.Execute(Replace:=2)
    else W.Selection.Find.Execute(Replace:=1)
  except
  end
end;

procedure TFormInput.Button1Click(Sender: TObject);
var
  WRun    : Boolean;
  W       : OLEVariant;
  i       : Integer;
  numstud : String;
  namefile: String;
  FIOstud : String;
  names_fl: array of String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование ИУЛ. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;
    SetLength(names_fl, n);
    for i:=1 to n do
    begin
      W.Documents.Open(fname_IUL);
      StringReplace(W, '{napr}', LabeledEdit4.Text, true);
      numstud:=IntToStr(i);
      if Length(numstud) = 1 then
        numstud:='0'+numstud;
      StringReplace(W, '{numstud}', numstud, true);
      StringReplace(W, '{theme}', studs[i-1].Theme, true);
      StringReplace(W, '{enddate}', DateToStr(DateTimePicker2.DateTime), true);
      StringReplace(W, '{begindt}', DateToStr(DateTimePicker1.DateTime), true);
      if studs[i-1].IO = '' then
        FIOstud:=studs[i-1].Fam
      else
        FIOstud:=studs[i-1].Fam + ' ' + studs[i-1].IO;
      StringReplace(W, '{fiostud}', FIOstud, true);
      StringReplace(W, '{ocenka}', studs[i-1].ocenka, true);
      StringReplace(W, '{ng}', LabeledEdit7.Text, true);
      StringReplace(W, '{fioruk}', LabeledEdit3.Text, true);
      StringReplace(W, '{group}', LabeledEdit1.Text, true);
      namefile:=path + LabeledEdit4.Text + '.' + LabeledEdit7.Text + numstud +
                ' ИУЛ ' + FIOstud + '.docx';
      names_fl[i-1]:=namefile;
      W.ActiveDocument.SaveAs(FileName:=namefile, FileFormat:=16);
      W.ActiveDocument.Close
    end;
    if CheckBox1.Checked then
    begin
      W.Documents.Open(names_fl[0]);
      for i:=2 to n do
      begin
        W.Selection.EndKey(Unit:=6);
        W.Selection.InsertBreak(Type:=7);
        W.Selection.InsertFile(FileName:=names_fl[i-1]);
      end;
      W.ActiveDocument.SaveAs(FileName:=path + LabeledEdit4.Text + ' ' +
                                        LabeledEdit1.Text + ' ИУЛ.docx',
                              FileFormat:=16);
      W.ActiveDocument.Close;
      for i:=1 to n do
        DeleteFile(names_fl[i-1])
    end;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end;
  names_fl:=nil
end;

procedure TFormInput.Button2Click(Sender: TObject);
var
  WRun              : Boolean;
  W                 : OLEVariant;
  i                 : Integer;
  numstud           : String;
  namefile          : String;
  FIOstud1, FIOstud2: String;
  names_fl          : array of String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование документов. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;
    SetLength(names_fl, n);
    for i:=1 to n do
    begin
      W.Documents.Open(fname_Perechen);
      StringReplace(W, '{napr}', LabeledEdit4.Text, true);
      numstud:=IntToStr(i);
      if Length(numstud) = 1 then
        numstud:='0'+numstud;
      StringReplace(W, '{numstud}', numstud, true);
      if studs[i-1].IO = '' then
      begin
        FIOstud1:=studs[i-1].Fam;
        FIOstud2:=studs[i-1].Fam
      end
      else
      begin
        FIOstud1:=studs[i-1].Fam + ' ' + studs[i-1].IO;
        FIOstud2:=studs[i-1].IO + ' ' + studs[i-1].Fam
      end;
      StringReplace(W, '{fiostud}', FIOstud2, true);
      StringReplace(W, '{page}', studs[i-1].page, true);
      StringReplace(W, '{ng}', LabeledEdit7.Text, true);
      StringReplace(W, '{fioruk}', LabeledEdit3.Text, true);
      StringReplace(W, '{fioarch}', LabeledEdit2.Text, true);
      namefile:=path + LabeledEdit4.Text + '.' + LabeledEdit7.Text + numstud +
                ' Перечень ' + FIOstud1 + '.docx';
      names_fl[i-1]:=namefile;
      W.ActiveDocument.SaveAs(FileName:=namefile, FileFormat:=16);
      W.ActiveDocument.Close
    end;
    if CheckBox1.Checked then
    begin
      W.Documents.Open(names_fl[0]);
      for i:=2 to n do
      begin
        W.Selection.EndKey(Unit:=6);
        W.Selection.InsertBreak(Type:=7);
        W.Selection.InsertFile(FileName:=names_fl[i-1]);
      end;
      W.ActiveDocument.SaveAs(FileName:=path + LabeledEdit4.Text + ' ' +
                                        LabeledEdit1.Text + ' Перечень.docx',
                              FileFormat:=16);
      W.ActiveDocument.Close;
      for i:=1 to n do
        DeleteFile(names_fl[i-1])
    end;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end;
  names_fl:=nil
end;

procedure TFormInput.FormClose(Sender: TObject; var Action: TCloseAction);
var
  f: TextFile;
begin
  studs:=nil;
  AssignFile(f, config_path);
  {$I-} Rewrite(f); {$I+}
  if IOResult = 0 then
  begin
    WriteLn(f, LabeledEdit1.Text);
    WriteLn(f, LabeledEdit7.Text);
    WriteLn(f, LabeledEdit2.Text);
    WriteLn(f, ComboBox1.Text);
    WriteLn(f, LabeledEdit3.Text);
    WriteLn(f, LabeledEdit4.Text);
    WriteLn(f, Edit1.Text);
    WriteLn(f, LabeledEdit5.Text);
    WriteLn(f, LabeledEdit6.Text);
    CloseFile(f)
  end
end;

procedure TFormInput.N4Click(Sender: TObject);
var
  f: TextFile;
begin
  studs:=nil;
  AssignFile(f, config_path);
  {$I-} Rewrite(f); {$I+}
  if IOResult = 0 then
  begin
    WriteLn(f, LabeledEdit1.Text);
    WriteLn(f, LabeledEdit7.Text);
    WriteLn(f, LabeledEdit2.Text);
    WriteLn(f, ComboBox1.Text);
    WriteLn(f, LabeledEdit3.Text);
    WriteLn(f, LabeledEdit4.Text);
    WriteLn(f, Edit1.Text);
    WriteLn(f, LabeledEdit5.Text);
    WriteLn(f, LabeledEdit6.Text);
    CloseFile(f)
  end;
  Application.Terminate
end;

procedure TFormInput.FormCreate(Sender: TObject);
var
  f: TextFile;
  s: String;
begin
  path:=ExtractFilePath(Application.ExeName);
  config_path:=path + 'config.txt';
  path:=path + 'docs\';

  ex_file_IUL:=FileExists('IUL.docx');
  ex_file_Perechen:=FileExists('archive.docx');
  ex_file_stick:=FileExists('sticker.docx');
  ex_file_ved:=FileExists('vedomost.docx');

  if ex_file_IUL then
    fname_IUL:=ExpandFileName('IUL.docx');
  if ex_file_Perechen then
    fname_Perechen:=ExpandFileName('archive.docx');
  if ex_file_stick then
    fname_stick:=ExpandFileName('sticker.docx');
  if ex_file_ved then
    fname_ved:=ExpandFileName('vedomost.docx');

  if not DirectoryExists(path) then
    CreateDir(path);

  AssignFile(f, config_path);
  {$I-} Reset(f); {$I+}
  if IOResult = 0 then
  begin
    ReadLn(f, s);
    LabeledEdit1.Text:=s;
    ReadLn(f, s);
    LabeledEdit7.Text:=s;
    ReadLn(f, s);
    LabeledEdit2.Text:=s;
    ReadLn(f, s);
    ComboBox1.Text:=s;
    ReadLn(f, s);
    LabeledEdit3.Text:=s;
    ReadLn(f, s);
    LabeledEdit4.Text:=s;
    ReadLn(f, s);
    Edit1.Text:=s;
    ReadLn(f, s);
    LabeledEdit5.Text:=s;
    ReadLn(f, s);
    LabeledEdit6.Text:=s;
    CloseFile(f)
  end
end;

procedure TFormInput.N2Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    CreteListStudents(OpenDialog1.FileName)
end;

procedure TFormInput.Button4Click(Sender: TObject);
var
  WRun    : Boolean;
  W       : OLEVariant;
  i       : Integer;
  numstud : String;
  namefile: String;
  FIOstud : String;
  names_fl: array of String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование этикеток. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;
    SetLength(names_fl, n);
    for i:=1 to n do
    begin
      W.Documents.Open(fname_stick);
      StringReplace(W, '{napr}', LabeledEdit4.Text, true);
      numstud:=IntToStr(i);
      if Length(numstud) = 1 then
        numstud:='0'+numstud;
      StringReplace(W, '{numstud}', numstud, true);
      if studs[i-1].IO = '' then
        FIOstud:=studs[i-1].Fam
      else
        FIOstud:=studs[i-1].Fam + ' ' + studs[i-1].IO;
      StringReplace(W, '{fiostud}', FIOstud, true);
      StringReplace(W, '{ng}', LabeledEdit7.Text, true);
      StringReplace(W, '{group}', LabeledEdit1.Text, true);
      StringReplace(W, '{theme}', studs[i-1].Theme, true);
      StringReplace(W, '{year}',FormatDateTime('yyyy',DateTimePicker2.DateTime),
                    true);
      namefile:=path + LabeledEdit4.Text + '.' + LabeledEdit7.Text + numstud +
                ' Этикетка ' + FIOstud + '.docx';
      names_fl[i-1]:=namefile;
      W.ActiveDocument.SaveAs(FileName:=namefile, FileFormat:=16);
      W.ActiveDocument.Close
    end;
    if CheckBox1.Checked then
    begin
      W.Documents.Open(names_fl[0]);
      for i:=2 to n do
      begin
        W.Selection.EndKey(Unit:=6);
        W.Selection.InsertBreak(Type:=7);
        W.Selection.InsertFile(FileName:=names_fl[i-1]);
      end;
      W.ActiveDocument.SaveAs(FileName:=path + LabeledEdit4.Text + ' ' +
                                        LabeledEdit1.Text + ' Этикетки.docx',
                              FileFormat:=16);
      W.ActiveDocument.Close;
      for i:=1 to n do
        DeleteFile(names_fl[i-1])
    end;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end;
  names_fl:=nil
end;

procedure TFormInput.Button3Click(Sender: TObject);
var
  WRun   : Boolean;
  W      : OLEVariant;
  i      : Integer;
  FIOstud: String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование ведомости. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;
    W.Documents.Open(fname_ved);
    StringReplace(W, '{napr}', LabeledEdit4.Text, true);
    StringReplace(W, '{naprname}', Edit1.Text, true);
    StringReplace(W, '{institute}', ComboBox1.Text, true);
    StringReplace(W, '{sem}', LabeledEdit6.Text, true);
    StringReplace(W, '{disciplina}', LabeledEdit5.Text, true);
    StringReplace(W, '{group}', LabeledEdit1.Text, true);
    StringReplace(W, '{fioruk}', LabeledEdit3.Text, true);
    if n>=1 then
    begin
      if studs[0].IO = '' then
        FIOstud:=studs[0].Fam
      else
        FIOstud:=studs[0].Fam + ' ' + studs[0].IO;
      StringReplace(W, '{nach}', FIOstud, false);
      W.Selection.MoveRight(Unit:=$C);
      W.Selection.TypeText(Text:=studs[0].Theme);
      for i:=2 to n do
      begin
        W.Selection.MoveRight(Unit:=$C);
        W.Selection.MoveRight(Unit:=$C);
        if studs[i-1].IO = '' then
          FIOstud:=studs[i-1].Fam
        else
          FIOstud:=studs[i-1].Fam + ' ' + studs[i-1].IO;
        W.Selection.TypeText(Text:=FIOstud);
        W.Selection.MoveRight(Unit:=$C);
        W.Selection.TypeText(Text:=studs[i-1].Theme)
      end
    end
    else
      StringReplace(W, '{nach}', '', false);
    W.ActiveDocument.SaveAs(FileName:=path + LabeledEdit4.Text + ' ' +
                                      LabeledEdit1.Text + ' Ведомость.docx',
                            FileFormat:=16);
    W.ActiveDocument.Close;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end
end;

end.
