unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Word2000, OleServer, ExtCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    WordDocument1: TWordDocument;
    WordApplication1: TWordApplication;
    OpenDialog1: TOpenDialog;
    WordParagraphFormat1: TWordParagraphFormat;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
   ItemIndex,FileName:OleVariant;
   vcol,vr: OleVariant;
   NLines,NF:integer;
   i,j,l,m:integer;
   S:array[1..200] of WideString;
   Format:array[1..100] of WideString;
   QN,ti:integer;
   K:integer;
   S1,S2:WideString;
begin
 Randomize;
 K:=StrToInt(LabeledEdit1.Text);
 QN:=StrToInt(LabeledEdit2.Text);
 Wordapplication1.Visible := True;
 OpenDialog1.Execute;
 FileName:=OpenDialog1.FileName;
 WordApplication1.Documents.Open(FileName,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 ItemIndex:=1;
 WordDocument1.ConnectTo(WordApplication1.Documents.Item(ItemIndex));
 Nlines:=WordDocument1.Paragraphs.Count;
 For i:=1 to NLines do
 Begin
   S[i]:=WordDocument1.Paragraphs.Item(i).Range.Text;
 End;
 WordDocument1.Disconnect;
 OpenDialog1.Execute;
 FileName:=OpenDialog1.FileName;
 WordApplication1.Documents.Open(FileName,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 ItemIndex:=1;
 WordDocument1.ConnectTo(WordApplication1.Documents.Item(ItemIndex));
 NF:=WordDocument1.Paragraphs.Count;
 For i:=1 to NF do
 Begin
   Format[i]:=WordDocument1.Paragraphs.Item(i).Range.Text;
 End;
 //WordApplication1.Sel
 //WordDocument1.Range.InsertAfter('оздаем объект Range1:');
 WordDocument1.Disconnect;
 ItemIndex:=1;
 vcol:=wdCollapseEnd;
 WordApplication1.Caption := 'Ёкзаменационные билеты';
 WordApplication1.Documents.Add(EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 WordDocument1.ConnectTo(WordApplication1.Documents.Item(ItemIndex));
 WordApplication1.Options.CheckSpellingAsYouType := False;
 WordApplication1.Options.CheckGrammarAsYouType := False;
 WordApplication1.Selection.Font.Size:=14;
 WordParagraphFormat1.ConnectTo(WordApplication1.Selection.ParagraphFormat);
 WordParagraphFormat1.Alignment:=wdAlignParagraphCenter;
 l:=NLines div QN;
 For i:=1 to K do
 Begin
   for j:=1 to NF do
   Begin
   vr:=wdAlignParagraphCenter;
     S1:=Format[j];
   If pos('&&&',format[j])>0 Then
   Begin
     S1:='';
     WordApplication1.Selection.Collapse(vcol);
     for ti:=1 to Qn do
     Begin
      m:=random(l)+1;
      WordApplication1.Selection.InsertAfter(IntToStr(ti)+'.  '+S[(ti-1)*l+m]);
     End;
     WordParagraphFormat1.ConnectTo(WordApplication1.Selection.ParagraphFormat);
     WordParagraphFormat1.Alignment:=wdAlignParagraphLeft;
   End
   Else
   begin
   If pos('##',format[j])>0 then
   Begin
    S2:=Format[j];
    s1:=Copy(S2,1,pos('##',S2)-1);
    Delete(S2,1,pos('##',S2)+1);
    s1:=s1+IntToStr(i)+Copy(S2,1,Length(S2));
    vr:=wdAlignParagraphCenter;
   End
   Else
   if pos('$',format[j])>0 then
   Begin
     Delete(S1,pos('$',S1),1);
     vr:=wdAlignParagraphRight;
   End
   else
   if pos('@',format[j])>0 then
   Begin
     Delete(S1,pos('@',S1),1);
     vr:=wdAlignParagraphLeft;
   End;
   WordApplication1.Selection.Collapse(vcol);
   WordApplication1.Selection.InsertAfter(S1);
   WordParagraphFormat1.ConnectTo(WordApplication1.Selection.ParagraphFormat);
   WordParagraphFormat1.Alignment:=vr
   End;
   End;
  End;
 WordDocument1.Disconnect;
 WordApplication1.Disconnect;
end;
end.
