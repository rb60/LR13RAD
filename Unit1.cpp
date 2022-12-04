//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SpeedButton1Click(TObject *Sender)
{
	OleContainer1->InsertObjectDialog();
    ComboBox1->Items = OleContainer1->ObjectVerbs;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SpeedButton2Click(TObject *Sender)
{
	OleContainer1->CreateObject("Word.Document", false);
	ComboBox1->Items = OleContainer1->ObjectVerbs;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SpeedButton3Click(TObject *Sender)
{
	if (OpenDialog1->Execute() == mrOk)
	{
		OleContainer1->CreateObjectFromFile(OpenDialog1->FileName, false);
		ComboBox1->Items = OleContainer1->ObjectVerbs;
	}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SpeedButton4Click(TObject *Sender)
{
	if (SaveDialog1->Execute() == mrOk)
	{
		OleContainer1->SaveToFile(SaveDialog1->FileName);
	}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SpeedButton5Click(TObject *Sender)
{
	if (SaveDialog1->Execute() == mrOk)
	{
		OleContainer1->SaveAsDocument(SaveDialog1->FileName);
	}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SpeedButton6Click(TObject *Sender)
{
	if (OpenDialog1->Execute() == mrOk)
	{
		OleContainer1->LoadFromFile(OpenDialog1->FileName);
        ComboBox1->Items = OleContainer1->ObjectVerbs;
	}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SpeedButton7Click(TObject *Sender)
{
    OleContainer1->DoVerb(ComboBox1->ItemIndex);
}
//---------------------------------------------------------------------------

