#include"Spire.Doc.o.h"
#include<vector>

using namespace Spire::Doc;
using namespace std;

void fill_MKK_form(	const wstring& name,
					const wstring& position,
					const wstring& date,
					const wstring& experience,
					vector<vector<wstring>> repairForms,
					vector<vector<wstring>> modificationForms)
{
	//Initialize an instance of the Document class
	intrusive_ptr<Document> doc = new Document();

	//Load a Word document
	doc->LoadFromFile(L"Направление на прохождение МКК_.doc");

	//Section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Table
	intrusive_ptr<ITable> table = section->GetTables()->GetItemInTableCollection(0);

	//************************************** Insert Name *************************************

	//Row
	intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(1);

	//Cell
	intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(0);

	//Paragrapgh
	intrusive_ptr<IParagraph> paragraph = cell->GetFirstParagraph();

	//Add text
	intrusive_ptr<TextRange> text = paragraph->AppendText(name.c_str());

	//Text format
	text->GetCharacterFormat()->SetFontName(L"Times New Roman");
	text->GetCharacterFormat()->SetFontSize(11);
	text->GetCharacterFormat()->SetBold(true);

	//************************************** Insert Position *********************************

	//Row
	row = table->GetRows()->GetItemInRowCollection(2);

	//Cell
	cell = row->GetCells()->GetItemInCellCollection(0);

	//Paragrapgh
	paragraph = cell->GetFirstParagraph();

	//Add text
	text = paragraph->AppendText(position.c_str());

	//Text format
	text->GetCharacterFormat()->SetFontName(L"Times New Roman");
	text->GetCharacterFormat()->SetFontSize(11);
	text->GetCharacterFormat()->SetBold(true);

	//Row
	row = table->GetRows()->GetItemInRowCollection(4);

	//Cell
	cell = row->GetCells()->GetItemInCellCollection(0);

	//Paragrapgh
	paragraph = cell->GetFirstParagraph();

	//Add text
	text = paragraph->AppendText(position.c_str());

	//Text format
	text->GetCharacterFormat()->SetFontName(L"Times New Roman");
	text->GetCharacterFormat()->SetFontSize(11);
	text->GetCharacterFormat()->SetBold(true);

	//************************************** Insert Date of Employment *********************************

	//Row
	row = table->GetRows()->GetItemInRowCollection(3);

	//Cell
	cell = row->GetCells()->GetItemInCellCollection(0);

	//Paragraph
	paragraph = cell->GetParagraphs()->GetItemInParagraphCollection(1);

	//Add text
	text = paragraph->AppendText(date.c_str());

	//Text format
	text->GetCharacterFormat()->SetFontName(L"Times New Roman");
	text->GetCharacterFormat()->SetFontSize(11);
	text->GetCharacterFormat()->SetBold(true);

	//************************************** Insert Experience *********************************

	//Row
	row = table->GetRows()->GetItemInRowCollection(3);

	//Cell
	cell = row->GetCells()->GetItemInCellCollection(0);

	//Paragraph
	paragraph = cell->GetLastParagraph();

	//Add text
	text = paragraph->AppendText(experience.c_str());

	//Text format
	text->GetCharacterFormat()->SetFontName(L"Times New Roman");
	text->GetCharacterFormat()->SetFontSize(11);
	text->GetCharacterFormat()->SetBold(true);

	//************************************** Insert Repair Forms *********************************

	for (int i = 0; i < repairForms.size(); i++)
	{
		//Row
		row = table->GetRows()->GetItemInRowCollection(8 + i);

		for (int j = 0; j < repairForms[i].size(); j++)
		{
			//Cell
			cell = row->GetCells()->GetItemInCellCollection(2);

			//Paragrapgh
			paragraph = cell->GetFirstParagraph();

			//Add text
			text = paragraph->AppendText(repairForms[i][j].c_str());

			//Text format
			text->GetCharacterFormat()->SetFontName(L"Times New Roman");
			text->GetCharacterFormat()->SetFontSize(11);

			if (j < repairForms[i].size() - 1)
			{
				//Add text
				text = paragraph->AppendText(L", ");

				//Text format
				text->GetCharacterFormat()->SetFontName(L"Times New Roman");
				text->GetCharacterFormat()->SetFontSize(11);
			}
		}
	}

	//************************************** Insert Modification Forms *********************************

	for (int i = 0; i < modificationForms.size(); i++)
	{
		//Row
		row = table->GetRows()->GetItemInRowCollection(8 + i);

		for (int j = 0; j < modificationForms[i].size(); j++)
		{
			//Cell
			cell = row->GetCells()->GetItemInCellCollection(3);

			//Paragrapgh
			paragraph = cell->GetFirstParagraph();

			//Add text
			text = paragraph->AppendText(modificationForms[i][j].c_str());

			//Text format
			text->GetCharacterFormat()->SetFontName(L"Times New Roman");
			text->GetCharacterFormat()->SetFontSize(11);

			if (j < modificationForms[i].size() - 1)
			{
				//Add text
				text = paragraph->AppendText(L", ");

				//Text format
				text->GetCharacterFormat()->SetFontName(L"Times New Roman");
				text->GetCharacterFormat()->SetFontSize(11);
			}
		}
	}

	//Result file name
	wstring resFileName = wstring(L"Направление на прохождение МКК_").append(name.c_str()).append(L".docx");

	//Save the result document
	doc->SaveToFile(resFileName.c_str(), FileFormat::Docx2013);
	doc->Close();
}

int main()
{
	vector<vector<wstring>> repairForms{
		{L"Ф09", L"Ф23"},
		{L"Ф09", L"Ф23"},
		{L"-"},
		{L"-"}
	};
	vector<vector<wstring>> modificationForms{
		{L"Ф09"},
		{L"Ф09"},
		{L"-"},
		{L"Ф07", L"Ф23"}
		};

	fill_MKK_form(	L"Волков Александр Юрьевич",
					L"Инженер-конструктор",
					L"19.12.2023 г.",
					L"9 лет.",
					repairForms,
					modificationForms);
	return 0;
}
