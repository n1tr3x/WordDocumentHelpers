using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Interface;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentHelpers;

public static class WordDocumentHelper
{
	public static int GetPagesCount(string filepath)
	{
		Word.Application wordApplication = new Word.Application();
		wordApplication.Visible = true;
		Word.Document wordDocument = wordApplication.Documents.Open(filepath);

		int result = (int)wordDocument.Words.Last.Information[Word.WdInformation.wdActiveEndPageNumber];
		wordApplication.Quit();
		return result;
	}

	public static int GetParagraphPageByText(this Word.Document document, string text)
	{
		Word.Paragraph? paragraph = document.Paragraphs.Cast<Word.Paragraph>()
			.FirstOrDefault(paragraph => paragraph.Range.Text.Contains(text));
		if (paragraph is null)
			return -1;
		return (int)paragraph.Range.Information[Word.WdInformation.wdActiveEndPageNumber];
	}

	public static int GetParagraphPageByTextParallel(this Word.Document document, string text)
	{
		return int.Parse(document.Paragraphs.Cast<Word.Paragraph>().First(paragraph => paragraph.Range.Text.Contains(text)).Range.Information[Word.WdInformation.wdActiveEndPageNumber].ToString());
	}

	public static int GetParagraphPageByTextFromIndex(this Word.Document document, string text, int startParagraph)
	{
		List<Word.Paragraph> l = document.Paragraphs.Cast<Word.Paragraph>().ToList();
		for (int i = startParagraph; i < l.Count; i++)
			if (l[i].Range.Text.Contains(text))
				return (int)l[i].Range.Information[Word.WdInformation.wdActiveEndPageNumber];
		return -1;
	}

	public static Dictionary<string, int> GetParagraphsPages(this Word.Document document)
	{
		Dictionary<string, int> result = new();
		for (int i = 1; i <= document.Paragraphs.Count; i++)
		{
			Word.Range range = document.Paragraphs[i].Range;
			result[range.Text] = (int)range.Information[Word.WdInformation.wdActiveEndPageNumber];
		}
		return result;
	}

	public static int GetParagraphIndex(this Microsoft.Office.Interop.Word.Paragraph paragraph)
	{
		return (int)paragraph.Range.Information[Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber];
	}

	public static bool ReplaceText(this Spire.Doc.Document document, string oldText, string newText)
	{
		bool flag = false;
		foreach (Section section in document.Sections)
		{
			foreach (Paragraph paragraph in section.Paragraphs)
			{
				if (paragraph.Text.Contains(oldText))
				{
					flag = true;
					paragraph.Text = paragraph.Text.Replace(oldText, newText);
				}
			}
			foreach (ITable table in section.Tables)
			{
				foreach (TableRow tableRow in table.Rows)
				{
					foreach (TableCell tableRowCell in tableRow.Cells)
					{
						foreach (Paragraph paragraph in tableRowCell.Paragraphs)
						{
							if (paragraph.Text.Contains(oldText))
							{
								flag = true;
								paragraph.Text = paragraph.Text.Replace(oldText, newText);
							}
						}
					}
				}
			}
		}

		return flag;
	}
}