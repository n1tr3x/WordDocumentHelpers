# WordDocumentHelpers

Библиотека extension-методов для работы с Word-документами. Таргет — `netstandard2.0`
(минимум, который требуют зависимости), потребляется из любого современного .NET.

Два независимых движка:

- **Microsoft.Office.Interop.Word** — методы, автоматизирующие установленный Word через COM
  (подсчёт страниц, поиск номера страницы абзаца по тексту). Требуют установленного Word и
  работают только на Windows. Interop-типы встраиваются (`EmbedInteropTypes`).
- **Spire.Doc (FreeSpire.Doc)** — замена текста без установленного Word.

## Установка (GitHub Packages)

```powershell
dotnet nuget add source https://nuget.pkg.github.com/n1tr3x/index.json `
  --name github-n1tr3x --username n1tr3x --password <GITHUB_PAT с read:packages>
dotnet add package WordDocumentHelpers
```

## Функции

Все методы — статические; помеченные `this` являются extension-методами.

### `GetPagesCount`
```csharp
public static int GetPagesCount(string filepath)
```
Открывает документ в новом экземпляре Word и возвращает **общее число страниц**.
Движок: Interop. Запускает Word (`Visible = true`) и закрывает его перед возвратом.

### `GetParagraphPageByText`
```csharp
public static int GetParagraphPageByText(this Word.Document document, string text)
```
Возвращает **номер страницы первого абзаца**, текст которого содержит `text`.
Если совпадений нет — возвращает `-1`. Движок: Interop.

### `GetParagraphPageByTextParallel`
```csharp
public static int GetParagraphPageByTextParallel(this Word.Document document, string text)
```
То же, что `GetParagraphPageByText`, но при отсутствии совпадения **бросает**
`InvalidOperationException` (использует `First`, а не `FirstOrDefault`). Движок: Interop.
> Имя историческое — параллельной обработки внутри нет.

### `GetParagraphPageByTextFromIndex`
```csharp
public static int GetParagraphPageByTextFromIndex(this Word.Document document, string text, int startParagraph)
```
Ищет первый абзац с текстом `text`, **начиная с 0-based индекса** `startParagraph`,
и возвращает номер его страницы. Если до конца документа совпадений нет — `-1`. Движок: Interop.

### `GetParagraphsPages`
```csharp
public static Dictionary<string, int> GetParagraphsPages(this Word.Document document)
```
Строит словарь **«текст абзаца → номер страницы»** по всем абзацам документа.
При одинаковом тексте у нескольких абзацев остаётся последнее значение (last-wins). Движок: Interop.

### `GetParagraphIndex`
```csharp
public static int GetParagraphIndex(this Word.Paragraph paragraph)
```
Возвращает **номер страницы**, на которой заканчивается переданный абзац
(`wdActiveEndPageNumber`). Движок: Interop.
> Имя историческое — метод отдаёт номер страницы, а не индекс абзаца.

### `ReplaceText`
```csharp
public static bool ReplaceText(this Spire.Doc.Document document, string oldText, string newText)
```
Заменяет `oldText` на `newText` **во всех секциях**: в абзацах и в ячейках таблиц.
Возвращает `true`, если хотя бы одна замена произошла. Движок: Spire.Doc.

## Примеры

```csharp
// Interop (нужен установленный Word)
var word = new Microsoft.Office.Interop.Word.Application();
var doc = word.Documents.Open(@"C:\contract.docx");
int total = WordDocumentHelper.GetPagesCount(@"C:\contract.docx");
int page  = doc.GetParagraphPageByText("Раздел 5");        // -1, если не найдено
var map   = doc.GetParagraphsPages();                       // текст абзаца -> страница

// Spire (без Word)
var sdoc = new Spire.Doc.Document(@"C:\template.docx");
bool changed = sdoc.ReplaceText("НОМЕР_ДЕЛА", "А40-12345/2026");
sdoc.SaveToFile(@"C:\out.docx");
```