# WordDocumentHelpers

Небольшая .NET 8 (`net8.0-windows`) библиотека с extension-методами для работы с Word-документами.

Два движка:
- **Microsoft.Office.Interop.Word** — методы, которым нужен установленный Word (подсчёт страниц, номер страницы абзаца по тексту). Interop-типы встраиваются (`EmbedInteropTypes`), чтобы не тянуть Office Core PIA в рантайме.
- **Spire.Doc (FreeSpire.Doc)** — замена плейсхолдеров без установленного Word.

## Установка (GitHub Packages)

Пакет публикуется в приватный GitHub Packages фид владельца `n1tr3x`:

```powershell
dotnet nuget add source https://nuget.pkg.github.com/n1tr3x/index.json `
  --name github-n1tr3x --username n1tr3x --password <GITHUB_PAT с read:packages>
dotnet add package WordDocumentHelpers
```

## API

| Метод | Движок | Назначение |
|---|---|---|
| `GetPagesCount(string filepath)` | Interop | Количество страниц в документе |
| `Document.GetParagraphPageByText(text)` | Interop | Номер страницы первого абзаца с текстом (`-1`, если не найдено) |
| `Document.GetParagraphPageByTextFromIndex(text, startParagraph)` | Interop | То же, начиная с 0-based индекса абзаца |
| `Document.GetParagraphsPages()` | Interop | Словарь «текст абзаца → страница» |
| `Spire.Doc.Document.ReplaceText(old, new)` | Spire | Замена текста по всем секциям и таблицам |

### Пример

```csharp
// Interop (нужен установленный Word)
var word = new Microsoft.Office.Interop.Word.Application();
var doc = word.Documents.Open(@"C:\contract.docx");
int page = doc.GetParagraphPageByText("Раздел 5");

// Spire (без Word)
var sdoc = new Spire.Doc.Document(@"C:\template.docx");
sdoc.ReplaceText("НОМЕР_ДЕЛА", "А40-12345/2026");
sdoc.SaveToFile(@"C:\out.docx");
```

## Замечания

- Interop-методы требуют установленного Microsoft Word.
- `FreeSpire.Doc` — бесплатная редакция со встроенными лимитами (страницы/абзацы) и возможным evaluation-водяным знаком на части операций.

## Публикация

Релиз публикуется автоматически: push git-тега вида `v1.2.3` запускает workflow
[`.github/workflows/publish.yml`](.github/workflows/publish.yml), который собирает
пакет с версией из тега и пушит его в GitHub Packages через `GITHUB_TOKEN`.
