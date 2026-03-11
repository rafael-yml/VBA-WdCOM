# VBA-WdCOM

VBA module for text extraction and image export from Word documents and PDFs via Word's COM interface. Entirely late-bound, no VBA references needed beyond a Word installation.

Pairs with [VBA-PdfTXT](https://github.com/rafael-yml/VBA-PdfTXT), [VBA-PdfWRT](https://github.com/rafael-yml/VBA-PdfWRT), and [VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR) as part of a self-contained document processing pipeline, but each module is independently usable.

---

## When to use this

Word's PDF converter is a last-resort path for PDFs. It handles files that native PDF parsers struggle with, particularly PDFs originally authored in Word, PDFs with complex layout, or files where stream-level text extraction produces garbled output.

It is **not** appropriate as a general-purpose PDF extractor because:
- It spawns a hidden Word instance (~2-3s startup cost per call)
- Word's PDF-to-DOCX conversion is lossy and unpredictable on complex layouts
- It can stall indefinitely on pathologically malformed files (see Known Limitations)

Use `PDF_WordCOM` selectively, after faster methods have been tried. The open-document functions (`Word_ExtractText`, `Word_GetImages`, etc.) are appropriate any time you already have a Word document open via COM.

---

## Functions

### Pre-flight

#### `WdCOM_IsWordInstalled()` -> `Boolean`

Returns `True` if Word can be started via COM. Caches a successful result so subsequent calls are instant. Failure is **never cached**: a transient issue (Word mid-update, hung COM surrogate) won't poison the session.

```vb
If Not WdCOM_IsWordInstalled() Then Exit Sub
```

---

#### `Word_IsWordDocument(sFilePath)` -> `Boolean`

Returns `True` if the file is a Word document. Uses magic bytes for the primary check and the file extension to disambiguate: `.docx`, `.xlsx`, and `.pptx` all share the same ZIP header, so the extension is required to tell them apart.

- `.docx` / `.docm`: ZIP header + doc extension
- `.doc`: OLE compound document header (D0 CF 11 E0)

Does not require Word to be installed. Cheap pre-flight before handing an unknown file to Word.

```vb
If Not Word_IsWordDocument(sPath) Then Exit Sub
```

---

### PDF entry point

#### `PDF_WordCOM(sFilePath)` -> `String`

Opens a PDF in an isolated hidden Word instance, extracts `Document.Content.Text`, runs a quality check, and returns the cleaned text. Returns `""` on any failure. Always spawns and quits its own isolated Word instance, never touches a user's open Word session.

```vb
Dim sText As String
sText = PDF_WordCOM("C:\docs\report.pdf")
If WdCOM_LastStatus() <> WDCOM_OK Then
    Debug.Print "Failed with code: " & WdCOM_LastStatus()
End If
```

**Quality checks applied to a sample of the first 100,000 characters:**
- Text shorter than 2 characters: `""` (`WDCOM_EMPTY`)
- Average word length above 25: `""` (`WDCOM_GARBLED`, catches encoding failures)

Documents longer than 100,000 characters are **not rejected**, only the sample is tested.

**Encrypted PDFs** are detected by scanning the PDF header and trailer for `/Encrypt` before Word is launched, avoiding startup cost entirely.

---

#### `WdCOM_LastStatus()` -> `Long`

Returns the status code from the most recent `PDF_WordCOM` or `Word_ExtractText` call. Compare against the `WDCOM_*` constants:

| Constant | Value | Meaning |
|---|---|---|
| `WDCOM_OK` | 0 | Text extracted successfully |
| `WDCOM_FILE_MISSING` | 1 | File not found |
| `WDCOM_ENCRYPTED` | 2 | PDF is encrypted, Word cannot open it |
| `WDCOM_WORD_FAILED` | 3 | Word not installed or failed to start |
| `WDCOM_OPEN_FAILED` | 4 | `Documents.Open` raised an error |
| `WDCOM_EMPTY` | 5 | Extracted text too short to be useful |
| `WDCOM_GARBLED` | 6 | Average word length too high, encoding failure |

---

### Open document operations

These functions take an already-open `Word.Document` passed as `Object`. The caller is responsible for opening and closing the document.

#### `Word_ExtractText(oDoc)` -> `String`

Returns the cleaned text of an open document directly. Applies the same credibility check and line-ending normalisation as `PDF_WordCOM`. Sets `WdCOM_LastStatus()` so the caller can distinguish an empty document from a garbled one.

```vb
Dim sText As String
sText = Word_ExtractText(oDoc)
If WdCOM_LastStatus() = WDCOM_GARBLED Then Debug.Print "Encoding failure"
```

---

#### `Word_PageCount(oDoc, [bFast])` -> `Long`

Returns the page count of an open document. Returns `0` on failure.

- `bFast = False` (default): `ComputeStatistics(2)`, accurate but repaginates the document. Can take several seconds on a large file.
- `bFast = True`: `BuiltInDocumentProperties(14)`, instant but reflects the last-saved count and may be stale if the document has been edited since opening.

```vb
Dim lPages As Long
lPages = Word_PageCount(oDoc)               ' accurate
lPages = Word_PageCount(oDoc, bFast:=True)  ' instant
```

---

#### `Word_ExportImages(oDoc, sOutputFolder)`

Exports all images from an open document to numbered PNG files in `sOutputFolder`. The folder is created if it does not exist. Output files are named `word_img_1.png`, `word_img_2.png`, etc.

```vb
Word_ExportImages oDoc, "C:\output\images\"
```

Covers:
- `InlineShapes`: images anchored in the text flow
- Floating `Shapes` of type 13 (msoPicture), 11 (msoLinkedPicture), and 7 (msoEmbeddedOLEObject)

Text boxes, SmartArt, charts, and other drawing objects are intentionally skipped. Shapes that cannot be copied via `CopyAsPicture` are silently skipped.

**Excel host required.** Uses a hidden `ChartObject` for clipboard-to-PNG export.

---

#### `Word_GetImages(oDoc)` -> `Collection`

Returns a `Collection` of `Byte()` arrays, one per image in document order, without writing permanent files to disk. Each element contains the raw PNG bytes of one image. Returns an empty `Collection` on complete failure, never returns `Nothing`.

Designed for direct passthrough to [VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR)'s `BytesToText` function, or any other in-memory consumer.

```vb
Dim oImages As Collection
Set oImages = Word_GetImages(oDoc)

Dim img As Variant
For Each img In oImages
    Dim aBytes() As Byte
    aBytes = img
    Dim sText As String
    sText = oOCR.BytesToText(aBytes)   ' VBA-WinOCR
Next img
```

Covers the same shape types as `Word_ExportImages`.

**Excel host required.**

---

## Known limitations

**Stall on corrupt files.** `Documents.Open` is synchronous with no timeout mechanism in VBA. `OpenAndRepair:=False`, `ReadOnly:=True`, and `DisplayAlerts:=0` prevent the most common cause (Word's repair/recovery dialog loop) but cannot guarantee termination on all malformed inputs. If stall tolerance is critical, run from a separate Excel instance that can be force-terminated.

**`Word_PageCount` with `bFast:=False` repaginates.** On a large document this can take several seconds on the first call after opening.

**`Word_GetImages` touches disk transiently.** `Chart.Export` has no stream-based API, so each image is written to `%TEMP%` and immediately read back and deleted. No permanent files are left by the function under normal operation.

---

## Requirements

- Microsoft Word installed (except `Word_IsWordDocument`, which only reads file bytes)
- Excel host for `Word_ExportImages` and `Word_GetImages`
- No VBA library references needed, fully late-bound

---

## License

MIT License.

Copyright (c) 2026, [rafael-yml](https://rafael-yml.lovable.app/)
