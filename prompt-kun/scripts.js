function listFiles(event) {
    // console.log("[DEBUG] listFiles called.");
    const files = event.target.files;
    const fileListContainer = document.getElementById("file-list-container");
    fileListContainer.innerHTML = ""; // リストをクリア

    // ファイルリストを配列に変換してソート
    const fileList = Array.from(files).sort((a, b) => {
        const textA = a.webkitRelativePath || a.name;
        const textB = b.webkitRelativePath || b.name;
        return textA.localeCompare(textB);
    });

    // console.log("[DEBUG] Sorted fileList:", fileList.map(f => f.name));

    for (const file of fileList) {
        const listItem = document.createElement("div");
        listItem.className = "file-item";
        // フルパスを表示するためにfile.webkitRelativePathを使用
        listItem.textContent = file.webkitRelativePath || file.name;
        listItem.onclick = () => {
            // console.log("[DEBUG] listItem clicked:", file.name);
            loadFile(file);
        };
        fileListContainer.appendChild(listItem);
    }
}

function loadFile(file) {
    // console.log("[DEBUG] loadFile called with file:", file.name);
    // ▼▼▼ バグ修正：FileReaderを2つ用いるようにし、onloadを上書きしない形へ修正 ▼▼▼
    const reader = new FileReader();
    reader.onload = function (e) {
        // console.log("[DEBUG] Initial FileReader onload for ArrayBuffer:", file.name);
        const arrayBuffer = e.target.result;
        const uint8Array = new Uint8Array(arrayBuffer);

        // ファイルのエンコーディングを検出
        const detectedEncoding = Encoding.detect(uint8Array);
        // console.log("[DEBUG] Detected encoding:", detectedEncoding, "for file:", file.name);

        // 新処理：別のFileReaderを用いてテキスト読み込みを行う
        const readerText = new FileReader();
        readerText.onload = function (e) {
            // console.log("[DEBUG] Second FileReader onload (readAsText):", file.name);
            let fileContent = e.target.result;
            const blocks = fileContent.split("★★★★★");
            // console.log("[DEBUG] File content blocks:", blocks.length, "blocks.");

            document.getElementById("prompt").value = blocks[0] ? blocks[0].trim() : "";
            document.getElementById("description").innerText = blocks[1] ? blocks[1].trim() : "";
            document.getElementById("input-files").innerText = blocks[2] ? blocks[2].trim() : "";
            document.getElementById("selected-file").textContent = file.name;
        };
        readerText.readAsText(file, detectedEncoding);
    };
    reader.onerror = function (err) {
        console.error("[DEBUG] Error in loadFile (ArrayBuffer phase):", err);
    };
    reader.readAsArrayBuffer(file);
    // ▲▲▲ バグ修正ここまで ▲▲▲
}

function dropHandler(event) {
    // console.log("[DEBUG] dropHandler called.");
    event.preventDefault();
    const files = event.dataTransfer.files;
    const droppedFiles = document.getElementById("dropped-files");
    for (const file of files) {
        // console.log("[DEBUG] File dropped:", file.name);
        const listItem = document.createElement("div");
        listItem.className = "file-item";
        listItem.textContent = file.name;
        listItem.file = file; // fileオブジェクトを要素に紐付ける
        const deleteBtn = document.createElement("span");
        deleteBtn.textContent = " [削除]";
        deleteBtn.className = "delete-btn";
        deleteBtn.onclick = function () {
            // console.log("[DEBUG] Delete clicked for:", file.name);
            droppedFiles.removeChild(listItem);
        };
        listItem.appendChild(deleteBtn);
        droppedFiles.appendChild(listItem);
    }
}

function dragOverHandler(event) {
    event.preventDefault();
    // console.log("[DEBUG] dragOverHandler called.");
}

// ▼▼▼ Excel, Word, PDFを判定するための拡張子判定関数 ▼▼▼
function isExcelFile(file) {
    return /\.(xlsx|xls|xlsm)$/i.test(file.name);
}
function isWordFile(file) {
    // 一般的にはdocx解析を想定。docの場合はライブラリで失敗することがあります。
    return /\.(doc|docx)$/i.test(file.name);
}
function isPDFFile(file) {
    return /\.pdf$/i.test(file.name);
}

// ▼▼▼ ★★★ 追加: PowerPointファイルの拡張子判定関数 ★★★
function isPowerPointFile(file) {
    return /\.(ppt|pptx)$/i.test(file.name);
}

// ▼▼▼ 修正版: 図形の中の文言抽出を "async" で行う ▼▼▼

// ★★★ 3-1. JSZipベースの場合に対応するため、asyncで文字列を取り出す関数を新設 ★★★
// ★ (変更前) getStringFromZipFileAsync 内の「unsupported fileObj structure」だった部分を改修
async function getStringFromZipFileAsync(fileObj) {
    if (!fileObj) {
      // console.log("[DEBUG] getStringFromZipFileAsync called with null fileObj.");
      return "";
    }

    // 1) JSZip形式なら async("string") が使えるケース
    if (typeof fileObj.async === "function") {
      // console.log("[DEBUG] getStringFromZipFileAsync: Using fileObj.async('string').");
      try {
        const str = await fileObj.async("string");
        return str;
      } catch (err) {
        console.error("[DEBUG] getStringFromZipFileAsync: Error in async('string')", err);
        return "";
      }
    }

    // 2) asNodeBuffer がある場合 (旧SheetJS構造)
    if (fileObj.asNodeBuffer) {
      // console.log("[DEBUG] getStringFromZipFileAsync: Using asNodeBuffer().");
      const buffer = fileObj.asNodeBuffer();
      return new TextDecoder().decode(buffer);
    }

    // 3) _data.getContent がある場合 (さらに古い構造)
    if (fileObj._data && fileObj._data.getContent) {
      // console.log("[DEBUG] getStringFromZipFileAsync: Using _data.getContent().");
      const buffer = fileObj._data.getContent();
      return new TextDecoder().decode(buffer);
    }

    // 4) ★ 今回の問題パターン: { name, type, content: Uint8Array, ... } の生圧縮データ
    if (fileObj.content && fileObj.content instanceof Uint8Array) {
      // console.log("[DEBUG] getStringFromZipFileAsync: Attempting pako.inflate on fileObj.content.");
      try {
        // pako.inflate で解凍 (バイナリ -> バイナリ)
        const inflated = pako.inflate(fileObj.content);
        // 解凍結果をテキスト化
        const xmlString = new TextDecoder().decode(inflated);
        return xmlString;
      } catch (inflateErr) {
        console.error("[DEBUG] pako.inflate failed. Maybe not compressed or unknown format:", inflateErr);
        // もし失敗したら、「そのままテキスト化」を試す
        // console.log("[DEBUG] Trying direct TextDecoder on raw content...");
        try {
          return new TextDecoder().decode(fileObj.content);
        } catch (decodeErr) {
          console.error("[DEBUG] Direct decode also failed:", decodeErr);
          return "";
        }
      }
    }

    // それでも該当しない場合
    // console.log("[DEBUG] getStringFromZipFileAsync: unsupported fileObj structure. Logging fileObj below:");
    // console.log(fileObj);
    return "";
}

// シート名マップを作成する関数を修正
async function createSheetNameMap(workbook) {
    const sheetNameMap = new Map();
    
    if (!workbook || !workbook.files) return sheetNameMap;
    
    try {
        const workbookFile = workbook.files['xl/workbook.xml'];
        if (!workbookFile) return sheetNameMap;
        
        const xmlString = await getStringFromZipFileAsync(workbookFile);
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
        
        // シート情報を取得
        const sheets = xmlDoc.getElementsByTagName('sheet');
        for (let i = 0; i < sheets.length; i++) {
            const sheet = sheets[i];
            const sheetName = sheet.getAttribute('name');
            // インデックスを0から始まる番号として使用（i）
            if (sheetName) {
                sheetNameMap.set(String(i), sheetName);
            }
        }
    } catch (err) {
        console.error("[DEBUG] Error creating sheet name map:", err);
    }
    
    return sheetNameMap;
}

// 図形テキスト抽出関数を修正
async function extractShapeTextFromWorkbookAsync(workbook) {
    let shapeText = "";
    if (!workbook || !workbook.files) return shapeText;

    // シート名マップを作成
    const sheetNameMap = await createSheetNameMap(workbook);
    
    const fileNames = Object.keys(workbook.files);

    for (const fileName of fileNames) {
        if (
            /^xl\/drawings\/drawing\d+\.xml$/i.test(fileName) ||
            /^xl\/drawings\/vmlDrawing\d+\.vml$/i.test(fileName)
        ) {
            const fileObj = workbook.files[fileName];
            if (!fileObj) continue;

            // drawingXMLとシートの関係を取得
            const drawingMatch = fileName.match(/drawing(\d+)\.(?:xml|vml)$/i);
            const drawingNum = drawingMatch ? drawingMatch[1] : null;
            let sheetName = "不明なシート";

            if (drawingNum) {
                sheetName = sheetNameMap.get(drawingNum) || "不明なシート";
            }

            const xmlString = await getStringFromZipFileAsync(fileObj);
            shapeText += parseShapeXml(xmlString, fileName, sheetName);
        }
    }

    return shapeText;
}

// parseShapeXml関数を修正
function parseShapeXml(xmlString, fileName, sheetName) {
    let result = "";
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "application/xml");

    if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
        console.warn(`XML parse error in ${fileName}`);
        return result;
    }

    const aTList = xmlDoc.getElementsByTagName("a:t");
    if (aTList.length > 0) {
        result += `【Shapes in "${sheetName}"】\n`;
        for (let i = 0; i < aTList.length; i++) {
            const text = aTList[i].textContent;
            if (text.trim()) {
                result += text + "\n";
            }
        }
        result += "\n";
    }

    // 2) VML (vmlDrawing.vml) 内の <v:shape>～<v:textbox>～ テキストを取り出す
    const shapeList = xmlDoc.getElementsByTagName("v:shape");
    if (shapeList.length > 0) {
        // console.log(`[DEBUG] Found <v:shape> elements in ${fileName}:`, shapeList.length);
        result += `【VML Shapes in ${fileName}】\n`;
        for (let i = 0; i < shapeList.length; i++) {
            const shape = shapeList[i];
            // テキストボックス要素を探す
            const textBoxList = shape.getElementsByTagName("v:textbox");
            if (textBoxList.length > 0) {
                // console.log(`[DEBUG] Found <v:textbox> in shape #${i}:`, textBoxList.length);
                // さらに中のdivなどをテキスト化
                for (let j = 0; j < textBoxList.length; j++) {
                    const tb = textBoxList[j];
                    const innerText = tb.textContent;
                    if (innerText.trim()) {
                        result += innerText + "\n";
                    }
                }
            }
        }
        result += "\n";
    }

    return result;
}

// ▼▼▼ 3-3. Excel ファイル読み込みを async 化し、図形内テキストも抽出 ▼▼▼
async function readExcelFile(file) {
    // console.log("[DEBUG] readExcelFile called (async):", file.name);
    const data = new Uint8Array(await file.arrayBuffer());  // FileReader不要でも読み込めるが、従来通りでもOK
    try {
        // コメント取得のため cellComments: true を指定
        // 図形内文言も取得するため bookFiles: true を指定
        // console.log("[DEBUG] About to XLSX.read (async) ...", file.name);
        const workbook = XLSX.read(data, {
            type: 'array',
            cellComments: true,
            bookFiles: true
        });
        // console.log("[DEBUG] XLSX.read complete:", file.name);

        let text = "";
        workbook.SheetNames.forEach(sheetName => {
            // console.log("[DEBUG] Processing sheet:", sheetName);
            const sheet = workbook.Sheets[sheetName];
            // シートごとにTSV化（区切り文字をタブに設定）
            const tsv = XLSX.utils.sheet_to_csv(sheet, { FS: '\t' });
            text += `【Sheet: ${sheetName}】\n${tsv}\n\n\n`;

            // シート内のコメントも出力する
            if (sheet["!comments"] && sheet["!comments"].length > 0) {
                // console.log("[DEBUG] Found comments in sheet:", sheetName);
                text += `【Comments in ${sheetName}】\n`;
                sheet["!comments"].forEach(comment => {
                    const author = comment.a || "unknown";
                    text += `Cell ${comment.ref} (by ${author}): ${comment.t}\n`;
                });
                text += "\n";
            }
        });

        // 図形(Shapes)の中のテキストを取得（非同期）
        const shapeText = await extractShapeTextFromWorkbookAsync(workbook);
        if (shapeText.trim()) {
            // console.log("[DEBUG] shapeText extracted length:", shapeText.length);
            text += `【Shapes Overall】\n${shapeText}\n`;
        } else {
            // console.log("[DEBUG] No shapeText extracted.");
        }

        return text;
    } catch (error) {
        console.error("[DEBUG] Error in readExcelFile:", error);
        throw error;
    }
}

// ▼▼▼ Word(docx) ファイルのテキストを、できる限りレイアウトを再現して抽出＋図形抽出＋空行削除＋階層インデント ▼▼▼
function readWordFile(file) {
    // console.log("[DEBUG] readWordFile called:", file.name);
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function (e) {
            // console.log("[DEBUG] FileReader onload for Word file:", file.name);
            try {
                const arrayBuffer = e.target.result;

                // 1) Mammoth でHTML化（段落・リスト・テーブル・見出しなどをなるべく保持）
                const result = await mammoth.convertToHtml({ arrayBuffer });
                // console.log("[DEBUG] Mammoth convertToHtml success:", file.name);
                let html = result.value || "";

                // 2) HTML を解析してテキスト化（段落/リスト/テーブル/見出し など階層構造をインデント）
                let text = convertMammothHtmlToTextPreserveLayout(html);

                // 3) Word特有の「1行おきに空行」を防ぐため、重複改行を削除
                text = text.replace(/\n\s*\n/g, "\n");

                // 4) docx 内の図形テキストを抽出
                const shapeText = await extractShapesFromDocx(arrayBuffer);
                if (shapeText.trim()) {
                    text += "\n【Shapes in Word】\n" + shapeText + "\n";
                }

                resolve(text);
            } catch (error) {
                console.error("[DEBUG] Error in readWordFile (mammoth or shape extraction):", error);
                reject(error);
            }
        };
        reader.onerror = function (error) {
            console.error("[DEBUG] FileReader error in readWordFile:", error);
            reject(error);
        };
        reader.readAsArrayBuffer(file);
    });
}

// ▼▼▼ docx の図形抽出用 (Word も zip 構造のためJSZipで drawingsを解析) ▼▼▼
async function extractShapesFromDocx(arrayBuffer) {
    let shapeText = "";
    try {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const fileNames = Object.keys(zip.files);
        for (const fileName of fileNames) {
            if (
                /^word\/drawings\/drawing\d+\.xml$/i.test(fileName) ||
                /^word\/drawings\/vmlDrawing\d+\.vml$/i.test(fileName)
            ) {
                // console.log("[DEBUG] Found docx drawing file:", fileName);
                const fileObj = zip.files[fileName];
                if (!fileObj) continue;
                const xmlString = await fileObj.async("string");
                shapeText += parseShapeXml(xmlString, fileName);
            }
        }
    } catch (err) {
        console.error("[DEBUG] Error in extractShapesFromDocx:", err);
    }
    return shapeText;
}

// ▼▼▼ HTML(段落/リスト/テーブル/見出し等)をなるべくテキストとして再現し、階層をインデントする ▼▼▼
function convertMammothHtmlToTextPreserveLayout(htmlString) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlString, "text/html");

    // 再帰的にノードをたどり、タグ階層を indentLevel で管理する
    function walk(node, indentLevel = 0) {
        let out = "";

        if (node.nodeType === Node.TEXT_NODE) {
            // テキストノードはそのまま返す
            return node.nodeValue;
        } else if (node.nodeType === Node.ELEMENT_NODE) {
            const tag = node.tagName.toLowerCase();
            switch (tag) {
                // 見出し(h1~h6)はタグに応じたレベルでインデントを変化させる
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6": {
                    const headingLevel = parseInt(tag.substring(1), 10);
                    // headingLevel - 1 でインデント幅を決定
                    const headingIndent = headingLevel - 1;
                    const content = Array.from(node.childNodes).map(n => walk(n, headingIndent)).join("").trim();
                    // 見出し前に改行を入れ、インデントして出力、さらに改行
                    out += "\n" + getIndentStr(headingIndent) + content + "\n";
                    break;
                }
                case "p": {
                    // 段落
                    const content = Array.from(node.childNodes).map(n => walk(n, indentLevel)).join("");
                    out += getIndentStr(indentLevel) + content + "\n";
                    break;
                }
                case "br":
                    out += "\n";
                    break;
                case "ul":
                case "ol": {
                    // リストの場合、子要素(li)は indentLevel+1
                    const children = Array.from(node.childNodes).map(n => walk(n, indentLevel + 1)).join("");
                    out += children;
                    out += "\n"; // リスト全体の後ろで改行
                    break;
                }
                case "li": {
                    // リスト項目。深い階層はさらにインデント
                    const content = Array.from(node.childNodes).map(n => walk(n, indentLevel)).join("");
                    out += getIndentStr(indentLevel) + "• " + content + "\n";
                    break;
                }
                case "table": {
                    // テーブル
                    const children = Array.from(node.childNodes).map(n => walk(n, indentLevel)).join("");
                    out += children + "\n";
                    break;
                }
                case "tr": {
                    // 行
                    // 各セルを | で区切って1行に
                    const cells = Array.from(node.children).map((cell) => walk(cell, indentLevel).trim());
                    out += getIndentStr(indentLevel) + cells.join(" | ") + "\n";
                    break;
                }
                case "td":
                case "th": {
                    // セル内の要素を連結
                    const content = Array.from(node.childNodes).map(n => walk(n, indentLevel)).join("");
                    out += content;
                    break;
                }
                default: {
                    // その他のタグ (span, strong, em, div, 等) は子要素を再帰処理
                    const content = Array.from(node.childNodes).map(n => walk(n, indentLevel)).join("");
                    out += content;
                    break;
                }
            }
        }
        return out;
    }

    function getIndentStr(level) {
        // インデント幅はレベル × 2スペース程度
        return "  ".repeat(level);
    }

    return walk(doc.body, 0).trim();
}

// PDF ファイルのテキストを抽出（PDFの元レイアウトを反映するよう改良）
function extractPDF(file) {
    // console.log("[DEBUG] extractPDF called:", file.name);
    return new Promise((resolve, reject) => {
        const fileReader = new FileReader();
        fileReader.onload = function () {
            // console.log("[DEBUG] FileReader onload for PDF file:", file.name);
            const typedarray = new Uint8Array(this.result);
            pdfjsLib.getDocument(typedarray).promise.then(pdf => {
                // console.log("[DEBUG] PDF loaded. numPages:", pdf.numPages, "File:", file.name);
                const maxPages = pdf.numPages;
                const pageTextPromises = [];
                for (let pageNum = 1; pageNum <= maxPages; pageNum++) {
                    pageTextPromises.push(
                        pdf.getPage(pageNum).then(page => {
                            return page.getTextContent().then(textContent => {
                                let textItems = textContent.items;
                                // テキストアイテムをY座標（降順）とX座標（昇順）でソート
                                textItems.sort((a, b) => {
                                    const yDiff = b.transform[5] - a.transform[5];
                                    if (Math.abs(yDiff) < 5) {
                                        return a.transform[4] - b.transform[4];
                                    }
                                    return yDiff;
                                });

                                // Y座標の近さでアイテムをグループ化して行ごとにまとめる
                                let lines = [];
                                let currentLine = [];
                                let currentY = null;
                                textItems.forEach(item => {
                                    const itemY = item.transform[5];
                                    if (currentY === null || Math.abs(itemY - currentY) < 5) {
                                        currentLine.push(item);
                                        if (currentY === null) currentY = itemY;
                                    } else {
                                        // 現在の行内をX座標でソート（昇順）して連結
                                        currentLine.sort((a, b) => a.transform[4] - b.transform[4]);
                                        let lineText = "";
                                        for (let i = 0; i < currentLine.length; i++) {
                                            if (i > 0) {
                                                const prevItem = currentLine[i - 1];
                                                const gap = currentLine[i].transform[4] - (prevItem.transform[4] + (prevItem.width || 0));
                                                if (gap > 5) {
                                                    const numSpaces = Math.max(1, Math.floor(gap / 5));
                                                    lineText += " ".repeat(numSpaces);
                                                }
                                            }
                                            lineText += currentLine[i].str;
                                        }
                                        lines.push(lineText);
                                        // 新しい行の開始
                                        currentLine = [item];
                                        currentY = itemY;
                                    }
                                });
                                // 残った最後の行を処理
                                if (currentLine.length > 0) {
                                    currentLine.sort((a, b) => a.transform[4] - b.transform[4]);
                                    let lineText = "";
                                    for (let i = 0; i < currentLine.length; i++) {
                                        if (i > 0) {
                                            const prevItem = currentLine[i - 1];
                                            const gap = currentLine[i].transform[4] - (prevItem.transform[4] + (prevItem.width || 0));
                                            if (gap > 5) {
                                                const numSpaces = Math.max(1, Math.floor(gap / 5));
                                                lineText += " ".repeat(numSpaces);
                                            }
                                        }
                                        lineText += currentLine[i].str;
                                    }
                                    lines.push(lineText);
                                }
                                // 行ごとに改行文字で連結してページのテキストとする
                                return lines.join("\n");
                            });
                        })
                    );
                }
                Promise.all(pageTextPromises).then(pagesText => {
                    // ページ間はダブル改行で区切る
                    const fullText = pagesText.join("\n\n");
                    // console.log("[DEBUG] PDF text extraction finished:", file.name);
                    resolve(fullText);
                }).catch(err => {
                    console.error("[DEBUG] Error in pageTextPromises:", err);
                    reject(err);
                });
            }).catch(err => {
                console.error("[DEBUG] Error in pdfjsLib.getDocument:", err);
                reject(err);
            });
        };
        fileReader.onerror = function (error) {
            console.error("[DEBUG] FileReader error in extractPDF:", error);
            reject(error);
        };
        fileReader.readAsArrayBuffer(file);
    });
}

// ▼▼▼ ★★★ 追加: PowerPointファイルの読み込み処理 (ppt / pptx) ★★★ ▼▼▼
function readPowerPointFile(file) {
    // console.log("[DEBUG] readPowerPointFile called:", file.name);
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            // console.log("[DEBUG] FileReader onload for PowerPoint file:", file.name);
            try {
                const arrayBuffer = e.target.result;
                const zip = await JSZip.loadAsync(arrayBuffer);

                // スライド・ノートのテキストを抽出
                let text = await extractSlidesAndNotesFromPptx(zip);

                // 図形のテキストを抽出
                const shapeText = await extractShapesFromPptx(zip);
                if (shapeText.trim()) {
                    text += "\n【Shapes in PowerPoint】\n" + shapeText + "\n";
                }

                resolve(text);
            } catch (err) {
                console.error("[DEBUG] readPowerPointFile error:", err);
                reject(err);
            }
        };
        reader.onerror = function (error) {
            console.error("[DEBUG] FileReader error in readPowerPointFile:", error);
            reject(error);
        };
        reader.readAsArrayBuffer(file);
    });
}

async function extractSlidesAndNotesFromPptx(zip) {
    let text = "";
    const fileNames = Object.keys(zip.files);

    // スライド (ppt/slides/slideN.xml)
    const slideFileNames = fileNames.filter(fn => /^ppt\/slides\/slide\d+\.xml$/i.test(fn));
    for (const fileName of slideFileNames) {
        const fileObj = zip.files[fileName];
        if (!fileObj) continue;
        const xmlString = await fileObj.async("string");
        const slideText = parseSlideXml(xmlString, fileName);
        if (slideText.trim()) {
            text += `【Slide: ${fileName}】\n${slideText}\n\n`;
        }
    }

    // ノート (ppt/notesSlides/notesSlideN.xml)
    const notesFileNames = fileNames.filter(fn => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(fn));
    for (const fileName of notesFileNames) {
        const fileObj = zip.files[fileName];
        if (!fileObj) continue;
        const xmlString = await fileObj.async("string");
        const notesText = parseSlideXml(xmlString, fileName);
        if (notesText.trim()) {
            text += `【Notes: ${fileName}】\n${notesText}\n\n`;
        }
    }

    return text;
}

function parseSlideXml(xmlString, fileName) {
    let result = "";
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "application/xml");

    if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
        console.warn(`XML parse error in ${fileName}`);
        return result;
    }

    // <a:t> 要素を抽出
    const aTList = xmlDoc.getElementsByTagName("a:t");
    for (let i = 0; i < aTList.length; i++) {
        const text = aTList[i].textContent;
        if (text.trim()) {
            // console.log("[DEBUG] text:", text);
            result += text + "\n";
        }
    }
    return result;
}

async function extractShapesFromPptx(zip) {
    let shapeText = "";
    try {
        const fileNames = Object.keys(zip.files);
        for (const fileName of fileNames) {
            if (
                /^ppt\/drawings\/drawing\d+\.xml$/i.test(fileName) ||
                /^ppt\/drawings\/vmlDrawing\d+\.vml$/i.test(fileName)
            ) {
                // console.log("[DEBUG] Found ppt drawing file:", fileName);
                const fileObj = zip.files[fileName];
                if (!fileObj) continue;
                const xmlString = await fileObj.async("string");
                shapeText += parseShapeXml(xmlString, fileName);
            }
        }
    } catch (err) {
        console.error("[DEBUG] Error in extractShapesFromPptx:", err);
    }
    return shapeText;
}

// ▼▼▼ copyToClipboard: ドロップされたファイルの読込結果をすべて連結してコピー ▼▼▼
function copyToClipboard() {
    // console.log("[DEBUG] copyToClipboard called.");


    const content = document.getElementById("prompt").value;
    // ●●●が含まれる場合は処理を中止
    if (content.includes("●●●")) {
        alert("プロンプトの「●●●」の部分を書き換えてください。");
        return;
    }
    const allContent = [content];
    const droppedFiles = document.querySelectorAll("#dropped-files .file-item");
    let filesToRead = droppedFiles.length;
    // console.log("[DEBUG] filesToRead:", filesToRead);

    if (filesToRead === 0) {
        // console.log("[DEBUG] No dropped files, copying only the prompt.");
        copyText(allContent.join("\n"));
    } else {
        droppedFiles.forEach((item) => {
            const file = item.file;
            // console.log("[DEBUG] Processing dropped file:", file.name);

            // ▼▼▼ 拡張子別に処理を振り分け ▼▼▼
            if (isExcelFile(file)) {
                // console.log("[DEBUG] Detected Excel file:", file.name);
                readExcelFile(file).then((fileContent) => {
                    allContent.push("");
                    allContent.push(`${file.name}`);
                    allContent.push("");
                    allContent.push(fileContent);
                    filesToRead--;
                    // console.log("[DEBUG] Excel done, remaining:", filesToRead);
                    if (filesToRead === 0) {
                        // console.log("[DEBUG] All files done. Copy to clipboard now.");
                        copyText(allContent.join("\n"));
                    }
                }).catch(e => {
                    console.error("[DEBUG] readExcelFile error:", e);
                    filesToRead--;
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                });
            } else if (isWordFile(file)) {
                // console.log("[DEBUG] Detected Word file:", file.name);
                readWordFile(file).then((fileContent) => {
                    allContent.push("");
                    allContent.push(`${file.name}`);
                    allContent.push("");
                    allContent.push(fileContent);
                    filesToRead--;
                    // console.log("[DEBUG] Word done, remaining:", filesToRead);
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                }).catch(e => {
                    console.error("[DEBUG] readWordFile error:", e);
                    filesToRead--;
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                });
            } else if (isPDFFile(file)) {
                // console.log("[DEBUG] Detected PDF file:", file.name);
                extractPDF(file).then((fileContent) => {
                    allContent.push("");
                    allContent.push(`${file.name}`);
                    allContent.push("");
                    allContent.push(fileContent);
                    filesToRead--;
                    // console.log("[DEBUG] PDF done, remaining:", filesToRead);
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                }).catch(e => {
                    console.error("[DEBUG] extractPDF error:", e);
                    filesToRead--;
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                });
            // ▼▼▼ ★★★ 追加: PowerPointファイルの場合 ★★★
            } else if (isPowerPointFile(file)) {
                // console.log("[DEBUG] Detected PowerPoint file:", file.name);
                readPowerPointFile(file).then((fileContent) => {
                    allContent.push("");
                    allContent.push(`${file.name}`);
                    allContent.push("");
                    allContent.push(fileContent);
                    filesToRead--;
                    // console.log("[DEBUG] PowerPoint done, remaining:", filesToRead);
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                }).catch(e => {
                    console.error("[DEBUG] readPowerPointFile error:", e);
                    filesToRead--;
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                });
            } else {
                // console.log("[DEBUG] Detected text/other file:", file.name);
                // ▼▼▼ 既存のテキストファイル読み取り処理はそのまま ▼▼▼
                const reader = new FileReader();
                reader.readAsArrayBuffer(file);
                reader.onload = function (e) {
                    // console.log("[DEBUG] TextFile ArrayBuffer loaded:", file.name);
                    const arrayBuffer = e.target.result;
                    const uint8Array = new Uint8Array(arrayBuffer);

                    // ファイルのエンコーディングを検出
                    const detectedEncoding = Encoding.detect(uint8Array);
                    // console.log("[DEBUG] Detected encoding (text file):", detectedEncoding, file.name);

                    reader.readAsText(file, detectedEncoding);
                    reader.onload = function (e) {
                        // console.log("[DEBUG] textFile readAsText complete:", file.name);
                        let fileContent = e.target.result;
                        // ファイルパスを追加
                        allContent.push("");
                        allContent.push(`${file.name}`); // ブラウザのセキュリティ上、フルパスは取得できない
                        allContent.push("");
                        allContent.push(fileContent);
                        filesToRead--;
                        // console.log("[DEBUG] Text file done, remaining:", filesToRead);
                        if (filesToRead === 0) {
                            copyText(allContent.join("\n"));
                        }
                    };
                };
                reader.onerror = function (err) {
                    console.error("[DEBUG] Error reading text/other file:", err);
                    filesToRead--;
                    if (filesToRead === 0) {
                        copyText(allContent.join("\n"));
                    }
                };
            }
        });
    }
}

// ▼▼▼ copyText: テキストをクリップボードにコピーし、Chat AI を開く ▼▼▼
function copyText(text) {
    // 生成されるプロンプトの冒頭に「今日の日付:YYYY/MM/DD」を出力する
    //const now = new Date();
    //const year = now.getFullYear();
    //const month = ("0" + (now.getMonth() + 1)).slice(-2);
    //const day = ("0" + now.getDate()).slice(-2);
    //const currentDate = `${year}/${month}/${day}`;
    
    //const header = `今日の日付:${currentDate}\n`;
    const header = "";
    const footer = `\n\n以上がインプット情報です。冒頭の指示に従ってください。\n\n`;
    let finalText = header + text + footer;

    // ▼▼▼ ここから追加：ハードコーディングされた置換処理 ▼▼▼
    const replacements = [
        { before: "置換前ワード1", after: "置換後ワード1" },
        { before: "置換前ワード2", after: "置換後ワード2" }
        // 必要な分だけ追加
    ];
    for (const { before, after } of replacements) {
        // 全ての出現箇所を置換
        finalText = finalText.split(before).join(after);
    }
    // ▲▲▲ 置換処理ここまで ▲▲▲

    // console.log("[DEBUG] copyText called. Final text length:", finalText.length);
    navigator.clipboard.writeText(finalText).then(
        function () {
            // console.log("[DEBUG] Successfully copied text to clipboard. Opening Chat AI.");
            window.open("https://v2.scsk-gai.jp/", "_blank");
        },
        function () {
            alert("コピーに失敗しました！");
        }
    );
}

// ▼▼▼ クリアボタンで初期化 ▼▼▼
function clearAll() {
    // console.log("[DEBUG] clearAll called.");
    document.getElementById("prompt").value = "";
    document.getElementById("dropped-files").innerHTML = "";
    document.querySelector('input[type="file"]').value = "";
    document.getElementById("selected-file").textContent = "";
    document.getElementById("description").innerText = "";
    document.getElementById("input-files").innerText = "";
}
