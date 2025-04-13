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

function parseShapeXml(xmlString, fileName) {
    const results = []; // { row: number, col: number, text: string } の配列
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "application/xml");

    // ★★★ 名前空間の定義 (Excel DrawingMLで一般的に使われるもの) ★★★
    const xdrNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const aNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

    const parserError = xmlDoc.getElementsByTagName("parsererror")[0];
    if (parserError) {
        // console.warn(`XML parse error in ${fileName}:`, parserError.textContent); // デバッグ用ログは維持
        return results; // 空配列を返す
    }

    // 1) XML (drawing.xml) 内の <xdr:sp> からテキストと位置を取得
    if (fileName.toLowerCase().endsWith('.xml')) {
        // ★★★ 名前空間を指定して <xdr:sp> 要素を取得 ★★★
        const shapes = xmlDoc.getElementsByTagNameNS(xdrNamespace, "sp");
        // console.log(`[DEBUG] Found ${shapes.length} <xdr:sp> elements in ${fileName}`);

        for (let i = 0; i < shapes.length; i++) { // デバッグしやすいように for ループに変更
            const shape = shapes[i];
            // console.log(`[DEBUG] Processing shape #${i}:`, shape.outerHTML.substring(0, 300)); // 要素の構造を一部表示

             // ★★★ 名前空間を指定して <a:t> 要素を取得 ★★★
            const tElements = shape.getElementsByTagNameNS(aNamespace, "t");
            const originalText = Array.from(tElements).map(el => el.textContent).join("");

            // 1. 改行・タブをスペースに置換
            let cleanedText = originalText.replace(/[\n\t]+/g, ' ');
            // 2. 連続する空白を1つに正規化
            cleanedText = cleanedText.replace(/\s{2,}/g, ' ');
            // 3. 前後の空白を削除
            cleanedText = cleanedText.trim();

            if (cleanedText) { // 整形後のテキストが空でない場合のみ追加
                // console.log(`[DEBUG] Found cleaned text in shape #${i}: "${cleanedText}"`);
                let row = Infinity, col = Infinity; // デフォルト値（位置情報がない場合）
                let anchorElement = null; // twoCellAnchor or oneCellAnchor or null
                let anchorType = null; // 'twoCell' or 'oneCell' or null

                // 位置情報の取得ロジック
                // shape要素の直接の親要素を取得 (通常は twoCellAnchor, oneCellAnchor, grpSp のいずれか)
                const parentElement = shape.parentElement;

                if (parentElement) {
                    const parentTagName = parentElement.tagName.toLowerCase();
                    // console.log(`[DEBUG] Parent element tag name for shape #${i}: ${parentTagName}`);

                    // 親がアンカー要素の場合
                    if (parentTagName === 'xdr:twocellanchor' || parentTagName === 'twoCellAnchor') {
                        anchorElement = parentElement;
                        anchorType = 'twoCell';
                        // console.log(`[DEBUG] Anchor found in direct parent (twoCellAnchor) for shape #${i}`);
                    } else if (parentTagName === 'xdr:onecellanchor' || parentTagName === 'oneCellAnchor') {
                        anchorElement = parentElement;
                        anchorType = 'oneCell';
                        // console.log(`[DEBUG] Anchor found in direct parent (oneCellAnchor) for shape #${i}`);
                    }
                    // 親がグループ要素の場合、さらにその親のアンカーを探す
                    else if (parentTagName === 'xdr:grpsp' || parentTagName === 'grpSp') {
                        // console.log(`[DEBUG] Shape #${i} is inside a group.`);
                        const grandParentElement = parentElement.parentElement;
                        if (grandParentElement) {
                             const grandParentTagName = grandParentElement.tagName.toLowerCase();
                             // console.log(`[DEBUG] Grandparent element tag name for shape #${i} (inside group): ${grandParentTagName}`);
                             if (grandParentTagName === 'xdr:twocellanchor' || grandParentTagName === 'twoCellAnchor') {
                                 anchorElement = grandParentElement;
                                 anchorType = 'twoCell';
                                 // console.log(`[DEBUG] Anchor found in grandparent (twoCellAnchor) for grouped shape #${i}`);
                             } else if (grandParentTagName === 'xdr:onecellanchor' || grandParentTagName === 'oneCellAnchor') {
                                 anchorElement = grandParentElement;
                                 anchorType = 'oneCell';
                                 // console.log(`[DEBUG] Anchor found in grandparent (oneCellAnchor) for grouped shape #${i}`);
                             }
                        }
                    }
                } else {
                    // console.log(`[DEBUG] Shape #${i} has no parentElement.`);
                }

                // アンカー要素が見つかった場合、座標を取得 (変更なし)
                if (anchorElement && anchorType) {
                    const fromEl = anchorElement.getElementsByTagNameNS(xdrNamespace, "from")[0];
                    let fromRow = Infinity, fromCol = Infinity;
                    if (fromEl) {
                        const rowEl = fromEl.getElementsByTagNameNS(xdrNamespace, "row")[0];
                        const colEl = fromEl.getElementsByTagNameNS(xdrNamespace, "col")[0];
                        fromRow = parseInt(rowEl?.textContent || Infinity, 10);
                        fromCol = parseInt(colEl?.textContent || Infinity, 10);
                        if (isNaN(fromRow)) fromRow = Infinity;
                        if (isNaN(fromCol)) fromCol = Infinity;
                        // console.log(`[DEBUG] Parsed 'from': row=${fromRow}, col=${fromCol}`);
                    } else {
                        // console.log(`[DEBUG] 'from' element not found in anchor for shape #${i}.`);
                    }

                    // twoCellAnchor の場合は 'to' も取得して中央値を計算
                    if (anchorType === 'twoCell') {
                        const toEl = anchorElement.getElementsByTagNameNS(xdrNamespace, "to")[0];
                        let toRow = Infinity, toCol = Infinity;
                        if (toEl) {
                            const rowEl = toEl.getElementsByTagNameNS(xdrNamespace, "row")[0];
                            const colEl = toEl.getElementsByTagNameNS(xdrNamespace, "col")[0];
                            toRow = parseInt(rowEl?.textContent || Infinity, 10);
                            toCol = parseInt(colEl?.textContent || Infinity, 10);
                            if (isNaN(toRow)) toRow = Infinity;
                            if (isNaN(toCol)) toCol = Infinity;
                            // console.log(`[DEBUG] Parsed 'to': row=${toRow}, col=${toCol}`);
                        } else {
                            // console.log(`[DEBUG] 'to' element not found in twoCellAnchor for shape #${i}.`);
                        }

                        // 中央値を計算 (両方が有効な場合のみ)
                        if (fromRow !== Infinity && toRow !== Infinity) {
                            row = Math.floor((fromRow + toRow) / 2);
                        } else {
                            row = fromRow; // to がなければ from を使う
                        }
                        if (fromCol !== Infinity && toCol !== Infinity) {
                            col = Math.floor((fromCol + toCol) / 2);
                        } else {
                            col = fromCol; // to がなければ from を使う
                        }
                         // console.log(`[DEBUG] Calculated center (twoCell): row=${row}, col=${col}`);

                    }
                    // oneCellAnchor の場合は 'from' の値を使う
                    else if (anchorType === 'oneCell') {
                        row = fromRow;
                        col = fromCol;
                        // console.log(`[DEBUG] Using 'from' (oneCell): row=${row}, col=${col}`);
                    }
                } else {
                     // console.log(`[DEBUG] Anchor element not found for text "${cleanedText}" in shape #${i}. Using default (Infinity, Infinity).`);
                }

                // 最終的な row, col が Infinity でないことを確認 (念のため)
                if (row === Infinity || col === Infinity) {
                    // console.log(`[DEBUG] Final position is Infinity for text "${cleanedText}" in shape #${i}.`);
                }
                 // console.log(`[DEBUG] Final position for text "${cleanedText}": row=${row}, col=${col}`);
                results.push({ row, col, text: cleanedText });
            } else {
                 // console.log(`[DEBUG] Shape #${i} has no text content after cleaning.`);
            }
        }
    }
    // 2) VML (vmlDrawing.vml) 内の <v:shape>～<v:textbox>～ テキストを取り出す
    else if (fileName.toLowerCase().endsWith('.vml')) {
        // VMLの名前空間は通常自動で解決されることが多いが、必要なら追加
        // const vNamespace = "urn:schemas-microsoft-com:vml";
        const shapeList = xmlDoc.getElementsByTagName("v:shape"); // VMLは名前空間なしでも取得できることが多い
        if (shapeList.length > 0) {
            // result += `【VML Shapes in ${fileName}】\n`; // ヘッダー削除
            for (let i = 0; i < shapeList.length; i++) {
                const shape = shapeList[i];
                // テキストボックス要素を探す
                const textBoxList = shape.getElementsByTagName("v:textbox");
                if (textBoxList.length > 0) {
                    // console.log(`[DEBUG] Found <v:textbox> in shape #${i}:`, textBoxList.length);
                    // さらに中のdivなどをテキスト化
                    for (let j = 0; j < textBoxList.length; j++) {
                        const tb = textBoxList[j];
                        const originalInnerText = tb.textContent;

                        // 1. 改行・タブをスペースに置換
                        let cleanedInnerText = originalInnerText.replace(/[\n\t]+/g, ' ');
                        // 2. 連続する空白を1つに正規化
                        cleanedInnerText = cleanedInnerText.replace(/\s{2,}/g, ' ');
                        // 3. 前後の空白を削除
                        cleanedInnerText = cleanedInnerText.trim();

                        if (cleanedInnerText) { // 整形後のテキストが空でない場合
                             // VMLは位置情報がないので、row/colはInfinityとして追加
                            results.push({ row: Infinity, col: Infinity, text: cleanedInnerText });
                            // result += innerText + "\n"; // 元の文字列結合ロジック削除
                        }
                    }
                }
            }
            // result += "\n"; // 末尾の改行削除
        }
    }

    // ★重要: parseShapeXml内ではソートしない。呼び出し元でシートごとにまとめた後ソートする。
    // ★重要: parseShapeXml内ではシート名ヘッダーを追加しない。

    return results; // { row, col, text } の配列を返す
}


async function extractShapeTextFromWorkbookAsync(workbook) {
    const shapesBySheet = new Map(); // キー: シート名, 値: [{ row, col, text }] の配列

    if (!workbook || !workbook.files) return "";

    // 1. シート名とインデックスのマッピングを作成
    const sheetNameMap = await createSheetNameMap(workbook);
    // console.log("[DEBUG] sheetNameMap:", sheetNameMap);

    // 2. シートと描画ファイルの関係を解析
    const sheetDrawingRelMap = new Map(); // キー: drawing#.xml (e.g., 'drawing1.xml'), 値: シートインデックス (文字列, e.g., '0')
    const sheetRelFiles = Object.keys(workbook.files).filter(fn => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/i.test(fn));
    // console.log("[DEBUG] Found sheet relation files:", sheetRelFiles);

    for (const relFileName of sheetRelFiles) {
        const sheetMatch = relFileName.match(/sheet(\d+)\.xml\.rels$/i);
        if (!sheetMatch) continue;
        const sheetIndex = parseInt(sheetMatch[1], 10) - 1; // シートインデックス (0始まり)
        // console.log(`[DEBUG] Processing relations for sheet index: ${sheetIndex} (from ${relFileName})`);

        const relFileObj = workbook.files[relFileName];
        if (!relFileObj) continue;

        try {
            const relXmlString = await getStringFromZipFileAsync(relFileObj);
            const parser = new DOMParser();
            const relXmlDoc = parser.parseFromString(relXmlString, 'application/xml');
            const relationships = relXmlDoc.getElementsByTagName('Relationship');

            for (let i = 0; i < relationships.length; i++) {
                const rel = relationships[i];
                const type = rel.getAttribute('Type');
                const target = rel.getAttribute('Target');

                // 描画ファイル (drawing.xml) への参照を探す
                if (type && type.endsWith('/drawing') && target) {
                    // target は "../drawings/drawing1.xml" のような形式
                    const drawingFileName = target.substring(target.lastIndexOf('/') + 1);
                    // console.log(`[DEBUG] Found drawing relation: ${drawingFileName} relates to sheet index ${sheetIndex}`);
                    sheetDrawingRelMap.set(drawingFileName, String(sheetIndex)); // マップに格納
                    break; // 通常、シートごとにdrawingは1つのはず
                }
                // VML (vmlDrawing.vml) への参照も考慮する (あれば)
                // 注: VMLの関連付けはdrawingより不明瞭なことが多い
                if (type && type.endsWith('/vmlDrawing') && target) {
                    const vmlFileName = target.substring(target.lastIndexOf('/') + 1);
                     // VMLの場合はシートインデックスとの関連付けが保証されない場合があるため、
                     // drawing が見つからなければ VML の関連を暫定的に使う、などの考慮が必要かもしれない。
                     // 一旦 drawing を優先し、VML の関連付けはここでは積極的に利用しない。
                    // console.log(`[DEBUG] Found VML drawing relation: ${vmlFileName} relates to sheet index ${sheetIndex} (Might be less reliable)`);
                    // sheetDrawingRelMap.set(vmlFileName, String(sheetIndex)); // 必要なら追加
                }

            }
        } catch (err) {
            console.error(`[DEBUG] Error processing sheet relation file ${relFileName}:`, err);
        }
    }
    // console.log("[DEBUG] sheetDrawingRelMap:", sheetDrawingRelMap);


    // 3. 図形ファイルを処理し、正しいシートに紐付ける
    const fileNames = Object.keys(workbook.files);
    for (const fileName of fileNames) {
        const isDrawingXml = /^xl\/drawings\/drawing\d+\.xml$/i.test(fileName);
        const isVmlDrawing = /^xl\/drawings\/vmlDrawing\d+\.vml$/i.test(fileName);

        if (isDrawingXml || isVmlDrawing) {
            const fileObj = workbook.files[fileName];
            if (!fileObj) continue;

            let sheetName = "不明なシート";
            const baseFileName = fileName.substring(fileName.lastIndexOf('/') + 1); // e.g., "drawing1.xml"

            // drawing.xml の場合は sheetDrawingRelMap を使ってシート名を特定
            if (isDrawingXml && sheetDrawingRelMap.has(baseFileName)) {
                const sheetIndexStr = sheetDrawingRelMap.get(baseFileName);
                if (sheetNameMap.has(sheetIndexStr)) {
                    sheetName = sheetNameMap.get(sheetIndexStr);
                    // console.log(`[DEBUG] Mapped ${baseFileName} to sheet: ${sheetName} (Index: ${sheetIndexStr})`);
                } else {
                     // console.log(`[DEBUG] Warning: sheetIndex ${sheetIndexStr} from rels not found in sheetNameMap for ${baseFileName}`);
                }
            }
            // vmlDrawing.vml の場合、または drawing.xml が rels になかった場合、
            // 従来のファイル名からの推測ロジックをフォールバックとして使用
            else {
                // console.log(`[DEBUG] Using fallback (filename based) sheet association for ${fileName}`);
                const drawingMatch = fileName.match(/(?:drawing|vmlDrawing)(\d+)\.(?:xml|vml)$/i);
                const drawingIndex = drawingMatch ? parseInt(drawingMatch[1], 10) - 1 : -1;
                if (drawingIndex >= 0 && sheetNameMap.has(String(drawingIndex))) {
                    sheetName = sheetNameMap.get(String(drawingIndex));
                    // console.log(`[DEBUG] Fallback association for ${fileName}: Sheet ${sheetName} (Index: ${drawingIndex})`);
                } else {
                     // console.log(`[DEBUG] Fallback association failed for ${fileName}`);
                }
            }


            const xmlString = await getStringFromZipFileAsync(fileObj);
            // parseShapeXml はシート名引数を内部では使っていないので、fileNameだけでOK
            const parsedShapes = parseShapeXml(xmlString, fileName); // { row, col, text } の配列を取得

            if (parsedShapes.length > 0) {
                // console.log(`[DEBUG] Adding ${parsedShapes.length} shapes from ${fileName} to sheet: ${sheetName}`);
                if (!shapesBySheet.has(sheetName)) {
                    shapesBySheet.set(sheetName, []);
                }
                shapesBySheet.get(sheetName).push(...parsedShapes);
            }
        }
    }

    // 4. 全ての drawing ファイルを処理した後、シートごとにグリッド化してテキストを結合
    let finalShapeText = "";
    const sortedSheetIndices = Array.from(sheetNameMap.keys()).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));

    // ソートされたシートインデックス順に処理
    for (const sheetIndexStr of sortedSheetIndices) {
        const sheetName = sheetNameMap.get(sheetIndexStr);
        if (shapesBySheet.has(sheetName)) {
            const shapes = shapesBySheet.get(sheetName);
            // 同じシート内で重複するテキストを持つシェイプを除去する（完全に同じテキストを持つシェイプが複数ある場合）
            // ※位置情報が異なる場合は別物として扱う
            const uniqueShapes = [];
            const seenTexts = new Set();
            shapes.forEach(shape => {
                // テキストと位置情報で一意性を判断（同一セル内の複数シェイプは別物）
                const key = `${shape.row}-${shape.col}-${shape.text}`;
                if (!seenTexts.has(key)) {
                    uniqueShapes.push(shape);
                    seenTexts.add(key);
                } else {
                     // console.log(`[DEBUG] Duplicate shape text removed: "${shape.text}" at row ${shape.row}, col ${col} on sheet "${sheetName}"`);
                }
            });


            // グリッドシェイプとその他のシェイプに分類
            const gridShapes = uniqueShapes.filter(s => s.row !== Infinity && s.col !== Infinity);
            const otherShapes = uniqueShapes.filter(s => s.row === Infinity || s.col === Infinity); // VMLや位置不明なもの

            let sheetGridText = "";
            if (gridShapes.length > 0) {
                // グリッドの最大行・列を計算
                let maxRow = 0;
                let maxCol = 0;
                gridShapes.forEach(s => {
                    maxRow = Math.max(maxRow, s.row);
                    maxCol = Math.max(maxCol, s.col);
                });

                // グリッドを初期化 (二次元配列、中身は文字列)
                const grid = Array(maxRow + 1).fill(null).map(() => Array(maxCol + 1).fill(""));

                // グリッドにテキストを配置 (同じセルは" "で連結)
                gridShapes.forEach(s => {
                    if (grid[s.row] && grid[s.row][s.col] !== undefined) {
                        grid[s.row][s.col] += (grid[s.row][s.col] ? " " : "") + s.text;
                    } else {
                         // 配列範囲外などのエラーを防ぐ（基本的には起こらないはずだが念のため）
                         console.warn(`[DEBUG] Shape at invalid grid position skipped: row=${s.row}, col=${s.col}, sheet=${sheetName}`);
                    }
                });

                sheetGridText += `【Shapes in "${sheetName}"】\n`;
                // console.log("[DEBUG] Original grid before slimming:", grid); // デバッグ用ログは維持

                // 1. 空列判定
                const emptyColumnIndices = new Set();
                for (let c = 0; c <= maxCol; c++) {
                    let isColEmpty = true;
                    for (let r = 0; r <= maxRow; r++) {
                        // grid[r] が存在し、かつ grid[r][c] が空文字でないことを確認
                        if (grid[r]?.[c] && grid[r][c] !== "") {
                            isColEmpty = false;
                            break;
                        }
                    }
                    if (isColEmpty) {
                        emptyColumnIndices.add(c);
                    }
                }
                // console.log("[DEBUG] Empty column indices:", emptyColumnIndices); // デバッグ用ログは維持

                // 2. グリッドをテキスト化 (空行・空列を削除、行末タブ削除)
                let gridContent = "";
                for (let r = 0; r <= maxRow; r++) {
                    let rowText = "";
                    let rowHasContent = false; // 行に内容があるかフラグ
                    let cellsInRow = []; // この行で実際に出力するセルを一時格納

                    for (let c = 0; c <= maxCol; c++) {
                        // 現在の列が空列でない場合のみ処理
                        if (!emptyColumnIndices.has(c)) {
                            const cellContent = grid[r]?.[c] || "";
                            cellsInRow.push(cellContent); // 空でない列のセル内容を配列に追加
                            // console.log("[DEBUG] cellContent", cellContent, "r=", r, "c=", c); // デバッグ用ログは維持
                            if (cellContent) {
                                 rowHasContent = true; // この行に内容があることを記録
                            }
                        }
                    }

                    // 内容のある行のみ、タブ区切りで結合して結果に追加
                    if (rowHasContent) {
                        rowText = cellsInRow.join("\t"); // 空でない列のセルだけをタブで結合
                        rowText = rowText.replace(/\t+$/, "");
                        gridContent += rowText + "\n";
                    }
                }

                // グリッド全体が空でなければ追加
                if (gridContent.trim()) {
                    sheetGridText += gridContent + "\n"; // グリッドの後にも改行
                } else {
                    // グリッドは存在したが中身がなかった場合（例えば空テキストの図形のみ、または空行/空列削除の結果）
                    // sheetGridText += "(No displayable shapes in grid)\n\n"; // ヘッダーのみ表示されるのは冗長なので、ヘッダーごと出力しない
                    sheetGridText = ""; // グリッドヘッダーも出力しない
                }
            }

            // VML/その他図形のテキストをリスト化
            let otherShapesText = "";
            if (otherShapes.length > 0) {
                // VMLなども元の出現順に近いように（不安定だが）ソートを試みる
                otherShapes.sort((a, b) => {
                     // もしVML内に順序を示す情報があればそれを使うべきだが、現状は特にないのでそのまま
                    return 0; // 元の配列順を維持（sortは不安定な場合がある）
                });
                otherShapesText += `【Other Shapes (e.g., VML) in "${sheetName}"】\n`;
                otherShapesText += otherShapes.map(s => s.text).join("\n") + "\n\n";
            }

            // シートごとのテキストを結合
            if (sheetGridText || otherShapesText) {
                 finalShapeText += sheetGridText + otherShapesText;
            }

        }
    }

    // "不明なシート" に分類されたものがあれば、同様に処理して最後に追加
    if (shapesBySheet.has("不明なシート")) {
        const shapes = shapesBySheet.get("不明なシート");
        // 重複除去
        const uniqueShapes = [];
        const seenTexts = new Set();
        shapes.forEach(shape => {
            const key = `${shape.row}-${shape.col}-${shape.text}`;
            if (!seenTexts.has(key)) {
                uniqueShapes.push(shape);
                seenTexts.add(key);
            } else {
                 // console.log(`[DEBUG] Duplicate shape text removed: "${shape.text}" at row ${shape.row}, col ${shape.col} on sheet "不明なシート"`);
            }
        });

        const gridShapes = uniqueShapes.filter(s => s.row !== Infinity && s.col !== Infinity);
        const otherShapes = uniqueShapes.filter(s => s.row === Infinity || s.col === Infinity);
        const sheetName = "不明なシート";

        let sheetGridText = "";
         if (gridShapes.length > 0) {
             // グリッド作成
             let maxRow = 0;
             let maxCol = 0;
             gridShapes.forEach(s => {
                 maxRow = Math.max(maxRow, s.row);
                 maxCol = Math.max(maxCol, s.col);
             });
             const grid = Array(maxRow + 1).fill(null).map(() => Array(maxCol + 1).fill(""));
             gridShapes.forEach(s => {
                  if (grid[s.row] && grid[s.row][s.col] !== undefined) {
                      grid[s.row][s.col] += (grid[s.row][s.col] ? " " : "") + s.text; // 元の改行連結
                  }
             });

             sheetGridText += `【Shapes in "${sheetName}"】\n`;

             // 1. 空列判定
             const emptyColumnIndices = new Set();
             for (let c = 0; c <= maxCol; c++) {
                 let isColEmpty = true;
                 for (let r = 0; r <= maxRow; r++) {
                     if (grid[r]?.[c] && grid[r][c] !== "") {
                         isColEmpty = false;
                         break;
                     }
                 }
                 if (isColEmpty) {
                     emptyColumnIndices.add(c);
                 }
             }

             // 2. グリッドをテキスト化 (空行・空列を削除、行末タブ削除)
             let gridContent = "";
             for (let r = 0; r <= maxRow; r++) {
                 let rowText = "";
                 let rowHasContent = false;
                 let cellsInRow = [];

                 for (let c = 0; c <= maxCol; c++) {
                     if (!emptyColumnIndices.has(c)) {
                         const cellContent = grid[r]?.[c] || "";
                         cellsInRow.push(cellContent);
                         if (cellContent) rowHasContent = true;
                     }
                 }
                 if (rowHasContent) {
                    rowText = cellsInRow.join("\t");
                    rowText = rowText.replace(/\t+$/, "");
                    gridContent += rowText + "\n";
                 }
             }

             if (gridContent.trim()) {
                sheetGridText += gridContent + "\n";
             } else {
                 // sheetGridText += "(No displayable shapes in grid)\n\n";
                 sheetGridText = ""; // ヘッダーごと出力しない
             }
        }

        // その他図形のテキスト化
        let otherShapesText = "";
         if (otherShapes.length > 0) {
             otherShapes.sort((a, b) => 0);
             otherShapesText += `【Other Shapes (e.g., VML) in "${sheetName}"】\n`;
             otherShapesText += otherShapes.map(s => s.text).join("\n") + "\n\n";
        }

         // テキスト結合
         if (sheetGridText || otherShapesText) {
            finalShapeText += sheetGridText + otherShapesText;
         }
    }

    return finalShapeText;
}



// ▼▼▼ 3-3. Excel ファイル読み込みを async 化し、図形内テキストも抽出 ▼▼▼
// ★★★ 改良：空行・空列削除、セル内整形処理を追加 ★★★
async function readExcelFile(file) {
    // console.log("[DEBUG] readExcelFile called (async):", file.name);
    const data = new Uint8Array(await file.arrayBuffer());  // FileReader不要でも読み込めるが、従来通りでもOK
    try {
        // コメント取得のため cellComments: true を指定
        // 図形内文言も取得するため bookFiles: true を指定
        // console.log("[DEBUG] About to XLSX.read (async) ...", file.name);
        const workbook = XLSX.read(data, {
            type: 'array',
            cellComments: true, // メモ（旧コメント）取得に必要
            bookFiles: true
        });
        // console.log("[DEBUG] XLSX.read complete:", file.name);

        let text = "";
        workbook.SheetNames.forEach(sheetName => {
            // console.log("[DEBUG] Processing sheet:", sheetName);
            const sheet = workbook.Sheets[sheetName];

            // 1. シートを2次元配列に変換 (空セルは "" とする)
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

            // 2. セル内容の前処理 & 空行削除
            const processedData = jsonData.map(row => {
                return row.map(cell => {
                    // セルの型チェックと前処理
                    let cellValue = cell;
                    if (cellValue === null || cellValue === undefined) {
                        cellValue = ""; // nullやundefinedは空文字列に
                    } else {
                        cellValue = String(cellValue); // 文字列に変換
                    }

                    // 改行とタブをスペースに置換
                    let cleanedCell = cellValue.replace(/[\n\t]+/g, ' ');
                    // 連続する空白を1つに正規化
                    cleanedCell = cleanedCell.replace(/\s{2,}/g, ' ');
                    // 前後の空白を削除
                    cleanedCell = cleanedCell.trim();
                    // 結果が空白のみの場合は空文字列にする
                    return cleanedCell;
                });
            }).filter(row => row.some(cell => cell !== "")); // 実質的に空でない行のみを残す

            // 3. 空列削除
            let finalData = [];
            if (processedData.length > 0) {
                const numCols = Math.max(...processedData.map(row => row.length)); // 最大列数を取得
                const colsToRemove = new Set();

                for (let j = 0; j < numCols; j++) {
                    let isColEmpty = true;
                    for (let i = 0; i < processedData.length; i++) {
                        // processedData[i][j] が範囲外 or 空文字列でないかチェック
                        if (processedData[i] && processedData[i][j] !== undefined && processedData[i][j] !== "") {
                            isColEmpty = false;
                            break;
                        }
                    }
                    if (isColEmpty) {
                        colsToRemove.add(j);
                    }
                }

                // 空でない列だけを含む新しいデータを作成
                finalData = processedData.map(row => {
                    const newRow = [];
                    const originalLength = row.length;
                    for (let j = 0; j < numCols; j++) { // 最大列数までループ
                        if (!colsToRemove.has(j)) {
                            // 元の行にその列が存在すれば値を追加、なければ空文字列
                            newRow.push(j < originalLength && row[j] !== undefined ? row[j] : "");
                        }
                    }
                    return newRow;
                });

                // 再度、空行が発生していないかチェック (空列削除の結果、行が空になる場合があるため)
                finalData = finalData.filter(row => row.some(cell => cell !== ""));
            }

            // 4. 最終的なデータをTSV化
            if (finalData.length > 0) {
                const newSheet = XLSX.utils.aoa_to_sheet(finalData);
                let tsv = XLSX.utils.sheet_to_csv(newSheet, { FS: '\t' });
                const lines = tsv.split('\n');
                const trimmedLines = lines.map(line => line.replace(/\t+$/, '')); // 正規表現で末尾の連続するタブを削除
                tsv = trimmedLines.join('\n');
                text += `【Sheet: ${sheetName}】\n${tsv}\n\n\n`;
            } else {
                 // シートが完全に空になった場合 (元のデータが空 or 空行/空列削除の結果)
                 text += `【Sheet: ${sheetName}】\n(シートは空、または有効なデータがありません)\n\n\n`;
            }


            // シート内のコメントも出力する
            // ★★★ コメントテキストのサニタイズ処理を追加 ★★★
            // ★★★ デバッグログ追加：コメント情報の確認 ★★★
            console.log(`[DEBUG] Checking comments for sheet: ${sheetName}`);
            console.log("[DEBUG] sheet['!comments'] object:", sheet["!comments"]); // オブジェクト自体をログ出力
            // ★★★ デバッグログ追加ここまで ★★★

            if (sheet["!comments"] && Array.isArray(sheet["!comments"]) && sheet["!comments"].length > 0) {
                 // console.log("[DEBUG] Found comments data in sheet:", sheetName, sheet["!comments"].length); // 元のログも維持
                text += `【Comments in ${sheetName}】\n`;
                sheet["!comments"].forEach(comment => {
                    // ★★★ デバッグログ追加：個々のコメント内容確認 ★★★
                    console.log("[DEBUG] Processing comment object:", comment);
                    // ★★★ デバッグログ追加ここまで ★★★

                    const author = comment.a || "unknown";
                    const originalCommentText = comment.t || ""; // コメントテキストを取得

                    // 1. 改行・タブをスペースに置換
                    let cleanedCommentText = originalCommentText.replace(/[\n\t]+/g, ' ');
                    // 2. 連続する空白を1つに正規化
                    cleanedCommentText = cleanedCommentText.replace(/\s{2,}/g, ' ');
                    // 3. 前後の空白を削除
                    cleanedCommentText = cleanedCommentText.trim();

                    if (cleanedCommentText) { // サニタイズ後のテキストが空でなければ出力
                        const cellRef = comment.ref || "unknown cell"; // セル参照も取得
                        text += `Cell ${cellRef} (by ${author}): ${cleanedCommentText}\n`;
                    } else {
                        console.log(`[DEBUG] Comment text was empty after cleaning for cell ${comment.ref}`); // 空になった場合のログ
                    }
                });
                text += "\n";
            } else {
                 // コメントが見つからなかった場合のログ
                 console.log(`[DEBUG] No comments data found or sheet["!comments"] is not a non-empty array for sheet: ${sheetName}`);
            }
        });

        // 図形(Shapes)の中のテキストを取得（非同期）
        const shapeText = await extractShapeTextFromWorkbookAsync(workbook);
        if (shapeText.trim()) {
            // console.log("[DEBUG] shapeText extracted length:", shapeText.length);
             // 図形テキストは extractShapeTextFromWorkbookAsync -> parseShapeXml でサニタイズ済み
            text += `${shapeText}\n`;
        } else {
            // console.log("[DEBUG] No shapeText extracted.");
        }

        return text;
    } catch (error) {
        console.error("[DEBUG] Error in readExcelFile:", error);
        throw error; // エラーを呼び出し元に伝播させる
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
            window.open("https://XXXXXXXXXX/", "_blank");
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
