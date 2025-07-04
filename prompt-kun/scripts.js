/*  ▼ 既に定義済みであれば再宣言を避けるため typeof チェックを入れる */

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: グループラベル結合用区切り文字 ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.GROUP_LABEL_SEPARATOR === "undefined") {
    window.GROUP_LABEL_SEPARATOR = "\\n"; // グループ化されたラベルの区切り文字
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: ノード位置出力モード ★★
    1: from/to の4点, 2: 中点, 3: from のみ
    ………………………………………………………………………………………………………………………… */
if (typeof window.NODE_POSITION_OUTPUT_MODE === "undefined") {
    window.NODE_POSITION_OUTPUT_MODE = 1;
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: 近傍自動コネクト用しきい値 (ユークリッド) ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.NEAR_SHAPE_THRESHOLD === "undefined") {
    /* 2以内なら自動で接続 – 必要に応じて変更可 */
    window.NEAR_SHAPE_THRESHOLD = 2;
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: 座標表示用係数 ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.ROW_COORD_MULTIPLIER === "undefined") {
    window.ROW_COORD_MULTIPLIER = 1; // 行座標の表示係数 (デフォルト: 1)
}
if (typeof window.COL_COORD_MULTIPLIER === "undefined") {
    window.COL_COORD_MULTIPLIER = 1; // 列座標の表示係数 (デフォルト: 1)
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: EMU → セル座標変換用定数 ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.EMU_PER_POINT === "undefined") {
    window.EMU_PER_POINT = 12700;               /* 1pt = 12 700 EMU */
}
if (typeof window.DEFAULT_ROW_HEIGHT_PT === "undefined") {
    window.DEFAULT_ROW_HEIGHT_PT = 15;          /* 既定行高 (pt)   */
}
if (typeof window.DEFAULT_ROW_HEIGHT_EMU === "undefined") {
    window.DEFAULT_ROW_HEIGHT_EMU = window.EMU_PER_POINT * window.DEFAULT_ROW_HEIGHT_PT; /* 190 500 EMU */
}
if (typeof window.EMU_PER_PIXEL === "undefined") {
    window.EMU_PER_PIXEL = 9525;                /* 1px = 9 525 EMU */
}
if (typeof window.DEFAULT_COL_WIDTH_PX === "undefined") {
    window.DEFAULT_COL_WIDTH_PX = 64;           /* 既定列幅 (px)   */
}
if (typeof window.DEFAULT_COL_WIDTH_EMU === "undefined") {
    window.DEFAULT_COL_WIDTH_EMU = window.EMU_PER_PIXEL * window.DEFAULT_COL_WIDTH_PX;   /* 609 600 EMU */
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: グループID付与ロジック切り替え ★★
    true: 旧ロジック (要素ごとに親のgrpSpを探してgroupIdを付与),
    false: 新ロジック (grpSpによる親集約)
    ………………………………………………………………………………………………………………………… */
if (typeof window.USE_LEGACY_GROUPING_LOGIC === "undefined") {
    window.USE_LEGACY_GROUPING_LOGIC = false;
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: 近傍補完の有効/無効切り替え ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.ENABLE_NEARBY_COMPLETION === "undefined") {
    window.ENABLE_NEARBY_COMPLETION = true;
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: Type 出力制御フラグ ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.OUTPUT_NODE_TYPE === "undefined") {
    window.OUTPUT_NODE_TYPE = false;
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: 非表示シート除外制御フラグ ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.EXCLUDE_HIDDEN_SHEETS === "undefined") {
    window.EXCLUDE_HIDDEN_SHEETS = true; // デフォルトで非表示シートを除外
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加: 空ノードへの接続を覆うノードへ付け替える機能のスイッチ ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.ENABLE_CONNECTOR_RETARGETING === "undefined") {
    window.ENABLE_CONNECTOR_RETARGETING = true;
}

/* ---------- Excel DrawingML 用 名前空間定義 (グローバル) ---------- */
if (typeof window.NS_MAIN === "undefined") {
    window.NS_MAIN               = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
}
if (typeof window.NS_RELATIONSHIPS === "undefined") {
    window.NS_RELATIONSHIPS      = "http://schemas.openxmlformats.org/package/2006/relationships";
}
if (typeof window.NS_DRAWINGML === "undefined") {
    window.NS_DRAWINGML          = "http://schemas.openxmlformats.org/drawingml/2006/main";
}
if (typeof window.NS_SPREADSHEETDRAWING === "undefined") {
    window.NS_SPREADSHEETDRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
}

/* ---------- 匿名コネクタ ID 用カウンタ ---------- */
if (typeof window.__CXN_AUTO_ID === "undefined") {
    window.__CXN_AUTO_ID = 0;
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加箇所 0-B :  抽出結果をシート別の構造化テキストへ整形  ★★
    ………………………………………………………………………………………………………………………… */
if (typeof window.formatToTextBySheet === "undefined") {
    window.formatToTextBySheet = function(resultsBySheet) {

        let output = "";

        const sheetNames = Object.keys(resultsBySheet).sort();  // 名前順で安定化
        for (const sheetName of sheetNames) {
            const data = resultsBySheet[sheetName];
            output += "【" +
                        `Shapes in ${sheetName}` +
                        "】\n";

            if (!data || (data.nodes.length === 0 && data.edges.length === 0)) {
                output += "(No relevant nodes or edges extracted for this sheet)\n\n";
                continue;
            }

            /* ---------- Nodes ---------- */
            output += "--- NODES ---\n";
            if (data.nodes.length > 0) {
                // 座標値をフォーマットするヘルパー関数 (係数適用は呼び出し元で行う)
                const formatCoord = (val) => (val !== null && val !== undefined && Number.isFinite(val)) ? Math.round(val) : "inf";

                data.nodes.forEach(node => {
                    // ▼▼▼ ここから修正 ▼▼▼
                    const typeInfo = window.OUTPUT_NODE_TYPE ? `Type:${node.type}, ` : ""; // Type情報を条件付きで生成
                    if (node.type === 'group') {
                        // Type: "group" の場合は ID, Type(条件付き), GroupID, ParentGroupID のみ出力
                        const groupInfo   = node.GroupID ? `GroupID:${node.GroupID}, `.replace(/drawing/g, '') : "";
                        const parentGroupInfo = node.ParentGroupID ? `ParentGroupID:${node.ParentGroupID}`.replace(/drawing/g, '') : "";
                        // TypeInfo を先頭に追加
                        output += `ID:${node.id.replace(/drawing/g, '')}, ${typeInfo}${groupInfo}${parentGroupInfo}\n`.replace(/,\s*$/,'') + `\n`; // 末尾の不要なカンマを削除
                    } else {
                        // Type: "group" 以外の場合 (既存のロジック)
                        const nodeLabel   = node.label || ""; // label が undefined の場合を考慮
                        const escaped     = nodeLabel.replace(/"/g, '""').replace(/\r?\n/g, "\\n");
                        // GroupID が null でない場合のみ出力
                        const groupInfo   = node.GroupID ? `, GroupID:${node.GroupID}`.replace(/drawing/g, '') : "";
                        // ★★ ParentGroupID が null でない場合のみ出力 ★★
                        const parentGroupInfo = node.ParentGroupID ? `, ParentGroupID:${node.ParentGroupID}`.replace(/drawing/g, '') : "";

                        // 位置情報の出力部分をモードに応じて切り替え
                        let positionInfo = "";
                        // node.type === 'groupAggregation' など他のタイプでは座標を出力する
                        switch (window.NODE_POSITION_OUTPUT_MODE) {
                            case 1: // from/to の4点
                                // 係数を適用してからフォーマット
                                const fromRow = node.fromRow * window.ROW_COORD_MULTIPLIER;
                                const fromCol = node.fromCol * window.COL_COORD_MULTIPLIER;
                                const toRow   = node.toRow   * window.ROW_COORD_MULTIPLIER;
                                const toCol   = node.toCol   * window.COL_COORD_MULTIPLIER;
                                const fromRowInfo = formatCoord(fromRow);
                                const fromColInfo = formatCoord(fromCol);
                                const toRowInfo   = formatCoord(toRow);
                                const toColInfo   = formatCoord(toCol);
                                positionInfo = `RowFrom:${fromRowInfo}, ColFrom:${fromColInfo}, RowTo:${toRowInfo}, ColTo:${toColInfo}`;
                                break;
                            case 3: // from のみ
                                // 係数を適用してからフォーマット
                                const fRow = node.fromRow * window.ROW_COORD_MULTIPLIER;
                                const fCol = node.fromCol * window.COL_COORD_MULTIPLIER;
                                const fRowInfo = formatCoord(fRow);
                                const fColInfo = formatCoord(fCol);
                                positionInfo = `Row:${fRowInfo}, Col:${fColInfo}`;
                                break;
                            case 2: // 中点 (デフォルト)
                            default:
                                let rowMid, colMid;
                                // from/to が両方有効な場合のみ中点を計算
                                if (Number.isFinite(node.fromRow) && Number.isFinite(node.toRow)) {
                                    rowMid = (node.fromRow + node.toRow) / 2;
                                } else {
                                    // 片方でも Infinity なら、有効な方を採用 (両方 Infinity なら Infinity)
                                    rowMid = Number.isFinite(node.fromRow) ? node.fromRow : node.toRow;
                                }
                                if (Number.isFinite(node.fromCol) && Number.isFinite(node.toCol)) {
                                    colMid = (node.fromCol + node.toCol) / 2;
                                } else {
                                    colMid = Number.isFinite(node.fromCol) ? node.fromCol : node.toCol;
                                }
                                // 係数を適用してからフォーマット
                                const midRow = rowMid * window.ROW_COORD_MULTIPLIER;
                                const midCol = colMid * window.COL_COORD_MULTIPLIER;
                                const midRowInfo = formatCoord(midRow);
                                const midColInfo = formatCoord(midCol);
                                positionInfo = `Row:${midRowInfo}, Col:${midColInfo}`;
                                break;
                        }
                        // 出力行の生成 (ParentGroupID を追加, TypeInfoを条件付きで追加)
                        output += `ID:${node.id.replace(/drawing/g, '')}, ${typeInfo}` +
                                    `Label:"${escaped}"${groupInfo}${parentGroupInfo}, ${positionInfo}\n`;
                    }
                    // ▲▲▲ 修正ここまで ▲▲▲
                });
            } else {
                output += "(No nodes extracted for this sheet)\n";
            }

            /* ---------- Edges ---------- */
            output += "\n--- EDGES ---\n";
            if (data.edges.length > 0) {
                data.edges.forEach(edge => {
                    const edgeLabel = edge.label || ""; // label が undefined の場合を考慮
                    const escaped   = edgeLabel.replace(/"/g, '""').replace(/\r?\n/g, "\\n");
                    // GroupID が null でない場合のみ出力
                    const groupInfo = edge.GroupID ? `, GroupID:${edge.GroupID}`.replace(/drawing/g, '') : "";
                    // Type情報を条件付きで生成
                    const typeInfo = window.OUTPUT_NODE_TYPE ? `Type:${edge.type}, ` : "";
                    // エッジには ParentGroupID は通常不要
                    output += `ID:${edge.id.replace(/drawing/g, '')}, ${typeInfo}Source:${edge.source}, Target:${edge.target}${groupInfo}`.replace(/drawing/g, '').replace(/,\s*$/,'') + `\n`; // 末尾の不要なカンマを削除
                });
            } else {
                output += "(No edges extracted for this sheet)\n";
            }
            output += "\n\n";
        }

        /* どのシートにもデータが無い場合のフォールバック */
        if (sheetNames.length === 0) {
            output = "# No sheets with associated drawings were found or processed.";
        }
        return output;
    };
}

/* ---------- Connector 共通 ---------- */
function processConnectorBase(cxnSp, prefix) {
    try {
        const cNvPr  = cxnSp.querySelector(":scope > nvCxnSpPr > cNvPr");
        if (!cNvPr) return null;

        const id      = `${prefix}_${cNvPr.getAttribute("id")}`;
        const cNvCxn  = cxnSp.querySelector(":scope > nvCxnSpPr > cNvCxnSpPr");
        const stId    = cNvCxn?.querySelector("a\\:stCxn, stCxn")?.getAttribute("id") || "0";
        const endId   = cNvCxn?.querySelector("a\\:endCxn, endCxn")?.getAttribute("id") || "0";

        const prstEl  = cxnSp.querySelector(":scope > spPr > a\\:prstGeom, :scope > spPr > prstGeom");
        const type    = prstEl ? prstEl.getAttribute("prst") || "customConnector" : "customConnector";
        const text    = extractTextFromElement(cxnSp) || "";

        return {
            id,
            source : stId === "0" ? null : `${prefix}_${stId}`,
            target : endId === "0" ? null : `${prefix}_${endId}`,
            type,
            label  : text,
            groupId: null
        };
    } catch (e) {
        console.error("[processConnectorBase]", e);
        return null;
    }
}


    
/*  …………………………………………………………………………………………………………………………
    ★★ 追加箇所 0-A :  extractStructuredShapesFromExcel で必要なヘルパー群  ★★
    ………………………………………………………………………………………………………………………… */

/* ---------- 1. workbook.xml を解析してシート情報(r:id ↔ name)を取得 ---------- */
function parseWorkbook(xmlString) {
    const sheets = {};
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "application/xml");
        const parserError = xmlDoc.querySelector("parsererror");
        if (parserError) {
            console.error(`XML 解析エラー (workbook.xml): ${parserError.textContent}`);
            return {};
        }

        const sheetElements = xmlDoc.getElementsByTagName("sheet");
        for (let i = 0; i < sheetElements.length; i++) {
            const sheet = sheetElements[i];
            const name   = sheet.getAttribute("name");
            const sheetId= sheet.getAttribute("sheetId");
            const rId    = sheet.getAttribute("r:id");
            if (name && sheetId && rId) sheets[rId] = { name, sheetId };
        }
    } catch (e) {
        console.error("[DEBUG] parseWorkbook error:", e);
    }
    return sheets;
}

/* ---------- 2. worksheet の rels をたどり sheet↔drawing のマップを生成 ---------- */
async function parseSheetRelationships(zip, sheets) {
    const sheetDrawingMap = {};
    const parser = new DOMParser();

    /* workbook.xml.rels から worksheet ファイルのパスを得る */
    const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    if (!wbRelsXml) return sheetDrawingMap;

    const wbRelsDoc = parser.parseFromString(wbRelsXml, "application/xml");
    const wbRels    = wbRelsDoc.getElementsByTagName("Relationship");
    const worksheetPathMap = {};
    for (const rel of wbRels) {
        const type   = rel.getAttribute("Type");
        const rId    = rel.getAttribute("Id");
        const target = rel.getAttribute("Target");
        if (type && rId && target && type.endsWith("/worksheet")) {
            worksheetPathMap[rId] = `xl/${target.replace(/^\/+/, "")}`;
        }
    }

    /* 各 worksheet の rels から drawing を探す */
    for (const rId in sheets) {
        const sheetInfo  = sheets[rId];
        const wsPath     = worksheetPathMap[rId];
        if (!wsPath) continue;

        const sheetDir        = wsPath.substring(0, wsPath.lastIndexOf("/"));          // 例: xl/worksheets
        const sheetFileName   = wsPath.substring(wsPath.lastIndexOf("/") + 1);         // 例: sheet1.xml
        const relsPath        = `${sheetDir}/_rels/${sheetFileName}.rels`;             // 例: xl/worksheets/_rels/sheet1.xml.rels
        const relsXml         = await zip.file(relsPath)?.async("string");
        if (!relsXml) continue;

        const relsDoc  = parser.parseFromString(relsXml, "application/xml");
        const rels     = relsDoc.getElementsByTagName("Relationship");
        for (const rel of rels) {
            const type   = rel.getAttribute("Type");
            const target = rel.getAttribute("Target");
            if (type && target && type.endsWith("/drawing")) {
                /* 重要: 基準パスを sheetDir に変更して正しい絶対パスを解決 */
                const drawingAbsPath = resolvePath(sheetDir, target);                  // ← ここを修正
                sheetDrawingMap[sheetInfo.name] = drawingAbsPath;
                break;
            }
        }
    }
    return sheetDrawingMap;
}




/* ---------- 4. Shape / Connector 共通の薄いラッパ  ---------- */
function processShapeBase(sp, prefix) {
    try {
        const cNvPr = sp.querySelector(":scope > nvSpPr > cNvPr");
        if (!cNvPr) return null; // 必須要素がなければnullを返す

        const id = `${prefix}_${cNvPr.getAttribute("id")}`;

        // タイプを取得 (prstGeomがあればそのprst属性、なければ"custom")
        const prstEl = sp.querySelector(":scope > spPr > a\\:prstGeom, :scope > spPr > prstGeom");
        const type = prstEl ? prstEl.getAttribute("prst") || "custom" : "custom";

        // テキスト抽出関数を呼び出し、結果をトリム
        const extractedText = extractTextFromElement(sp); // extractTextFromElement内でtrim()されている想定
        const label = extractedText;                      // 抽出したテキストをそのままlabelに設定

        // テキストの有無を判定して hasText フラグを設定
        const hasText = label !== "";                     // 空文字列でなければtrue

        // デバッグログ (必要に応じてコメントアウトまたは削除)
        // console.log("[processShapeBase] ID:", id, "Type:", type, "Label:", `"${label}"`, "HasText:", hasText);

        return {
            id: id,
            type: type,
            label: label,    // 抽出されたテキスト (なければ空文字列)
            hasText: hasText, // テキスト有無フラグ
            groupId: null   // groupId は後続のグループ処理で設定される可能性があるため、ここでは null 初期化
        };
    } catch (e) {
        console.error("[DEBUG] processShapeBase error:", e);
        return null; // エラー発生時もnullを返す
    }
}

/* ---------- 直線 (<xdr:sp> line) を Edge として扱う ---------- */
function processLineShape(sp, prefix) {
    try {
        const cNvPr = sp.querySelector(":scope > nvSpPr > cNvPr");
        if (!cNvPr) return null;

        const id     = `${prefix}_${cNvPr.getAttribute("id")}`;
        const prstEl = sp.querySelector(":scope > spPr > a\\:prstGeom, :scope > spPr > prstGeom");
        const type   = prstEl ? prstEl.getAttribute("prst") || "line" : "line";
        const text   = extractTextFromElement(sp) || "";

        /* 線は端点未接続として初期化 – 後段の近傍探索で source/target を補完 */
        return {
            id,
            source : null,
            target : null,
            type   : type,
            label  : text,
            groupId: null
        };
    } catch (e) {
        console.error("[processLineShape error]", e);
        return null;
    }
}

function processConnectorBase(cxnSp, prefix) {
    try {
        /* ① ID 取得 ― 欠落している場合は自動採番 */
        const cNvPr = cxnSp.querySelector(":scope > nvCxnSpPr > cNvPr");
        let id;
        if (cNvPr) {
            id = `${prefix}_${cNvPr.getAttribute("id")}`;
        } else {
            id = `${prefix}_anonCxn_${window.__CXN_AUTO_ID++}`;
            // console.log(`[UNNAMED] cxnSp without cNvPr captured → ${id}`);
        }

        /* ② 端点取得（未接続は null） */
        const cNvCxn   = cxnSp.querySelector(":scope > nvCxnSpPr > cNvCxnSpPr");
        const stIdAttr = cNvCxn?.querySelector("a\\:stCxn, stCxn")?.getAttribute("id") || "0";
        const endIdAttr= cNvCxn?.querySelector("a\\:endCxn, endCxn")?.getAttribute("id") || "0";

        const source = stIdAttr === "0" ? null : `${prefix}_${stIdAttr}`;
        const target = endIdAttr === "0" ? null : `${prefix}_${endIdAttr}`;

        /* ③ 種類 & ラベル */
        const prstEl = cxnSp.querySelector(":scope > spPr > a\\:prstGeom, :scope > spPr > prstGeom");
        const type   = prstEl ? prstEl.getAttribute("prst") || "customConnector" : "customConnector";
        const label  = extractTextFromElement(cxnSp) || "";

        return { id, source, target, type, label, groupId: null };
    } catch (e) {
        console.error("[processConnectorBase error]", e);
        return null;
    }
}


function getGroupId(grpSp, prefix) {
    try {
        const cNvPr = grpSp.querySelector(":scope > nvGrpSpPr > cNvPr");
        return cNvPr ? `${prefix}_${cNvPr.getAttribute("id")}` : null;
    } catch { return null; }
}

/* ---------- Anchor 端点取得  (EMU → セル座標を含む) ---------- */
function getAnchorEndpoint(elem, endpoint /* 'from' | 'to' */) {
    let anchor = elem;
    while (
        anchor &&
        !(
            anchor.namespaceURI === NS_SPREADSHEETDRAWING &&
            (anchor.localName === "twoCellAnchor" || anchor.localName === "oneCellAnchor")
        )
    ) {
        anchor = anchor.parentElement;
    }
    if (!anchor) return { row: Infinity, col: Infinity };

    const endEl = anchor.getElementsByTagNameNS(NS_SPREADSHEETDRAWING, endpoint)[0];
    if (!endEl) return { row: Infinity, col: Infinity };

    const rowNode    = endEl.getElementsByTagNameNS(NS_SPREADSHEETDRAWING, "row")[0];
    const colNode    = endEl.getElementsByTagNameNS(NS_SPREADSHEETDRAWING, "col")[0];
    const rowOffNode = endEl.getElementsByTagNameNS(NS_SPREADSHEETDRAWING, "rowOff")[0];
    const colOffNode = endEl.getElementsByTagNameNS(NS_SPREADSHEETDRAWING, "colOff")[0];

    const baseRow = parseInt(rowNode?.textContent || "0", 10);
    const baseCol = parseInt(colNode?.textContent || "0", 10);
    const rowOff  = parseInt(rowOffNode?.textContent || "0", 10);
    const colOff  = parseInt(colOffNode?.textContent || "0", 10);

    const row = isNaN(baseRow)
        ? Infinity
        : baseRow + rowOff / window.DEFAULT_ROW_HEIGHT_EMU;
    const col = isNaN(baseCol)
        ? Infinity
        : baseCol + colOff / window.DEFAULT_COL_WIDTH_EMU;

    return { row, col };
}

/* Helper function to find the closest parent grpSp ID */
function findParentGroupId(element, prefix) {
    try {
        let p = element.parentElement;
        while (p) {
            if (p.namespaceURI === window.NS_SPREADSHEETDRAWING && p.localName === "grpSp") {
                // Found the closest parent grpSp, return its ID
                const parentGrpId = getGroupId(p, prefix); // getGroupId should handle cases where ID is missing
                return parentGrpId;
            }
            p = p.parentElement;
        }
    } catch (e) {
        console.error("[findParentGroupId] Error:", e, "Element:", element);
    }
    return null; // No parent grpSp found
}

/* ---------- drawing.xml から nodes / edges を抽出 ---------- */
function extractStructure(xmlDoc, drawingPath) {

    /* ===== 0. 初期セットアップ ========================================== */
    const nodes            = [];
    const edges            = [];
    const elementMap       = {}; // elementMap には sp, cxnSp, grpSp をすべて入れる
    const connectedNodeIds = new Set(); // これは近傍補完で使う
    const prefix           = drawingPath.split("/").pop().replace(".xml", "");

    /* ===== 1. 距離関数 (長方形–点 最短距離) ============================== */
    const dist = (rect, p) => {
        // console.log("rect", rect);
        // console.log("p", p);
        if (p.row === Infinity || p.col === Infinity) return Infinity;
        const dx = p.col < rect.minCol ? rect.minCol - p.col :
                  p.col > rect.maxCol ? p.col - rect.maxCol : 0;
        const dy = p.row < rect.minRow ? rect.minRow - p.row :
                  p.row > rect.maxRow ? p.row - rect.maxRow : 0;
        // console.log("dx", dx);
        // console.log("dy", dy);
        return Math.sqrt(dx * dx + dy * dy);
    };

    /* ===== 2. 1st パス: <sp> / <cxnSp> / <grpSp> 収集 ===================== */
    const allElems = xmlDoc.getElementsByTagName("*");
    const groupElements = []; // grpSp 要素を一時的に保持

    for (const el of allElems) {
        if (el.namespaceURI !== NS_SPREADSHEETDRAWING) continue;

        if (el.localName === "sp") {
            const prstEl   = el.querySelector(":scope > spPr > a\\:prstGeom, :scope > spPr > prstGeom");
            const shapeTyp = prstEl ? prstEl.getAttribute("prst") || "custom" : "custom";
            let item = null;
            if (shapeTyp === "line" || shapeTyp === "lineInv") {
                item = processLineShape(el, prefix);
                if (item) edges.push(item);
            } else {
                item = processShapeBase(el, prefix);
                if (item) nodes.push(item);
            }
            // item が null でないことを確認してから elementMap に追加
            if (item && item.id) elementMap[item.id] = el;
        } else if (el.localName === "cxnSp") {
            const e = processConnectorBase(el, prefix);
            // e が null でないことを確認してから処理
            if (e && e.id) {
                edges.push(e);
                elementMap[e.id] = el;
                if (e.source) connectedNodeIds.add(e.source);
                if (e.target) connectedNodeIds.add(e.target);
                 if (!e.source || !e.target) {
                     // console.log(`[UNATTACHED] cxnSp captured: ${e.id} (src=${e.source}, tgt=${e.target})`);
                 }
            }
        } else if (el.localName === "grpSp") {
            // grpSp 要素は後でまとめて処理するため、ここでは収集のみ
            groupElements.push(el);
            // elementMap にも追加しておく（ID取得のため）
            const grpId = getGroupId(el, prefix);
            if (grpId) {
                elementMap[grpId] = el;
            }
        }
    }

    /* ===== 3. ノード矩形マップ & from/to 座標格納 (グループ化前) ======== */
    const nodeRectMap = {};
    // nodes 配列にはこの時点では sp 由来のノードのみが含まれる
    nodes.forEach(n => {
        const el = elementMap[n.id]; // sp 由来の要素のみのはず
        if (el && el.localName === "sp") { //念のため sp 要素か確認
            const pFrom = getAnchorEndpoint(el, "from");
            const pTo   = getAnchorEndpoint(el, "to");
            n.fromRow = pFrom.row;
            n.fromCol = pFrom.col;
            n.toRow = pTo.row;
            n.toCol = pTo.col;
            nodeRectMap[n.id] = {
                minRow: Math.min(pFrom.row, pTo.row),
                maxRow: Math.max(pFrom.row, pTo.row),
                minCol: Math.min(pFrom.col, pTo.col),
                maxCol: Math.max(pFrom.col, pTo.col)
            };
        } else {
            // 万が一 sp 以外が含まれていた場合のフォールバック
            n.fromRow = Infinity; n.fromCol = Infinity;
            n.toRow = Infinity; n.toCol = Infinity;
            if (!nodeRectMap[n.id]) { // 既存のエントリがなければ初期化
                 nodeRectMap[n.id] = {minRow: Infinity, maxRow: Infinity, minCol: Infinity, maxCol: Infinity};
            }
        }
    });

    /* ===== 4. 近傍補完 =================================== */
    if (window.ENABLE_NEARBY_COMPLETION) {
        // edges 配列内のコネクタ (cxnSp または line sp) に対して実行
        edges.forEach(e => {
            const edgeElement = elementMap[e.id]; // cxnSp または line sp 要素を取得
            if (!edgeElement) return;

            /* --- source --- */
            if (!e.source) {
                const pos = getAnchorEndpoint(edgeElement, "from");
                // console.log("source edgePos", pos);

                let best = Infinity, nearest = null;
                // 接続先候補は sp 由来のノードのみとする (nodeRectMap にはその矩形のみが存在)
                for (const nid in nodeRectMap) {
                     const nodeElement = elementMap[nid];
                     // 接続候補が sp であること、かつ矩形情報が存在することを確認
                     if (nodeElement && nodeElement.localName === "sp" && nodeRectMap[nid]) {
                         const d = dist(nodeRectMap[nid], pos);
                         if (d < best) { best = d; nearest = nid; }
                     }
                }
                // console.log("nearest", nearest);
                // console.log("best", best);
                if (nearest && best <= window.NEAR_SHAPE_THRESHOLD) {
                    e.source = nearest;
                    connectedNodeIds.add(nearest); // 接続された sp ノード ID を記録
                    // console.log(`[FIXED] ${e.id} source -> ${nearest} (d=${best})`);
                }
            }
            /* --- target --- */
            if (!e.target) {
                const pos = getAnchorEndpoint(edgeElement, "to");
                // console.log("target edgePos", pos);
                let best = Infinity, nearest = null;
                for (const nid in nodeRectMap) {
                     const nodeElement = elementMap[nid];
                     if (nodeElement && nodeElement.localName === "sp" && nodeRectMap[nid]) {
                        const d = dist(nodeRectMap[nid], pos);
                        if (d < best) { best = d; nearest = nid; }
                    }
                }
                // console.log("nearest", nearest);
                // console.log("best", best);
                if (nearest && best <= window.NEAR_SHAPE_THRESHOLD) {
                    e.target = nearest;
                    connectedNodeIds.add(nearest); // 接続された sp ノード ID を記録
                    // console.log(`[FIXED] ${e.id} target -> ${nearest} (d=${best})`);
                }
            }
        });
    }

    /* ===== 5. グループ処理 (ロジック切り替え) ============================ */
    if (window.USE_LEGACY_GROUPING_LOGIC) {
        /* ----------【旧ロジック】 Type: "group", GroupID/ParentGroupID 付与 ---------- */
        // console.log("[INFO] Using legacy grouping logic (Type: group, hierarchy).");

        // 1. グループノード (Type: group) を生成して nodes 配列に追加
        const groupNodesToAdd = [];
        groupElements.forEach(grpSp => {
            const grpId = getGroupId(grpSp, prefix);
            if (grpId) {
                const cNvPr = grpSp.querySelector(":scope > nvGrpSpPr > cNvPr");
                const label = cNvPr ? (cNvPr.getAttribute("name") || cNvPr.getAttribute("descr") || "") : "";
                const parentGroupId = findParentGroupId(grpSp, prefix);

                groupNodesToAdd.push({
                    id: grpId,
                    type: "group", // ★ 旧ロジックでは Type: "group"
                    label: label.trim(),
                    hasText: label.trim() !== "",
                    ParentGroupID: parentGroupId,
                    GroupID: null,
                    fromRow: Infinity, fromCol: Infinity, toRow: Infinity, toCol: Infinity
                });
                nodeRectMap[grpId] = {minRow: Infinity, maxRow: Infinity, minCol: Infinity, maxCol: Infinity};
            }
        });
        nodes.push(...groupNodesToAdd);

        // 2. 既存のノード (sp由来) とエッジに GroupID を設定
        const elementsToProcess = [...nodes, ...edges].filter(item => item.type !== "group"); // group ノードは除く
        elementsToProcess.forEach(item => {
            const element = elementMap[item.id];
            if (element) {
                let p = element.parentElement;
                let directParentGrpId = null;
                while (p) {
                    if (p.namespaceURI === NS_SPREADSHEETDRAWING && p.localName === "grpSp") {
                        directParentGrpId = getGroupId(p, prefix);
                        break;
                    }
                    p = p.parentElement;
                }
                item.GroupID = directParentGrpId;
            }
        });

    } else {
        /* ----------【新ロジック】 Type: "groupAggregation", 親集約 ---------- */
        // console.log("[INFO] Using new grouping logic (Type: groupAggregation, aggregation).");
        const nodeToTopGroup  = {};
        const groupAggregates = {};

        // nodes 配列 (sp由来ノードのみ) を対象に実行
        nodes.forEach(n => {
            const el = elementMap[n.id];
            if (!el || el.localName !== "sp") return;

            let p = el.parentElement, topGrp = null;
            while (p) {
                if (p.namespaceURI === NS_SPREADSHEETDRAWING && p.localName === "grpSp")
                    topGrp = p;
                p = p.parentElement;
            }
            if (topGrp) {
                const gid = getGroupId(topGrp, prefix);
                if (gid) {
                    nodeToTopGroup[n.id] = gid;
                    if (!groupAggregates[gid])
                        groupAggregates[gid] = { labelParts: [], childIds: [] };
                    if (n.label && n.hasText)
                        groupAggregates[gid].labelParts.push(n.label);
                    groupAggregates[gid].childIds.push(n.id);
                }
            }
        });

        /* --- 親グループノード (Type: groupAggregation) 生成 --- */
        const groupNodes = [];
        for (const gid in groupAggregates) {
            const label = groupAggregates[gid].labelParts.join(window.GROUP_LABEL_SEPARATOR);
            groupNodes.push({
                id: gid,
                type: "groupAggregation", // ★ 新ロジックでは Type: "groupAggregation"
                label: label,
                hasText: label !== "",
                ParentGroupID: null, // 新ロジックでは階層は追わない
                GroupID: null,
                fromRow: Infinity, fromCol: Infinity, toRow: Infinity, toCol: Infinity
             });
        }

        /* --- 端点を親へ付け替え --- */
        edges.forEach(e => {
            if (e.source && nodeToTopGroup[e.source]) e.source = nodeToTopGroup[e.source];
            if (e.target && nodeToTopGroup[e.target]) e.target = nodeToTopGroup[e.target];
        });

        /* --- 子ノード削除・親ノード追加 --- */
        const childSet = new Set(Object.keys(nodeToTopGroup));
        for (let i = nodes.length - 1; i >= 0; i--) {
            // type が 'group' でない (つまり sp 由来の) ノードで、
            // かつ nodeToTopGroup に ID が存在するものを削除
            if (nodes[i].type !== 'group' && nodes[i].type !== 'groupAggregation' && childSet.has(nodes[i].id)) {
                 nodes.splice(i, 1);
            }
        }
        nodes.push(...groupNodes); // 生成した groupAggregation ノードを追加

        /* --- グループ矩形 = 子のBBox & グループノードに from/to 座標格納 --- */
        for (const gid in groupAggregates) {
            let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
            groupAggregates[gid].childIds.forEach(cid => {
                const r = nodeRectMap[cid]; // グループ化前の矩形情報 (sp由来ノードのみ)
                if (!r) return;
                minR = Math.min(minR, r.minRow); maxR = Math.max(maxR, r.maxRow);
                minC = Math.min(minC, r.minCol); maxC = Math.max(maxC, r.maxCol);
            });
            const groupRect = (minR === Infinity)
                ? { minRow: Infinity, maxRow: Infinity, minCol: Infinity, maxCol: Infinity }
                : { minRow: minR, maxRow: maxR, minCol: minC, maxCol: maxC };
            nodeRectMap[gid] = groupRect;

            // nodes 配列から該当の groupAggregation ノードを探す
            const groupNode = nodes.find(gn => gn.id === gid && gn.type === 'groupAggregation');
            if (groupNode) {
                groupNode.fromRow = groupRect.minRow;
                groupNode.fromCol = groupRect.minCol;
                groupNode.toRow = groupRect.maxRow;
                groupNode.toCol = groupRect.maxCol;
            }
        }
    } // <-- End of if/else (grouping logic)

    /* ===== 5-B. ★★ 追加: 空ノードへの接続を包含ノードへ付け替え ★★ ===== */
    if (window.ENABLE_CONNECTOR_RETARGETING) {
        // console.log("[INFO] Applying connector retargeting logic.");

        // nodes 配列を ID で検索できるように Map を作成
        const nodesById = new Map(nodes.map(n => [n.id, n]));

        edges.forEach(edge => {
            // --- Source 側の付け替えチェック ---
            if (edge.source) {
                const sourceNode = nodesById.get(edge.source);
                // ソースノードが存在し、テキストが空で、座標が有効な場合
                if (sourceNode && !sourceNode.hasText &&
                    Number.isFinite(sourceNode.fromRow) && Number.isFinite(sourceNode.fromCol) &&
                    Number.isFinite(sourceNode.toRow)   && Number.isFinite(sourceNode.toCol))
                {
                    let coveringNode = null;
                    // 他の全てのノードをチェック
                    for (const potentialParentNode of nodes) {
                        // 自分自身、テキストが空、座標が無効なノードはスキップ
                        if (potentialParentNode.id === sourceNode.id || !potentialParentNode.hasText ||
                            !Number.isFinite(potentialParentNode.fromRow) || !Number.isFinite(potentialParentNode.fromCol) ||
                            !Number.isFinite(potentialParentNode.toRow)   || !Number.isFinite(potentialParentNode.toCol))
                        {
                            continue;
                        }

                        // 包含関係をチェック (potentialParentNode が sourceNode を覆うか)
                        if (potentialParentNode.fromRow <= sourceNode.fromRow &&
                            potentialParentNode.fromCol <= sourceNode.fromCol &&
                            potentialParentNode.toRow   >= sourceNode.toRow   &&
                            potentialParentNode.toCol   >= sourceNode.toCol)
                        {
                            // 最初に見つかった包含ノードを採用 (より洗練させるなら面積比較など)
                            coveringNode = potentialParentNode;
                            break;
                        }
                    }

                    // 覆うノードが見つかった場合、接続先を付け替え
                    if (coveringNode) {
                        // console.log(`[RETARGET] Edge ${edge.id} source: ${edge.source} (blank) -> ${coveringNode.id} (covering)`);
                        edge.source = coveringNode.id;
                        connectedNodeIds.delete(sourceNode.id); // 元の空ノードへの接続参照を削除
                        connectedNodeIds.add(coveringNode.id); // 新しい接続先を追加
                    }
                }
            }

            // --- Target 側の付け替えチェック (Source側と同様のロジック) ---
            if (edge.target) {
                const targetNode = nodesById.get(edge.target);
                // ターゲットノードが存在し、テキストが空で、座標が有効な場合
                if (targetNode && !targetNode.hasText &&
                    Number.isFinite(targetNode.fromRow) && Number.isFinite(targetNode.fromCol) &&
                    Number.isFinite(targetNode.toRow)   && Number.isFinite(targetNode.toCol))
                {
                    let coveringNode = null;
                    for (const potentialParentNode of nodes) {
                        if (potentialParentNode.id === targetNode.id || !potentialParentNode.hasText ||
                            !Number.isFinite(potentialParentNode.fromRow) || !Number.isFinite(potentialParentNode.fromCol) ||
                            !Number.isFinite(potentialParentNode.toRow)   || !Number.isFinite(potentialParentNode.toCol))
                        {
                            continue;
                        }
                        if (potentialParentNode.fromRow <= targetNode.fromRow &&
                            potentialParentNode.fromCol <= targetNode.fromCol &&
                            potentialParentNode.toRow   >= targetNode.toRow   &&
                            potentialParentNode.toCol   >= targetNode.toCol)
                        {
                            coveringNode = potentialParentNode;
                            break;
                        }
                    }
                    if (coveringNode) {
                        // console.log(`[RETARGET] Edge ${edge.id} target: ${edge.target} (blank) -> ${coveringNode.id} (covering)`);
                        edge.target = coveringNode.id;
                        connectedNodeIds.delete(targetNode.id);
                        connectedNodeIds.add(coveringNode.id);
                    }
                }
            }
        });
    }

    // console.log("[DEBUG] nodes", nodes);
    // console.log("[DEBUG] edges", edges);
    // console.log("[DEBUG] connectedNodeIds", connectedNodeIds);

    /* ===== 6. 幽霊端点を除去 ================================ */
    for (let i = edges.length - 1; i >= 0; i--) {
        const e = edges[i];
        if (!e.source || !e.target) { edges.splice(i, 1); continue; }
        // nodes 配列に存在するIDかチェック (sp, group, groupAggregation が含まれる)
        const sourceExists = nodes.some(n => n.id === e.source);
        const targetExists = nodes.some(n => n.id === e.target);
        if (!sourceExists || !targetExists) edges.splice(i, 1);
    }

    /* ===== 7. 最終フィルタリング ======================================= */
    const finalNodes = nodes.filter(n =>
        !(!n.hasText && !connectedNodeIds.has(n.id))
    );
    const finalEdges = edges.filter(e =>
        (true)
    );

    // console.log("[DEBUG] finalNodes", finalNodes);
    // console.log("[DEBUG] finalEdges", finalEdges);

    return { nodes: finalNodes, edges: finalEdges };
}    


/* ---------- 6. 図形内テキスト抽出（Shape / Connector で共用） ---------- */
function extractTextFromElement(element) {
    try {
        const txBody = element.getElementsByTagNameNS(NS_SPREADSHEETDRAWING, "txBody")[0];
        if (!txBody) return "";
        const paragraphs = txBody.getElementsByTagNameNS(NS_DRAWINGML, "p");
        const lines = [];
        for (const p of paragraphs) {
            const runs = p.getElementsByTagNameNS(NS_DRAWINGML, "r");
            let line   = "";
            for (const r of runs) {
                const t = r.getElementsByTagNameNS(NS_DRAWINGML, "t")[0];
                if (t) line += t.textContent;
            }
            lines.push(line);
        }
        return lines.join("\n").trim();
    } catch {
        return "";
    }
}


/*  …………………………………………………………………………………………………………………………
    ★★ 追加箇所 1  :  Excel ファイルから図形を構造化テキストで抽出する関数  ★★
    ………………………………………………………………………………………………………………………… */
/**
 *  引数: arrayBuffer   – File から取得した ArrayBuffer
 *  戻値: string        – formatToTextBySheet() で生成した構造化テキスト
 */
async function extractStructuredShapesFromExcel(arrayBuffer) {
    try {
        /* JSZip で XLSX を展開 */
        const zip = await JSZip.loadAsync(arrayBuffer);

        /* workbook.xml 解析 → シート一覧取得 */
        const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
        if (!workbookXml) return "";
        const sheets = parseWorkbook(workbookXml);
        if (Object.keys(sheets).length === 0) return "";

        /* workbook / worksheet 間の rels 解析 → sheet ↔ drawing マップを作成 */
        const sheetDrawingMap = await parseSheetRelationships(zip, sheets);
        const drawingSet      = new Set(Object.values(sheetDrawingMap));

        /* drawing XML 読込 */
        const drawingXmlsMap = {};
        for (const drawingPath of drawingSet) {
            const xml = await zip.file(drawingPath)?.async("string");
            if (xml) drawingXmlsMap[drawingPath] = xml;
        }

        /*  各 Drawing を解析して nodes / edges を蓄積  */
        const resultsBySheet   = {};
        const processedDrawings = new Map();

        for (const sheetName in sheetDrawingMap) {
            const drawingPath = sheetDrawingMap[sheetName];
            if (!drawingXmlsMap[drawingPath]) continue;

            if (!processedDrawings.has(drawingPath)) {
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(drawingXmlsMap[drawingPath], "application/xml");
                const { nodes, edges } = extractStructure(xmlDoc, drawingPath);
                processedDrawings.set(drawingPath, { nodes, edges });
            }

            const { nodes, edges } = processedDrawings.get(drawingPath);
            if (nodes.length === 0 && edges.length === 0) continue;

            if (!resultsBySheet[sheetName]) resultsBySheet[sheetName] = { nodes: [], edges: [] };
            resultsBySheet[sheetName].nodes.push(...nodes);
            resultsBySheet[sheetName].edges.push(...edges);
        }

        if (Object.keys(resultsBySheet).length === 0) return "";
        return formatToTextBySheet(resultsBySheet);
    } catch (err) {
        console.error("[DEBUG] extractStructuredShapesFromExcel error:", err);
        return "";
    }
}

/*  …………………………………………………………………………………………………………………………
    ★★ 追加箇所 1  :  操作ログ送信用ユーティリティ関数  ★★
    ………………………………………………………………………………………………………………………… */
function sendOperationLog(logData) {
    // console.log("[DEBUG] sendOperationLog called. data=", logData);
    fetch("https://XXXXXXXXXX", {
        method: "POST",
        headers: { 
            "Content-Type": "application/json",
            "x-api-key": "XXXXXXXXXX"
        },
        body: JSON.stringify(logData),
        keepalive: true           // ページ遷移直前でも投げ切れるように
    }).catch(err => {
        console.error("[DEBUG] sendOperationLog error:", err);
    });
}

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

// ▼▼▼ 修正版: グループ化対応強化 (Anchor起点に変更) ▼▼▼
function parseShapeXml(xmlString, fileName) {
    const results = []; // { row: number, col: number, text: string } の配列
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "application/xml");

    // ★★★ 名前空間の定義 (Excel DrawingMLで一般的に使われるもの) ★★★
    const xdrNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const aNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const vNamespace = "urn:schemas-microsoft-com:vml"; // VML用

    const parserError = xmlDoc.getElementsByTagName("parsererror")[0];
    if (parserError) {
        // console.warn(`XML parse error in ${fileName}:`, parserError.textContent); // デバッグ用ログは維持
        return results; // 空配列を返す
    }

    // --- DrawingML (.xml) のパース ---
    if (fileName.toLowerCase().endsWith('.xml')) {
        // 1. アンカー要素 (twoCellAnchor, oneCellAnchor) をすべて取得
        const twoCellAnchors = xmlDoc.getElementsByTagNameNS(xdrNamespace, "twoCellAnchor");
        const oneCellAnchors = xmlDoc.getElementsByTagNameNS(xdrNamespace, "oneCellAnchor");
        const allAnchors = [...Array.from(twoCellAnchors), ...Array.from(oneCellAnchors)];
        // console.log(`[DEBUG] Found ${allAnchors.length} anchor elements in ${fileName}`);

        // 2. 各アンカー要素を処理
        for (const anchorElement of allAnchors) {
            const anchorType = anchorElement.tagName.toLowerCase().includes('twocell') ? 'twoCell' : 'oneCell';
            // console.log(`[DEBUG] Processing anchor (${anchorType}):`, anchorElement.outerHTML.substring(0, 100));

            // 3. アンカー要素内の図形 (<xdr:sp>) とグループ (<xdr:grpSp>) を再帰的に処理する関数
            const processNode = (node, currentAnchor) => {
                const shapesFound = []; // このノード以下で見つかった { text, node: spElement }

                // 直接の子要素である <xdr:sp> を探す
                // 名前空間プレフィックスを取得 (なければデフォルトを試す)
                const prefix = node.lookupPrefix(xdrNamespace) || 'xdr'; // xdr が一般的だがファイルによる可能性
                const spSelector = `:scope > ${prefix}\\:sp, :scope > sp`; // プレフィックスありとなしを試す
                const directShapes = node.querySelectorAll(spSelector);
                 // console.log(`[DEBUG] Found ${directShapes.length} direct shapes in node:`, node.tagName, `using selector "${spSelector}"`);
                for (const shape of directShapes) {
                     // ★★★ 名前空間を指定して <a:t> 要素を取得 ★★★
                    const tElements = shape.getElementsByTagNameNS(aNamespace, "t");
                    const originalText = Array.from(tElements).map(el => el.textContent).join("");
                    let cleanedText = originalText.replace(/[\n\t]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
                    if (cleanedText) {
                         // console.log(`[DEBUG] Found direct shape text: "${cleanedText}"`);
                        shapesFound.push({ text: cleanedText, node: shape });
                    }
                }

                // 直接の子要素である <xdr:grpSp> を探し、再帰処理
                const grpSpSelector = `:scope > ${prefix}\\:grpSp, :scope > grpSp`;
                const groupShapes = node.querySelectorAll(grpSpSelector);
                 // console.log(`[DEBUG] Found ${groupShapes.length} group shapes in node:`, node.tagName, `using selector "${grpSpSelector}"`);
                for (const group of groupShapes) {
                    // console.log(`[DEBUG] Recursively processing group:`, group.tagName);
                    shapesFound.push(...processNode(group, currentAnchor)); // 再帰呼び出し
                }
                return shapesFound;
            };

            // 現在のアンカー要素から処理を開始
            const shapesDataInAnchor = processNode(anchorElement, anchorElement);
             // console.log(`[DEBUG] Found ${shapesDataInAnchor.length} shapes total in anchor (${anchorType})`);

            // 4. 見つかった各図形データに対して座標を取得し、結果に追加
            if (shapesDataInAnchor.length > 0) {
                // アンカー要素から座標情報を取得 (一度だけ行う)
                let row = Infinity, col = Infinity;
                const fromEl = anchorElement.getElementsByTagNameNS(xdrNamespace, "from")[0];
                let fromRow = Infinity, fromCol = Infinity;
                if (fromEl) {
                    const rowEl = fromEl.getElementsByTagNameNS(xdrNamespace, "row")[0];
                    const colEl = fromEl.getElementsByTagNameNS(xdrNamespace, "col")[0];
                    fromRow = parseInt(rowEl?.textContent || Infinity, 10);
                    fromCol = parseInt(colEl?.textContent || Infinity, 10);
                    if (isNaN(fromRow)) fromRow = Infinity;
                    if (isNaN(fromCol)) fromCol = Infinity;
                }

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
                    }
                    // 中央値を計算
                    if (fromRow !== Infinity && toRow !== Infinity) row = Math.floor((fromRow + toRow) / 2);
                    else row = fromRow;
                    if (fromCol !== Infinity && toCol !== Infinity) col = Math.floor((fromCol + toCol) / 2);
                    else col = fromCol;
                } else { // oneCell
                    row = fromRow;
                    col = fromCol;
                }
                 // console.log(`[DEBUG] Anchor position: row=${row}, col=${col}`);

                // アンカー内で見つかったすべての図形に同じ座標を適用
                for (const shapeData of shapesDataInAnchor) {
                     // console.log(`[DEBUG] Adding result: row=${row}, col=${col}, text="${shapeData.text}"`);
                    results.push({ row, col, text: shapeData.text });
                }
            }
        }
    }
    // --- VML (.vml) のパース --- (変更なし)
    else if (fileName.toLowerCase().endsWith('.vml')) {
        // VMLの名前空間は通常自動で解決されることが多いが、必要なら追加
        const shapeList = xmlDoc.getElementsByTagName("v:shape"); // VMLは名前空間なしでも取得できることが多い
        if (shapeList.length > 0) {
            for (let i = 0; i < shapeList.length; i++) {
                const shape = shapeList[i];
                // テキストボックス要素を探す
                const textBoxList = shape.getElementsByTagName("v:textbox");
                if (textBoxList.length > 0) {
                    for (let j = 0; j < textBoxList.length; j++) {
                        const tb = textBoxList[j];
                        const originalInnerText = tb.textContent;
                        let cleanedInnerText = originalInnerText.replace(/[\n\t]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
                        if (cleanedInnerText) {
                            results.push({ row: Infinity, col: Infinity, text: cleanedInnerText });
                        }
                    }
                }
                 // <v:textpath string="..."> も考慮 (図形に沿ったテキストなど)
                 const textPathList = shape.getElementsByTagName("v:textpath");
                 for (let k = 0; k < textPathList.length; k++) {
                     const tp = textPathList[k];
                     const text = tp.getAttribute("string");
                     if (text) {
                         let cleanedText = text.replace(/[\n\t]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
                         if (cleanedText) {
                             results.push({ row: Infinity, col: Infinity, text: cleanedText });
                         }
                     }
                 }
            }
        }
    }

    // 重複除去 (最終段階で) - 同じ位置・同じテキストのものを除去
    const uniqueResults = [];
    const seenKeys = new Set();
    for (const item of results) {
        // row, col が Infinity の場合も考慮してキーを作成
        const rowKey = item.row === Infinity ? 'inf' : item.row;
        const colKey = item.col === Infinity ? 'inf' : item.col;
        const key = `${rowKey}-${colKey}-${item.text}`;
        if (!seenKeys.has(key)) {
            uniqueResults.push(item);
            seenKeys.add(key);
        } else {
             // console.log(`[DEBUG] Duplicate shape removed (final stage): ${key}`);
        }
    }

    return uniqueResults; // { row, col, text } の配列を返す
}


// ▼▼▼ 修正版: 図形の中の文言抽出を "async" で行う (Python版の関連付けロジックに寄せた修正) ▼▼▼
async function extractShapeTextFromWorkbookAsync(workbook) {
    const shapesBySheet = new Map(); // キー: シート名, 値: [{ row, col, text }] の配列
    const drawingPathToSheetNameMap = new Map(); // キー: 図形ファイルの絶対パス, 値: シート名
    const sheetNameMapByIndex = new Map(); // キー: シートインデックス(文字列), 値: シート名 (元の createSheetNameMap の結果も保持)
    const sheetRIdToNameMap = new Map(); // キー: シートの rId, 値: シート名
    const sheetRIdToFilePathMap = new Map(); // キー: シートの rId, 値: シートのファイルパス (e.g., xl/worksheets/sheet1.xml)
    const processedDrawingPaths = new Set(); // 処理済みの描画ファイルパスを追跡

    if (!workbook || !workbook.files) return "";

    // ★★★ 名前空間の定義 (XML解析用) ★★★
    const nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const nsRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    const nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const parser = new DOMParser();

    // --- 1. workbook.xml と workbook.xml.rels を解析 ---
    try {
        // 1a. workbook.xml からシート名とrIdを取得
        const workbookFile = workbook.files['xl/workbook.xml'];
        if (workbookFile) {
            const workbookXmlString = await getStringFromZipFileAsync(workbookFile);
            const workbookXmlDoc = parser.parseFromString(workbookXmlString, 'application/xml');
            // ★★★ 名前空間を指定して <sheet> 要素を取得 ★★★
            const sheetNodes = workbookXmlDoc.getElementsByTagNameNS(nsMain, 'sheet');
            for (let i = 0; i < sheetNodes.length; i++) {
                const sheetNode = sheetNodes[i];
                const name = sheetNode.getAttribute('name');
                // ★★★ 名前空間を指定して r:id 属性を取得 ★★★
                const rId = sheetNode.getAttributeNS(nsR, 'id');
                const sheetIndexStr = String(i); // 0始まりのインデックス
                if (name) {
                    sheetNameMapByIndex.set(sheetIndexStr, name); // 元のMapも作成
                    if (rId) {
                        sheetRIdToNameMap.set(rId, name);
                    }
                }
            }
             // console.log("[DEBUG] sheetRIdToNameMap:", sheetRIdToNameMap);
             // console.log("[DEBUG] sheetNameMapByIndex:", sheetNameMapByIndex); // 元のMapもログ出力
        } else {
            console.warn("[DEBUG] xl/workbook.xml not found.");
            // workbook.xml がないとシート名が分からないので、処理継続は難しい場合がある
        }

        // 1b. workbook.xml.rels からシートの rId とファイルパスを取得
        const workbookRelsFile = workbook.files['xl/_rels/workbook.xml.rels'];
        if (workbookRelsFile) {
            const workbookRelsXmlString = await getStringFromZipFileAsync(workbookRelsFile);
            const workbookRelsXmlDoc = parser.parseFromString(workbookRelsXmlString, 'application/xml');
             // ★★★ 名前空間を指定して <Relationship> 要素を取得 ★★★
            const relNodes = workbookRelsXmlDoc.getElementsByTagNameNS(nsRel, 'Relationship');
            for (let i = 0; i < relNodes.length; i++) {
                const relNode = relNodes[i];
                const rId = relNode.getAttribute('Id');
                const target = relNode.getAttribute('Target');
                const type = relNode.getAttribute('Type');
                // シートへのリレーションシップの場合
                if (rId && target && type && type.endsWith('/worksheet')) {
                    // target は "worksheets/sheet1.xml" のような形式
                    // ZIP内の絶対パスに正規化 (先頭に "xl/" をつける)
                    const filePath = resolvePath('xl', target); // resolvePath ヘルパー関数を使用
                    sheetRIdToFilePathMap.set(rId, filePath);
                }
            }
             // console.log("[DEBUG] sheetRIdToFilePathMap:", sheetRIdToFilePathMap);
        } else {
            console.warn("[DEBUG] xl/_rels/workbook.xml.rels not found.");
            // workbook.xml.rels がないとシートファイルパスが分からない
        }

    } catch (err) {
        console.error("[DEBUG] Error parsing workbook or its rels:", err);
        // エラーが発生しても、可能な限り処理を継続することを試みる
    }

    // --- 2. 各シートの rels を解析し、描画ファイルとの関連付けを行う ---
    for (const [rId, sheetFilePath] of sheetRIdToFilePathMap.entries()) {
        const sheetName = sheetRIdToNameMap.get(rId);
        if (!sheetName) {
             // console.log(`[DEBUG] Sheet name not found for rId ${rId}, skipping rels.`);
            continue;
        }

        const sheetDir = sheetFilePath.substring(0, sheetFilePath.lastIndexOf('/')); // 例: xl/worksheets
        const sheetFileName = sheetFilePath.substring(sheetFilePath.lastIndexOf('/') + 1); // 例: sheet1.xml
        const sheetRelsFilePath = `${sheetDir}/_rels/${sheetFileName}.rels`; // 例: xl/worksheets/_rels/sheet1.xml.rels

        const sheetRelsFile = workbook.files[sheetRelsFilePath];
        if (sheetRelsFile) {
            try {
                const sheetRelsXmlString = await getStringFromZipFileAsync(sheetRelsFile);
                const sheetRelsXmlDoc = parser.parseFromString(sheetRelsXmlString, 'application/xml');
                 // ★★★ 名前空間を指定して <Relationship> 要素を取得 ★★★
                const relNodes = sheetRelsXmlDoc.getElementsByTagNameNS(nsRel, 'Relationship');

                for (let i = 0; i < relNodes.length; i++) {
                    const relNode = relNodes[i];
                    const target = relNode.getAttribute('Target');
                    const type = relNode.getAttribute('Type');

                    // 図形ファイル (drawing.xml or vmlDrawing.vml) への参照を探す
                    if (target && type && (type.endsWith('/drawing') || type.endsWith('/vmlDrawing'))) {
                        // target は "../drawings/drawing1.xml" のような形式
                        // 現在のrelsファイルのディレクトリを基準に絶対パスを解決
                        const drawingAbsPath = resolvePath(sheetDir, target); // resolvePath ヘルパー関数を使用
                        const isVml = type.endsWith('/vmlDrawing');

                        // drawing -> sheet_name マップ (drawing優先、VMLは上書きしない)
                        // drawingAbsPath が空文字列や null でないことを確認
                        if (drawingAbsPath && (!drawingPathToSheetNameMap.has(drawingAbsPath) || !isVml)) {
                             drawingPathToSheetNameMap.set(drawingAbsPath, sheetName);
                             // console.log(`[DEBUG] Mapped drawing ${drawingAbsPath} to sheet: ${sheetName}`);
                        } else if (drawingAbsPath && isVml) {
                            // console.log(`[DEBUG] VML ${drawingAbsPath} association skipped (already mapped or drawing exists). Sheet: ${sheetName}`);
                        } else if (!drawingAbsPath) {
                            console.warn(`[DEBUG] Resolved drawing path is invalid for target "${target}" in ${sheetRelsFilePath}`);
                        }
                    }
                }
            } catch (err) {
                console.error(`[DEBUG] Error processing sheet relation file ${sheetRelsFilePath}:`, err);
            }
        } else {
             // console.log(`[DEBUG] Sheet rels file not found: ${sheetRelsFilePath}`);
        }
    }
     // console.log("[DEBUG] drawingPathToSheetNameMap:", drawingPathToSheetNameMap);

    // --- 3. 関連付けられた描画ファイルを処理 ---
    for (const [drawingPath, sheetName] of drawingPathToSheetNameMap.entries()) {
        // drawingPath が有効か再確認
        if (!drawingPath || typeof drawingPath !== 'string') {
             console.warn(`[DEBUG] Invalid drawing path found in map key: ${drawingPath}, skipping.`);
            continue;
        }
        const fileObj = workbook.files[drawingPath];
        if (!fileObj) {
             console.warn(`[DEBUG] Drawing file referenced but not found in zip: ${drawingPath}`);
            continue;
        }
        processedDrawingPaths.add(drawingPath); // 処理済みとしてマーク

        try {
            const xmlString = await getStringFromZipFileAsync(fileObj);
            // parseShapeXml はシート名引数を内部では使っていないので、drawingPathだけでOK
            const parsedShapes = parseShapeXml(xmlString, drawingPath); // { row, col, text } の配列を取得

            if (parsedShapes.length > 0) {
                // console.log(`[DEBUG] Adding ${parsedShapes.length} shapes from ${drawingPath} to sheet: ${sheetName}`);
                if (!shapesBySheet.has(sheetName)) {
                    shapesBySheet.set(sheetName, []);
                }
                shapesBySheet.get(sheetName).push(...parsedShapes);
            }
        } catch (err) {
            console.error(`[DEBUG] Error parsing drawing file ${drawingPath} for sheet ${sheetName}:`, err);
        }
    }

    // --- 4. ZIP内の全描画ファイルをチェックし、未処理なら "不明なシート" へ ---
     // console.log("[DEBUG] Checking for any remaining unassociated drawing files...");
    const allFileNames = Object.keys(workbook.files);
    for (const fileName of allFileNames) {
        // 正規表現で描画ファイルのパスパターンにマッチするかチェック
        // ★★★ パスの区切り文字を / に統一してチェック ★★★
        const normalizedFileName = fileName.replace(/\\/g, '/');
        if (/^xl\/drawings\/(drawing\d+\.xml|vmlDrawing\d+\.vml)$/i.test(normalizedFileName)) {
            if (!processedDrawingPaths.has(normalizedFileName)) { // ★★★ 正規化後のパスでチェック ★★★
                 // console.log(`[DEBUG] Found unprocessed drawing file: ${normalizedFileName}. Assigning to "不明なシート".`);
                const fileObj = workbook.files[fileName]; // 元のファイル名でファイルオブジェクトを取得
                if (!fileObj) continue;

                try {
                    const xmlString = await getStringFromZipFileAsync(fileObj);
                    // parseShapeXml はシート名引数を内部では使っていないので、fileNameだけでOK
                    const parsedShapes = parseShapeXml(xmlString, fileName);
                    if (parsedShapes.length > 0) {
                        const unknownSheetName = "不明なシート";
                        if (!shapesBySheet.has(unknownSheetName)) {
                            shapesBySheet.set(unknownSheetName, []);
                        }
                        shapesBySheet.get(unknownSheetName).push(...parsedShapes);
                        // processedDrawingPaths.add(normalizedFileName); // ここで追加する必要はない
                    }
                } catch (err) {
                     console.error(`[DEBUG] Error processing unassociated drawing file ${fileName}:`, err);
                }
            }
        }
    }


    // --- 5. 全ての drawing ファイルを処理した後、シートごとにグリッド化してテキストを結合 ---
    // (この部分は既存のロジックをほぼ流用し、シート名の順序付けを改善)
    let finalShapeText = "";
    // 取得したシート名の順序を元にソート (workbook.xml の順序)
    const sortedSheetNames = Array.from(sheetNameMapByIndex.values()); // 元のインデックス順マップを使用

    // shapesBySheet に存在するが sortedSheetNames にないシート名（"不明なシート"など）を追加
    const existingSheetNames = new Set(sortedSheetNames);
    for(const sheetName of shapesBySheet.keys()){
        if(!existingSheetNames.has(sheetName)){
            sortedSheetNames.push(sheetName); // 末尾に追加
        }
    }
    // "不明なシート" があれば最後に移動する (もし中間に入っていた場合)
    if (sortedSheetNames.includes("不明なシート")) {
        const index = sortedSheetNames.indexOf("不明なシート");
        if (index !== -1 && index !== sortedSheetNames.length - 1) {
             sortedSheetNames.splice(index, 1); // 一旦削除
             sortedSheetNames.push("不明なシート"); // 末尾に追加
        }
    }


    // ソートされたシート名の順に処理
    for (const sheetName of sortedSheetNames) {
         if (shapesBySheet.has(sheetName)) {
             const shapes = shapesBySheet.get(sheetName);
             // 重複除去ロジック (既存のまま)
             const uniqueShapes = [];
             const seenTexts = new Set();
             shapes.forEach(shape => {
                 // テキストと位置情報で一意性を判断（同一セル内の複数シェイプは別物）
                 const key = `${shape.row}-${shape.col}-${shape.text}`;
                 if (!seenTexts.has(key)) {
                     uniqueShapes.push(shape);
                     seenTexts.add(key);
                 } else {
                      // console.log(`[DEBUG] Duplicate shape text removed: "${shape.text}" at row ${shape.row}, col ${shape.col} on sheet "${sheetName}"`);
                 }
             });

             // グリッド/その他分類 (既存のまま)
             const gridShapes = uniqueShapes.filter(s => s.row !== Infinity && s.col !== Infinity);
             const otherShapes = uniqueShapes.filter(s => s.row === Infinity || s.col === Infinity); // VMLや位置不明なもの

             let sheetGridText = "";
             let otherShapesText = "";

             // グリッド処理 (既存のまま)
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

                 // console.log("[DEBUG] Original grid before slimming:", grid); // デバッグ用ログは維持

                 // 1. 空列判定 (既存のまま)
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

                 // 2. グリッドをテキスト化 (空行・空列を削除、行末タブ削除) (既存のまま)
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

                 // グリッド全体が空でなければ追加 (既存のまま)
                 if (gridContent.trim()) {
                     sheetGridText += `【Shapes in "${sheetName}"】\n${gridContent}\n`; // グリッドの後には改行を1つだけにする
                 } else {
                     // グリッドは存在したが中身がなかった場合
                     sheetGridText = ""; // ヘッダーごと出力しない
                 }
             }

             // VML/その他図形のテキストをリスト化 (既存のまま)
             if (otherShapes.length > 0) {
                 // VMLなども元の出現順に近いように（不安定だが）ソートを試みる
                 otherShapes.sort((a, b) => 0); // 元の配列順を維持
                 // 一意なテキストのみをリスト化
                 const uniqueOtherTexts = [...new Set(otherShapes.map(s => s.text))];
                 if (uniqueOtherTexts.length > 0) {
                     otherShapesText += `【Other Shapes (e.g., VML) in "${sheetName}"】\n`;
                     otherShapesText += uniqueOtherTexts.join("\n") + "\n\n";
                 }
             }

             // シートごとのテキストを結合 (既存のまま)
             if (sheetGridText || otherShapesText) {
                  finalShapeText += sheetGridText + otherShapesText;
             }
         }
    }

    return finalShapeText.trim(); // 末尾の余分な改行を除去
}


// ▼▼▼ ヘルパー関数: パス解決 (簡易版) - 変更なし ▼▼▼
// basePath: 例 "xl/worksheets"
// relativePath: 例 "../drawings/drawing1.xml"
function resolvePath(basePath, relativePath) {
    // 先頭や末尾の / を除去して一貫性を保つ
    basePath = basePath.replace(/^\/+|\/+$/g, '');
    relativePath = relativePath.replace(/^\/+|\/+$/g, '');

    // Windowsパス区切り文字 \ を / に置換
    basePath = basePath.replace(/\\/g, '/');
    relativePath = relativePath.replace(/\\/g, '/');

    const baseParts = basePath.split('/').filter(part => part !== ''); // 空の要素を除去
    const relativeParts = relativePath.split('/');

    // ベースパスがファイルの場合、最後の要素（ファイル名）を削除してディレクトリパスにする
    // 例: xl/worksheets/sheet1.xml -> xl/worksheets
    // ただし、relsファイルからの相対パス解決なので、basePathは常にディレクトリのはず
    // if (basePath.includes('.') && !basePath.endsWith('/')) {
    //     baseParts.pop();
    // }

    let newParts = [...baseParts]; // コピーを作成

    for (const part of relativeParts) {
        if (part === '..') {
            if (newParts.length > 0) {
                newParts.pop(); // 親ディレクトリへ移動 (ルートより上には行かない)
            }
        } else if (part !== '.' && part !== '') {
            newParts.push(part); // サブディレクトリまたはファイル名を追加
        }
    }

    let resolved = newParts.join('/');

    // 絶対パスでない場合 (例: ../../file.xml のような解決結果)、元のrelativePathを返す方が安全かもしれない
    // ここでは単純な結合のみを行う
    // ZIP内のパスは通常 / で始まらないので、先頭の / は不要

    return resolved;
}


/*  …………………………………………………………………………………………………………………………
    ★★ 改めて修正: readExcelFile ★★
      1. セル抽出ロジックは "改修前" のまま完全再現
      2. structuredShapeText をシートごとに分割し、
         セル出力の直後に同シート分だけ連結
      3. シート名を判定できなかった図形 ("不明なシート" 等) は
         ループ後にまとめて出力
    ………………………………………………………………………………………………………………………… */
async function readExcelFile(file) {
    /* ① ArrayBuffer を保持（図形抽出ロジックにも渡すため） */
    const arrayBuffer = await file.arrayBuffer();
    const data        = new Uint8Array(arrayBuffer);

    try {
        /* ② SheetJS で Workbook 読み込み（元コード通り） */
        const workbook = XLSX.read(data, {
            type:        'array',
            cellComments:true,
            bookFiles:   true,
            cellNF:      true,
            cellDates:   true
        });

        let text = "";

        /* ③ workbook.xml から直接非表示シート情報を取得 */
        const hiddenSheetsSet = new Set();
        if (window.EXCLUDE_HIDDEN_SHEETS) {
            try {
                const zip = await JSZip.loadAsync(arrayBuffer);
                const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
                if (workbookXml) {
                    const parser = new DOMParser();
                    const xmlDoc = parser.parseFromString(workbookXml, "application/xml");
                    
                    // 名前空間を考慮して sheet 要素を取得
                    const sheetElements = xmlDoc.querySelectorAll("sheet, workbook > sheets > sheet");
                    console.log(`[DEBUG] Found ${sheetElements.length} sheet elements in workbook.xml`);
                    
                    for (let i = 0; i < sheetElements.length; i++) {
                        const sheet = sheetElements[i];
                        const name = sheet.getAttribute("name");
                        const state = sheet.getAttribute("state");
                        console.log(`[DEBUG] Sheet in XML: name="${name}", state="${state}"`);
                        
                        // state属性があり、かつhiddenまたはveryHiddenの場合のみ非表示シートとして扱う
                        if (name && state && (state === "hidden" || state === "veryHidden")) {
                            hiddenSheetsSet.add(name);
                            console.log(`[DEBUG] Found hidden sheet via XML: ${name} (state: ${state})`);
                        }
                    }
                }
            } catch (e) {
                console.warn("[DEBUG] Could not parse workbook.xml for hidden sheets:", e);
            }
        }

        /* ④ 先に図形を一括抽出し "シート名 → 図形テキスト" へ整形 */
        const structuredShapeText = await extractStructuredShapesFromExcel(arrayBuffer);
        const shapeTextMap        = {};   // { sheetName : string }

        if (structuredShapeText.trim()) {
            const re = /【Shapes in ([^】]+)】([\s\S]*?)(?=【Shapes in [^】]+】|\s*$)/g;
            let m;
            while ((m = re.exec(structuredShapeText)) !== null) {
                const sheetName = m[1];
                const body      = m[2];
                shapeTextMap[sheetName] = `【Shapes in ${sheetName}】${body}\n`;
            }
        }

        /* ⑤ シート順にセル情報 → コメント → 図形を出力 */
        workbook.SheetNames.forEach((sheetName, sheetIndex) => {
            // デバッグ用：ワークブック情報とシート情報をログ出力
            if (window.EXCLUDE_HIDDEN_SHEETS) {
                console.log(`[DEBUG] Processing sheet: ${sheetName}, index: ${sheetIndex}`);
                console.log(`[DEBUG] Available sheet names:`, workbook.SheetNames);
                console.log(`[DEBUG] Workbook.Sheets available:`, !!workbook.Sheets);
                console.log(`[DEBUG] Workbook.Workbook available:`, !!workbook.Workbook);
                if (workbook.Workbook && workbook.Workbook.Sheets) {
                    console.log(`[DEBUG] Sheet info for ${sheetName}:`, workbook.Workbook.Sheets[sheetIndex]);
                }
            }

            // 非表示シートをスキップ（設定で有効な場合のみ）
            if (window.EXCLUDE_HIDDEN_SHEETS) {
                // まずXMLから直接取得した非表示シート情報でチェック
                if (hiddenSheetsSet.has(sheetName)) {
                    console.log(`[DEBUG] Skipping hidden sheet (via XML): ${sheetName}`);
                    return; // 非表示シートをスキップ
                }
                
                // 次にSheetJSの情報でチェック
                if (workbook.Workbook && workbook.Workbook.Sheets) {
                    const sheetInfo = workbook.Workbook.Sheets[sheetIndex];
                    if (sheetInfo && (sheetInfo.Hidden === 1 || sheetInfo.Hidden === 2 || sheetInfo.state === 'hidden' || sheetInfo.state === 'veryHidden')) {
                        console.log(`[DEBUG] Skipping hidden sheet (via SheetJS): ${sheetName}`);
                        return; // 非表示シートをスキップ
                    }
                }
            }
            const sheet    = workbook.Sheets ? workbook.Sheets[sheetName] : null;
            if (!sheet) {
                console.warn(`[DEBUG] Sheet not found: ${sheetName}`);
                return;
            }
            const sheetRef = sheet['!ref'];
            let   jsonData = [];

            /* ---------- 4-1. セルの TSV 抽出（非表示セル除外対応） ---------- */
            if (sheetRef) {
                const range = XLSX.utils.decode_range(sheetRef);
                jsonData    = Array(range.e.r + 1).fill(null)
                                .map(() => Array(range.e.c + 1).fill(""));

                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {

                        const cellAddr = { r: R, c: C };
                        const cellRef  = XLSX.utils.encode_cell(cellAddr);
                        const cell     = sheet[cellRef];

                        let cellValue  = "";
                        if (cell && cell.v !== undefined) {
                            cellValue = cell.w || XLSX.utils.format_cell(cell);
                        }

                        if (cellValue === null || cellValue === undefined) {
                            cellValue = "";
                        } else {
                            cellValue = String(cellValue);
                        }
                        let cleanedCell = cellValue
                                            .replace(/[\n\t]+/g, " ")
                                            .replace(/\s{2,}/g, " ")
                                            .trim();
                        jsonData[R][C] = cleanedCell;
                    }
                }
            }

            /* 空行削除 */
            const processedData = jsonData.filter(
                row => row && row.some(cell => cell !== "")
            );

            /* 空列削除 */
            let finalData = [];
            if (processedData.length > 0) {
                const numCols = Math.max(
                    0,
                    ...processedData.map(row => (row ? row.length : 0))
                );
                const colsToRemove = new Set();

                for (let j = 0; j < numCols; j++) {
                    let isEmpty = true;
                    for (let i = 0; i < processedData.length; i++) {
                        if (processedData[i] && processedData[i][j] !== "") {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty) colsToRemove.add(j);
                }

                finalData = processedData.map(row => {
                    const newRow = [];
                    if (!row) return newRow;
                    const origLen = row.length;
                    for (let j = 0; j < numCols; j++) {
                        if (!colsToRemove.has(j)) {
                            newRow.push(j < origLen && row[j] !== undefined ? row[j] : "");
                        }
                    }
                    return newRow;
                }).filter(row => row && row.some(cell => cell !== ""));
            }

            if (finalData.length > 0) {
                let tsv = finalData.map(row => row.join('\t')).join('\n');
                /* 行末タブを念のため削除（元ロジック） */
                const lines         = tsv.split('\n');
                const trimmedLines  = lines.map(l => l.replace(/\t+$/, ""));
                tsv = trimmedLines.join('\n');

                text += `【Sheet: ${sheetName}】\n${tsv}\n\n\n`;
            } else {
                text += `【Sheet: ${sheetName}】\n(シートは空、または有効なデータがありません)\n\n\n`;
            }

            /* ---------- 4-2. コメント出力（元ロジックそのまま） ---------- */
            if (sheet['!comments'] && Array.isArray(sheet['!comments']) && sheet['!comments'].length > 0) {
                text += `【Comments in ${sheetName}】\n`;
                sheet['!comments'].forEach(comment => {
                    const author = comment.a || "unknown";
                    const body   = String(comment.t || "")
                                        .replace(/[\n\t]+/g, " ")
                                        .replace(/\s{2,}/g, " ")
                                        .trim();
                    if (body) {
                        text += `Cell ${comment.ref || "unknown cell"} (by ${author}): ${body}\n`;
                    }
                });
                text += "\n";
            }

            /* ---------- 4-3. 同シートの図形情報をすぐ後ろに ---------- */
            if (shapeTextMap[sheetName]) {
                text += `${shapeTextMap[sheetName]}\n`;
                delete shapeTextMap[sheetName];  // 消し込む
            }
        });

        /* ⑤ シート名が取れなかった図形を最後にまとめる */
        Object.keys(shapeTextMap).forEach(restName => {
            text += `${shapeTextMap[restName]}\n`;
        });

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

// ▼▼▼ copyText: テキストをクリップボードにコピーし、Chat AI を開く ▼▼▼ */
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

    /* ------------------------------------------------------------------
       ★★ 追加箇所 2 :  操作ログ収集 & 送信
       ------------------------------------------------------------------ */
    try {
        const logData = {
            sending_side: "prompt-kun",
            prompt_file: document.getElementById("selected-file").textContent || "",
            attached_files: Array.from(
                document.querySelectorAll("#dropped-files .file-item")
            ).map(el => (el.file ? el.file.name : el.textContent)),
            datetime: new Date().toISOString(),
            machine_name: navigator.userAgent,
            // 送信したい情報を任意に記載
            memo: "XXXXXXXXXX"
        };
        // sendOperationLog(logData);
    } catch (err) {
        console.error("[DEBUG] failed to prepare/send operation log:", err);
    }
    /* ------------------------------------------------------------------ */

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
