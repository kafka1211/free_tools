## Excel DrawingML (`drawing*.xml`) から Nodes/Edges を抽出する仕様解説

Excelファイル (.xlsx) 内の図形描画情報（DrawingML）を解析し、シート上の図形（Node）とそれらを繋ぐ線（Edge）を構造化されたテキスト形式で抽出することを目的としています。ここでは、その抽出ロジック、特に `drawing*.xml` の仕様とコードの関連について、ナレッジとして活用できるように詳しく解説します。

### 1. Excel DrawingML (`drawing*.xml`) の概要

**役割とシートとの関連:**

*   Excelの各シートに含まれる図形（オートシェイプ、テキストボックス、コネクタ、画像、グラフなど）の情報は、`xl/drawings/` ディレクトリ内の `drawing<N>.xml` ファイルにXML形式で記述されます。（古い形式の図形は `vmlDrawing<N>.vml` に記述されることもあります。）
*   どのシートがどの `drawing*.xml` を参照するかは、各シートファイル (`xl/worksheets/sheet<N>.xml`) のリレーションシップファイル (`xl/worksheets/_rels/sheet<N>.xml.rels`) によって定義されています。
    *   コード内の `parseSheetRelationships` 関数がこの関連付けを行っています。

**基本的な構造と主要要素:**

*   `drawing*.xml` のルート要素は通常 `<xdr:wsDr>` (Worksheet Drawing) です。
*   図形やグループは、その位置を定義する **アンカー (Anchor)** 要素の中に配置されます。
    *   `<xdr:twoCellAnchor>`: 図形が2つのセルにまたがって配置される場合。開始セル (`<xdr:from>`) と終了セル (`<xdr:to>`) を指定します。サイズ変更や移動の挙動も属性で定義されます。
    *   `<xdr:oneCellAnchor>`: 図形が1つのセル内に配置される場合。開始セル (`<xdr:from>`) のみを指定します。
    *   `<xdr:absoluteAnchor>`: セルに関係なく、ページ上の絶対位置（EMU単位）で配置される場合。（コードでは主に `twoCellAnchor` と `oneCellAnchor` を扱います）
*   アンカー要素内には、個々の図形要素が配置されます。
    *   `<xdr:sp>` (Shape): オートシェイプ、テキストボックス、画像などの基本的な図形。
    *   `<xdr:cxnSp>` (Connection Shape): 図形同士を接続するコネクタ線。
    *   `<xdr:grpSp>` (Group Shape): 複数の図形をまとめたグループ。再帰的に `<xdr:sp>`, `<xdr:cxnSp>`, `<xdr:grpSp>` を含むことができます。
    *   `<xdr:graphicFrame>`: グラフやSmartArtなど、より複雑な描画オブジェクト。 (このコードでは主に `sp`, `cxnSp`, `grpSp` に注目)
    *   `<xdr:pic>`: 画像。

**名前空間:**

*   DrawingMLでは主に以下のXML名前空間が使用されます。
    *   `xdr` (または `xsd` など、ファイルによってプレフィックスは異なる): `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` - Excel固有の描画要素（アンカーなど）を定義。 (コード内 `NS_SPREADSHEETDRAWING`)
    *   `a` (または `a14` など): `http://schemas.openxmlformats.org/drawingml/2006/main` - Office共通の描画要素（図形の形状、色、テキストなど）を定義。 (コード内 `NS_DRAWINGML`)
    *   `r`: `http://schemas.openxmlformats.org/officeDocument/2006/relationships` - リレーションシップID（画像ファイルへのリンクなど）を指定。 (コード内 `NS_RELATIONSHIPS`)

### 2. Node (図形: `xdr:sp`) の抽出仕様

**対象要素:**

*   主に `<xdr:sp>` 要素がNodeとして抽出されます。ただし、`prst="line"` などの直線形状はEdgeとして扱われます（後述）。

**属性の取得方法:**

*   **ID:**
    *   各 `<xdr:sp>` 要素内には、`<xdr:nvSpPr>` (Non-Visual Shape Properties) があり、その中に `<xdr:cNvPr>` (Non-Visual Drawing Properties) が含まれます。
    *   `<xdr:cNvPr>` の `id` 属性が一意な識別子となります。
    *   コードでは `processShapeBase` 関数内で `prefix + "_" + cNvPr.getAttribute("id")` として `drawing*.xml` ファイル名をプレフィックスに付与したIDを生成します。
*   **Type (種類):**
    *   `<xdr:spPr>` (Shape Properties) 内の `<a:prstGeom>` (Preset Geometry) 要素の `prst` 属性値が図形の種類（例: `rectangle`, `ellipse`, `roundRectangle`, `line`）を示します。
    *   `prst` 属性が存在しない場合や `<a:prstGeom>` がない場合は、コードでは `"custom"` または `"line"` (直線の場合) として扱います。
*   **Label (テキスト):**
    *   `<xdr:txBody>` (Text Body) 要素内に図形のテキストコンテンツが含まれます。
    *   テキストは `<a:p>` (Paragraph) 要素の中に `<a:r>` (Run) 要素があり、その中の `<a:t>` (Text) 要素に実際の文字列が記述されます。一つの図形内に複数の `<a:p>` や `<a:r>` が存在することがあります（改行や書式変更のため）。
    *   コード内の `extractTextFromElement` 関数が、これらの `<a:t>` 要素のテキストを連結し、改行を保持して抽出します（`formatToTextBySheet` で `\n` に置換）。`trim()` で前後の空白は除去されます。
*   **Position (位置):**
    *   Nodeの位置は、その `<xdr:sp>` を内包する親のアンカー要素 (`<xdr:twoCellAnchor>` または `<xdr:oneCellAnchor>`) によって決まります。
    *   アンカー要素内の `<xdr:from>` および `<xdr:to>` (twoCellAnchorの場合) 要素が、図形のバウンディングボックスの開始/終了セルとセル内オフセットを定義します。
    *   **座標系の詳細 (Anchor, EMU, セル座標変換):**
        *   `<xdr:from>` / `<xdr:to>` 内には以下の要素があります。
            *   `<xdr:col>`: 列インデックス (0始まり)
            *   `<xdr:row>`: 行インデックス (0始まり)
            *   `<xdr:colOff>`: 列内のオフセット (EMU単位)
            *   `<xdr:rowOff>`: 行内のオフセット (EMU単位)
        *   **EMU (English Metric Units):** Office文書で使われる単位で、1インチ = 914400 EMU、1ポイント = 12700 EMU です。非常に細かい精度で位置を指定できます。
        *   コード内の `getAnchorEndpoint` 関数が、これらの値とグローバル定数 (`DEFAULT_ROW_HEIGHT_EMU`, `DEFAULT_COL_WIDTH_EMU`) を用いて、EMUオフセットを考慮した**小数値のセル座標** (`row`, `col`) を計算します。
            *   `row = baseRow + rowOff / DEFAULT_ROW_HEIGHT_EMU`
            *   `col = baseCol + colOff / DEFAULT_COL_WIDTH_EMU`
        *   計算された `fromRow`, `fromCol`, `toRow`, `toCol` がNodeオブジェクトに格納されます。
    *   **`NODE_POSITION_OUTPUT_MODE` の解説:**
        *   `formatToTextBySheet` 関数内で、Nodeの位置情報をどのようにテキスト出力するかを制御します。
            *   `1`: `RowFrom`, `ColFrom`, `RowTo`, `ColTo` の4つの座標を出力（バウンディングボックス）。
            *   `2` (デフォルト): `RowFrom`/`RowTo`、`ColFrom`/`ColTo` の中点 (`Row`, `Col`) を出力。
            *   `3`: `RowFrom`, `ColFrom` の座標のみ (`Row`, `Col`) を出力。
        *   出力前に座標値に係数 (`ROW_COORD_MULTIPLIER`, `COL_COORD_MULTIPLIER`) が乗算され、`Math.round()` で整数化されます。

### 3. Edge (コネクタ: `xdr:cxnSp` / 直線: `xdr:sp`) の抽出仕様

**対象要素:**

*   `<xdr:cxnSp>` (Connection Shape): 明示的なコネクタ要素。
*   `<xdr:sp>` で `<a:prstGeom prst="line">` または `prst="lineInv"` を持つもの: 単純な直線もEdgeとして扱われます。

**属性の取得方法:**

*   **ID:**
    *   `<xdr:cxnSp>` の場合: `<xdr:nvCxnSpPr><xdr:cNvPr id="...">` から取得。
    *   `<xdr:sp>` (直線) の場合: `<xdr:nvSpPr><xdr:cNvPr id="...">` から取得。
    *   コードでは `processConnectorBase` / `processLineShape` でプレフィックス付きIDを生成します。コネクタにIDがない場合 (`cNvPr` が欠落) は `anonCxn_N` という形式で自動採番されます (`__CXN_AUTO_ID` カウンタ使用)。
*   **Type (種類):**
    *   `<xdr:cxnSp>` / `<xdr:sp>` 内の `<xdr:spPr><a:prstGeom prst="...">` 属性値。コネクタの場合は `bentConnector3`, `curvedConnector2` など、直線の場合は `line` などになります。指定がない場合は `customConnector` や `line` になります。
*   **Label (テキスト):**
    *   Nodeと同様に `<xdr:txBody>` 内のテキストを `extractTextFromElement` で抽出します。コネクタにラベルが付いている場合に使われます。
*   **Source (接続元) / Target (接続先):**
    *   **`<xdr:cxnSp>` の場合:**
        *   `<xdr:nvCxnSpPr><xdr:cNvCxnSpPr>` (Non-Visual Connection Shape Properties) 内の要素で定義されます。
        *   `<a:stCxn>` (Start Connection): 接続元の情報を持ち、`id` 属性に接続先 **Node (図形)** のID (数値) が入ります。
        *   `<a:endCxn>` (End Connection): 接続先の情報を持ち、`id` 属性に接続先 **Node (図形)** のID (数値) が入ります。
        *   `id="0"` の場合は、その端点がどの図形にも接続されていないことを意味します。
        *   コード (`processConnectorBase`) では、これらの `id` にプレフィックスを付けて `source`, `target` プロパティに格納します (`0` の場合は `null`)。
    *   **`<xdr:sp>` (直線) の場合:**
        *   XML構造上、直線 (`<xdr:sp>`) には接続情報 (`<a:stCxn>`, `<a:endCxn>`) が含まれません。
        *   そのため、コード (`processLineShape`) では `source`, `target` を初期値 `null` として生成します。
        *   後続の **近傍補完ロジック** (`extractStructure` 内の "近傍補完" セクション) で、直線の端点の座標 (`getAnchorEndpoint` で取得) と、既存のNode (sp由来) の矩形 (`nodeRectMap`) とのユークリッド距離を計算します。
        *   距離がしきい値 (`NEAR_SHAPE_THRESHOLD`) 以下で最も近いNodeが見つかれば、そのNodeのIDを `source` / `target` に設定します。これにより、視覚的に接続されているように見える直線もEdgeとして扱えるようになります。
        *   コネクタ (`cxnSp`) で `id="0"` (未接続) の場合も、同様に近傍補完が試みられます。

### 4. グループ (`xdr:grpSp`) の処理仕様

**グループ要素の認識:**

*   `<xdr:grpSp>` (Group Shape) 要素がグループを表します。
*   グループ自体もIDを持ちます: `<xdr:nvGrpSpPr><xdr:cNvPr id="...">` から取得 (`getGroupId` 関数)。

**2つのグループ化ロジック (`USE_LEGACY_GROUPING_LOGIC`):**

Excelのグループは階層構造を持つことができます。コードでは、このグループ情報をどのようにNodes/Edgesリストに反映させるかについて、2つの異なるロジックを提供しています。これはグローバル変数 `USE_LEGACY_GROUPING_LOGIC` (boolean) によって切り替えられます。

*   **旧ロジック (`true`): 階層型 (Type: "group")**
    *   各 `<xdr:grpSp>` 要素を、`type: "group"` という種類の **独立したNode** として `nodes` 配列に追加します。
    *   グループNodeには自身のIDと、オプションで親グループのID (`ParentGroupID`) が付与されます (`findParentGroupId` で探索)。
    *   グループ内に直接含まれる `<xdr:sp>` や `<xdr:cxnSp>` 要素 (Node/Edge) には、それらが属する直接の親グループのIDが `GroupID` プロパティとして設定されます。
    *   **メリット:** XMLの階層構造を比較的忠実に表現します。
    *   **デメリット:** グループ自体がNodeとなるため、視覚的な要素（グループ内の図形）と構造的な要素（グループNode）が混在し、グラフ構造としては扱いにくい場合があります。グループ内の図形のテキストは個々のNodeに残ります。

*   **新ロジック (`false`, デフォルト): 集約型 (Type: "groupAggregation")**
    *   グループ化されている `<xdr:sp>` (Node) を特定し、それらが属する **最上位の** `<xdr:grpSp>` (祖先グループ) を見つけます。
    *   最上位のグループごとに、`type: "groupAggregation"` という種類の **単一のNode** を生成します。
    *   **子Nodeの集約:**
        *   元のグループ内の `<xdr:sp>` (テキストを持つもの) のラベル (`label`) を、区切り文字 (`GROUP_LABEL_SEPARATOR`) で連結し、生成された `groupAggregation` Nodeのラベルとします。
        *   元のグループ内の `<xdr:sp>` Nodeは `nodes` 配列から **削除** されます。
    *   **エッジの付け替え:**
        *   グループ内の子Nodeに接続されていたEdge (`source` または `target` が子Node IDだったもの) は、接続先が親である `groupAggregation` NodeのIDに **付け替え** られます。
    *   グループの座標 (`fromRow`, `fromCol`, `toRow`, `toCol`) は、削除された子Node群全体のバウンディングボックスとして計算され、`groupAggregation` Nodeに設定されます。
    *   `GroupID`, `ParentGroupID` は基本的には使用されません（`null`）。
    *   **メリット:** グラフ構造がシンプルになり、視覚的なグループを一つのまとまったNodeとして扱えます。グループ全体のテキスト情報も集約されます。
    *   **デメリット:** 元の細かい階層構造や、グループ内の個々の図形の情報は失われます。

**グループIDと親子関係:**

*   旧ロジックでは `GroupID` (直接の親) と `ParentGroupID` (グループ自身の親) を使って階層を表現します。
*   新ロジックでは `groupAggregation` Nodeが最上位グループを表し、親子関係は明示的には保持されません。

### 5. 座標と単位について

*   **EMU (English Metric Units):** 前述の通り、Office文書内部での標準単位。`1pt = 12700 EMU`。
*   **アンカー要素内の座標情報:** `<xdr:from>`, `<xdr:to>` 内の `<xdr:row>`, `<xdr:col>`, `<xdr:rowOff>`, `<xdr:colOff>` が基本情報。
*   **EMUからセル座標への変換:**
    *   `getAnchorEndpoint` 関数内で、オフセット(EMU)をデフォルトの行高/列幅(EMU)で割ることで、セル内での相対位置（0.0～1.0）を計算し、ベースの行/列インデックスに加算します。
    *   使用される定数:
        *   `EMU_PER_POINT` (12700)
        *   `DEFAULT_ROW_HEIGHT_PT` (15) → `DEFAULT_ROW_HEIGHT_EMU` (190500)
        *   `EMU_PER_PIXEL` (9525)
        *   `DEFAULT_COL_WIDTH_PX` (64) → `DEFAULT_COL_WIDTH_EMU` (609600)
    *   **注意:** これらのデフォルト値はExcelの標準設定に基づいていますが、ユーザーが変更した行高や列幅には対応していません。そのため、実際の表示と計算結果にずれが生じる可能性があります。より正確な座標を得るには、シートの行高・列幅情報 (`<row>`, `<col>` 要素の属性) を別途解析する必要があります。
*   **座標表示係数:**
    *   `ROW_COORD_MULTIPLIER`, `COL_COORD_MULTIPLIER`: 最終的なテキスト出力時に、計算されたセル座標に乗算される係数。例えば、座標のスケールを変更したい場合などに使用できます (デフォルトは 1)。

### 6. コード内の主要な処理フロー解説

1.  **`extractStructuredShapesFromExcel(arrayBuffer)`:**
    *   XLSXファイルをJSZipで展開。
    *   `xl/workbook.xml` を `parseWorkbook` で解析し、シート名とr:Idを取得。
    *   `xl/_rels/workbook.xml.rels` と各シートの `_rels/*.xml.rels` を `parseSheetRelationships` で解析し、シート名と対応する `drawing*.xml` のパスをマッピング (`sheetDrawingMap`)。
    *   各 `drawing*.xml` を読み込み、DOMParserでXMLドキュメントに変換。
    *   各XMLドキュメントと描画パスを `extractStructure` に渡し、Nodes/Edgesリストを取得。
    *   結果をシートごとに集約 (`resultsBySheet`)。
    *   `formatToTextBySheet` で最終的な構造化テキストを生成して返す。
2.  **`extractStructure(xmlDoc, drawingPath)`:**
    *   `drawing*.xml` ドキュメントを受け取り、Nodes/Edgesを抽出するコアロジック。
    *   **1stパス:** `xmlDoc` 内の `<xdr:sp>`, `<xdr:cxnSp>`, `<xdr:grpSp>` を走査。
        *   `sp` は `processShapeBase` (Node) または `processLineShape` (Edge) で処理。
        *   `cxnSp` は `processConnectorBase` (Edge) で処理。
        *   `grpSp` は要素を収集 (`groupElements`)。
        *   各要素を `elementMap` に ID と共に追加。
    *   **座標取得:** `nodes` 配列 (現時点では `sp` 由来のみ) の各要素について `getAnchorEndpoint` を呼び出し、`fromRow/Col`, `toRow/Col` を設定。同時に矩形情報 `nodeRectMap` を作成。
    *   **近傍補完:** `edges` 配列を走査し、`source` または `target` が `null` のものについて、端点座標と `nodeRectMap` を比較し、近ければ接続を補完。
    *   **グループ処理:** `USE_LEGACY_GROUPING_LOGIC` の値に応じて、旧ロジックまたは新ロジックを実行。
        *   旧: `grpSp` から `type: "group"` Node生成、子要素に `GroupID`/`ParentGroupID` 付与。
        *   新: 最上位 `grpSp` ごとに `type: "groupAggregation"` Node生成、子Node集約＆削除、エッジ付け替え、座標再計算。
    *   **フィルタリング:** 自己ループ (`source === target`) や、存在しないNodeを指すEdge (`source`/`target` が `nodes` 配列にない) を `edges` 配列から除去。
    *   最終的な `nodes` と `edges` を含むオブジェクトを返す。
3.  **Helper Functions:**
    *   `processShapeBase`, `processConnectorBase`, `processLineShape`: 各要素タイプから基本情報を抽出してオブジェクト化。
    *   `getAnchorEndpoint`: 要素を遡ってアンカーを見つけ、指定された端点 (`from` or `to`) のセル座標を計算。
    *   `extractTextFromElement`: 要素内の `<a:t>` テキストを抽出・結合。
    *   `getGroupId`, `findParentGroupId`: グループIDや親グループIDを取得。
    *   `formatToTextBySheet`: 抽出結果を指定形式のテキストに整形。

### 7. 注意点と限界

*   **DrawingMLの複雑性:** ネストしたグループ、複雑な図形（SmartArt, Chart）、代替テキスト、効果などは完全には解釈されません。VML形式の図形 (`*.vml`) の扱いは限定的です (テキスト抽出のみ試行)。
*   **座標精度の限界:** 前述の通り、行高・列幅の変更に対応していないため、座標の精度には限界があります。`NEAR_SHAPE_THRESHOLD` の値は、この不正確さも考慮して調整する必要があります。
*   **接続補完の精度:** 近傍探索による接続補完は、図形が密集している場合に意図しない接続をする可能性があります。
*   **テキスト抽出:** 書式情報（色、太字、サイズなど）は失われ、プレーンテキストのみが抽出されます。

### まとめ

このコードは、Excelの `drawing*.xml` ファイルを解析し、図形 (Node) とコネクタ/直線 (Edge) を抽出するための比較的高度なロジックを実装しています。特に、アンカー情報からの座標計算、EMU単位の扱い、コネクタ接続情報の解析、近傍探索による接続補完、そして2種類のグループ化戦略（階層型 vs 集約型）の選択肢を提供している点が特徴です。Excel DrawingMLの仕様を理解することで、コードの挙動や設定値（定数、フラグ）の意味をより深く把握し、必要に応じたカスタマイズやデバッグが可能になります。