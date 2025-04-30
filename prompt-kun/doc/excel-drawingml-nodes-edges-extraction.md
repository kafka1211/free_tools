## Excel DrawingML (`drawing*.xml`) から Nodes/Edges を抽出する仕様解説

### 機能概要

このドキュメントで解説するコードは、Excelファイル (.xlsx) に含まれる図形描画情報 (DrawingML) を解析し、シート上の図形とそれらを繋ぐ線を構造化されたテキストデータとして抽出する機能を提供します。主な機能は以下の通りです。

*   **Node (図形) 抽出:** シート上のオートシェイプ、テキストボックスなどを「Node」として認識し、ID、種類 (rectangle, ellipse など)、テキストラベル、セル座標に基づいた位置情報を抽出します。
*   **Edge (コネクタ/直線) 抽出:** 図形間を接続するコネクタや、単純な直線を「Edge」として認識し、ID、種類、テキストラベル、接続元 (Source) と接続先 (Target) のNode IDを抽出します。
*   **接続情報の解析:** コネクタ要素 (`<xdr:cxnSp>`) が持つ接続先ID (`<a:stCxn>`, `<a:endCxn>`) を解析します。
*   **近傍接続補完 (オプション):** 直線要素 (`<xdr:sp prst="line">`) や未接続のコネクタ端点について、座標的に最も近い図形 (Node) を探索し、接続情報を補完します (`ENABLE_NEARBY_COMPLETION` フラグ)。
*   **グループ処理 (選択可能):** 図形のグループ (`<xdr:grpSp>`) を処理する2つのロジックを提供します。
    *   **階層型:** グループ自体を独立したNode (`type: "group"`) として扱い、XMLの階層を保持します。
    *   **集約型 (デフォルト):** グループ内の図形を一つの集約Node (`type: "groupAggregation"`) にまとめ、テキストを集約し、接続されていたEdgeを付け替えます (`USE_LEGACY_GROUPING_LOGIC` フラグ)。
*   **座標計算:** DrawingML内のEMU単位の位置情報を、シートのセル座標 (行/列の小数値) に変換します。出力形式 (バウンディングボックス、中心点、左上点) は選択可能です (`NODE_POSITION_OUTPUT_MODE` フラグ)。
*   **空ノード接続付け替え (オプション):** テキストを持たない図形 (空ノード) への接続を、その空ノードを包含する別の意味のあるノードへと自動的に付け替えます (`ENABLE_CONNECTOR_RETARGETING` フラグ)。
*   **フィルタリング:** 接続情報を持たない孤立した空ノードや、存在しないノードを指すエッジ（幽霊端点）などを除去します。
*   **構造化テキスト出力:** 抽出したNodesとEdgesの情報を、指定されたフォーマットのテキストとして出力します。

以下のセクションでは、これらの機能を実現するためのDrawingMLの仕様と、コード内の具体的な処理ロジックについて詳しく解説します。

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

*   DrawingMLでは主に以下のXML名前空間が使用されます。コード内ではグローバル変数として定義されています。
    *   `xdr` (または `xsd` など、ファイルによってプレフィックスは異なる): `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` - Excel固有の描画要素（アンカーなど）を定義。 (コード内 `NS_SPREADSHEETDRAWING`)
    *   `a` (または `a14` など): `http://schemas.openxmlformats.org/drawingml/2006/main` - Office共通の描画要素（図形の形状、色、テキストなど）を定義。 (コード内 `NS_DRAWINGML`)
    *   `r`: `http://schemas.openxmlformats.org/officeDocument/2006/relationships` - リレーションシップID（画像ファイルへのリンクなど）を指定。 (コード内 `NS_RELATIONSHIPS`)

### 2. Node (図形: `xdr:sp`) の抽出仕様

**対象要素:**

*   主に `<xdr:sp>` 要素がNodeとして抽出されます。ただし、`prst="line"` などの直線形状はEdgeとして扱われます（後述）。
*   グループ化ロジックによっては `<xdr:grpSp>` もNode (`type: "group"` または `type: "groupAggregation"`) として扱われます（詳細はセクション4参照）。

**属性の取得方法:**

*   **ID:**
    *   各 `<xdr:sp>` 要素内には、`<xdr:nvSpPr>` (Non-Visual Shape Properties) があり、その中に `<xdr:cNvPr>` (Non-Visual Drawing Properties) が含まれます。
    *   `<xdr:cNvPr>` の `id` 属性が一意な識別子となります。
    *   コードでは `processShapeBase` 関数内で `prefix + "_" + cNvPr.getAttribute("id")` として `drawing*.xml` ファイル名をプレフィックスに付与したIDを生成します。
*   **Type (種類):**
    *   `<xdr:spPr>` (Shape Properties) 内の `<a:prstGeom>` (Preset Geometry) 要素の `prst` 属性値が図形の種類（例: `rectangle`, `ellipse`, `roundRectangle`, `line`）を示します。
    *   `prst` 属性が存在しない場合や `<a:prstGeom>` がない場合は、コードでは `"custom"` として扱います。
    *   最終的な出力テキストに `Type:` 情報を含めるかどうかは、`OUTPUT_NODE_TYPE` フラグで制御されます。
*   **Label (テキスト):**
    *   `<xdr:txBody>` (Text Body) 要素内に図形のテキストコンテンツが含まれます。
    *   テキストは `<a:p>` (Paragraph) 要素の中に `<a:r>` (Run) 要素があり、その中の `<a:t>` (Text) 要素に実際の文字列が記述されます。一つの図形内に複数の `<a:p>` や `<a:r>` が存在することがあります（改行や書式変更のため）。
    *   コード内の `extractTextFromElement` 関数が、これらの `<a:t>` 要素のテキストを連結し、改行を保持して抽出します（`formatToTextBySheet` で `window.GROUP_LABEL_SEPARATOR` (デフォルト `\n`) に置換）。`trim()` で前後の空白は除去されます。
*   **HasText (テキスト有無):**
    *   `processShapeBase` 関数内で、抽出された `label` が空文字列でないかどうかを判定し、`hasText` (boolean) フラグを設定します。これは後のフィルタリングや接続先付け替えロジックで使用されます。
*   **Position (位置):**
    *   Nodeの位置は、その `<xdr:sp>` (または `<xdr:grpSp>`) を内包する親のアンカー要素 (`<xdr:twoCellAnchor>` または `<xdr:oneCellAnchor>`) によって決まります。
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
        *   計算された `fromRow`, `fromCol`, `toRow`, `toCol` がNodeオブジェクトに格納されます (`type: "group"` のNodeでは `Infinity` になります)。
    *   **`NODE_POSITION_OUTPUT_MODE` の解説:**
        *   `formatToTextBySheet` 関数内で、Nodeの位置情報をどのようにテキスト出力するかを制御します (`type: "group"` 以外の場合)。
            *   `1`: `RowFrom`, `ColFrom`, `RowTo`, `ColTo` の4つの座標を出力（バウンディングボックス）。
            *   `2` (デフォルト): `RowFrom`/`RowTo`、`ColFrom`/`ColTo` の中点 (`Row`, `Col`) を出力。
            *   `3`: `RowFrom`, `ColFrom` の座標のみ (`Row`, `Col`) を出力。
        *   出力前に座標値に係数 (`ROW_COORD_MULTIPLIER`, `COL_COORD_MULTIPLIER`) が乗算され、`Math.round()` で整数化されます。`Infinity` は `"inf"` として出力されます。
*   **GroupID / ParentGroupID:**
    *   グループ化ロジックによって設定されます（詳細はセクション4参照）。

### 3. Edge (コネクタ: `xdr:cxnSp` / 直線: `xdr:sp`) の抽出仕様

**対象要素:**

*   `<xdr:cxnSp>` (Connection Shape): 明示的なコネクタ要素。
*   `<xdr:sp>` で `<a:prstGeom prst="line">` または `prst="lineInv"` を持つもの: 単純な直線もEdgeとして扱われます。

**属性の取得方法:**

*   **ID:**
    *   `<xdr:cxnSp>` の場合: `<xdr:nvCxnSpPr><xdr:cNvPr id="...">` から取得。
    *   `<xdr:sp>` (直線) の場合: `<xdr:nvSpPr><xdr:cNvPr id="...">` から取得。
    *   コードでは `processConnectorBase` / `processLineShape` でプレフィックス付きIDを生成します。
    *   コネクタ要素 (`<xdr:cxnSp>`) に `<xdr:cNvPr>` がなくIDが取得できない場合、コードは `prefix_anonCxn_N` 形式で一意なIDを自動採番します (`window.__CXN_AUTO_ID` グローバルカウンタを使用)。
*   **Type (種類):**
    *   `<xdr:cxnSp>` / `<xdr:sp>` 内の `<xdr:spPr><a:prstGeom prst="...">` 属性値。コネクタの場合は `bentConnector3`, `curvedConnector2` など、直線の場合は `line` などになります。指定がない場合は `customConnector` や `line` になります。
    *   最終的な出力テキストに `Type:` 情報を含めるかどうかは、`OUTPUT_NODE_TYPE` フラグで制御されます。
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
        *   この近傍補完ロジックは、`ENABLE_NEARBY_COMPLETION` フラグが `true` の場合にのみ実行されます。
*   **GroupID:**
    *   旧グループ化ロジック (`USE_LEGACY_GROUPING_LOGIC = true`) の場合、属する親グループのIDが設定されます。新ロジックでは通常 `null` です。

### 4. グループ (`xdr:grpSp`) の処理仕様

**グループ要素の認識:**

*   `<xdr:grpSp>` (Group Shape) 要素がグループを表します。
*   グループ自体もIDを持ちます: `<xdr:nvGrpSpPr><xdr:cNvPr id="...">` から取得 (`getGroupId` 関数)。

**2つのグループ化ロジック (`USE_LEGACY_GROUPING_LOGIC`):**

Excelのグループは階層構造を持つことができます。コードでは、このグループ情報をどのようにNodes/Edgesリストに反映させるかについて、2つの異なるロジックを提供しています。これはグローバル変数 `USE_LEGACY_GROUPING_LOGIC` (boolean) によって切り替えられます。

*   **旧ロジック (`true`): 階層型 (Type: "group")**
    *   各 `<xdr:grpSp>` 要素を、`type: "group"` という種類の **独立したNode** として `nodes` 配列に追加します。
    *   グループNodeには自身のID (`id`) と、オプションで親グループのID (`ParentGroupID`) が付与されます (`findParentGroupId` で探索)。自身のグループIDを示す `GroupID` は持ちません (`null`)。
    *   グループ内に直接含まれる `<xdr:sp>` や `<xdr:cxnSp>` 要素 (Node/Edge) には、それらが属する直接の親グループのIDが `GroupID` プロパティとして設定されます。
    *   **メリット:** XMLの階層構造を比較的忠実に表現します。
    *   **デメリット:** グループ自体がNodeとなるため、視覚的な要素（グループ内の図形）と構造的な要素（グループNode）が混在し、グラフ構造としては扱いにくい場合があります。グループ内の図形のテキストは個々のNodeに残ります。出力形式も `type: "group"` 専用のものが適用されます。

*   **新ロジック (`false`, デフォルト): 集約型 (Type: "groupAggregation")**
    *   グループ化されている `<xdr:sp>` (Node) を特定し、それらが属する **最上位の** `<xdr:grpSp>` (祖先グループ) を見つけます。
    *   最上位のグループごとに、`type: "groupAggregation"` という種類の **単一のNode** を生成します。
    *   **子Nodeの集約:**
        *   元のグループ内の `<xdr:sp>` (テキストを持つもの, `hasText: true`) のラベル (`label`) を、区切り文字 (`GROUP_LABEL_SEPARATOR`) で連結し、生成された `groupAggregation` Nodeのラベルとします。
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

*   **EMU (English Metric Units):** 前述の通り、Office文書内部での標準単位。`1pt = 12700 EMU`。コード内定数 `EMU_PER_POINT` で定義。
*   **アンカー要素内の座標情報:** `<xdr:from>`, `<xdr:to>` 内の `<xdr:row>`, `<xdr:col>`, `<xdr:rowOff>`, `<xdr:colOff>` が基本情報。
*   **EMUからセル座標への変換:**
    *   `getAnchorEndpoint` 関数内で、オフセット(EMU)をデフォルトの行高/列幅(EMU)で割ることで、セル内での相対位置（0.0～1.0）を計算し、ベースの行/列インデックスに加算します。
    *   使用される定数:
        *   `EMU_PER_POINT` (12700)
        *   `DEFAULT_ROW_HEIGHT_PT` (15) → `DEFAULT_ROW_HEIGHT_EMU` (190500)
        *   `EMU_PER_PIXEL` (9525)
        *   `DEFAULT_COL_WIDTH_PX` (64) → `DEFAULT_COL_WIDTH_EMU` (609600)
    *   **注意:** これらのデフォルト値はExcelの標準設定に基づいていますが、ユーザーが変更した行高や列幅には対応していません。そのため、実際の表示と計算結果にずれが生じる可能性があります。より正確な座標を得るには、シートの行高・列幅情報 (`<row>`, `<col>` 要素の属性) を別途解析する必要があります（現在のコードでは未対応）。
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
        *   `cxnSp` は `processConnectorBase` (Edge) で処理。ID欠落時は自動採番。
        *   `grpSp` は要素を収集 (`groupElements`)。
        *   各要素を `elementMap` に ID と共に追加。接続されたNode IDを `connectedNodeIds` に記録。
    *   **座標取得:** `nodes` 配列 (現時点では `sp` 由来のみ) の各要素について `getAnchorEndpoint` を呼び出し、`fromRow/Col`, `toRow/Col` を設定。同時に矩形情報 `nodeRectMap` を作成。
    *   **近傍補完 (オプション):** `ENABLE_NEARBY_COMPLETION` が `true` の場合、`edges` 配列を走査し、`source` または `target` が `null` のものについて、端点座標と `nodeRectMap` を比較し、近ければ (`NEAR_SHAPE_THRESHOLD` 内) 接続を補完。
    *   **グループ処理:** `USE_LEGACY_GROUPING_LOGIC` の値に応じて、旧ロジックまたは新ロジックを実行。
        *   旧: `grpSp` から `type: "group"` Node生成、子要素に `GroupID`/`ParentGroupID` 付与。
        *   新: 最上位 `grpSp` ごとに `type: "groupAggregation"` Node生成、子Node集約＆削除、エッジ付け替え、座標再計算。
    *   **接続先付け替え (オプション):** `ENABLE_CONNECTOR_RETARGETING` が `true` の場合、テキストを持たない空のNode (`hasText` が `false`) への接続を、その空Nodeを包含する別のNodeへと付け替える。
    *   **幽霊端点除去:** `edges` 配列を走査し、`source` または `target` が `null` (補完・付け替え後も未接続) または存在しないNode IDを指している Edge を除去。
    *   **最終フィルタリング:**
        *   **Nodes:** テキストを持たず (`hasText` が `false`)、かつどのEdgeからも接続されていない (`connectedNodeIds` に含まれない) Node (幽霊ノード) を除去。
        *   **Edges:** (現在は特別なフィルタリングなし。幽霊端点除去で対応済み)
    *   最終的な `nodes` と `edges` を含むオブジェクトを返す。
3.  **Helper Functions:**
    *   `processShapeBase`, `processConnectorBase`, `processLineShape`: 各要素タイプから基本情報を抽出してオブジェクト化。`processShapeBase` は `hasText` フラグも設定。
    *   `getAnchorEndpoint`: 要素を遡ってアンカーを見つけ、指定された端点 (`from` or `to`) のセル座標 (小数値) を計算。
    *   `extractTextFromElement`: 要素内の `<a:t>` テキストを抽出・結合。
    *   `getGroupId`, `findParentGroupId`: グループIDや親グループIDを取得。
    *   `formatToTextBySheet`: 抽出結果を指定形式のテキストに整形。`NODE_POSITION_OUTPUT_MODE` や `OUTPUT_NODE_TYPE` フラグ、座標係数などを考慮。`type: "group"` の場合とそれ以外で出力形式を分ける。

### 7. 注意点と限界

*   **DrawingMLの複雑性:** ネストしたグループ、複雑な図形（SmartArt, Chart）、代替テキスト、効果などは完全には解釈されません。VML形式の図形 (`*.vml`) の扱いは限定的です (テキスト抽出のみ試行、座標なし)。
*   **座標精度の限界:** 前述の通り、行高・列幅のユーザー変更に対応していないため、座標の精度には限界があります。`NEAR_SHAPE_THRESHOLD` の値は、この不正確さも考慮して調整する必要があります。
*   **接続補完/付け替えの精度:** 近傍探索による接続補完や、包含関係による接続先付け替えは、図形が密集・重なり合っている場合に意図しない結果を生む可能性があります。しきい値 (`NEAR_SHAPE_THRESHOLD`) や包含判定ロジックの調整が必要になる場合があります。
*   **テキスト抽出:** 書式情報（色、太字、サイズなど）は失われ、プレーンテキストのみが抽出されます。

### 8. まとめ

このコードは、Excelの `drawing*.xml` ファイルを解析し、図形 (Node) とコネクタ/直線 (Edge) を抽出するための比較的高度なロジックを実装しています。特に、アンカー情報からの座標計算、EMU単位の扱い、コネクタ接続情報の解析、近傍探索による接続補完、空ノードへの接続付け替え (`ENABLE_CONNECTOR_RETARGETING`)、そして2種類のグループ化戦略（階層型 vs 集約型、`USE_LEGACY_GROUPING_LOGIC`）の選択肢を提供している点が特徴です。各種フラグ (`ENABLE_NEARBY_COMPLETION`, `NODE_POSITION_OUTPUT_MODE`, `OUTPUT_NODE_TYPE` など) により、抽出や出力の挙動を制御できます。Excel DrawingMLの仕様を理解することで、コードの挙動や設定値（定数、フラグ）の意味をより深く把握し、必要に応じたカスタマイズやデバッグが可能になります。

### 9. 補足 具体的なXMLサンプルとその解説

Excel の `drawing*.xml` ファイルがどのような構造になっているか、そして提供されたコードがそのXMLをどのように解析して Nodes/Edges 情報を抽出しているかをイメージしやすくするために、具体的なXMLサンプルとその解説を以下に示します。

**注意点:**

*   実際の `drawing*.xml` は、スタイル情報 (色、線種、効果など) や他のプロパティを含み、非常に長くなることがあります。以下のサンプルは、Node/Edge 抽出ロジックの理解に必要な要素に焦点を当て、簡略化しています。
*   名前空間のプレフィックス (`xdr:`, `a:` など) はファイルによって異なる場合がありますが、通常はこの形式です。
*   `id` 属性の値はExcelが自動で割り振る数値です。コードではファイル名のプレフィックスを付けて一意性を確保します（例: `drawing1_1`）。

---

**`xl/drawings/drawing1.xml` のサンプル:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <!-- ルート要素: Worksheet Drawing -->

    <!-- ===== 例1: 単純な四角形 (Node) - twoCellAnchor ===== -->
    <!-- 図形が複数のセル範囲に固定される -->
    <xdr:twoCellAnchor editAs="oneCell">
        <!-- 開始位置: 列B (1), 行2 (1) の左上から -->
        <xdr:from>
            <xdr:col>1</xdr:col>       <!-- 列インデックス (0始まり) -->
            <xdr:colOff>0</xdr:colOff> <!-- 列内のオフセット (EMU単位) -->
            <xdr:row>1</xdr:row>       <!-- 行インデックス (0始まり) -->
            <xdr:rowOff>0</xdr:rowOff> <!-- 行内のオフセット (EMU単位) -->
        </xdr:from>
        <!-- 終了位置: 列D (3), 行4 (3) の途中まで -->
        <xdr:to>
            <xdr:col>3</xdr:col>
            <xdr:colOff>285750</xdr:colOff> <!-- 約 0.47列 -->
            <xdr:row>3</xdr:row>
            <xdr:rowOff>95250</xdr:rowOff>  <!-- 約 0.5行 -->
        </xdr:to>
        <!-- ▼ 図形本体 (Shape) - これがNodeになる ▼ -->
        <xdr:sp macro="" textlink="">
            <!-- 非表示プロパティ (IDなど) -->
            <xdr:nvSpPr>
                <!-- ★ コードはここの 'id' を Node ID として使用 (例: drawing1_1) -->
                <xdr:cNvPr id="1" name="Rectangle 1"/>
                <xdr:cNvSpPr/> <!-- 図形固有の非表示プロパティ -->
            </xdr:nvSpPr>
            <!-- 表示プロパティ (形状、スタイル、テキストなど) -->
            <xdr:spPr>
                <!-- ★ 図形の形状 (Preset Geometry) - これが Node の Type になる -->
                <a:prstGeom prst="rectangle"> <!-- Type: rectangle -->
                    <a:avLst/> <!-- 調整値リスト (空) -->
                </a:prstGeom>
                <!-- ... ここに色、線などのスタイル情報が大量に入る (省略) ... -->
            </xdr:spPr>
            <!-- ★ 図形内のテキスト - これが Node の Label になり、hasText=true となる -->
            <xdr:txBody>
                <a:bodyPr/> <!-- テキストボックス全体のプロパティ -->
                <a:lstStyle/> <!-- リストスタイル (箇条書きなど) -->
                <a:p> <!-- 段落 (Paragraph) -->
                    <a:r> <!-- 書式ラン (Run) - 同じ書式のテキストまとまり -->
                        <a:rPr lang="ja-JP"/> <!-- 言語設定など -->
                        <a:t>開始ノード</a:t> <!-- ★ テキスト本体 (Label: "開始ノード") -->
                    </a:r>
                </a:p>
            </xdr:txBody>
        </xdr:sp>
        <xdr:clientData/> <!-- アプリケーション固有データ -->
    </xdr:twoCellAnchor>

    <!-- ===== 例2: 円 (Node) - oneCellAnchor ===== -->
    <!-- 図形が1つのセルに固定され、サイズは絶対値で指定 -->
    <xdr:oneCellAnchor>
        <!-- 開始位置: 列F (5), 行2 (1) の途中から -->
        <xdr:from>
            <xdr:col>5</xdr:col>
            <xdr:colOff>95250</xdr:colOff>
            <xdr:row>1</xdr:row>
            <xdr:rowOff>47625</xdr:rowOff>
        </xdr:from>
        <!-- サイズ (EMU単位) -->
        <xdr:ext cx="914400" cy="914400"/> <!-- cx:幅, cy:高さ -->
        <!-- ▼ 図形本体 (Shape) - Node ▼ -->
        <xdr:sp macro="" textlink="">
            <xdr:nvSpPr>
                <!-- ★ Node ID: drawing1_2 -->
                <xdr:cNvPr id="2" name="Oval 2"/>
                <xdr:cNvSpPr/>
            </xdr:nvSpPr>
            <xdr:spPr>
                <!-- ★ Node Type: ellipse -->
                <a:prstGeom prst="ellipse">
                    <a:avLst/>
                </a:prstGeom>
                <!-- ... スタイル情報 (省略) ... -->
            </xdr:spPr>
            <!-- ★ Node Label: "終了ノード", hasText=true -->
            <xdr:txBody>
                <a:bodyPr/> <a:lstStyle/>
                <a:p><a:r><a:t>終了ノード</a:t></a:r></a:p>
            </xdr:txBody>
        </xdr:sp>
        <xdr:clientData/>
    </xdr:oneCellAnchor>

    <!-- ===== 例3: コネクタ (Edge) - 開始ノードと終了ノードを接続 ===== -->
    <xdr:twoCellAnchor> <!-- コネクタ自身もアンカー内に配置されるが、経路を示すためのもので接続とは直接関係ない -->
        <xdr:from>
            <xdr:col>2</xdr:col> <xdr:colOff>0</xdr:colOff>
            <xdr:row>2</xdr:row> <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:to>
            <xdr:col>5</xdr:col> <xdr:colOff>0</xdr:colOff>
            <xdr:row>2</xdr:row> <xdr:rowOff>0</xdr:rowOff>
        </xdr:to>
        <!-- ▼ コネクタ本体 (Connection Shape) - これがEdgeになる ▼ -->
        <xdr:cxnSp macro="" textlink="">
            <!-- 非表示プロパティ (ID、接続情報) -->
            <xdr:nvCxnSpPr>
                <!-- ★ Edge ID: drawing1_3 -->
                <xdr:cNvPr id="3" name="Bent Connector 3"/>
                <!-- ★ コネクタ固有の非表示プロパティ (接続情報) -->
                <xdr:cNvCxnSpPr>
                    <!-- ★ 始点接続情報 (Start Connection) -->
                    <!-- 'id' 属性が接続先の Node ID (数値) を示す -->
                    <a:stCxn id="1" idx="3"/> <!-- ★ Source: Node ID 1 (drawing1_1) -->
                    <!-- ★ 終点接続情報 (End Connection) -->
                    <a:endCxn id="2" idx="1"/> <!-- ★ Target: Node ID 2 (drawing1_2) -->
                </xdr:cNvCxnSpPr>
            </xdr:nvCxnSpPr>
            <!-- 表示プロパティ (形状、スタイル) -->
            <xdr:spPr>
                <!-- ★ コネクタの形状 - これが Edge の Type になる -->
                <a:prstGeom prst="bentConnector3"> <!-- Type: bentConnector3 -->
                    <a:avLst/>
                </a:prstGeom>
                <!-- ... 線のスタイル情報 (省略) ... -->
            </xdr:spPr>
            <!-- コネクタにテキストが付く場合 (Label, hasText=false if empty) -->
             <xdr:txBody/> <!-- 空のテキストボディ -->
        </xdr:cxnSp>
        <xdr:clientData/>
    </xdr:twoCellAnchor>

     <!-- ===== 例4: グループ化された図形 ===== -->
    <xdr:twoCellAnchor> <!-- グループ全体のアンカー -->
        <xdr:from>
            <xdr:col>1</xdr:col> <xdr:row>5</xdr:row>
            <!-- ... (省略) ... -->
        </xdr:from>
        <xdr:to>
            <xdr:col>4</xdr:col> <xdr:row>8</xdr:row>
            <!-- ... (省略) ... -->
        </xdr:to>
        <!-- ▼ グループ本体 (Group Shape) ▼ -->
        <xdr:grpSp>
            <xdr:nvGrpSpPr>
                 <!-- ★ グループ自体のID (旧ロジックのGroup Node ID / 新ロジックの groupAggregation Node ID) -->
                <xdr:cNvPr id="10" name="Group 10"/>
                <xdr:cNvGrpSpPr/> <!-- グループ固有の非表示プロパティ -->
            </xdr:nvGrpSpPr>
            <xdr:grpSpPr> <!-- グループ全体の表示プロパティ (位置、サイズなど) -->
                <a:xfrm> <!-- グループ内の子要素に対する変換情報 -->
                    <!-- ... (省略) ... -->
                </a:xfrm>
            </xdr:grpSpPr>

            <!-- ▼ グループ内の図形1 (Node) ▼ -->
            <xdr:sp>
                <xdr:nvSpPr>
                    <!-- ★ Node ID: drawing1_11 -->
                    <!-- 【旧ロジック】このNodeは GroupID=drawing1_10 を持つ -->
                    <!-- 【新ロジック】このNodeは削除され、情報は groupAggregation(ID:10) に集約される -->
                    <xdr:cNvPr id="11" name="Rectangle 11"/>
                </xdr:nvSpPr>
                <xdr:spPr><a:prstGeom prst="roundRectangle"/></xdr:spPr>
                <xdr:txBody><a:p><a:r><a:t>要素 A</a:t></a:r></a:p></xdr:txBody> <!-- Label: 要素 A, hasText=true (新ロジックでの集約対象) -->
            </xdr:sp>

            <!-- ▼ グループ内の図形2 (Node) ▼ -->
            <xdr:sp>
                <xdr:nvSpPr>
                    <!-- ★ Node ID: drawing1_12 -->
                    <xdr:cNvPr id="12" name="Triangle 12"/>
                </xdr:nvSpPr>
                <xdr:spPr><a:prstGeom prst="triangle"/></xdr:spPr>
                <xdr:txBody><a:p><a:r><a:t>要素 B</a:t></a:r></a:p></xdr:txBody> <!-- Label: 要素 B, hasText=true (新ロジックでの集約対象) -->
            </xdr:sp>

            <!-- ▼ グループ内のコネクタ (Edge) - 要素Aと要素Bを接続 ▼ -->
            <xdr:cxnSp>
                <xdr:nvCxnSpPr>
                    <!-- ★ Edge ID: drawing1_13 -->
                    <xdr:cNvPr id="13" name="Connector 13"/>
                    <xdr:cNvCxnSpPr>
                        <a:stCxn id="11" idx="1"/> <!-- Source: drawing1_11 -->
                        <a:endCxn id="12" idx="1"/> <!-- Target: drawing1_12 -->
                         <!-- 【旧ロジック】このEdgeは GroupID=drawing1_10 を持つ -->
                         <!-- 【新ロジック】Source/Targetが groupAggregation(ID:10) に付け替えられる -->
                    </xdr:cNvCxnSpPr>
                </xdr:nvCxnSpPr>
                 <xdr:spPr><a:prstGeom prst="straightConnector1"/></xdr:spPr>
                 <xdr:txBody/> <!-- Label: "", hasText=false -->
            </xdr:cxnSp>

        </xdr:grpSp> <!-- グループ要素終了 -->
        <xdr:clientData/>
    </xdr:twoCellAnchor>

    <!-- ===== 例5: 単純な直線 (Edgeとして近傍補完対象) ===== -->
     <xdr:twoCellAnchor>
        <xdr:from> <!-- 直線の始点座標 -->
            <xdr:col>1</xdr:col> <xdr:colOff>200000</xdr:colOff>
            <xdr:row>4</xdr:row> <xdr:rowOff>100000</xdr:rowOff>
        </xdr:from>
        <xdr:to> <!-- 直線の終点座標 -->
            <xdr:col>2</xdr:col> <xdr:colOff>50000</xdr:colOff>
            <xdr:row>5</xdr:row> <xdr:rowOff>20000</xdr:rowOff>
        </xdr:to>
        <!-- ▼ 直線も <xdr:sp> 要素で表現される ▼ -->
        <xdr:sp macro="" textlink="">
            <xdr:nvSpPr>
                <!-- ★ Edge ID: drawing1_4 -->
                <xdr:cNvPr id="4" name="Line 4"/>
            </xdr:nvSpPr>
            <xdr:spPr>
                <!-- ★ Edge Type: line -->
                <a:prstGeom prst="line">
                    <a:avLst/>
                </a:prstGeom>
                 <!-- ★★★ 直線要素には、コネクタのような明示的な接続情報 (<a:stCxn>, <a:endCxn>) は無い ★★★ -->
                 <!-- ★★★ コードは、この直線の from/to 座標と他のNodeの矩形を比較し、近ければ接続を補完する (ENABLE_NEARBY_COMPLETION=trueの場合) ★★★ -->
            </xdr:spPr>
             <xdr:txBody/> <!-- Label: "", hasText=false -->
        </xdr:sp>
        <xdr:clientData/>
    </xdr:twoCellAnchor>

</xdr:wsDr>
```

---

**コードによる解析の流れ (上記サンプルに対応):**

1.  **ファイル読み込み:** `extractStructuredShapesFromExcel` が `drawing1.xml` を読み込み、`extractStructure` に渡します。プレフィックスは `drawing1` となります。
2.  **1stパス:**
    *   `xdr:sp` (ID: 1, 2, 11, 12) を発見 → `processShapeBase` で Node オブジェクト生成 (ID: `drawing1_1`, `drawing1_2`, `drawing1_11`, `drawing1_12`)。`type`, `label`, `hasText` を抽出。
    *   `xdr:sp` (ID: 4) を発見 → `prst="line"` なので `processLineShape` で Edge オブジェクト生成 (ID: `drawing1_4`, Type: `line`, Source/Target: `null`, Label: `""`, hasText: `false`)。
    *   `xdr:cxnSp` (ID: 3, 13) を発見 → `processConnectorBase` で Edge オブジェクト生成 (ID: `drawing1_3`, `drawing1_13`)。`type`, `label`, `hasText` を抽出し、`stCxn`/`endCxn` から `source`/`target` を設定 (例: Edge 3 の Source=`drawing1_1`, Target=`drawing1_2`)。
    *   `xdr:grpSp` (ID: 10) を発見 → グループ要素として一時リストに保持。
    *   `elementMap` に全要素、`connectedNodeIds` に Node 1, 2, 11, 12 が追加される。
3.  **座標取得:** 各 Node (ID 1, 2, 11, 12) について、親アンカーの `from`/`to` 要素から `getAnchorEndpoint` で `fromRow/Col`, `toRow/Col` を計算し、Node オブジェクトと `nodeRectMap` に格納。
4.  **近傍補完 (`ENABLE_NEARBY_COMPLETION = true` の場合):** Edge 4 (直線) の `from`/`to` 座標を `getAnchorEndpoint` で計算。Node 1, 2, 11, 12 の矩形 (`nodeRectMap`) との距離を比較。もし `NEAR_SHAPE_THRESHOLD` 内に Node 1 やグループ (ID 10) の矩形があれば、Edge 4 の `source` や `target` が設定される可能性があります。
5.  **グループ処理 (`USE_LEGACY_GROUPING_LOGIC` による分岐):**
    *   **旧 (`true`):** `type="group"`, `id="drawing1_10"` のNode生成。Node 11, 12, Edge 13 に `GroupID="drawing1_10"` 付与。
    *   **新 (`false`):** `type="groupAggregation"`, `id="drawing1_10"` のNode生成。Label は `"要素 A\n要素 B"` (区切り文字が `\n` の場合)。Node 11, 12 を削除 (`nodes` 配列から)。Edge 13 の Source を `drawing1_11` -> `drawing1_10`、Target を `drawing1_12` -> `drawing1_10` に変更。Node 10 の座標を Node 11, 12 の範囲で計算。
6.  **接続先付け替え (`ENABLE_CONNECTOR_RETARGETING = true` の場合):** もし Node 1 などがテキストを持たない (`hasText=false`) 場合、それに接続している Edge 3 などは、Node 1 を包含する別の Node (もしあれば) に接続先が付け替えられる可能性があります（このサンプルでは該当なし）。
7.  **幽霊端点除去:** Edge 3, 4, 13 について、`source` または `target` が `null` (補完・付け替え後も未接続) または存在しないNode IDを指しているか確認。
8.  **最終フィルタリング:**
    *   **Nodes:** テキストを持たず、どこからも接続されていないNodeがあれば削除（このサンプルでは該当なし）。旧ロジックの場合、`type: "group"` の Node 10 は残る。新ロジックの場合、`type: "groupAggregation"` の Node 10 は残る。
    *   **Edges:** 最終的に残った Edge (例: Edge 3、補完されていれば Edge 4) がリストに含まれる。
9.  **テキスト整形:** `formatToTextBySheet` が最終的なNodes/Edgesリストを指定形式のテキストに整形して出力。`NODE_POSITION_OUTPUT_MODE`, `OUTPUT_NODE_TYPE` などのフラグに従う。

---