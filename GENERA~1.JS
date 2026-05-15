// GATEN職 会員機能 通常サイト統合 提案書
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, Header, Footer, PageOrientation,
  TabStopType, TabStopPosition,
} = require('docx');

// ============== 共通スタイル ==============
const JP_FONT = "Yu Gothic"; // 游ゴシック（Windows標準・Mac標準で日本語対応）

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

// ヘルパー: 通常段落
const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 320 },
  ...opts,
  children: [new TextRun({ text, font: JP_FONT, size: 22 })],
});

// ヘルパー: 太字段落
const PB = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 320 },
  ...opts,
  children: [new TextRun({ text, font: JP_FONT, size: 22, bold: true })],
});

// ヘルパー: 複数Run段落
const PRun = (runs, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 320 },
  ...opts,
  children: runs.map(r => new TextRun({
    font: JP_FONT, size: 22, ...r,
  })),
});

// ヘルパー: 箇条書き
const Bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: "bullets", level },
  spacing: { after: 80, line: 300 },
  children: [new TextRun({ text, font: JP_FONT, size: 22 })],
});

// ヘルパー: 番号付き
const Num = (text) => new Paragraph({
  numbering: { reference: "numbers", level: 0 },
  spacing: { after: 80, line: 300 },
  children: [new TextRun({ text, font: JP_FONT, size: 22 })],
});

// ヘルパー: テーブルセル
const Cell = (text, opts = {}) => new TableCell({
  borders,
  width: { size: opts.width || 4680, type: WidthType.DXA },
  shading: opts.header ? { fill: "1F4E79", type: ShadingType.CLEAR } : undefined,
  margins: { top: 100, bottom: 100, left: 140, right: 140 },
  children: [new Paragraph({
    spacing: { after: 0, line: 280 },
    children: [new TextRun({
      text,
      font: JP_FONT,
      size: 20,
      bold: !!opts.header,
      color: opts.header ? "FFFFFF" : "000000",
    })],
  })],
});

// 空行
const BR = () => new Paragraph({ children: [new TextRun({ text: "", font: JP_FONT })] });

// 区切り線
const HR = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
  spacing: { before: 60, after: 200 },
  children: [],
});

// ============== ドキュメント構築 ==============
const children = [];

// ----- 表紙 -----
children.push(new Paragraph({
  spacing: { before: 1800, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: "GATEN職 会員機能 統合改善 提案書",
    font: JP_FONT, size: 44, bold: true, color: "1F4E79",
  })],
}));
children.push(new Paragraph({
  spacing: { after: 1200 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: "会員専用機能の通常サイト側への展開について",
    font: JP_FONT, size: 26, color: "595959",
  })],
}));

children.push(new Paragraph({
  spacing: { before: 2400, after: 120 },
  alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text: "作成日：2026年5月15日", font: JP_FONT, size: 22 })],
}));
children.push(new Paragraph({
  spacing: { after: 120 },
  alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text: "対象サイト：https://gaten.info/", font: JP_FONT, size: 22 })],
}));
children.push(new Paragraph({
  spacing: { after: 120 },
  alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text: "Version 1.0", font: JP_FONT, size: 22 })],
}));

children.push(new Paragraph({
  pageBreakBefore: true,
  children: [new TextRun({ text: "", font: JP_FONT })],
}));

// ----- 1. はじめに / エグゼクティブサマリー -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "1. エグゼクティブサマリー", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

children.push(P(
  "本提案書は、当社運営の求人ポータルサイト「GATEN職（https://gaten.info/）」における、会員機能のユーザー体験改善に関する方針を整理したものである。"
));
children.push(P(
  "現状、会員登録ユーザーは管理画面（/member 配下）で各種会員機能を利用できるが、企業検索を行うと検索結果・企業詳細ページが管理画面外の通常サイトに遷移し、会員専用機能が利用できなくなる課題がある。"
));
children.push(P(
  "本提案では、ログイン状態を通常サイト側でも保持・活用し、サイト全域で会員専用機能を利用可能にする方針を示す。優先機能は「お気に入り・キープ機能」および「閲覧履歴・おすすめ表示」とする。"
));

children.push(BR());
children.push(PB("【期待される効果】"));
children.push(Bullet("ユーザー離脱率の低下（管理画面と通常サイトを行き来する負担の解消）"));
children.push(Bullet("お気に入り保存・閲覧履歴の蓄積による再訪率向上"));
children.push(Bullet("会員価値の体感機会増加 → 会員登録CVRの向上"));
children.push(Bullet("おすすめ表示によるマッチング精度向上 → 応募CVRの向上"));

// ----- 2. 現状の課題 -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "2. 現状の課題", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

children.push(new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 160 },
  children: [new TextRun({ text: "2.1 ユーザー体験上の課題", font: JP_FONT, size: 28, bold: true, color: "1F4E79" })],
}));
children.push(P("会員ユーザーが企業検索を行う際の典型的な動線で、以下の問題が発生している。"));
children.push(BR());
children.push(Num("会員ログイン後、管理画面（/member 配下）から企業検索を開始する。"));
children.push(Num("検索結果ページ・企業詳細ページが、管理画面外（通常サイト側のページ）に遷移する。"));
children.push(Num("通常サイト側ではログイン状態のUIや会員専用機能（お気に入り保存、閲覧履歴、応募ボタン等）が表示されないため、ユーザーは「機能が使えなくなった」と感じる。"));
children.push(Num("管理画面に戻るための導線も限定的で、操作が分断される。"));

children.push(BR());
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 160 },
  children: [new TextRun({ text: "2.2 ビジネス上の影響", font: JP_FONT, size: 28, bold: true, color: "1F4E79" })],
}));
children.push(Bullet("会員機能の利用機会が限定的なため、会員価値が実感されにくい"));
children.push(Bullet("検索動線途中での離脱が発生し、応募・問い合わせCVRが伸び悩む"));
children.push(Bullet("閲覧履歴やお気に入りといった行動データが管理画面内に閉じており、レコメンデーション等の発展的施策に活用できない"));
children.push(Bullet("非会員ユーザーから見ても「会員登録のメリット」が伝わりにくい構造"));

children.push(BR());
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 160 },
  children: [new TextRun({ text: "2.3 構造的な原因（推定）", font: JP_FONT, size: 28, bold: true, color: "1F4E79" })],
}));
children.push(P("同一ドメイン・同一システム上で運用されているにも関わらず、UI実装上は次のように分離していると推定される。"));
children.push(BR());
children.push(Bullet("管理画面（/member 配下）：会員ログイン状態を前提とした専用テンプレートで描画"));
children.push(Bullet("通常サイト（企業検索結果・企業詳細等）：非会員向けの公開ページとして実装されており、ログイン状態を参照していない／参照していてもUIに反映していない"));
children.push(Bullet("結果として、認証セッション自体は通信されているが、UI層で「会員向け表示」が分岐しておらず、機能が表出していない可能性が高い"));

// ----- 3. 目指す姿 -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "3. 目指す姿（To-Be）", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

children.push(P(
  "通常サイト全域において、ログイン中の会員ユーザーには会員向けUIと機能を表示し、非ログインユーザーには従来通りの公開UIを表示する「ログイン状態を意識した統合UI」を実現する。"
));
children.push(BR());

children.push(PB("【主要なユーザー体験の変化】"));

const beforeAfterTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2340, 3510, 3510],
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        Cell("シーン", { header: true, width: 2340 }),
        Cell("現状（As-Is）", { header: true, width: 3510 }),
        Cell("改善後（To-Be）", { header: true, width: 3510 }),
      ],
    }),
    new TableRow({ children: [
      Cell("企業検索", { width: 2340 }),
      Cell("管理画面で検索 → 結果は通常サイトに遷移し機能消失", { width: 3510 }),
      Cell("どこから検索しても会員機能を維持したまま検索可能", { width: 3510 }),
    ]}),
    new TableRow({ children: [
      Cell("企業詳細閲覧", { width: 2340 }),
      Cell("お気に入り保存ボタンが表示されない", { width: 3510 }),
      Cell("ハート/星アイコンでワンクリック保存可能", { width: 3510 }),
    ]}),
    new TableRow({ children: [
      Cell("再訪時", { width: 2340 }),
      Cell("自分の保存企業・履歴は管理画面内でしか確認できない", { width: 3510 }),
      Cell("トップページ・検索ページで「最近見た企業」「あなたへのおすすめ」が表示", { width: 3510 }),
    ]}),
    new TableRow({ children: [
      Cell("ヘッダー", { width: 2340 }),
      Cell("ログインしていてもユーザー名等の表示が限定的", { width: 3510 }),
      Cell("全ページ共通でログイン状態・マイページ導線を表示", { width: 3510 }),
    ]}),
  ],
});
children.push(beforeAfterTable);

// ----- 4. 解決方針 -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "4. 解決方針", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

children.push(new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 160 },
  children: [new TextRun({ text: "4.1 基本方針", font: JP_FONT, size: 28, bold: true, color: "1F4E79" })],
}));
children.push(P(
  "同一ドメイン・同一システムであるという前提を活かし、認証セッションは現状を維持したまま、通常サイト側のテンプレート／コンポーネントを「ログイン状態を判定して表示を分岐させる」形に拡張する。"
));
children.push(P("管理画面（/member）と通常サイトのUIを統合し、サイト全域を「会員ファースト」設計に転換する。"));

children.push(BR());
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 160 },
  children: [new TextRun({ text: "4.2 実装の3層構造", font: JP_FONT, size: 28, bold: true, color: "1F4E79" })],
}));

const layerTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1800, 2500, 5060],
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        Cell("層", { header: true, width: 1800 }),
        Cell("対応内容", { header: true, width: 2500 }),
        Cell("具体的な実装イメージ", { header: true, width: 5060 }),
      ],
    }),
    new TableRow({ children: [
      Cell("① 認証層", { width: 1800 }),
      Cell("セッション維持の確認", { width: 2500 }),
      Cell("同一ドメイン下では既にCookieベースのセッションが共有されているはず。通常サイト側のテンプレートからも会員IDを取得できる状態か確認・必要に応じ調整。", { width: 5060 }),
    ]}),
    new TableRow({ children: [
      Cell("② UI層", { width: 1800 }),
      Cell("会員向け表示分岐の実装", { width: 2500 }),
      Cell("ヘッダー、検索結果カード、企業詳細ページに「ログイン中なら表示する会員向けUI」を追加。テンプレート分岐 or 共通コンポーネント化。", { width: 5060 }),
    ]}),
    new TableRow({ children: [
      Cell("③ 機能層", { width: 1800 }),
      Cell("会員機能APIの整備", { width: 2500 }),
      Cell("お気に入り保存/解除、閲覧履歴記録、おすすめ取得などのエンドポイントを通常サイトから呼び出し可能にする（Ajax/Fetch）。", { width: 5060 }),
    ]}),
  ],
});
children.push(layerTable);

children.push(BR());
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 160 },
  children: [new TextRun({ text: "4.3 優先実装機能", font: JP_FONT, size: 28, bold: true, color: "1F4E79" })],
}));

children.push(PB("【優先機能①】お気に入り・キープ機能"));
children.push(Bullet("企業一覧カード／企業詳細ページに「キープする」ボタン（ハート or 星アイコン）を追加"));
children.push(Bullet("非ログイン時はクリック時にログイン誘導モーダルを表示（CVR向上の副次効果）"));
children.push(Bullet("ログイン時はAjaxでお気に入りテーブルに保存／解除"));
children.push(Bullet("管理画面の既存「キープリスト」と同一データベースを参照（既存機能との連続性を保つ）"));

children.push(BR());
children.push(PB("【優先機能②】閲覧履歴・おすすめ表示"));
children.push(Bullet("ログインユーザーが企業詳細ページを閲覧した際、閲覧履歴テーブルへ記録"));
children.push(Bullet("トップページ・検索結果ページに「最近見た企業」セクションを表示"));
children.push(Bullet("閲覧履歴のタグ・職種・エリアに基づく「あなたにおすすめの企業」セクションを表示（初期はシンプルな類似ロジックで十分）"));
children.push(Bullet("おすすめロジックは将来的に協調フィルタリング等への発展余地を残す設計"));

// ----- 5. 段階的実装プラン -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "5. 段階的実装プラン", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

const phaseTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1400, 2400, 3400, 2160],
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        Cell("フェーズ", { header: true, width: 1400 }),
        Cell("内容", { header: true, width: 2400 }),
        Cell("主な作業", { header: true, width: 3400 }),
        Cell("目安期間", { header: true, width: 2160 }),
      ],
    }),
    new TableRow({ children: [
      Cell("Phase 0", { width: 1400 }),
      Cell("現状把握・要件定義", { width: 2400 }),
      Cell("認証セッションの仕様確認／既存テンプレート構造の調査／要件・受け入れ基準の確定", { width: 3400 }),
      Cell("約1〜2週間", { width: 2160 }),
    ]}),
    new TableRow({ children: [
      Cell("Phase 1", { width: 1400 }),
      Cell("ヘッダー統合", { width: 2400 }),
      Cell("全ページ共通ヘッダーにログイン状態表示／マイページ導線追加。最小コストで効果体感", { width: 3400 }),
      Cell("約1〜2週間", { width: 2160 }),
    ]}),
    new TableRow({ children: [
      Cell("Phase 2", { width: 1400 }),
      Cell("お気に入り機能", { width: 2400 }),
      Cell("企業一覧／詳細にキープボタン追加。Ajaxエンドポイント整備。既存DBと連携", { width: 3400 }),
      Cell("約2〜3週間", { width: 2160 }),
    ]}),
    new TableRow({ children: [
      Cell("Phase 3", { width: 1400 }),
      Cell("閲覧履歴", { width: 2400 }),
      Cell("閲覧記録の自動保存／トップ・検索ページに「最近見た企業」表示", { width: 3400 }),
      Cell("約2週間", { width: 2160 }),
    ]}),
    new TableRow({ children: [
      Cell("Phase 4", { width: 1400 }),
      Cell("おすすめ表示", { width: 2400 }),
      Cell("閲覧履歴ベースの類似企業レコメンドの実装と表示", { width: 3400 }),
      Cell("約2〜3週間", { width: 2160 }),
    ]}),
    new TableRow({ children: [
      Cell("Phase 5", { width: 1400 }),
      Cell("計測・改善", { width: 2400 }),
      Cell("各機能のクリック率・CVR計測。A/Bテストで効果検証", { width: 3400 }),
      Cell("継続実施", { width: 2160 }),
    ]}),
  ],
});
children.push(phaseTable);

children.push(BR());
children.push(P("※ 期間はあくまで目安。実装会社のリソース、既存コードベースの状態により変動する。"));
children.push(P("※ Phase 1〜2 を先行リリースすることで、約1〜1.5ヶ月で主要なユーザー体験課題は解消される見込み。"));

// ----- 6. リスク・留意点 -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "6. リスク・留意点", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

const riskTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2800, 3280, 3280],
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        Cell("リスク", { header: true, width: 2800 }),
        Cell("内容", { header: true, width: 3280 }),
        Cell("対策", { header: true, width: 3280 }),
      ],
    }),
    new TableRow({ children: [
      Cell("セキュリティ（CSRF等）", { width: 2800 }),
      Cell("お気に入り追加・閲覧履歴記録などのPOST処理が公開ページから呼び出されるためCSRF対策が必要", { width: 3280 }),
      Cell("既存の認証トークン／CSRFトークンを通常サイト側のフォームにも埋め込み、検証フローを統一", { width: 3280 }),
    ]}),
    new TableRow({ children: [
      Cell("キャッシュ", { width: 2800 }),
      Cell("通常サイトは公開ページとしてCDN／フルキャッシュされている場合、ログイン状態の表示が反映されない", { width: 3280 }),
      Cell("ログイン状態に依存する部分のみクライアントサイドJSで描画／ESI等の部分キャッシュを利用", { width: 3280 }),
    ]}),
    new TableRow({ children: [
      Cell("既存機能との二重実装", { width: 2800 }),
      Cell("管理画面側のお気に入り画面と、通常サイト側のキープボタンが別実装になると保守コストが増加", { width: 3280 }),
      Cell("API／データモデルは1つに統一。UIコンポーネント単位で共通化する設計を最初に決める", { width: 3280 }),
    ]}),
    new TableRow({ children: [
      Cell("非ログインユーザーへの影響", { width: 2800 }),
      Cell("既存の非会員ユーザー体験を損なわないこと", { width: 3280 }),
      Cell("ログイン状態フラグでUIを切り替える設計とし、非会員はこれまで通りの画面を維持。リリース前にデグレード検証を実施", { width: 3280 }),
    ]}),
    new TableRow({ children: [
      Cell("プライバシー", { width: 2800 }),
      Cell("閲覧履歴の保存・利用について利用規約／プライバシーポリシーの整合性確認が必要", { width: 3280 }),
      Cell("規約改定の要否を確認し、必要であれば事前に対応", { width: 3280 }),
    ]}),
  ],
});
children.push(riskTable);

// ----- 7. 開発会社への依頼・確認事項 -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "7. 開発会社への確認・依頼事項", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

children.push(P("外部開発会社へ実装を依頼する場合、以下の事項を確認・依頼することを推奨する。"));
children.push(BR());

children.push(PB("【ヒアリング事項（Phase 0）】"));
children.push(Bullet("会員認証セッションの仕様（Cookie名、有効期間、SameSite設定など）"));
children.push(Bullet("通常サイト側のページ（検索結果・企業詳細）のレンダリング方式（SSR／キャッシュの有無）"));
children.push(Bullet("既存のお気に入り・キープ機能のDBスキーマ、API仕様"));
children.push(Bullet("CSS／JSのアセット管理方式（共通化のしやすさ）"));
children.push(Bullet("リリース手順（ステージング環境の有無、リリースサイクル）"));

children.push(BR());
children.push(PB("【依頼事項】"));
children.push(Bullet("本書のPhase 0〜2を最優先で実施するスコープ・見積もり・スケジュールの提示"));
children.push(Bullet("実装にあたっての技術的な実現方式（テンプレート分岐／コンポーネント分割／API設計）の提案"));
children.push(Bullet("リリース後の効果計測ポイント（イベント計測の埋め込み）も合わせて設計に含めること"));

// ----- 8. まとめ -----
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 200, after: 200 },
  children: [new TextRun({ text: "8. まとめ", font: JP_FONT, size: 32, bold: true, color: "1F4E79" })],
}));
children.push(HR());

children.push(P(
  "本課題の本質は「同一システムでありながら、UI／機能のレイヤーで会員向けと公開向けが分断されていること」にある。認証基盤の作り変えのような大規模改修ではなく、UI層を中心とした拡張により、比較的低リスク・短期間で改善が可能と考えられる。"
));
children.push(P(
  "Phase 1（ヘッダー統合）・Phase 2（お気に入り機能）の先行リリースだけでも体感的なUX改善が大きく、優先機能であるキープ機能と閲覧履歴・おすすめ表示は段階的に展開していくことが現実的である。"
));
children.push(P(
  "本提案を基に社内検討および開発会社との詳細協議を進め、3〜4ヶ月程度での全フェーズ完了を目指したい。"
));

children.push(BR());
children.push(BR());
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text: "以上", font: JP_FONT, size: 22 })],
}));

// ============== Document設定 ==============
const doc = new Document({
  creator: "GATEN職",
  title: "GATEN職 会員機能 統合改善 提案書",
  description: "会員専用機能の通常サイト側への展開に関する提案",
  styles: {
    default: { document: { run: { font: JP_FONT, size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: JP_FONT, color: "1F4E79" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: JP_FONT, color: "1F4E79" },
        paragraph: { spacing: { before: 200, after: 160 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            text: "GATEN職 会員機能 統合改善 提案書",
            font: JP_FONT, size: 18, color: "808080",
          })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "- ", font: JP_FONT, size: 18, color: "808080" }),
            new TextRun({ children: [PageNumber.CURRENT], font: JP_FONT, size: 18, color: "808080" }),
            new TextRun({ text: " -", font: JP_FONT, size: 18, color: "808080" }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("gaten_proposal.docx", buffer);
  console.log("Created: gaten_proposal.docx");
});
