import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_excel_report():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    FONT_NAME = "Meiryo UI"
    
    font_title = Font(name=FONT_NAME, size=15, bold=True, color="FFFFFF")
    font_subtitle = Font(name=FONT_NAME, size=9.5, italic=True, color="CCCCCC")
    font_section = Font(name=FONT_NAME, size=11.5, bold=True, color="1F1F1F")
    font_tbl_header = Font(name=FONT_NAME, size=9.5, bold=True, color="FFFFFF")
    font_data = Font(name=FONT_NAME, size=9, bold=False, color="262626")
    font_data_bold = Font(name=FONT_NAME, size=9, bold=True, color="262626")
    font_badge_crit = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")

    # Fills (Monotone Palette: Charcoal / Slate / Ash / White)
    fill_title = PatternFill(start_color="1F1F1F", end_color="1F1F1F", fill_type="solid")      # Deep Charcoal
    fill_tbl_header = PatternFill(start_color="333333", end_color="333333", fill_type="solid") # Dark Charcoal
    fill_sub_header = PatternFill(start_color="4D4D4D", end_color="4D4D4D", fill_type="solid") # Slate
    fill_zebra_light = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")# White
    fill_zebra_alt = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")  # Ultra Light Gray
    fill_highlight = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")  # Soft Gray
    fill_callout = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")    # Accent Gray

    # Borders
    thin_border_side = Side(border_style="thin", color="CCCCCC")
    thick_border_side = Side(border_style="medium", color="333333")
    double_border_side = Side(border_style="double", color="333333")

    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thick_border_side, bottom=thick_border_side)
    border_total = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_border_side)

    # Alignments
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_center_top = Alignment(horizontal="center", vertical="top", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # =========================================================================
    # SHEET 1: エグゼクティブサマリ (Executive Summary)
    # =========================================================================
    ws1 = wb.create_sheet(title="エグゼクティブサマリ")
    ws1.views.sheetView[0].showGridLines = True

    # Title block
    ws1.merge_cells("A1:G2")
    ws1["A1"] = "【インフラ性能診断報告書】契約管理アプリ テーブルダンプ時におけるCPU高負荷事象の分析と対策提言"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_title
    ws1["A1"].alignment = align_header

    ws1["A3"] = "発行日: 2026年10月 | 診断対象: IT環境バッチ (c7a.2xlarge: 8 vCPU / 16 GiB / Up to 12.5Gbps) / 本番環境 (c7a.8xlarge: 32 vCPU / 64 GiB) | 分類: インフラ性能評価・アーキテクチャ設計"
    ws1["A3"].font = font_subtitle
    ws1["A3"].fill = fill_sub_header
    ws1.merge_cells("A3:G3")
    ws1["A3"].alignment = align_center

    # Section 1: 結論サマリ (Verdict)
    ws1["A5"] = "■ 1. 総合評価と結論（エグゼクティブサマリ）"
    ws1["A5"].font = font_section

    summary_cards = [
        ("設問 1: CPU > 100%の発生有無", "【極めて一般的な正常事象】\nLinuxのtopコマンド（Irixモード）は1コア=100%として積算するため、マルチスレッド動作するJava JVMプロセスが複数コアを消費すると100%（実測125%等）を超えるのは完全な仕様通りの正常動作です。またバッチ処理がCPUを上限まで使い切るのは、ハードウェア資産を最大活用する健全な動作です。"),
        ("設問 2: 現在のEC2監視設定推測", "【OLTP用しきい値が流用されている可能性大】\n一般的なクラウド標準（Web/AP向け）の Warning: 70~80%、Critical: 85~90%（評価期間5分）が適用されていると推測されます。バッチ処理と常時サービスを同一基準で監視しているため、正常なバッチ高負荷でクリティカルが誤発報しています。"),
        ("設問 3: サイジングと監視域値の妥当性", "【サイジング: 妥当 (c7a.2xlarge) / 監視域値: 不適切】\n最新世代のc7a.2xlarge（AMD Genoa, Zen 4, 8vCPU/16GB, Up to 12.5Gbps）は極めて高性能であり、メモリも16GB中5.5GB使用・空き9.8GBと潤沢です。CPU一時飽和は「多重度設定」に起因しており、スペック増強ではなくJP1マルチ数と監視域値の見直しが必須です。"),
        ("設問 4: JVM数 10 vs 8 (IT環境)", "【「8」（またはベスト推奨「7」）への変更を強く推奨】\n8vCPUに対して10JVMを動かすと、CPU要求量（約12コア分）が供給量（8コア）を上回り、コンテキストスイッチとCPUキャッシュ破棄が激増してLoad Avgが17.59まで悪化します。並列度を7〜8に絞った方が、実は全体のバッチ完了時間は短縮されます。"),
        ("設問 5: 本番 58JVM / 32vCPU / 32マルチ", "【32マルチは上限ギリギリ、安全推奨値は「24〜28マルチ」】\n58ジョブを32並列で流す運用は、各JVM内部スレッド（GC/JIT）の重複によりCPUが100%に張り付くほか、最も危険なのは「RDS/Auroraに対する32本同時の大規模SELECTによるDB側の枯渇・遅延」です。安全推奨値は24〜28並列から開始することです。"),
        ("設問 6: クリティカル検知への対策方針", "【想定範囲の一時的高負荷だが、運用放置は厳禁】\n計算機科学・バッチ特性としては「想定内」ですが、アラート放置は「狼少年化（真の障害の見落とし）」やOS無応答を招くため対策が不可欠です。「①JP1並列度適正化」「②JVM GCスレッド抑制」「③バッチ時間帯の監視しきい値緩和」の3段構えで恒久対策を行います。")
    ]

    row = 6
    for title, desc in summary_cards:
        ws1.merge_cells(start_row=row, start_column=1, end_row=row+1, end_column=2)
        ws1.merge_cells(start_row=row, start_column=3, end_row=row+1, end_column=7)
        c_title = ws1.cell(row=row, column=1, value=title)
        c_desc = ws1.cell(row=row, column=3, value=desc)
        
        c_title.font = font_data_bold
        c_title.fill = fill_highlight
        c_title.alignment = align_header
        
        c_desc.font = font_data
        c_desc.fill = fill_zebra_light
        c_desc.alignment = align_left_top
        
        for r in range(row, row+2):
            for c in range(1, 8):
                ws1.cell(row=r, column=c).border = border_cell
        row += 2

    # Section 2: 6つの問いに対する明確なインフラ回答マトリクス
    row += 1
    ws1.cell(row=row, column=1, value="■ 2. 課題・論点別の詳細回答マトリクス").font = font_section
    row += 1

    headers_m = ["No", "論点 / 質問項目", "現状・診断結果", "リスク・影響度", "推奨アクション・設定値", "優先度"]
    for col_idx, h in enumerate(headers_m, start=1):
        c = ws1.cell(row=row, column=col_idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    matrix_data = [
        (1, "CPU 100%超えは一般的に発生しうるか", "Linux topの仕様（Irixモード）により、マルチスレッドJVMが1プロセスで100%（1コア分）を超えるのは極めて一般的。バッチ全体でもリソース最大利用は設計通りの健全な動作。", "低（正常仕様）", "動作原理を関係者へ周知。Solarisモード表示またはCloudWatch（100%満点スケール）での評価基準を統一する。", "中"),
        (2, "現在のEC2監視設定と妥当性", "推測：CloudWatch CPUUtilization Warning: 80% / Critical: 90%（5分間）。短時間のバッチ高負荷で即報知される過敏な設定。", "中（アラート形骸化）", "バッチ実行時間帯（夜間等）のアラート抑止（EventBridge等）または評価期間を15~30分へ延長。", "高"),
        (3, "インスタンスサイジングの妥当性", "IT環境はc7a.2xlarge（8vCPU / 16GB / Up to 12.5Gbps）。メモリ使用5.5GB、空き9.8GBと潤沢。CPUのみ一時的過密状態。", "低（サイジング自体は妥当）", "現行スペックのままで十分継続運用可能。スペック増強（課金増）ではなく並列度（JP1マルチ数）の制御で解決。", "高"),
        (4, "JVM数 10 vs 8 (8vCPU環境)", "10JVM同時実行によりLoad Average 17.59を記録。CPU要求量（約12コア分）が供給量（8コア）を上回りコンテキストスイッチが多発。", "中（処理遅延・スループット頭打ち）", "JP1マルチ数を「8」またはベスト推奨「7」に変更。1コアをOS/管理用に空けることで全体のスループットが最大化。", "最高"),
        (5, "本番 58JVM / 32vCPU / 32マルチの可否", "32並列実行はコア数と等倍だが、GCスレッド重複やDB接続集中（32コネクション同時フルスキャン）でDB枯渇リスクあり。", "高（DB遅延・バッチ遅延）", "JP1マルチ数を「24〜28」で本番検証開始。JVM起動時オプションに -XX:ParallelGCThreads=2 などを設定。", "高"),
        (6, "CPUクリティカル検知への対策方針", "一時的高負荷としては「想定範囲内」だが、監視アラートが発報し続ける運用はSRE/運用保守上NG。", "中（運用破綻リスク）", "「並列数適正化」＋「GC抑制」＋「バッチ時間帯の監視域値変更」の恒久対策3点をセットで実施。", "最高"),
    ]

    for item in matrix_data:
        fill_cur = fill_zebra_alt if item[0] % 2 == 0 else fill_zebra_light
        ws1.cell(row=row, column=1, value=item[0]).alignment = align_center
        ws1.cell(row=row, column=2, value=item[1]).alignment = align_left
        ws1.cell(row=row, column=3, value=item[2]).alignment = align_left
        ws1.cell(row=row, column=4, value=item[3]).alignment = align_center
        ws1.cell(row=row, column=5, value=item[4]).alignment = align_left
        ws1.cell(row=row, column=6, value=item[5]).alignment = align_center

        for col_idx in range(1, 7):
            c = ws1.cell(row=row, column=col_idx)
            c.font = font_data
            c.fill = fill_cur
            c.border = border_cell
        row += 1

    # Set column widths
    col_widths_ws1 = [6, 28, 36, 20, 36, 10, 10]
    for i, w in enumerate(col_widths_ws1, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 2: メトリクス詳細分析 (Telemetry Deep Dive)
    # =========================================================================
    ws2 = wb.create_sheet(title="メトリクス詳細分析")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:G2")
    ws2["A1"] = "【実測メトリクス完全解剖】IT環境 (c7a.2xlarge) テーブルダンプ実行時の負荷状況とボトルネック特定"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_title
    ws2["A1"].alignment = align_header

    ws2["A4"] = "■ 1. 取得メトリクスの項目別詳細解析（ハードウェア・OS・ランタイム）"
    ws2["A4"].font = font_section

    headers_m2 = ["指標カテゴリ", "実測値", "基準値 / 許容上限", "ステータス", "インフラ技術者の診断解説と動作実態"]
    row = 5
    for idx, h in enumerate(headers_m2, 1):
        c = ws2.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    metrics_analysis = [
        ("Load Average (1m)", "17.59", "8.0 (vCPU数と同等)", "過負荷 (220%)", "8 vCPUに対して実行待ち・実行中タスクが平均約17.6個存在。急激な上昇（15分前 1.99 → 5分前 5.03 → 1分前 17.59）であり、JP1から一斉に10ジョブが起動された瞬間の急峻なキュー滞留を示している。"),
        ("Tasks 内訳", "843 total (1 run, 841 sleep)", "通常 200〜400 程度", "スレッド多数生成", "843タスク中841がsleep（I/Oや同期待ち）、1がrunning。Java JVM 10プロセスがそれぞれ数十本のスレッド（GC、ワーカー、DB接続等）をOSネイティブに生成し、大量のLWP（Lightweight Process）が存在。"),
        ("%Cpu(s) us (User)", "95.9%", "70.0%以下（通常運用）", "高負荷（正常動作）", "CPU時間の96%がJavaアプリケーションのユーザ空間処理（SQLフェッチ、データ型変換、文字列整形、CSV/JSONフォーマット、メモリ展開、GC処理）に正しく使われている。"),
        ("%Cpu(s) sy (System)", "3.1%", "15.0%以下", "極めて健全", "OSカーネル空間の消費は3.1%と極小。コンテキストスイッチやシステムコール（write/epoll等）によるOS層の過剰なオーバーヘッドはまだ致命的レベルに達していない。"),
        ("%Cpu(s) wa (I/O Wait)", "0.8%", "5.0%以下", "極めて健全", "ディスクやEBSへの書き込み待ちによるCPUブロックはほぼゼロ。ダンプデータはLinuxのページキャッシュ（buff/cache: 8.9GB）に非同期に書き込まれているため、I/Oボトルネックではない。"),
        ("%Cpu(s) hi / si (割り込み)", "0.1% / 0.0%", "1.0%以下", "極めて健全", "ハードウェア割り込み（hi）およびソフトウェア割り込み（si）は皆無。c7aの高性能SR-IOVネットワーク（Up to 12.5Gbps）とNitroシステムが安定処理している。"),
        ("%Cpu(s) st (Steal Time)", "0.0%", "0.0% (完全専用)", "完璧（健全）", "AWS NitroハイパーバイザーによるCPUスチール（他VMへの割り当て奪取やCPUバースト枯渇）は一切なし。c7aインスタンスへ割り当てられた物理CPU能力が100%専有提供されている。"),
        ("Memory Used", "5,557.3 MiB", "15,358 MiB (全体)", "余裕あり (36%)", "実使用メモリは約5.4GB。10JVM起動時でも1JVMあたり約500MB程度の消費に収まっており、メモリ逼迫やOOM（Out of Memory）の懸念は皆無。"),
        ("Memory Buff/Cache", "8,939.6 MiB", "動的再利用可能", "極めて効果的", "メモリの約58%（約8.7GB）がOSページキャッシュとして機能。テーブルダンプで生成されたファイルデータが高速にRAMキャッシュされており、ディスクI/O遅延を防いでいる。"),
        ("Swap Used", "0.2 MiB", "3,814 MiB中 0.2MiB", "完全未使用", "スワップ発生率0.005%。スラッシング（メモリ不足によるディスクスワップ頻発）は全く起きておらず、パフォーマンス低下要因から除外できる。"),
        ("プロセス別 %CPU", "125.5%, 122.4%, 115.1%...", "単一コア=100%", "マルチスレッド動作", "各JVMプロセスが1.0〜1.3コア分のCPU能力を同時消費。JVM内部でバッチワーカースレッド、G1GC並列スレッド、JITコンパイラがマルチコアで並行稼働している証拠。")
    ]

    for item in metrics_analysis:
        ws2.cell(row=row, column=1, value=item[0]).alignment = align_center
        ws2.cell(row=row, column=2, value=item[1]).alignment = align_center
        ws2.cell(row=row, column=3, value=item[2]).alignment = align_center
        ws2.cell(row=row, column=4, value=item[3]).alignment = align_center
        ws2.cell(row=row, column=5, value=item[4]).alignment = align_left_top

        for c_idx in range(1, 6):
            cell = ws2.cell(row=row, column=c_idx)
            cell.font = font_data
            cell.border = border_cell
            if "過負荷" in item[3]:
                cell.fill = fill_highlight
            else:
                cell.fill = fill_zebra_light
        row += 1

    # Section 2: なぜ100%を超えるのか？（topの動作モード対比）
    row += 2
    ws2.cell(row=row, column=1, value="■ 2. topコマンドにおける「%CPU 100%超え」の技術的背景（Irix vs Solarisモード）").font = font_section
    row += 1

    headers_mode = ["表示モード", "キー操作", "100%の定義", "最大値（c7a.2xlarge: 8 vCPU）", "用途・特徴", "今回の観測結果の解釈"]
    for idx, h in enumerate(headers_mode, 1):
        c = ws2.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    mode_data = [
        ("Irix モード (デフォルト)", "Shift + I で切替", "「1つのCPUコアの100%」", "800% (8コア × 100%)", "Linux topの標準。プロセス単位の消費量をコア数積算で表示。マルチスレッドアプリの並列度把握に最適。", "125.5%は「1つのJVMが約1.25コア分の計算資源を同時に消費した」という意味。マルチスレッド動作として完全正常。"),
        ("Solaris モード", "Shift + I で切替", "「全CPU合計に対する割合」", "100% (全コア合算を100%と正規化)", "システム全体の総処理能力に対するプロセス占有率を表示。CloudWatchのCPUUtilizationと同じスケール。", "125.5%をSolarisモードで換算すると、125.5 ÷ 8 = 約 15.7% のマシン全体占有率となる。")
    ]

    for item in mode_data:
        for c_idx in range(1, 7):
            cell = ws2.cell(row=row, column=c_idx, value=item[c_idx-1])
            cell.font = font_data
            cell.border = border_cell
            cell.fill = fill_zebra_alt if item[0].startswith("Irix") else fill_zebra_light
            cell.alignment = align_left if c_idx in [1, 5, 6] else align_center
        row += 1

    # Section 3: Tasks 843（running 1, sleeping 841）の深掘り
    row += 2
    ws2.cell(row=row, column=1, value="■ 3. Tasks 843 (1 running, 841 sleeping) のOS内部スケジューリング構造").font = font_section
    row += 1

    headers_task = ["タスク状態", "実測タスク数", "OS内部の動作原理", "バッチ処理における意味と評価"]
    for idx, h in enumerate(headers_task, 1):
        c = ws2.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    task_data = [
        ("running (実行中)", "1", "CPUコア上でまさにその瞬間に命令を実行しているタスク。コア数（8個）のうち、topコマンド自身またはJVMスレッドがサンプリング瞬間に補足されたもの。", "topコマンド実行瞬間のスナップショット値。実際には8個のコアがミリ秒単位で交代しながら常時8タスクを実行中。"),
        ("sleeping (休止・待機)", "841", "TASK_INTERRUPTIBLE / TASK_UNINTERRUPTIBLE。DBからのソケット受信待ち、ミューテックス待機、GC待機、スリープ待機のスレッド群。", "10個のJVM内部スレッド（1JVMあたり30〜50スレッド）とOSデーモン。大部分はイベント駆動待機しており、メモリのみ消費しCPUは消費しない。"),
        ("stopped / zombie", "1 stopped / 0 zombie", "停止シグナル（SIGSTOP）を受けたプロセスが1件、ゾンビプロセスはゼロ。", "ゾンビプロセス（孤児プロセス）ゼロであり、プロセスのクリーンアップは正常に行われている。")
    ]

    for item in task_data:
        ws2.cell(row=row, column=1, value=item[0]).alignment = align_center
        ws2.cell(row=row, column=2, value=item[1]).alignment = align_center
        ws2.cell(row=row, column=3, value=item[2]).alignment = align_left_top
        ws2.cell(row=row, column=4, value=item[3]).alignment = align_left_top

        for c_idx in range(1, 5):
            cell = ws2.cell(row=row, column=c_idx)
            cell.font = font_data
            cell.border = border_cell
            cell.fill = fill_zebra_light
        row += 1

    # Set column widths
    col_widths_ws2 = [18, 14, 20, 16, 50, 10, 10]
    for i, w in enumerate(col_widths_ws2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 3: 動作原理と歴史 (Principles, History & Analogies)
    # =========================================================================
    ws3 = wb.create_sheet(title="動作原理と歴史・小学生向け解説")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:G2")
    ws3["A1"] = "【深掘り解説】テーブルダンプ処理の動作原理・技術の進化史・小学生向け直感モデル"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_title
    ws3["A1"].alignment = align_header

    ws3["A4"] = "■ 1. 小学生にも1発でわかる！「最新超火力8口コンロ厨房と10人の二刀流料理人」の例え話"
    ws3["A4"].font = font_section

    analogies = [
        ("IT用語 / 実態", "厨房のたとえ（直感イメージ）", "何が起きているのか？（わかりやすい解説）"),
        ("最新サーバ (c7a.2xlarge)", "ピカピカの最高級レストラン厨房", "最新型の超高速調理器具（AMD Zen 4）を備えた超強力な仕事場です。"),
        ("8 vCPU (仮想コア)", "8口の最新IHコンロ（調理台）", "同時にフライパンを火にかけられる場所は、世界中どこを探しても「8個」しかありません。"),
        ("Up to 12.5Gbps ネットワーク", "超特急の食材配達レーン", "倉庫（DB）から契約データの食材が超ハイスピードで途切れず届くため、料理人は食材待ちになりません。"),
        ("10 JVM (Javaバッチ)", "10人のプロ料理人", "JP1という店長から「今すぐ全員で10種類の契約データ炒めを作れ！」と命令されて厨房に殺到しました。"),
        ("CPU 125%消費", "二刀流でフライパンを振る料理人", "1人の料理人が両手を使って、コンロを1.25個分同時に使って猛スピードで炒め物をしています（マルチスレッド）。"),
        ("Load Average 17.59", "コンロ待ちの大行列（18人）", "コンロが8個しかないのに、10人の二刀流料理人（12.5口分）が押し寄せ、後ろに行列ができて大混雑しています。"),
        ("コンテキストスイッチ", "「ちょっとそこどいて！」の押し問答", "コンロの取り合いになり、フライパンを置いたり持ち上げたり場所を譲り合う無駄な時間（ロス）ばかりが増えています。"),
        ("解決策: 並列度を8(7)にする", "コンロの数に合わせて順番に調理", "料理人を7人に減らし、1口は店長や皿洗い用に空けます。押し問答ゼロになり、実はこの方が全員早く帰れます！")
    ]

    row = 5
    for item in analogies:
        is_hdr = (item[0] == "IT用語 / 実態")
        c1 = ws3.cell(row=row, column=1, value=item[0])
        c2 = ws3.cell(row=row, column=2, value=item[1])
        c3 = ws3.cell(row=row, column=3, value=item[2])
        ws3.merge_cells(start_row=row, start_column=3, end_row=row, end_column=7)

        for col_idx, cell in enumerate([c1, c2, c3], 1):
            if is_hdr:
                cell.font = font_tbl_header
                cell.fill = fill_tbl_header
                cell.alignment = align_header
                cell.border = border_header
            else:
                cell.font = font_data_bold if col_idx < 3 else font_data
                cell.fill = fill_zebra_alt if row % 2 == 0 else fill_zebra_light
                cell.alignment = align_center if col_idx < 3 else align_left
                cell.border = border_cell
        
        if not is_hdr:
            for c_idx in range(4, 8):
                ws3.cell(row=row, column=c_idx).border = border_cell
                ws3.cell(row=row, column=c_idx).fill = fill_zebra_alt if row % 2 == 0 else fill_zebra_light
        row += 1

    # Section 2: コンピュータ技術の歴史と背景
    row += 2
    ws3.cell(row=row, column=1, value="■ 2. なぜこの仕組みになったのか？ ITアーキテクチャ登場の歴史と背景").font = font_section
    row += 1

    headers_hist = ["時代・年代", "歴史的背景と直面した壁", "誕生した技術イノベーション", "現代（本件）への影響と教訓"]
    for idx, h in enumerate(headers_hist, 1):
        c = ws3.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    history_data = [
        ("1960~80年代\nメインフレーム時代", "計算機が国家予算並みに超高価だった時代。昼間は窓口業務、夜間は計算機が空くため、翌朝までに一括で大量データを処理する必要があった。", "【バッチ処理とジョブ管理の哲学】\n「高価なCPUを遊ばせるな」という思想から、CPUを100%限界まで使い切って最短時間で処理を終わらせる設計思想が確立された。", "バッチ処理中にCPUが100%近くに達するのは「ハードウェア資産を限界まで使い切っている」証拠であり、本来正常で優秀な設計。"),
        ("2000年代初頭\nクロック周波数の壁", "Pentium 4時代にCPUの動作周波数が4GHzに達した際、熱暴走と消費電力の急増（物理の壁）に衝突し、単一コアの高速化が限界を迎えた。", "【マルチコア・SMP（対称型マルチプロセッサ）】\n1つのCPUを速くするのを諦め、1つのチップに複数の「コア（脳）」を載せて並列処理する道へと人類は舵を切った。", "現在の8 vCPUや32 vCPUの前提。並列度を上げないとCPU能力を引き出せないが、コア数以上に並列化すると管理ロスが発生する。"),
        ("Javaスレッドの進化\n(1995~2020年代)", "初期Java（Green Threads）はOSスレッドを使わずJVM内で疑似並行実行していたため、CPUが何コアあっても1コアしか使えなかった。", "【ネイティブスレッド (1:1モデル)】\nJavaスレッド1本をOSのpthread 1本に直結。これによりJVMが複数コアを同時に消費し「125%」が出せるようになった。", "1つのJavaバッチが内部で複数スレッド（GC/JIT/ワーカー）を動かすため、JVM数＝CPU使用数ではなく「JVM数 × スレッド係数」となる。"),
        ("クラウド仮想化\n(XenからAWS Nitro)", "従来の仮想マシン（Xen等）はCPU割り当て時にハイパーバイザーの介入遅延やCPU横取り（Steal Time）が発生した。", "【AWS Nitro System (第5世代 / c7a)】\nvCPUは物理コアのハイパースレッディング（HT）として100%専有保証され、Steal Time 0.0%を実現。", "今回の計測でSteal Time 0.0%が実証。インフラの外部要因ではなく、純粋にゲストOS内のプロセス配置バランスの問題と特定できる。"),
        ("ジョブ運用の進化\n(JP1/AJSの役割)", "数百本ものバッチジョブを無制限に一斉実行すると、CPUもメモリも瞬時にパンクしシステム全体がクラッシュする問題が発生。", "【ジョブ実行多重度（キュー制御）】\nマシンの物理コア数に合わせて「同時に走らせるジョブの数」を厳密に絞り込み、順次実行するキューイング機構が必須となった。", "本件の根本解決策。物理コア数（8 vCPU）に対してJP1の多重度を適切に絞る（7〜8）ことで、システムを最速かつ安全に稼働させる。")
    ]

    for item in history_data:
        ws3.cell(row=row, column=1, value=item[0]).alignment = align_center
        ws3.cell(row=row, column=2, value=item[1]).alignment = align_left_top
        ws3.cell(row=row, column=3, value=item[2]).alignment = align_left_top
        ws3.cell(row=row, column=4, value=item[3]).alignment = align_left_top

        for c_idx in range(1, 5):
            cell = ws3.cell(row=row, column=c_idx)
            cell.font = font_data
            cell.fill = fill_zebra_alt if row % 2 == 0 else fill_zebra_light
            cell.border = border_cell
        row += 1

    # Section 3: テーブルダンプ実行時のOS内部動作パイプライン
    row += 2
    ws3.cell(row=row, column=1, value="■ 3. テーブルダンプ処理実行時のデータフローとCPU消費パイプライン").font = font_section
    row += 1

    headers_pipe = ["フェーズ", "処理内容", "主な消費リソース", "今回の実測値との整合性", "最適化の着眼点"]
    for idx, h in enumerate(headers_pipe, 1):
        c = ws3.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    pipeline_data = [
        ("① DBフェッチ", "JDBC経由でDBから契約テーブルレコードを受信", "ネットワークI/O (Up to 12.5Gbps)", "ネットワーク帯域が広大で待たされない（CPU sy 3.1%）", "Fetch Sizeのチューニング（デフォルト10件→1000件等）"),
        ("② オブジェクト展開", "受信バイナリをJava Entityオブジェクトへデシリアライズ", "メモリヒープ割当、CPU (us)", "JVMヒープに一時格納（used 5.5GBと安定）", "短命オブジェクトの過剰生成抑制"),
        ("③ 文字列整形・変換", "契約データをCSV/TSV/JSON等のダンプ形式へ文字コード変換", "CPU (us 95.9%) ← 最大ボトルネック", "%CPU 125%の正体。文字列連結・文字コード変換の計算負荷", "StringBuilder利用、効率的なCSVシリアライザの採用"),
        ("④ ガベージコレクション", "出力完了した旧世代オブジェクトのメモリ回収 (G1GC/Parallel)", "CPU (us / 並列スレッド)", "JVMが100%超を消費する要因（バックグラウンドGC稼働）", "-XX:ParallelGCThreads の明示的制限"),
        ("⑤ ディスク出力", "ファイルへダンプ結果をwriteシステムコールで書き出し", "OSページキャッシュ (buff/cache)", "wa 0.8%と極小。8.9GBのbuff/cacheがディスクI/Oを全吸収", "バッファ付きストリーム（BufferedOutputStream）の適用")
    ]

    for item in pipeline_data:
        ws3.cell(row=row, column=1, value=item[0]).alignment = align_center
        ws3.cell(row=row, column=2, value=item[1]).alignment = align_left
        ws3.cell(row=row, column=3, value=item[2]).alignment = align_left
        ws3.cell(row=row, column=4, value=item[3]).alignment = align_left
        ws3.cell(row=row, column=5, value=item[4]).alignment = align_left

        for c_idx in range(1, 6):
            cell = ws3.cell(row=row, column=c_idx)
            cell.font = font_data
            cell.fill = fill_zebra_alt if row % 2 == 0 else fill_zebra_light
            cell.border = border_cell
        row += 1

    # Set column widths
    col_widths_ws3 = [18, 30, 32, 34, 30, 10, 10]
    for i, w in enumerate(col_widths_ws3, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 4: サイジングと並列度比較 (Sizing & Concurrency Simulation)
    # =========================================================================
    ws4 = wb.create_sheet(title="サイジングと並列度比較")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:G2")
    ws4["A1"] = "【徹底比較】IT環境 (c7a.2xlarge) vs 本番環境 (c7a.8xlarge) 並列度とCPU・スループット相関検証"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_title
    ws4["A1"].alignment = align_header

    ws4["A4"] = "■ 1. IT環境（c7a.2xlarge: 8 vCPU / 16 GiB）におけるJVM並列数シミュレーション"
    ws4["A4"].font = font_section

    headers_sim1 = ["設定パターン", "JP1マルチ数", "総消費CPU期待値", "Load Average予測", "コンテキストスイッチ", "全体処理完了時間", "総合判定と推奨度"]
    row = 5
    for idx, h in enumerate(headers_sim1, 1):
        c = ws4.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    sim_it = [
        ("現行設定", "10 並列", "約 12.0 vCPU相当 (要求過多)", "16.0 ~ 18.0 (実測 17.59)", "激甚 (CPUキャッシュ破棄多発)", "基準時間 (100%)", "非推奨 (CPU過密飽和・OS遅延・アラート頻発)"),
        ("コア等倍設定", "8 並列", "約 9.6 vCPU相当 (やや過多)", "8.5 ~ 10.0", "中程度 (適正範囲)", "約 90% ~ 95% (短縮)", "妥当 (実効スループット向上)"),
        ("ベスト推奨設定", "7 並列 (8-1)", "約 8.4 vCPU相当 (上限付近)", "7.0 ~ 8.0", "最小 (CPUキャッシュ高効率)", "約 90% (最短終了)", "【最も推奨】(OS/監視用に1コア確保・安定最速)"),
        ("安全重視設定", "6 並列 (8-2)", "約 7.2 vCPU相当 (余力あり)", "5.5 ~ 6.5", "極小 (完全スムーズ)", "約 105% (ごく僅かに延伸)", "推奨 (アラート絶対抑止・安定稼働最優先)")
    ]

    for item in sim_it:
        for c_idx in range(1, 8):
            cell = ws4.cell(row=row, column=c_idx, value=item[c_idx-1])
            cell.font = font_data_bold if "推奨" in item[6] else font_data
            cell.border = border_cell
            cell.alignment = align_left if c_idx in [1, 7] else align_center
            if "最も推奨" in item[6]:
                cell.fill = fill_highlight
            elif "非推奨" in item[6]:
                cell.fill = fill_zebra_alt
            else:
                cell.fill = fill_zebra_light
        row += 1

    # Section 2: 本番環境（c7a.8xlarge: 32 vCPU / 64 GiB / 58 JVM予定）のシミュレーション
    row += 2
    ws4.cell(row=row, column=1, value="■ 2. 本番環境（c7a.8xlarge: 32 vCPU / 64 GiB / 58 JVM予定）におけるマルチ数検証").font = font_section
    row += 1

    headers_sim2 = ["設定パターン", "JP1マルチ数", "同時稼働JVM数", "CPU負荷予測", "DB接続負荷 (RDS/Aurora)", "メモリ消費推計", "本番運用の判定とコメント"]
    for idx, h in enumerate(headers_sim2, 1):
        c = ws4.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    sim_prod = [
        ("無制限一斉実行 (NG例)", "58 並列", "58 JVM", "100% 張り付き (Load Avg 60+)", "58本同時フルスキャン (DB崩壊リスク)", "約 45 ~ 60 GB (メモリ枯渇危険)", "絶対不可 (DB枯渇・OSハングの危険触発)"),
        ("コア等倍実行 (質問のケース)", "32 並列", "32 JVM", "95 ~ 100% (Load Avg 35~42)", "32本同時実行 (DB側IOPS急増懸念)", "約 25 ~ 35 GB (64GB搭載なら可)", "注意 (上限ギリギリ。DB側のRead IOPS耐性検証が必須)"),
        ("本番推奨設定 (実務ベスト)", "24 ~ 28 並列", "24 ~ 28 JVM", "80 ~ 90% (Load Avg 24~28)", "適正範囲 (キューイングにより平準化)", "約 20 ~ 28 GB (極めて安全)", "【本番推奨】(CPU余力確保・DB安定・バッチ滞留なし)"),
        ("保守的設定 (DB保護優先)", "16 並列 (半数)", "16 JVM", "50 ~ 65% (Load Avg 16~18)", "低負荷 (DBへの影響最小限)", "約 15 ~ 20 GB (完全安全)", "安全 (ただしバッチ全体の突き抜け・終了時刻延伸に注意)")
    ]

    for item in sim_prod:
        for c_idx in range(1, 8):
            cell = ws4.cell(row=row, column=c_idx, value=item[c_idx-1])
            cell.font = font_data_bold if "本番推奨" in item[6] else font_data
            cell.border = border_cell
            cell.alignment = align_left if c_idx in [1, 7] else align_center
            if "本番推奨" in item[6]:
                cell.fill = fill_highlight
            elif "絶対不可" in item[6]:
                cell.fill = fill_zebra_alt
            else:
                cell.fill = fill_zebra_light
        row += 1

    # Section 3: なぜ「並列度を減らすと速くなる」のか？（アムダールの法則とオーバーヘッド）
    row += 2
    ws4.cell(row=row, column=1, value="■ 3. 並列処理のパラドックス：並列数を増やしても性能が頭打ち・低下する力学").font = font_section
    row += 1

    headers_paradox = ["ボトルネック要因", "並列数過多（10並列/8vCPU）で起きること", "適正並列数（7~8並列）で起きること", "物理的メカニズム"]
    for idx, h in enumerate(headers_paradox, 1):
        c = ws4.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    paradox_data = [
        ("CPU L1/L2 キャッシュ破棄", "10個のJVMが頻繁に切り替わるため、CPU内部の超高速キャッシュ（L1/L2）が毎回クリアされメインメモリへの再読込が発生。", "同一コアで同一スレッドが連続実行されるため、キャッシュヒット率が最大化し計算が高速化。", "キャッシュスラッシング防止\n(L1アクセス: 1ns vs メモリ: 100ns)"),
        ("Linux CFS スケジューラ競合", "実行可能キューに10個以上のタスクが常に並び、スケジューラが頻繁にタイマー割り込みとレジスタ退避/復元を実行。", "キュー滞留が最小化し、カーネルのコンテキストスイッチコスト（sy時間）が極小化。", "プロセス切替コストの低減\n(コンテキストスイッチ損失排除)"),
        ("JVM ガベージコレクション (GC)", "10個のJVMが一斉にGCスレッドを立ち上げるため、GCスレッド同士がCPUコアを取り合ってStop The Worldが延伸。", "GCの実行が時間差で分散し、各JVMのGCが短時間で即座に完了。", "スレッド爆発の防止\n(-XX:ParallelGCThreads競合解消)"),
        ("DB / ストレージ 競合", "10本同時にDBからSELECTすると、DBのバッファプール追い出しやテーブルロック競合が発生。", "適度な順次実行により、DBのReadキャッシュが有効活用されトータルスループット向上。", "キューイング理論（M/M/sモデル）\nによる全体待ち時間最適化")
    ]

    for item in paradox_data:
        ws4.cell(row=row, column=1, value=item[0]).alignment = align_center
        ws4.cell(row=row, column=2, value=item[1]).alignment = align_left
        ws4.cell(row=row, column=3, value=item[2]).alignment = align_left
        ws4.cell(row=row, column=4, value=item[3]).alignment = align_left

        for c_idx in range(1, 5):
            cell = ws4.cell(row=row, column=c_idx)
            cell.font = font_data
            cell.fill = fill_zebra_alt if row % 2 == 0 else fill_zebra_light
            cell.border = border_cell
        row += 1

    col_widths_ws4 = [18, 14, 28, 24, 26, 24, 38]
    for i, w in enumerate(col_widths_ws4, 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 5: 監視設計と対策ロードマップ (Monitoring & Action Plan)
    # =========================================================================
    ws5 = wb.create_sheet(title="監視設計と対策ロードマップ")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:G2")
    ws5["A1"] = "【インフラ実践提言】EC2監視設計の再構築・しきい値見直し・具体的対策ロードマップ"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_title
    ws5["A1"].alignment = align_header

    ws5["A4"] = "■ 1. Web/APサーバ vs バッチサーバの監視思想の違いとあるべきしきい値設計"
    ws5["A4"].font = font_section

    headers_mon = ["サーバ種別", "特性・目的", "推奨 Warning しきい値", "推奨 Critical しきい値", "評価期間 (Evaluation Period)", "検知時の意味・運用アクション"]
    row = 5
    for idx, h in enumerate(headers_mon, 1):
        c = ws5.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    mon_data = [
        ("Web / AP サーバ\n(OLTP系)", "不特定のユーザリクエストに応答。突発スパイクに耐えるための余力確保が必須。", "70% ~ 75%", "85%", "5分 (1分×5回 または 5分×1回)", "突発トラフィック増またはアプリループ。オートスケール発動または調査が必要。"),
        ("バッチサーバ\n(通常時間帯)", "昼間やジョブ非実行時間。OS待機・JP1エージェント常駐。", "60% ~ 70%", "80%", "5分", "予期せぬプロセス暴走やゾンビプロセスの存在を疑う。"),
        ("バッチサーバ\n(バッチ実行窓：現行)", "夜間テーブルダンプ等の一括実行中。短時間で終わらせるためCPUを使う。", "80% (誤発報要因)", "90% (誤発報要因)", "5分 (短すぎる)", "【現状】正常な高負荷なのにクリティカル発報して運用が狼少年化している。"),
        ("バッチサーバ\n(バッチ実行窓：推奨案)", "高CPU利用を「正常」と位置づけ、真のハングアップ（無限ループ・停止）のみを検知。", "90% (15分連続)", "95% (20分連続)", "15 ~ 30分 (連続継続)", "バッチ想定完了時間を超過してCPUが張り付いている「真の異常」のみを的確に通報。")
    ]

    for item in mon_data:
        for c_idx in range(1, 7):
            cell = ws5.cell(row=row, column=c_idx, value=item[c_idx-1])
            cell.font = font_data_bold if "推奨案" in item[0] else font_data
            cell.border = border_cell
            cell.alignment = align_left if c_idx in [1, 2, 6] else align_center
            if "推奨案" in item[0]:
                cell.fill = fill_highlight
            elif "現行" in item[0]:
                cell.fill = fill_zebra_alt
            else:
                cell.fill = fill_zebra_light
        row += 1

    # Section 2: 監視の高度化（CloudWatch Composite Alarms / 静観運用）
    row += 2
    ws5.cell(row=row, column=1, value="■ 2. クラウドインフラにおける監視改善アプローチ比較").font = font_section
    row += 1

    headers_cloud = ["対策方式", "概要・実装手法", "メリット", "デメリット・考慮事項", "推奨度"]
    for idx, h in enumerate(headers_cloud, 1):
        c = ws5.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    cloud_data = [
        ("① 評価期間の延長\n(即効性No.1)", "CloudWatch Alarmの評価期間を「5分」から「15分×2回（計30分）」等に延長する。", "設定変更がコンソール/Terraformで即座に可能。追加コストゼロ。", "バッチ実行時間が30分を超える場合は依然として発報する可能性あり。", "◎ (即座に実施)"),
        ("② バッチ時間帯の\nアラート静観運用", "JP1ジョブネット起動前後、またはEventBridgeスケジュールで夜間帯のアラート通知を無効化（Action無効化）。", "夜間の正常な高負荷による当番担当者の呼出・狼少年化を完全にゼロ化。", "バッチ実行中に本物の障害（プロセス死など）が発生した場合に検知が遅れるリスク（JP1側監視でカバー必須）。", "○ (有力選択肢)"),
        ("③ 複合アラーム\n(Composite Alarm)", "「CPU > 90%」AND「プロセス応答なし OR Load Avg > 30」など複数メトリクスの論理積で発報。", "誤検知を極小化し、真のサーバー死のみを100%ピンポイント検知。", "CloudWatch Composite Alarmsの追加設定（僅かな月額費用）とCloudWatch Agent設計が必要。", "◎ (本番向け推奨)"),
        ("④ JP1側監視への\n主軸移行", "インフラのCPU%監視は参考値とし、バッチの正常性はJP1の「ジョブ終了コード」「実行遅延監視（SLA）」に一元化。", "業務影響（ジョブが時間内に正常終了したか）に直結した監視体系となる。", "インフラ単体障害（OSハングでJP1マネージャーへの応答途絶）を検知する死活監視（Ping/SSM）は別途維持が必要。", "◎ (運用ベストプラクティス)")
    ]

    for item in cloud_data:
        for c_idx in range(1, 6):
            cell = ws5.cell(row=row, column=c_idx, value=item[c_idx-1])
            cell.font = font_data_bold if "◎" in item[4] else font_data
            cell.border = border_cell
            cell.alignment = align_left if c_idx in [1, 2, 3, 4] else align_center
            cell.fill = fill_highlight if "◎" in item[4] else fill_zebra_light
        row += 1

    # Section 3: 具体的対策ロードマップ
    row += 2
    ws5.cell(row=row, column=1, value="■ 3. 恒久対策の具体的実行ロードマップ（4フェーズ）").font = font_section
    row += 1

    headers_road = ["フェーズ", "対策項目", "具体的設定内容・コマンド・パラメータ", "担当領域", "期待効果"]
    for idx, h in enumerate(headers_road, 1):
        c = ws5.cell(row=row, column=idx, value=h)
        c.font = font_tbl_header
        c.fill = fill_tbl_header
        c.alignment = align_header
        c.border = border_header
    row += 1

    road_data = [
        ("Phase 1\n(即時対応)", "JP1並列数の適正化\n(IT環境)", "JP1ジョブネットのマルチ数（同時実行数）を「10」→「7」または「8」へ変更。", "ジョブ運用 / アプリ", "コンテキストスイッチ激減。Load Avgが17.59→8.0前後に適正化。CPU競合解消。"),
        ("Phase 2\n(即時対応)", "監視しきい値の暫定修正", "CloudWatch Alarmの評価期間を15分へ延長、またはバッチ実行時間帯のSNS通知を抑止。", "インフラ / SRE", "夜間帯の不要なクリティカルアラート誤発報（狼少年化）を直ちにストップ。"),
        ("Phase 3\n(本番準備)", "JVMスレッド制限の組込", "バッチ起動時Javaオプションに以下を付与:\n-XX:ParallelGCThreads=2 -XX:CICompilerCount=2", "アプリ開発", "1つのJVMが不必要に全コア（32コア等）を奪うGCスパイクを物理的に防止。"),
        ("Phase 4\n(本番検証)", "本番32vCPU負荷試験", "JP1マルチ数 24 → 28 → 32 と段階的に引き上げ、バッチ終了時刻とDB Read IOPSを計測。", "総合テストPJ", "本番環境での安全マージンとDB限界値を見極め、本番運用最適値（推奨24~28）を確定。")
    ]

    for item in road_data:
        for c_idx in range(1, 6):
            cell = ws5.cell(row=row, column=c_idx, value=item[c_idx-1])
            cell.font = font_data_bold if "Phase 1" in item[0] else font_data
            cell.border = border_cell
            cell.alignment = align_left if c_idx in [2, 3, 5] else align_center
            cell.fill = fill_zebra_alt if "Phase 1" in item[0] else fill_zebra_light
        row += 1

    col_widths_ws5 = [16, 26, 38, 16, 38, 10, 10]
    for i, w in enumerate(col_widths_ws5, 1):
        ws5.column_dimensions[get_column_letter(i)].width = w

    output_path = "EC2_Performance_Analysis_Batch_Dump.xlsx"
    wb.save(output_path)
    print(f"Excel report saved successfully to {output_path}")

if __name__ == "__main__":
    build_excel_report()
