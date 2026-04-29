import streamlit as st
from notion_client import Client
import os
from dotenv import load_dotenv
import requests
import json
import openai
import base64
from docx import Document 
from docx.shared import Inches, Pt
from io import BytesIO
import re
from datetime import datetime
import cv2
import tempfile

# PDF解析ライブラリのインポート（エラー回避策付き）
try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

# 1. 環境設定の読み込み
load_dotenv()
NOTION_TOKEN = os.getenv("NOTION_API_KEY")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
raw_api_key = os.getenv("OPENAI_API_KEY")

# OpenAI API Keyの設定
if raw_api_key:
    openai.api_key = raw_api_key.strip().strip('"').strip("'")

# --- セッション状態の初期化 ---
if 'manual_text' not in st.session_state:
    st.session_state['manual_text'] = ""
if 'checklist_text' not in st.session_state:
    st.session_state['checklist_text'] = ""
if 'source_files' not in st.session_state:
    st.session_state['source_files'] = []
if 'processed_images_bytes' not in st.session_state:
    st.session_state['processed_images_bytes'] = []

# Notionクライアントの初期化
if "notion" not in st.session_state:
    if NOTION_TOKEN:
        try:
            st.session_state.notion = Client(auth=NOTION_TOKEN)
        except Exception:
            st.session_state.notion = None
    else:
        st.session_state.notion = None

# --- 共通関数 ---

def encode_image(image_bytes):
    return base64.b64encode(image_bytes).decode('utf-8')

def add_to_notion(name, definition):
    if not st.session_state.notion or not DATABASE_ID:
        return False, "Notion設定が不足しています"
    try:
        st.session_state.notion.pages.create(
            parent={"database_id": DATABASE_ID},
            properties={
                "名称": {"title": [{"text": {"content": name.strip()}}]},
                "意味": {"rich_text": [{"text": {"content": definition.strip()}}]}
            }
        )
        return True, f"「{name}」を登録しました！"
    except Exception as e:
        return False, f"Notion登録エラー: {str(e)}"

def get_notion_data(search_query="", mode="名称のみ"):
    if not NOTION_TOKEN or not DATABASE_ID:
        return []
    filter_data = None
    if search_query:
        if mode == "名称のみ":
            filter_data = {"property": "名称", "title": {"contains": search_query}}
        else:
            filter_data = {
                "or": [
                    {"property": "名称", "title": {"contains": search_query}},
                    {"property": "意味", "rich_text": {"contains": search_query}}
                ]
            }
    payload = {}
    if filter_data: payload["filter"] = filter_data
    if not search_query:
        payload["page_size"] = 5
        payload["sorts"] = [{"timestamp": "created_time", "direction": "descending"}]

    try:
        if st.session_state.notion:
            response = st.session_state.notion.databases.query(database_id=DATABASE_ID, **payload)
            return response.get("results", [])
    except Exception: pass
    try:
        url = f"https://api.notion.com/v1/databases/{DATABASE_ID}/query"
        headers = {"Authorization": f"Bearer {NOTION_TOKEN}", "Notion-Version": "2022-06-28", "Content-Type": "application/json"}
        res = requests.post(url, headers=headers, data=json.dumps(payload))
        if res.status_code == 200: return res.json().get("results", [])
    except Exception: pass
    return []

# --- 各ページのコンテンツ定義 ---

def page_manual_creator():
    st.header("📝 AIマニュアル作成アシスタント V13")
    
    uploaded_files = st.file_uploader("写真または動画をアップロード（順番通りに）", 
                                      type=['png', 'jpg', 'jpeg', 'mp4', 'mov'], 
                                      accept_multiple_files=True)

    if uploaded_files:
        current_file_names = [f.name for f in uploaded_files]
        previous_file_names = [f.name for f in st.session_state['source_files']]
        
        if current_file_names != previous_file_names:
            st.session_state['manual_text'] = ""
            st.session_state['checklist_text'] = ""
            st.session_state['source_files'] = uploaded_files
            st.session_state['processed_images_bytes'] = []
            
            temp_processed_bytes = []
            with st.spinner("ファイルを読み込み・解析中..."):
                for file in uploaded_files:
                    file_ext = file.name.split('.')[-1].lower()
                    if file_ext in ['mp4', 'mov']:
                        st.info(f"🎥 動画「{file.name}」から画像を抽出中...")
                        with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_ext}') as tfile:
                            tfile.write(file.read())
                            temp_path = tfile.name
                        cap = cv2.VideoCapture(temp_path)
                        fps = cap.get(cv2.CAP_PROP_FPS) or 30
                        frame_interval = int(fps * 3)
                        count = 0
                        while cap.isOpened():
                            ret, frame = cap.read()
                            if not ret: break
                            if count % frame_interval == 0:
                                _, buffer = cv2.imencode('.jpg', frame)
                                temp_processed_bytes.append(buffer.tobytes())
                            count += 1
                        cap.release()
                        os.remove(temp_path)
                    else:
                        temp_processed_bytes.append(file.getvalue())
                st.session_state['processed_images_bytes'] = temp_processed_bytes

    if st.session_state['processed_images_bytes']:
        st.write("### 📸 AIが解析する画像リスト")
        col_pre = st.columns(min(len(st.session_state['processed_images_bytes']), 5))
        for idx, img_bytes in enumerate(st.session_state['processed_images_bytes']):
            with col_pre[idx % 5]:
                st.image(img_bytes, caption=f"画像{idx+1}", width=150)

    all_images_base64 = [encode_image(b) for b in st.session_state['processed_images_bytes']]

    st.markdown("---")
    col_btn1, col_btn2 = st.columns([1, 1])
    
    with col_btn1:
        if st.button("マニュアルを新規生成する", use_container_width=True):
            if not all_images_base64:
                st.warning("画像をアップロードしてください。")
            else:
                with st.spinner("AIがマニュアルを作成中..."):
                    prompt_text = """
                    あなたはプロの業務マニュアル作成者です。提供された画像（[画像1], [画像2]...）を解析し、以下のルールでマニュアルを作成してください。

                    ### 出力フォーマット構成
                    1. 1行目：必ず「タイトル：[作業名]」のみを出力してください。挨拶や謝罪は絶対に入れないでください。
                    2. 手順の記述：
                       手順1：[画像から読み取った操作内容の記述]
                       [画像1]
                       
                       手順2：[画像から読み取った操作内容の記述]
                       [画像2]
                       ...という形式で、すべての手順を記述してください。

                    ### 画像解析ルール
                    - **優先順位**: 画像内に「○で囲まれた数字（①、②など）」がある場合は、その数字の順番を手順の番号として優先してください。
                    - **重要事項**: 画像内に手書きのメモ、赤枠、○囲み、矢印などがある場合、その内容は必ず手順の中に「※重要：[内容]」として記述してください。
                    - **メタ情報の排除**: 「画像タグ：」「: ファイル選択画面」「[画像1], [画像2]」といったリストだけの行は作成しないでください。必ず「手順X：」の直後に配置してください。
                    - **構成**: 1つの手順に対して、対応する画像タグ（例：[画像1]）を1つ、必ずその手順のすぐ下に配置してください。
                    - **免責事項・注意書きの絶対禁止**: 「画像の内容に基づいており実際の操作と異なる場合があります」「解析できませんでしたがガイドラインを提供します」などのAIとしての断り書き、謝罪、前置き、後書きは **絶対に** 出力しないでください。
                    """
                    
                    content_payload = [{"type": "text", "text": prompt_text}]
                    for img in all_images_base64:
                        content_payload.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img}"}})

                    try:
                        response = openai.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": content_payload}],
                            max_tokens=2500,
                            temperature=0.3,
                        )
                        st.session_state['manual_text'] = response.choices[0].message.content
                    except Exception as e:
                        st.error(f"エラー: {e}")

    with col_btn2:
        if st.button("確認用チェックリストを作成", use_container_width=True):
            if not st.session_state['manual_text']:
                st.warning("先にマニュアルを生成してください。")
            else:
                with st.spinner("チェックリストを分析中..."):
                    try:
                        response = openai.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": f"以下のマニュアルに基づき、作業者がミスをしないための「確認用チェックリスト」を作成せよ。\n\n【ルール】\n・各項目の先頭は「□ 」にすること。\n・「確認用チェックリスト」というタイトルや、前置き・後書きの説明文（「以下はチェックリストです」など）は一切出力しないでください。\n・箇条書きのリスト部分のみを出力してください。\n\n{st.session_state['manual_text']}"}],
                            max_tokens=1000,
                        )
                        st.session_state['checklist_text'] = response.choices[0].message.content
                    except Exception as e:
                        st.error(f"エラー: {e}")

    # 結果表示
    if st.session_state['manual_text']:
        lines = st.session_state['manual_text'].strip().split('\n')
        title_line = lines[0].strip()
        title_name = title_line.replace("タイトル：", "").replace("タイトル:", "").strip()
        
        if not title_name or any(x in title_name for x in ["申し訳", "できません", "不明", "ガイドライン", "提供"]):
            doc_title = f"{datetime.now().strftime('%Y%m%d')}_業務マニュアル"
        else:
            doc_title = title_name

        safe_file_name = re.sub(r'[\\/:*?"<>|]', '', doc_title)

        st.divider()
        col_res1, col_res2 = st.columns([3, 1])
        with col_res1: st.success("マニュアルが表示可能です。")
        with col_res2:
            doc = Document()
            doc.add_heading(doc_title, 0)
            for line in lines[1:]:
                clean_line = line.strip()
                if not clean_line: continue
                
                matches = re.findall(r'\[画像(\d+)\]', clean_line)
                if matches:
                    for m in matches:
                        idx = int(m) - 1
                        if 0 <= idx < len(st.session_state['processed_images_bytes']):
                            doc.add_picture(BytesIO(st.session_state['processed_images_bytes'][idx]), width=Inches(4.5))
                else:
                    if "画像タグ" in clean_line and ":" in clean_line: continue
                    if "実際の操作手順は異なる場合" in clean_line or "画像の内容に基づいており" in clean_line: continue
                    
                    clean_text = clean_line.replace('**', '').replace('#', '').strip()
                    if clean_text:
                        p = doc.add_paragraph()
                        run = p.add_run(clean_text)
                        if clean_text.startswith("手順"):
                            run.bold = True
                            run.font.size = Pt(12)
            
            if st.session_state['checklist_text']:
                doc.add_page_break()
                doc.add_heading("確認用チェックリスト", level=1)
                for cl in st.session_state['checklist_text'].split('\n'):
                    clean_cl = cl.replace('**', '').replace('#', '').strip()
                    if clean_cl and "確認用" not in clean_cl and "チェックリスト" not in clean_cl and "以下は" not in clean_cl:
                        if not clean_cl.startswith('□'): clean_cl = f"□ {clean_cl}"
                        p = doc.add_paragraph(clean_cl)
                        p.paragraph_format.left_indent = Inches(0.2)
            
            bio = BytesIO()
            doc.save(bio)
            st.download_button("📥 Wordファイルで保存", bio.getvalue(), file_name=f"{safe_file_name}.docx")

        st.markdown(f"# {doc_title}")
        for line in lines[1:]:
            clean_line = line.strip()
            if not clean_line: continue
            
            matches = re.findall(r'\[画像(\d+)\]', clean_line)
            if matches:
                img_cols = st.columns(len(matches))
                for idx, m in enumerate(matches):
                    img_idx = int(m) - 1
                    if 0 <= img_idx < len(st.session_state['processed_images_bytes']):
                        with img_cols[idx]:
                            st.image(st.session_state['processed_images_bytes'][img_idx], width=300)
            else:
                if "画像タグ" in clean_line and ":" in clean_line: continue
                if "実際の操作手順は異なる場合" in clean_line or "画像の内容に基づいており" in clean_line: continue
                
                display_line = clean_line.replace('#', '').replace('**', '').strip()
                if display_line.startswith("手順"):
                    st.markdown(f"<div style='font-size: 1.15em; font-weight: bold; margin-top: 15px; margin-bottom: 5px;'>{display_line}</div>", unsafe_allow_html=True)
                else:
                    st.markdown(display_line)

        if st.session_state['checklist_text']:
            st.divider()
            st.markdown(f"## {doc_title}：チェックリスト")
            st.markdown("""
                <div style='border: 1px solid #ccc; padding: 5px 10px; margin-bottom: 15px; width: auto; display: inline-block; font-weight: bold;'>
                    確認用チェックリスト
                </div>
                """, unsafe_allow_html=True)
                
            for cl in st.session_state['checklist_text'].split('\n'):
                c = cl.replace('**', '').replace('#', '').strip()
                if c and "確認用" not in c and "チェックリスト" not in c and "以下は" not in c:
                    if not c.startswith('□'): c = f"□ {c}"
                    st.markdown(f"<div style='margin-left: 1.5em; margin-bottom: 0.5em;'>{c}</div>", unsafe_allow_html=True)

def page_glossary_search():
    st.header("🔍 用語検索")
    col1, col2 = st.columns([3, 1])
    with col1: q = st.text_input("検索ワード", placeholder="キーワードを入力...", key="search_input")
    with col2: m = st.selectbox("検索対象", ["名称のみ", "全体（意味を含む）"], key="search_mode")
    st.divider()
    with st.spinner("取得中..."): items = get_notion_data(q, m)
    if items:
        for item in items:
            try:
                props = item.get("properties", {})
                name = props.get("名称", {}).get("title", [])[0]["plain_text"]
                definition = props.get("意味", {}).get("rich_text", [])[0]["plain_text"]
                with st.expander(f"📌 {name}"):
                    st.write(definition)
            except: continue

def page_glossary_registration():
    st.header("📥 用語の登録")
    
    tab1, tab2, tab3 = st.tabs(["手動入力", "一括登録（コピペ）", "PDF解析"])
    
    with tab1:
        st.subheader("1件ずつ登録")
        with st.form("single_form", clear_on_submit=True):
            name = st.text_input("用語の名称")
            definition = st.text_area("意味・解説")
            submitted = st.form_submit_button("Notionへ登録")
            if submitted:
                if name and definition:
                    with st.spinner("登録中..."):
                        success, msg = add_to_notion(name, definition)
                        if success: st.success(msg)
                        else: st.error(msg)
                else:
                    st.warning("名称と意味を入力してください。")

    with tab2:
        st.subheader("一括登録")
        st.write("「用語 > 意味」の形式で入力（※改行で複数を一括登録可）")
        st.caption("※区切り記号は「>」または「＞」が使用可能です。")
        bulk_text = st.text_area("貼り付けエリア", height=300, placeholder="API > アプリ間の窓口\nUI > ユーザーインターフェース")
        
        if st.button("まとめて登録を実行"):
            if bulk_text:
                lines = bulk_text.strip().split("\n")
                success_count = 0
                with st.spinner("一括登録中..."):
                    for line in lines:
                        sep = ">" if ">" in line else "＞" if "＞" in line else None
                        if sep:
                            parts = line.split(sep, 1)
                            if len(parts) == 2:
                                ok, _ = add_to_notion(parts[0], parts[1])
                                if ok: success_count += 1
                st.success(f"{success_count}件登録しました。")
            else:
                st.warning("テキストを入力してください。")

    with tab3:
        st.subheader("PDFから用語を自動抽出・登録")
        
        # PyPDF2が利用可能かチェック
        if PyPDF2 is None:
            st.error("⚠️ PDF解析ライブラリ(PyPDF2)がインストールされていません。requirements.txtに記載してデプロイし直してください。")
            return

        st.write("参考書や資料などのPDFをアップロードすると、AIが重要な用語を読み取ってNotionに自動登録します。")
        st.caption("※PCのフォルダから複数のPDFをマウスで囲んで、一気にアップロードすることができます。")
        
        uploaded_pdfs = st.file_uploader("PDFファイルをアップロード（複数選択可）", type="pdf", accept_multiple_files=True)
        
        chunk_size = st.slider("1回あたりの解析ページ数", min_value=10, max_value=100, value=40, help="一度に数百ページを解析するとAIが用語を見落としやすくなります。推奨の30〜50ページ単位で自動分割して処理します。")
        
        if st.button("PDF解析と自動登録を開始"):
            if not uploaded_pdfs:
                st.warning("PDFファイルをアップロードしてください。")
            elif not raw_api_key:
                st.error("OpenAI API Keyが設定されていません。")
            else:
                total_files = len(uploaded_pdfs)
                for idx, pdf_file in enumerate(uploaded_pdfs):
                    st.markdown(f"**📄 {pdf_file.name}** の解析を開始します ({idx+1}/{total_files})")
                    
                    try:
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        total_pages = len(pdf_reader.pages)
                        st.info(f"総ページ数: {total_pages}ページ（{chunk_size}ページずつ分割して処理します）")
                        
                        progress_bar = st.progress(0)
                        
                        for i, start_page in enumerate(range(0, total_pages, chunk_size)):
                            end_page = min(start_page + chunk_size, total_pages)
                            with st.spinner(f"ページ {start_page+1} 〜 {end_page} を解析・登録中..."):
                                
                                text_chunk = ""
                                for page_num in range(start_page, end_page):
                                    extracted_text = pdf_reader.pages[page_num].extract_text()
                                    if extracted_text:
                                        text_chunk += extracted_text + "\n"
                                
                                if not text_chunk.strip():
                                    st.warning(f"ページ {start_page+1}〜{end_page}: テキストを抽出できませんでした（スキャン画像などの可能性があります）。")
                                    progress_bar.progress((end_page) / total_pages)
                                    continue
                                
                                prompt = f"""
                                以下のテキストから、重要な専門用語やキーワードを抽出し、その名称と意味を簡潔にまとめてください。
                                出力は必ず以下のJSON形式のみで行ってください。他の文章や挨拶は一切含めないでください。
                                {{
                                    "terms": [
                                        {{"名称": "用語名", "意味": "意味の説明"}},
                                        ...
                                    ]
                                }}
                                
                                【テキスト】
                                {text_chunk}
                                """
                                
                                try:
                                    response = openai.chat.completions.create(
                                        model="gpt-4o-mini",
                                        response_format={"type": "json_object"},
                                        messages=[{"role": "user", "content": prompt}],
                                        temperature=0.3
                                    )
                                    
                                    result_json = json.loads(response.choices[0].message.content)
                                    terms = result_json.get("terms", [])
                                    
                                    if terms:
                                        success_count = 0
                                        for term in terms:
                                            name = term.get("名称", "").strip()
                                            definition = term.get("意味", "").strip()
                                            if name and definition:
                                                ok, _ = add_to_notion(name, definition)
                                                if ok: success_count += 1
                                                
                                        st.success(f"✅ ページ {start_page+1}〜{end_page}: {success_count}件の用語を登録しました！")
                                    else:
                                        st.info(f"ページ {start_page+1}〜{end_page}: 登録すべき用語は見つかりませんでした。")
                                        
                                except Exception as e:
                                    st.error(f"AI解析エラー (ページ {start_page+1}〜{end_page}): {e}")
                                
                            progress_bar.progress((end_page) / total_pages)
                            
                    except Exception as e:
                        st.error(f"ファイル読み込みエラー ({pdf_file.name}): {e}")
                        
                st.balloons()
                st.success("🎉 すべてのPDFの解析とNotionへの登録が完了しました！")

# --- メインレイアウト ---
st.set_page_config(page_title="お仕事支援マルチツール", layout="wide")

with st.sidebar:
    st.title("🛠️ Menu")
    selection = st.radio("機能選択", ["AIマニュアル作成", "用語検索", "用語登録"])
    st.divider()
    if NOTION_TOKEN and DATABASE_ID: st.success("Notion Connected")
    else: st.error("Notion Disconnected")
    if raw_api_key: st.success("OpenAI Ready")
    else: st.error("OpenAI API Key Missing")

if selection == "AIマニュアル作成":
    page_manual_creator()
elif selection == "用語検索":
    page_glossary_search()
elif selection == "用語登録":
    page_glossary_registration()