import streamlit as st
import streamlit.components.v1 as components
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
import unicodedata # ★新規追加：全角・半角を統一して検索するためのライブラリ

# --- Google Drive API 用に追加 ---
from google.oauth2 import service_account
from googleapiclient.discovery import build

# PDF解析用ライブラリ
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# 1. 環境設定の読み込み
load_dotenv()
NOTION_TOKEN = os.getenv("NOTION_API_KEY")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID") # 用語集用DB
MANUAL_DB_ID = os.getenv("NOTION_MANUAL_DB_ID") # マニュアル用DB
raw_api_key = os.getenv("OPENAI_API_KEY")

# --- Google Drive 設定 ---
# --- Google Drive 設定 ---
def get_gdrive_service():
    # 1. Streamlit Secrets (クラウド用)
    if "gcp_service_account" in st.secrets:
        creds = service_account.Credentials.from_service_account_info(st.secrets["gcp_service_account"])
    # 2. ローカルファイル (Macでの開発用)
    else:
        try:
            creds = service_account.Credentials.from_service_account_file("googledrive_key.json")
        except FileNotFoundError:
            return None
    
    return build('drive', 'v3', credentials=creds)

# サービスを使える状態にする
drive_service = get_gdrive_service()
GDRIVE_FOLDER_ID = "1v5wXyLbX85AiYwCAcQCk0x3ID09bGQAO"
GDRIVE_FOLDER_ID = "1v5wXyLbX85AiYwCAcQcKOx3ID09bGQAP" # ★ここにご自身のフォルダIDを入力してください

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

@st.cache_data(ttl=60)
def get_notion_total_count(token, db_id):
    if not token or not db_id:
        return 0
    count = 0
    has_more = True
    next_cursor = None
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }
    url = f"https://api.notion.com/v1/databases/{db_id}/query"
    
    while has_more:
        payload = {"page_size": 100}
        if next_cursor:
            payload["start_cursor"] = next_cursor
        try:
            res = requests.post(url, headers=headers, json=payload)
            if res.status_code == 200:
                data = res.json()
                count += len(data.get("results", []))
                has_more = data.get("has_more", False)
                next_cursor = data.get("next_cursor", None)
            else:
                break
        except Exception:
            break
    return count

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
        payload["sorts"] = [{"timestamp": "last_edited_time", "direction": "descending"}]

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

# --- NotionのマニュアルDBからメタデータを取得 ---
@st.cache_data(ttl=60)
def get_notion_manual_metadata(token, db_id):
    """マニュアル管理用Notion DBから、タイトルと種類のリストを取得する"""
    if not token or not db_id: return []
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }
    url = f"https://api.notion.com/v1/databases/{db_id}/query"
    manuals_meta = []
    
    has_more = True
    next_cursor = None
    
    while has_more:
        payload = {"page_size": 100}
        if next_cursor: payload["start_cursor"] = next_cursor
        try:
            res = requests.post(url, headers=headers, json=payload)
            if res.status_code == 200:
                data = res.json()
                for item in data.get("results", []):
                    props = item.get("properties", {})
                    
                    # 内容（タイトル）の取得
                    title_prop = props.get("内容", {}).get("title", [])
                    name = "".join([t["plain_text"] for t in title_prop]) if title_prop else "名称未設定"
                    
                    # 種類（セレクトプロパティ等を想定）の取得
                    category = "未分類"
                    cat_prop = props.get("種類", {})
                    if "select" in cat_prop and cat_prop["select"]:
                        category = cat_prop["select"].get("name", "未分類")
                    elif "multi_select" in cat_prop and cat_prop["multi_select"]:
                        category = cat_prop["multi_select"][0].get("name", "未分類") # 複数ある場合は1つ目を採用
                    elif "rich_text" in cat_prop and cat_prop["rich_text"]:
                        category = cat_prop["rich_text"][0]["plain_text"]
                        
                    manuals_meta.append({"name": name, "category": category})
                
                has_more = data.get("has_more", False)
                next_cursor = data.get("next_cursor", None)
            else:
                break
        except Exception:
            break
            
    return manuals_meta

def add_to_notion(name, definition):
    if not NOTION_TOKEN or not DATABASE_ID: return False, "Notion設定が不足しています"
    name = name.strip()
    definition = definition.strip()

    existing_items = get_notion_data(name, mode="名称のみ")
    target_page_id = None
    existing_meaning = ""
    found_name = ""

    for item in existing_items:
        try:
            props = item.get("properties", {})
            item_name = props.get("名称", {}).get("title", [])[0]["plain_text"]
            if item_name == name:
                target_page_id = item["id"]
                found_name = item_name
                break
        except Exception: continue

    if not target_page_id:
        for item in existing_items:
            try:
                props = item.get("properties", {})
                item_name = props.get("名称", {}).get("title", [])[0]["plain_text"]
                if item_name.startswith(f"{name} (") or f"({name})" in item_name:
                    target_page_id = item["id"]
                    found_name = item_name
                    break
            except Exception: continue

    if target_page_id:
        try:
            page_detail = st.session_state.notion.pages.retrieve(page_id=target_page_id)
            rich_text_arr = page_detail.get("properties", {}).get("意味", {}).get("rich_text", [])
            existing_meaning = "".join([t["plain_text"] for t in rich_text_arr])
        except Exception: pass

    headers = {"Authorization": f"Bearer {NOTION_TOKEN}", "Notion-Version": "2022-06-28", "Content-Type": "application/json"}

    if target_page_id:
        if definition in existing_meaning: return True, f"「{found_name}」は既に同じ意味が登録済のためスキップしました"
        if "・" not in existing_meaning and existing_meaning.strip(): updated_meaning = f"・{existing_meaning}\n・{definition}"
        else: updated_meaning = f"{existing_meaning}\n・{definition}"
        updated_meaning = updated_meaning[:2000]

        payload = {"properties": {"意味": {"rich_text": [{"text": {"content": updated_meaning}}]}}}
        try:
            if st.session_state.notion:
                st.session_state.notion.pages.update(page_id=target_page_id, **payload)
                return True, f"「{found_name}」に新たな意味を追記しました！"
        except Exception: pass
        return False, "Notion更新に失敗しました"
    else:
        payload = {
            "parent": {"database_id": DATABASE_ID},
            "properties": {
                "名称": {"title": [{"text": {"content": name}}]},
                "意味": {"rich_text": [{"text": {"content": definition[:2000]}}]}
            }
        }
        try:
            if st.session_state.notion:
                st.session_state.notion.pages.create(**payload)
                get_notion_total_count.clear() 
                return True, f"「{name}」を新規登録しました！"
        except Exception: pass
        return False, "Notion登録に失敗しました"

@st.cache_data(ttl=60)
def get_gdrive_manuals():
    if not os.path.exists(GOOGLE_SERVICE_ACCOUNT_JSON):
        st.error(f"認証ファイル {GOOGLE_SERVICE_ACCOUNT_JSON} が見つかりません。同じフォルダに配置してください。")
        return {}
    try:
        scopes = ['https://www.googleapis.com/auth/drive.readonly']
        creds = service_account.Credentials.from_service_account_file(GOOGLE_SERVICE_ACCOUNT_JSON, scopes=scopes)
        service = build('drive', 'v3', credentials=creds)

        query = f"'{GDRIVE_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.document' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        return {f['name']: f['id'] for f in files}
    except Exception as e:
        st.error(f"Googleドライブのデータ取得中にエラーが発生しました: {e}")
        return {}

# --- ページコンテンツ ---

def page_manual_creator():
    st.header("📝 AIマニュアル作成アシスタント V14")
    
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
                    1. 冒頭に必ず以下の3行のメタデータを出力してください。
                       局名：YTV
                       種類：[画像内容から推測したカテゴリ（例: マスター、ファイリング、事務、その他）]
                       タイトル：[作業名]
                    2. 空白行を1行あけて、手順の記述を開始してください：
                       手順1：[操作内容]
                       [画像1]
                       ...という形式で記述してください。
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
            if st.session_state['manual_text']:
                with st.spinner("作成中..."):
                    try:
                        response = openai.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": f"以下からチェックリストを作成せよ。\n\n{st.session_state['manual_text']}"}],
                        )
                        st.session_state['checklist_text'] = response.choices[0].message.content
                    except Exception as e: st.error(f"エラー: {e}")

    if st.session_state['manual_text']:
        lines = st.session_state['manual_text'].strip().split('\n')
        doc_title = "業務マニュアル"
        for line in lines[:5]:
            if "タイトル：" in line or "タイトル:" in line:
                doc_title = line.replace("タイトル：", "").replace("タイトル:", "").strip()
                break
        
        st.divider()
        st.markdown(f"# {doc_title}")
        st.write(st.session_state['manual_text'])

def page_manual_viewer():
    st.header("📚 マニュアル閲覧")
    
    if not MANUAL_DB_ID:
        st.error("環境変数 `NOTION_MANUAL_DB_ID` が設定されていません。")
        return

    with st.spinner("マニュアル一覧をNotionから取得中..."):
        notion_meta = get_notion_manual_metadata(NOTION_TOKEN, MANUAL_DB_ID)
        gdrive_manuals = get_gdrive_manuals()

    if not notion_meta:
        st.info("Notionの「仕事メモ」データベースにマニュアルが見つかりません。")
        return

    # --- 検索とフィルタリングUI ---
    st.markdown("### 🔍 マニュアルを探す")
    
    unique_categories = set([m['category'] for m in notion_meta])
    category_options = ["すべて"] + sorted(list(unique_categories))
    
    col_search, col_filter = st.columns([2, 1])
    with col_search:
        search_query = st.text_input("マニュアル名検索 (部分一致)", placeholder="例: 切符, マスター...")
    with col_filter:
        selected_category = st.selectbox("種類フィルタ", category_options)

    # ★変更点：検索ワードを全角・半角・大文字・小文字を統一（正規化）して比較する処理を追加
    filtered_manuals = []
    
    # 検索窓の文字をNFKC正規化＆小文字化
    if search_query:
        norm_query = unicodedata.normalize('NFKC', search_query).lower()
    else:
        norm_query = ""

    for m in notion_meta:
        # Notionから取得したマニュアル名もNFKC正規化＆小文字化して比較
        if norm_query:
            norm_name = unicodedata.normalize('NFKC', m['name']).lower()
            match_q = norm_query in norm_name
        else:
            match_q = True
            
        match_c = (selected_category == "すべて" or m['category'] == selected_category)
        
        if match_q and match_c:
            filtered_manuals.append(m)

    st.divider()

    # --- 結果表示UI ---
    if not filtered_manuals:
        st.warning("条件に一致するマニュアルが見つかりませんでした。")
        return

    col1, col2 = st.columns([3, 1])
    with col1:
        manual_names = [m['name'] for m in filtered_manuals]
        selected_manual_name = st.selectbox("閲覧するファイルを選択", ["選択してください"] + manual_names)
    with col2:
        edit_mode = st.toggle("✏️ 編集モード", value=False)

    if selected_manual_name != "選択してください":
        if selected_manual_name in gdrive_manuals:
            doc_id = gdrive_manuals[selected_manual_name]
            mode_suffix = "edit" if edit_mode else "preview"
            url = f"https://docs.google.com/document/d/{doc_id}/{mode_suffix}"
            
            st.success(f"Googleドライブから「{selected_manual_name}」を読み込みました。")
            components.iframe(url, height=900, scrolling=True)
        else:
            st.info(f"「{selected_manual_name}」はNotionに登録されていますが、Googleドライブ上に同名のドキュメントが見つかりませんでした。")
            st.caption("※Googleドライブにドキュメントを作成し、Notionの「内容」と完全一致するファイル名にするとここに表示されます。")

def page_glossary_search():
    st.header("🔍 用語検索")
    total_count = get_notion_total_count(NOTION_TOKEN, DATABASE_ID)
    st.caption(f"📚 現在の用語登録総数: **{total_count}** 件")
    
    col1, col2 = st.columns([3, 1])
    with col1: q = st.text_input("検索ワード", placeholder="キーワードを入力...", key="search_input")
    with col2: m = st.selectbox("検索対象", ["名称のみ", "全体（意味を含む）"], key="search_mode")
    st.divider()
    if not q:
        st.write("🕒 **直近の編集履歴（最新5件）**")
    with st.spinner("取得中..."): items = get_notion_data(q, m)
    if items:
        for item in items:
            try:
                props = item.get("properties", {})
                name = props.get("名称", {}).get("title", [])[0]["plain_text"]
                rich_text_arr = props.get("意味", {}).get("rich_text", [])
                definition = "".join([t["plain_text"] for t in rich_text_arr])
                with st.expander(f"📌 {name}"):
                    st.write(definition)
            except: continue

def page_glossary_registration():
    st.header("📥 用語の登録")
    tab1, tab2, tab3 = st.tabs(["手動入力", "一括登録", "PDF解析"])
    
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
                else: st.warning("名称と意味を入力してください。")

    with tab2:
        st.subheader("一括登録")
        st.write("「用語 > 意味」の形式で入力")
        bulk_text = st.text_area("貼り付けエリア", placeholder="API > アプリ間の窓口", height=300)
        
        if st.button("まとめて登録を実行"):
            if bulk_text:
                with st.spinner("一括登録中..."):
                    lines = bulk_text.strip().split("\n")
                    success_count = 0
                    for line in lines:
                        sep = ">" if ">" in line else "＞" if "＞" in line else None
                        if sep:
                            parts = line.split(sep, 1)
                            if len(parts) == 2:
                                ok, _ = add_to_notion(parts[0], parts[1])
                                if ok: success_count += 1
                    st.success(f"{success_count}件 登録・追記されました。")
            else: st.warning("登録するテキストを入力してください。")

    with tab3:
        st.subheader("PDFから用語を自動抽出・登録")
        if pdfplumber is None:
            st.error("⚠️ PDF解析用ライブラリ(pdfplumber)がインストールされていません。")
        else:
            uploaded_pdfs = st.file_uploader("PDFファイルをアップロード", type="pdf", accept_multiple_files=True)
            chunk_size = st.slider("1回あたりの解析ページ数", 1, 50, 40)
            
            if st.button("PDF解析と自動登録を開始"):
                if not uploaded_pdfs: st.warning("PDFファイルをアップロードしてください。")
                elif not raw_api_key: st.error("OpenAI API Keyが設定されていません。")
                else:
                    for pdf_file in uploaded_pdfs:
                        status_area = st.empty()
                        status_area.info(f"📄 {pdf_file.name} を読み込み中...")
                        try:
                            with pdfplumber.open(pdf_file) as pdf:
                                total_pages = len(pdf.pages)
                                total_success_count = 0
                                successful_term_names = [] 
                                progress_bar = st.progress(0)
            
                                for i in range(0, total_pages, chunk_size):
                                    end_page = min(i + chunk_size, total_pages)
                                    text_chunk = ""
                                    for page_num in range(i, end_page):
                                        page_text = pdf.pages[page_num].extract_text(layout=True)
                                        if page_text: text_chunk += page_text + "\n"
                                    
                                    if not text_chunk.strip(): continue
                                    
                                    with st.spinner(f"{i+1}〜{end_page}ページ目を解析中..."):
                                        prompt = f"""
                                        あなたはテレビ局の技術運用・マスター・ファイリング業務に精通した専門職員です。
                                        提供されたテキストから重要な「専門用語」を抽出し、その「意味・解説」をNotionの用語集として作成してください。
                                        【絶対条件】
                                        1. テレビ局の職員として、専門的かつ詳細に回答してください。
                                        2. その用語が「何であるか」「どのような役割を果たすか」を詳しく記述してください。
                                        3. 略語がある場合は、『略語（正式名称）』の形式に統一してください。
                                        【出力形式】
                                        {{
                                            "terms": [
                                                {{"名称": "用語名", "意味": "詳細な解説テキスト"}},
                                                ...
                                            ]
                                        }}
                                        ### 解析対象テキスト
                                        {text_chunk}
                                        """
                                        response = openai.chat.completions.create(
                                            model="gpt-4o",
                                            response_format={"type": "json_object"},
                                            messages=[{"role": "system", "content": "JSON形式で出力してください。"}, {"role": "user", "content": prompt}],
                                        )
                                        try:
                                            result = json.loads(response.choices[0].message.content)
                                            terms = result.get("terms", [])
                                            for t in terms:
                                                ok, msg = add_to_notion(t["名称"], t["意味"])
                                                if ok and "スキップ" not in msg: 
                                                    total_success_count += 1
                                                    successful_term_names.append(t["名称"]) 
                                        except Exception as e: st.warning(f"一部の解析フォーマットエラー: {e}")
                                    progress_bar.progress(end_page / total_pages)
                            status_area.success(f"✅ {pdf_file.name} の処理完了！ 合計 {total_success_count}件 の用語を登録・追記しました。")
                            if successful_term_names:
                                st.markdown("**📝 登録・追記された用語一覧:**")
                                st.write(", ".join(successful_term_names))
                        except Exception as e: st.error(f"エラーが発生しました ({pdf_file.name}): {e}")

# --- メインレイアウト ---
st.set_page_config(page_title="お仕事支援マルチツール", layout="wide")

with st.sidebar:
    st.title("🛠️ Menu")
    selection = st.radio("機能選択", ["AIマニュアル作成", "マニュアル閲覧", "用語検索", "用語登録"])
    st.divider()
    if NOTION_TOKEN and DATABASE_ID and MANUAL_DB_ID: st.success("Notion Connected")
    elif NOTION_TOKEN and DATABASE_ID: st.warning("Notion Partial (Manual DB Missing)")
    else: st.error("Notion Disconnected")
    if raw_api_key: st.success("OpenAI Ready")
    else: st.error("OpenAI API Key Missing")

if selection == "AIマニュアル作成": page_manual_creator()
elif selection == "マニュアル閲覧": page_manual_viewer()
elif selection == "用語検索": page_glossary_search()
elif selection == "用語登録": page_glossary_registration()