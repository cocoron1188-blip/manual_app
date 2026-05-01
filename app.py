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

# PDF解析用ライブラリ
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

def add_to_notion(name, definition):
    """
    Notionに用語を登録する。既存データがある場合は意味を箇条書きで追記する。
    """
    if not NOTION_TOKEN or not DATABASE_ID:
        return False, "Notion設定が不足しています"
        
    name = name.strip()
    definition = definition.strip()

    # 1. 既存データの完全一致チェック
    existing_items = get_notion_data(name, mode="名称のみ")
    target_page_id = None
    existing_meaning = ""

    for item in existing_items:
        try:
            props = item.get("properties", {})
            item_name = props.get("名称", {}).get("title", [])[0]["plain_text"]
            if item_name == name:  # 完全一致した場合のみ対象とする
                target_page_id = item["id"]
                rich_text_arr = props.get("意味", {}).get("rich_text", [])
                existing_meaning = "".join([t["plain_text"] for t in rich_text_arr])
                break
        except Exception:
            continue

    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }

    # 2. 追記 または 新規登録 の処理
    if target_page_id:
        # --- 追記処理 ---
        # 既存テキストと全く同じ解説が来たらスキップ
        if definition in existing_meaning:
            return True, f"「{name}」は既に同じ意味が登録済のためスキップしました"

        # 既存の意味が箇条書きになっていなければ、箇条書きに整形して追記
        if "・" not in existing_meaning and existing_meaning.strip():
            updated_meaning = f"・{existing_meaning}\n・{definition}"
        else:
            updated_meaning = f"{existing_meaning}\n・{definition}"
            
        updated_meaning = updated_meaning[:2000] # 文字数制限対策

        payload = {
            "properties": {
                "意味": {"rich_text": [{"text": {"content": updated_meaning}}]}
            }
        }
        
        # ライブラリ経由でのアップデート
        try:
            if st.session_state.notion:
                st.session_state.notion.pages.update(page_id=target_page_id, **payload)
                return True, f"「{name}」に新たな意味を追記しました！"
        except Exception:
            pass

        # フォールバック（直接リクエスト）でのアップデート
        try:
            url = f"https://api.notion.com/v1/pages/{target_page_id}"
            res = requests.patch(url, headers=headers, data=json.dumps(payload))
            if res.status_code == 200:
                return True, f"「{name}」に新たな意味を追記しました！"
            else:
                return False, f"Notion追記エラー: {res.text}"
        except Exception as e:
            return False, f"Notion追記通信エラー: {str(e)}"

    else:
        # --- 新規登録処理 ---
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
                return True, f"「{name}」を新規登録しました！"
        except Exception:
            pass

        try:
            url = "https://api.notion.com/v1/pages"
            res = requests.post(url, headers=headers, data=json.dumps(payload))
            if res.status_code == 200:
                return True, f"「{name}」を新規登録しました！"
            else:
                return False, f"Notion登録エラー: {res.text}"
        except Exception as e:
            return False, f"Notion登録通信エラー: {str(e)}"

# --- ページコンテンツ ---

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
                    1. 1行目：必ず「タイトル：[作業名]」のみを出力してください。
                    2. 手順の記述：
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
        title_line = lines[0].strip()
        title_name = title_line.replace("タイトル：", "").replace("タイトル:", "").strip()
        doc_title = title_name if title_name else "業務マニュアル"
        
        st.divider()
        st.markdown(f"# {doc_title}")
        st.write(st.session_state['manual_text'])

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
        st.write("「用語 > 意味」の形式で入力（※改行で複数を一括登録可）")
        st.write("※区切り記号は「>」または「＞」が使用可能です。")
        bulk_text = st.text_area("貼り付けエリア", placeholder="API > アプリ間の窓口\nUI > ユーザーインターフェース", height=300)
        
        if st.button("まとめて登録を実行"):
            if bulk_text:
                # 修正：スピナー（動作中表示）を追加
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
                    # 修正：メッセージの表現を変更
                    st.success(f"{success_count}件 登録・追記されました。")
            else:
                st.warning("登録するテキストを入力してください。")

    with tab3:
        st.subheader("PDFから用語を自動抽出・登録")
        if PyPDF2 is None:
            st.error("⚠️ PDF解析用ライブラリ(PyPDF2)がインストールされていません。")
            return

        st.write("参考書や資料などのPDFをアップロードすると、AIが重要な用語を読み取ってNotionに自動登録します。")
        
        uploaded_pdfs = st.file_uploader("PDFファイルをアップロード（複数選択可）", type="pdf", accept_multiple_files=True)
        chunk_size = st.slider("1回あたりの解析ページ数", 1, 50, 40)
        
        if st.button("PDF解析と自動登録を開始"):
            if not uploaded_pdfs:
                st.warning("PDFファイルをアップロードしてください。")
            elif not raw_api_key:
                st.error("OpenAI API Keyが設定されていません。")
            else:
                for pdf_file in uploaded_pdfs:
                    status_area = st.empty()
                    status_area.info(f"📄 {pdf_file.name} を読み込み中...")
                    
                    try:
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        total_pages = len(pdf_reader.pages)
                        total_success_count = 0
                        successful_term_names = [] 
                        progress_bar = st.progress(0)
                        
                        for i in range(0, total_pages, chunk_size):
                            end_page = min(i + chunk_size, total_pages)
                            text_chunk = ""
                            for page_num in range(i, end_page):
                                text_chunk += pdf_reader.pages[page_num].extract_text() + "\n"
                            
                            if not text_chunk.strip(): continue
                            
                            with st.spinner(f"{i+1}〜{end_page}ページ目を解析中..."):
                                prompt = f"""
                                あなたはテレビ局の技術運用・マスター・ファイリング業務に精通した専門職員です。
                                提供されたテキストから重要な「専門用語」を抽出し、その「意味・解説」をNotionの用語集として作成してください。

                                【絶対条件】
                                1. テレビ局の職員として、専門的かつ詳細に回答してください。
                                2. 単なる1行程度の要約ではなく、その用語が「何であるか」に加えて「どのような役割を果たすか」「技術的にどのような背景があるか」を可能な限り詳しく記述してください。
                                3. 意味のところにわからない名称（略称）があった場合（例：DSやAPSなど）は、推測で展開せず、そのままの略称（DS、APS等）を用いて解説してください。
                                   （※前提知識：当環境においてDSはData Server、APSはAutomatic Program Control Systemを指すことが多いですが、AIが勝手に書き換えず、原文のニュアンスを維持してください）
                                4. 推測は厳禁ですが、提供されたテキスト内にある情報は余さず反映させてください。

                                【出力形式】
                                必ず以下のJSON形式で出力してください。
                                {{
                                    "terms": [
                                        {{"名称": "用語名", "意味": "詳細な解説テキスト（詳しく記述すること）"}},
                                        ...
                                    ]
                                }}

                                ### 解析対象テキスト
                                {text_chunk}
                                """
                                
                                response = openai.chat.completions.create(
                                    model="gpt-4o",
                                    response_format={"type": "json_object"},
                                    messages=[
                                        {"role": "system", "content": "JSON形式で出力してください。"},
                                        {"role": "user", "content": prompt}
                                    ],
                                )
                                
                                try:
                                    result = json.loads(response.choices[0].message.content)
                                    terms = result.get("terms", [])
                                    for t in terms:
                                        ok, msg = add_to_notion(t["名称"], t["意味"])
                                        if ok and "スキップ" not in msg: 
                                            total_success_count += 1
                                            successful_term_names.append(t["名称"]) 
                                except Exception as e:
                                    st.warning(f"一部の解析フォーマットエラー: {e}")
                            
                            progress_bar.progress(end_page / total_pages)
                        
                        status_area.success(f"✅ {pdf_file.name} の処理完了！ 合計 {total_success_count}件 の用語を登録・追記しました。")
                        
                        if successful_term_names:
                            st.markdown("**📝 登録・追記された用語一覧:**")
                            st.write(", ".join(successful_term_names))
                            
                    except Exception as e:
                        st.error(f"エラーが発生しました ({pdf_file.name}): {e}")
                st.balloons()

# --- メインレイアウト ---
st.set_page_config(page_title="お仕事支援マルチツール", layout="wide")

with st.sidebar:
    st.title("🛠️ Menu")
    selection = st.radio("機能選択", ["AIマニュアル作成", "用語検索", "用語登録"])
    
    st.divider()
    if NOTION_TOKEN and DATABASE_ID:
        st.success("Notion Connected")
    else:
        st.error("Notion Disconnected")
        
    if raw_api_key:
        st.success("OpenAI Ready")
    else:
        st.error("OpenAI API Key Missing")

if selection == "AIマニュアル作成": page_manual_creator()
elif selection == "用語検索": page_glossary_search()
elif selection == "用語登録": page_glossary_registration()