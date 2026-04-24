import streamlit as st
import openai
import os
from dotenv import load_dotenv
import base64
from docx import Document 
from docx.shared import Inches, Pt
from io import BytesIO
import re
from datetime import datetime
import cv2
import tempfile

# --- 設定ファイルの読み込み ---
env_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(env_path):
    with open(env_path, encoding='utf-8') as f:
        load_dotenv(stream=f, override=True)

raw_api_key = os.getenv("OPENAI_API_KEY")
if raw_api_key:
    openai.api_key = raw_api_key.strip().strip('"').strip("'")

st.set_page_config(page_title="AIマニュアル作成アシスタント V13", layout="wide")
st.title("🚀 AIマニュアル作成アシスタント V13")

# セッション状態の初期化
if 'manual_text' not in st.session_state:
    st.session_state['manual_text'] = ""
if 'checklist_text' not in st.session_state:
    st.session_state['checklist_text'] = ""
if 'source_files' not in st.session_state:
    st.session_state['source_files'] = []
if 'processed_images_bytes' not in st.session_state:
    st.session_state['processed_images_bytes'] = []

def encode_image(image_bytes):
    return base64.b64encode(image_bytes).decode('utf-8')

# --- 画面構成 ---
# 【変更】動画ファイル（mp4, mov）も許可するように変更
uploaded_files = st.file_uploader("写真または動画をアップロード（順番通りに）", 
                                  type=['png', 'jpg', 'jpeg', 'mp4', 'mov'], 
                                  accept_multiple_files=True)

if uploaded_files != st.session_state['source_files']:
    st.session_state['manual_text'] = ""
    st.session_state['checklist_text'] = ""
    st.session_state['source_files'] = uploaded_files
    st.session_state['processed_images_bytes'] = [] # リセット

if uploaded_files:
    has_video = False
    temp_processed_bytes = []
    
    # ファイルの処理（画像はそのまま、動画は3秒ごとに抽出）
    if not st.session_state['processed_images_bytes']:
        with st.spinner("ファイルを読み込み・解析中...（動画の場合は数秒かかります）"):
            for file in uploaded_files:
                file_ext = file.name.split('.')[-1].lower()
                
                if file_ext in ['mp4', 'mov']:
                    has_video = True
                    st.video(file)
                    st.info(f"🎥 動画「{file.name}」から3秒ごとに画像を自動抽出しています...")
                    
                    # 一時ファイルとして動画を保存（OpenCVで読み込むため）
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{file_ext}') as tfile:
                        tfile.write(file.read())
                        temp_path = tfile.name

                    # OpenCVで動画読み込み
                    cap = cv2.VideoCapture(temp_path)
                    fps = cap.get(cv2.CAP_PROP_FPS)
                    if fps == 0: fps = 30 # フォールバック
                    frame_interval = int(fps * 3) # 3秒ごとのフレーム数
                    
                    count = 0
                    while cap.isOpened():
                        ret, frame = cap.read()
                        if not ret:
                            break
                        # 3秒ごとのフレームを保存
                        if count % frame_interval == 0:
                            _, buffer = cv2.imencode('.jpg', frame)
                            temp_processed_bytes.append(buffer.tobytes())
                        count += 1
                        
                    cap.release()
                    os.remove(temp_path) # 一時ファイルの削除
                else:
                    # 写真の場合はそのまま追加
                    temp_processed_bytes.append(file.getvalue())
            
            st.session_state['processed_images_bytes'] = temp_processed_bytes

    # 抽出・アップロードされたすべての画像をプレビュー表示
    if st.session_state['processed_images_bytes']:
        st.write("### 📸 AIが解析する画像リスト")
        col_pre = st.columns(min(len(st.session_state['processed_images_bytes']), 5))
        for idx, img_bytes in enumerate(st.session_state['processed_images_bytes']):
            with col_pre[idx % 5]:
                st.image(img_bytes, caption=f"画像{idx+1}", width=150)

    # ベース64エンコード（AI送信用）
    all_images_base64 = [encode_image(b) for b in st.session_state['processed_images_bytes']]

    st.markdown("---")
    col_btn1, col_btn2 = st.columns([1, 1])
    
    with col_btn1:
        if st.button("マニュアルを新規生成する", use_container_width=True):
            # トークン制限の簡易チェック（画像が多すぎる場合の警告）
            if len(all_images_base64) > 30:
                st.warning("⚠️ 抽出された画像が30枚を超えています。AIの処理制限に引っかかる可能性があるため、少し時間がかかるかエラーになる場合があります。")

            with st.spinner("AIがマニュアルを作成中..."):
                # 動画が含まれていたかどうかでAIへの指示を自動分岐
                has_video_in_session = any(f.name.split('.')[-1].lower() in ['mp4', 'mov'] for f in uploaded_files)
                
                if has_video_in_session:
                    optimization_instruction = """
                    【重要：動画からの抽出画像の最適化】
                    提供されている画像群には、動画から一定間隔で自動抽出された連続写真が含まれています。視覚的に明確な変化がない（ほぼ同じ画面である）画像は、マニュアルの手順文や画像タグから自動的に除外し、代表的な1枚だけを採用しなさい。マニュアル全体をスッキリさせ、似たような画像が並ばないように配置しなさい。
                    """
                else:
                    optimization_instruction = """
                    【重要：画像の完全使用（写真モード）】
                    提供されている画像はユーザーが意図して選定した写真です。似たような画像であっても勝手に省略・除外せず、必ずすべての画像（[画像1], [画像2]...）を適切な手順の箇所に配置しなさい。
                    """

                content_payload = [{
                    "type": "text", 
                    "text": f"""
                    提供された画像群から、業務マニュアルを作成してください。以下のルールを厳守してください：
                    1. 1行目は「タイトル：[システム名や作業名]」としてください（判断不能なら「タイトル：不明」）。
                    2. 手順は意味のある大きなまとまり（「1. システムの起動」「2. データチェック」など）でグループ化し、その中に具体的な手順を箇条書きで記載してください。
                    3. グループ化された手順の最後に、そのステップで参照すべきすべての画像タグをまとめて記載してください。例：[画像1][画像2]
                    
                    {optimization_instruction}

                    4. 手書き文字は【重要】として反映してください。
                    """
                }]
                
                for img in all_images_base64:
                    content_payload.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img}"}})

                try:
                    response = openai.chat.completions.create(
                        model="gpt-4o",
                        messages=[{"role": "user", "content": content_payload}],
                        max_tokens=2000,
                    )
                    st.session_state['manual_text'] = response.choices[0].message.content
                except Exception as e:
                    st.error(f"エラーが発生しました（画像が多すぎるか、通信エラーの可能性があります）: {e}")

    with col_btn2:
        if st.button("確認用チェックリストを作成", use_container_width=True):
            if not st.session_state['manual_text']:
                st.warning("先にマニュアルを生成してください。")
            else:
                with st.spinner("チェックリストを分析中..."):
                    content_payload = [{
                        "type": "text", 
                        "text": f"""
                        以下のマニュアル内容に基づき、作業者がミスをしないための「確認用チェックリスト」を作成してください。
                        
                        【マニュアル内容】
                        {st.session_state['manual_text']}
                        
                        【ルール】
                        ・箇条書き形式で、確認すべき重要項目を抽出してください。
                        ・印刷してペンでチェックできるよう、各項目の先頭は必ず「□ 」（四角記号と半角スペース）にしてください。Wordの箇条書き機能は不要です。
                        """
                    }]
                    try:
                        response = openai.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": content_payload}],
                            max_tokens=1000,
                        )
                        st.session_state['checklist_text'] = response.choices[0].message.content
                    except Exception as e:
                        st.error(f"エラー: {e}")

# --- 結果表示とWord保存 ---
if st.session_state['manual_text']:
    lines = st.session_state['manual_text'].strip().split('\n')
    first_line = lines[0].strip()
    
    title_name = first_line.replace("タイトル：", "").replace("タイトル:", "").replace('業務マニュアル：', '').strip()
    today_str = datetime.now().strftime("%Y%m%d")
    if not title_name or title_name == "不明" or len(title_name) > 30:
        doc_title = f"{today_str}_業務マニュアル"
    else:
        doc_title = title_name

    safe_file_name = re.sub(r'[\\/:*?"<>|]', '', doc_title)

    col_res1, col_res2 = st.columns([3, 1])
    with col_res1:
        st.success("作成が完了しました。")
    with col_res2:
        # Word作成
        doc = Document()
        doc.add_heading(doc_title, 0)
        
        for i in range(1, len(lines)):
            line = lines[i]
            matches = re.findall(r'\[画像(\d+)\]', line)
            clean_line = re.sub(r'\[画像\d+\]', '', line).replace('**', '').replace('#', '').strip()
            
            if clean_line: 
                doc.add_paragraph(clean_line)
            
            # Wordファイルへの画像挿入ロジック（抽出された全画像リストから取得）
            if matches:
                for match in matches:
                    img_idx = int(match) - 1
                    if 0 <= img_idx < len(st.session_state['processed_images_bytes']):
                        doc.add_picture(BytesIO(st.session_state['processed_images_bytes'][img_idx]), width=Inches(3.0))

        # --- チェックリスト追加 ---
        if st.session_state['checklist_text']:
            doc.add_page_break()
            doc.add_heading(f"{doc_title}：チェックリスト", level=1)
            
            doc.add_paragraph("----------------------------------------------------------------------------------------------------")
            
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            cell = table.cell(0, 0)
            p = cell.paragraphs[0]
            p.alignment = 1

            run = p.add_run("確認用チェックリスト")
            run.bold = True
            run.font.size = Pt(12)
            
            doc.add_paragraph("")

            check_lines = st.session_state['checklist_text'].split('\n')
            for cl in check_lines:
                clean_cl = cl.replace('**', '').replace('#', '').strip()
                if clean_cl: 
                    if not clean_cl.startswith('□'):
                        clean_cl = f"□ {clean_cl}"
                    doc.add_paragraph(clean_cl)

        bio = BytesIO()
        doc.save(bio)
        st.download_button(label="📥 Wordファイルで保存", data=bio.getvalue(), file_name=f"{safe_file_name}.docx")

    # プレビュー表示
    st.markdown("---")
    st.markdown(f"# {doc_title}")
    for i in range(1, len(lines)):
        line = lines[i]
        matches = re.findall(r'\[画像(\d+)\]', line)
        preview_text = re.sub(r'\[画像\d+\]', '', line).strip()
        
        if preview_text: 
            st.markdown(preview_text)
            
        # プレビューへの画像挿入ロジック
        if matches:
            img_cols = st.columns(len(matches))
            for idx, match in enumerate(matches):
                img_idx = int(match) - 1
                if 0 <= img_idx < len(st.session_state['processed_images_bytes']):
                    with img_cols[idx]:
                        st.image(st.session_state['processed_images_bytes'][img_idx], width=300)
    
    # チェックリストのプレビュー
    if st.session_state['checklist_text']:
        st.markdown("---")
        st.markdown(f"## {doc_title}：チェックリスト")
        st.markdown("<hr style='border: 1px solid #ccc;'>", unsafe_allow_html=True)

        st.markdown("""
        <div style='border: 1px solid #ccc; padding: 5px 10px; margin-bottom: 15px; width: fit-content; font-weight: bold; text-align: center; display: inline-block;'>
            確認用チェックリスト
        </div>
        """, unsafe_allow_html=True)
        
        check_lines = st.session_state['checklist_text'].split('\n')
        preview_checklist = ""
        for cl in check_lines:
            clean_cl = cl.replace('**', '').replace('#', '').strip()
            if clean_cl:
                if not clean_cl.startswith('□'):
                    clean_cl = f"□ {clean_cl}"
                preview_checklist += f"{clean_cl}\n\n"
        
        st.markdown(preview_checklist)