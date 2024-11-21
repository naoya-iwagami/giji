import os  
import openai  
import docx  
import streamlit as st  

# ページの設定をwideにする  
st.set_page_config(layout="wide")  

# 環境変数からAPIキーとエンドポイントを取得  
api_key = os.getenv("AZURE_OPENAI_KEY")  
api_base = os.getenv("AZURE_OPENAI_ENDPOINT")

if api_key is None:  
    st.error("APIキーが設定されていません。環境変数OPENAI_API_KEYを設定してください。")  
    st.stop()

if api_base is None:
    st.error("APIエンドポイントが設定されていません。環境変数OPENAI_API_BASEを設定してください。")  
    st.stop()

# OpenAI APIの設定  
openai.api_type = "azure"  
openai.api_base = api_base  # 環境変数から取得したAPIエンドポイントを設定
openai.api_version = "2023-03-15-preview"  
openai.api_key = api_key

def read_docx(file):
    try:
        doc = docx.Document(file)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        st.error(f"Error reading docx file: {e}")
        return None

def create_summary(full_text, system_message, temperature):  
    try:  
        user_message = f"全文:\n{full_text}"

        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": user_message}  
            ],  
            temperature=temperature,  # 温度を設定
            max_tokens=16000  # 最大トークン数に合わせて調整
        )  
        summary = response["choices"][0]["message"]["content"]
        return summary
    except Exception as e:  
        st.error(f"Unexpected error in create_summary: {e}")  
        return full_text

# デフォルトのシステムメッセージ
default_system_message = (
    "あなたはプロフェッショナルな議事録作成者です。"
    "以下の全文を基にできるだけ詳細かつ具体的な議事録を作成してください。"
    "議事録には以下の項目を含めてください:\n"
    "1. 会議開催日\n"
    "2. 会議概要\n"
    "3. 詳細な議論内容（具体的な発言や数値データを含む）\n"
    "4. アクションアイテム\n"
    "これら以外の項目は含めないでください。"
    "重要なポイントや具体例を盛り込み、できるだけ箇条書きではなく文章で深みのある議事録を作成してください。"
)

# StreamlitのUI設定  
st.title("議事録アプリ")  
st.write("docxファイルをアップロードして、その内容から議事録を作成します。")  

# システムメッセージの入力欄
system_message = st.text_area(
    "システムメッセージを入力してください（デフォルトのメッセージが初期値として設定されています）:",
    value=default_system_message,
    height=250
)

# 温度のスライダー設定
temperature = st.slider(
    "生成する文章のランダム性（温度）を設定してください。0は厳密、1はクリエイティブに近い挙動を示します。",
    min_value=0.0, 
    max_value=1.0, 
    value=0.5, 
    step=0.1
)

# 複数ファイルのアップロード
uploaded_files = st.file_uploader("ファイルをアップロード (最大3つ)", type="docx", accept_multiple_files=True)  

if uploaded_files:  
    if len(uploaded_files) > 3:  
        st.error("最大3つのファイルまでアップロードできます。")  
    else:  
        total_char_count = 0
        file_char_counts = []
        all_texts = ""  # 全文テキストを保存する変数

        for uploaded_file in uploaded_files:
            text = read_docx(uploaded_file)
            if text:
                char_count = len(text)
                total_char_count += char_count
                file_char_counts.append((uploaded_file.name, char_count))
                all_texts += text + "\n"
            else:
                st.error(f"{uploaded_file.name} の読み込みに失敗しました。")

        # 各ファイルの文字数と合計文字数を表示
        st.write("アップロードされたファイルの文字数")
        for file_name, char_count in file_char_counts:
            st.write(f"- {file_name}: {char_count}文字")
        st.write(f"**合計文字数: {total_char_count}文字**")

        # 議事録作成のボタンを配置
        if st.button('議事録を作成'):
            if all_texts:
                # 議事録を作成  
                summary = create_summary(all_texts, system_message, temperature)
                st.write(summary)
            else:  
                st.error("ファイルの読み込みに失敗しました。")  
