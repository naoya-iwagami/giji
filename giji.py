import os  
import openai  
import docx  
import streamlit as st  
  
# ページの設定をwideにする  
st.set_page_config(layout="wide")  
  
# 環境変数からAPIキーを取得  
api_key = os.getenv("OPENAI_API_KEY")  
if not api_key:  
    st.error("APIキーが設定されていません。環境変数OPENAI_API_KEYを設定してください。")  
    st.stop()  
  
# OpenAI APIの設定  
openai.api_type = "azure"  
openai.api_base = "https://test-chatgpt-pm-1.openai.azure.com/"  # ご自身のエンドポイントに合わせて変更してください  
openai.api_version = "2024-10-01-preview"  # エンジンに合わせてAPIバージョンを指定  
openai.api_key = api_key  
  
def read_docx(file):  
    try:  
        doc = docx.Document(file)  
        full_text = [para.text for para in doc.paragraphs]  
        return '\n'.join(full_text)  
    except Exception as e:  
        st.error(f"Error reading docx file: {e}")  
        return None  
  
def create_summary(full_text, system_message):  
    try:  
        # ユーザーメッセージを作成  
        user_message = f"{system_message}\n\n以下はテキストです。\n\n{full_text}"  
  
        messages = [  
            {"role": "user", "content": user_message}  
        ]  
  
        response = openai.ChatCompletion.create(  
            engine='o1-preview',  
            messages=messages,  
            max_completion_tokens=32768  # 最大トークン数を指定  
        )  
  
        final_summary = response["choices"][0]["message"]["content"]  
        return final_summary  
  
    except Exception as e:  
        st.error(f"Unexpected error in create_summary: {e}")  
        return None  
  
# StreamlitのUI設定  
st.title("議事録アプリ")  
st.write("docxファイルをアップロードして、その内容から議事録を作成します。")  
  
# ユーザーが指定できる system_message（デフォルトの文章を設定）  
default_system_message = (  
    "あなたはプロフェッショナルな議事録作成者です。"  
    "与えられたテキストを基に、できるだけ詳細かつ具体的な議事録を作成してください。"  
    "議事録には以下の項目を含めてください:\n"  
    "1. 会議開催日\n"  
    "2. 会議概要\n"  
    "3. 詳細な議論内容（具体的な発言や数値データを含む）\n"  
    "4. アクションアイテム\n"  
    "これら以外の項目は含めないでください。"  
    "重要なポイントや具体例を盛り込み、深みのある議事録を作成してください。"  
    "参加者の名前や発言者の情報は含めないでください。"  
    "全体で2000文字程度にまとめてください。"  
)  
  
system_message = st.text_area("システムメッセージを入力してください:", default_system_message, height=300)  
  
# 複数ファイルのアップロード  
uploaded_files = st.file_uploader(  
    "ファイルをアップロード (最大3つ)",  
    type="docx",  
    accept_multiple_files=True  
)  
  
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
        st.write("アップロードされたファイルの文字数:")  
        for file_name, char_count in file_char_counts:  
            st.write(f"- {file_name}: {char_count}文字")  
        st.write(f"**合計文字数: {total_char_count}文字**")  
  
        # 議事録作成のボタンを配置  
        if st.button('議事録を作成'):  
            if all_texts:  
                # 議事録を作成  
                summary = create_summary(all_texts, system_message)  
  
                if summary is None:  
                    st.error("議事録の生成に失敗しました。")  
                    st.stop()  
  
                st.header("作成された議事録")  
                st.write(summary)  
  
            else:  
                st.error("ファイルの読み込みに失敗しました。")  