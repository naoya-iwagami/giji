import os  
import openai  
import docx  
import streamlit as st  
import tempfile  
  
# ページの設定をwideにする  
st.set_page_config(layout="wide")  
  
# 環境変数からAPIキーを取得  
api_key = os.getenv("OPENAI_API_KEY")  
if api_key is None:  
    st.error("APIキーが設定されていません。環境変数OPENAI_API_KEYを設定してください。")  
    st.stop()  
  
# OpenAI APIの設定  
openai.api_type = "azure"  
openai.api_base = "https://test-chatgpt-pm-1.openai.azure.com/"  
openai.api_version = "2023-03-15-preview"  
openai.api_key = api_key  
  
def read_docx(file_path):  
    try:  
        doc = docx.Document(file_path)  
        full_text = []  
        for para in doc.paragraphs:  
            full_text.append(para.text)  
        return '\n'.join(full_text)  
    except Exception as e:  
        st.error(f"Error reading docx file: {e}")  
        return None  
  
def summarize_text(text):  
    try:  
        system_message = "あなたはプロフェッショナルな議事録作成者です。会議の文字起こしを基に、重要な議論内容、結論、アクションアイテム（課題や次のステップ）を明確に整理し、議事録を作成してください。参加者の名前や発言者が特定できる場合はそれも含めてください。"  
        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  # 正しいエンジン名を指定  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"以下は、会議の文字起こしです。これを基に議事録を作成してください。: {text}"}  
            ],  
            temperature=0.5,  
            max_tokens=3000  
        )  
        return response["choices"][0]["message"]["content"]  
    except openai.error.AuthenticationError as e:  
        st.error(f"AuthenticationError: {e}")  
    except openai.error.APIConnectionError as e:  
        st.error(f"APIConnectionError: {e}")  
    except openai.error.InvalidRequestError as e:  
        st.error(f"InvalidRequestError: {e}")  
    except openai.error.OpenAIError as e:  
        st.error(f"OpenAIError: {e}")  
    except Exception as e:  
        st.error(f"Unexpected error: {e}")  
  
def hiroyuki_comments(summary):  
    try:  
        system_message = "あなたは論破王ひろゆきです。以下の議事録について、指摘やコメントをしてください。"  
        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  # 正しいエンジン名を指定  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"以下は、議事録です。これについて皮肉も交えながらひろゆきっぽく各項目に指摘を入れてください。: {summary}"}  
            ],  
            temperature=0.5,  
            max_tokens=3000  
        )  
        return response["choices"][0]["message"]["content"]  
    except openai.error.AuthenticationError as e:  
        st.error(f"AuthenticationError: {e}")  
    except openai.error.APIConnectionError as e:  
        st.error(f"APIConnectionError: {e}")  
    except openai.error.InvalidRequestError as e:  
        st.error(f"InvalidRequestError: {e}")  
    except openai.error.OpenAIError as e:  
        st.error(f"OpenAIError: {e}")  
    except Exception as e:  
        st.error(f"Unexpected error: {e}")  
  
def consulting_advice(summary):  
    try:  
        system_message = "あなたはプロフェッショナルなコンサルタントです。以下の議事録を基に、ビジネスの改善点や次のステップについてアドバイスをしてください。"  
        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  # 正しいエンジン名を指定  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"以下は、議事録です。これを基にアドバイスをしてください。: {summary}"}  
            ],  
            temperature=0.5,  
            max_tokens=3000  
        )  
        return response["choices"][0]["message"]["content"]  
    except openai.error.AuthenticationError as e:  
        st.error(f"AuthenticationError: {e}")  
    except openai.error.APIConnectionError as e:  
        st.error(f"APIConnectionError: {e}")  
    except openai.error.InvalidRequestError as e:  
        st.error(f"InvalidRequestError: {e}")  
    except openai.error.OpenAIError as e:  
        st.error(f"OpenAIError: {e}")  
    except Exception as e:  
        st.error(f"Unexpected error: {e}")  
  
# StreamlitのUI設定  
st.title("議事録アプリ")  
st.write("docxファイルをアップロードして、その内容から議事録を作成します。")  
  
# ひろゆきの指摘を受けるかどうかのチェックボックス  
with_hiroyuki = st.checkbox("ひろゆきの指摘を受ける", value=False)  
  
# コンサルティングのアドバイスを受けるかどうかのチェックボックス  
with_consulting = st.checkbox("コンサルティングのアドバイスを受ける", value=False)  
  
uploaded_file = st.file_uploader("ファイルをアップロード", type="docx")  
  
if uploaded_file is not None:  
    # 一時ファイルに保存  
    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:  
        tmp_file.write(uploaded_file.read())  
        tmp_file_path = tmp_file.name  
  
    # docxファイルの内容を読み込む  
    text = read_docx(tmp_file_path)  
  
    if text:  
        # テキストを要約  
        summary = summarize_text(text)  
        if summary:  
            col1, col2, col3 = st.columns(3)  
            with col1:  
                st.header("議事録")  
                st.write(summary)  
            if with_hiroyuki:  
                with col2:  
                    st.header("ひろゆきの指摘")  
                    comments = hiroyuki_comments(summary)  
                    if comments:  
                        st.write(comments)  
                    else:  
                        st.error("ひろゆきの指摘の生成に失敗しました。")  
            if with_consulting:  
                with col3:  
                    st.header("コンサルティングのアドバイス")  
                    advice = consulting_advice(summary)  
                    if advice:  
                        st.write(advice)  
                    else:  
                        st.error("コンサルティングのアドバイスの生成に失敗しました。")  
        else:  
            st.error("要約の生成に失敗しました。")  
    else:  
        st.error("ファイルの読み込みに失敗しました。")  
