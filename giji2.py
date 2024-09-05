import os  
import openai  
import docx  
import streamlit as st  
import tempfile  
import pandas as pd

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
  
# 辞書ファイルのパスをプログラム内に記載
dictionary_path = "dictionary.csv"

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

def modify_text_with_dictionary(text, dictionary_df):  
    try:  
        # 辞書情報をリスト化し、ChatGPTに渡す準備
        dictionary_info = "\n".join([f"'{row['読み方']}' -> '{row['専門用語']}'" for _, row in dictionary_df.iterrows()])
        
        system_message = (
            "あなたはプロフェッショナルな文章修正家です。以下の辞書情報を使用して、文章中の誤記を修正し、適切な専門用語に置換してください。"
            "辞書情報には、'読み方' -> '専門用語'の形式で文字列が入っており、各ペアを改行で区切った一つの大きな文字列を作成しています。"
            "例えば、'せいまく' という読み方が '製膜' という専門用語に対応してます。この場合、成膜というワードが文章にあれば製膜に置き換えてほしい。"
            "また、プロとしてここはこういう意味ではと思うところがあればどんどん文章を修正してください"
        )
        
        user_message = (
            f"文章:\n{text}\n\n"
            f"辞書情報:\n{dictionary_info}\n\n"
            "文章中に誤記があれば、辞書情報を参考に正しい専門用語に修正してください。"
        )
        
        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  # 正しいエンジン名を指定  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": user_message}  
            ],  
            temperature=0.5,
            max_tokens=8000
        )
        
        modified_text = response["choices"][0]["message"]["content"]
        
        return modified_text
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
        return text

def create_summary(full_text, include_names):  
    try:  
        # 更新されたシステムメッセージ
        system_message = (
            "あなたはプロフェッショナルな議事録作成者です。以下の全文を基に、詳細な議事録を作成してください。"
            "議事録には以下の要素を必ず含めてください:\n"
            "会議概要:\n"
            "詳細な議論内容\n"
            "決定事項とその理由\n"
            "保留事項\n"
            "アクションアイテム\n\n"
            "また、可能であれば以下の要素も含めてください:\n"
            "議論のプロセスや意見の対立点\n"
            "重要な数値データや具体例\n\n"
            "できるだけ具体的かつ詳細に記述し、表面的な要約は避けてください。"
        )

        if include_names:  
            system_message += "参加者の名前や発言者が特定できる場合はそれも含めてください。"  
        else:  
            system_message += "参加者の名前や発言者の情報は含めないでください。"  

        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  # 正しいエンジン名を指定  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"全文:\n{full_text}"}  
            ],  
            temperature=0.5,
            max_tokens=8000
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
        return full_text

def hiroyuki_comments(summary):  
    try:  
        system_message = "あなたは論破王ひろゆきです。以下の議事録について皮肉や論理的な切り口で少し挑発的かつ軽い皮肉を交えて指摘を行ってください。また、話の流れが少しおかしい部分や矛盾がありそうな部分を鋭く指摘し相手を諭すようなトーンでコメントしてください。"  

        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  # 正しいエンジン名を指定  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"以下は、議事録です。各項目に指摘を入れてください。: {summary}"}  
            ],  
            temperature=0.5,
            max_tokens=8000
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
        system_message = "あなたはプロフェッショナルなコンサルタントです。以下の議事録を基にビジネスの改善点や次のステップについてアドバイスをしてください。"  

        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  # 正しいエンジン名を指定  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"以下は、議事録です。これを基にアドバイスをしてください。: {summary}"}  
            ],  
            temperature=0.5,
            max_tokens=8000
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

# 辞書ファイルを読み込み
if os.path.exists(dictionary_path):
    dictionary_df = pd.read_csv(dictionary_path)
else:
    st.error("指定された辞書ファイルが存在しません。")

# StreamlitのUI設定  
st.title("議事録アプリ")  
st.write("docxファイルをアップロードして、その内容から議事録を作成します。")  
  
# 参加者の名前や発言者情報を含めるかどうかのチェックボックス  
include_names = st.checkbox("参加者の名前や発言者の情報を含める", value=True)  

# ひろゆきの指摘を受けるかどうかのチェックボックス  
with_hiroyuki = st.checkbox("ひろゆきの指摘を受ける", value=False)  
  
# コンサルティングのアドバイスを受けるかどうかのチェックボックス  
with_consulting = st.checkbox("コンサルティングのアドバイスを受ける", value=False)  
  
# 複数ファイルのアップロード
uploaded_files = st.file_uploader("ファイルをアップロード (最大3つ)", type="docx", accept_multiple_files=True)  
  
if uploaded_files:  
    if len(uploaded_files) > 3:  
        st.error("最大3つのファイルまでアップロードできます。")  
    else:  
        # 複数のファイルの内容を結合  
        all_texts = ""  
        for uploaded_file in uploaded_files:  
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:  
                tmp_file.write(uploaded_file.read())  
                tmp_file_path = tmp_file.name  
            text = read_docx(tmp_file_path)  
            if text:  
                all_texts += text + "\n"  
  
        if all_texts:  
            # 語句の修正を実施
            modified_text = modify_text_with_dictionary(all_texts, dictionary_df)

            # 議事録を作成  
            summary = create_summary(modified_text, include_names)

            # レイアウトの決定
            if with_hiroyuki and with_consulting:  
                col1, col2, col3 = st.columns(3)  
            elif with_hiroyuki or with_consulting:  
                col1, col2 = st.columns(2)  
            else:  
                col1 = st.container()  

            with col1:  
                st.write(summary)

            if with_hiroyuki:  
                with col2 if 'col2' in locals() else col1:  
                    st.header("ひろゆきの指摘")  
                    comments = hiroyuki_comments(summary)  
                    if comments:  
                        st.write(comments)  
                    else:  
                        st.error("ひろゆきの指摘の生成に失敗しました。")  

            if with_consulting:  
                with col3 if 'col3' in locals() else col2:  
                    st.header("コンサルティングのアドバイス")  
                    advice = consulting_advice(summary)  
                    if advice:  
                        st.write(advice)  
                    else:  
                        st.error("コンサルティングのアドバイスの生成に失敗しました。")  

        else:  
            st.error("ファイルの読み込みに失敗しました。")  
