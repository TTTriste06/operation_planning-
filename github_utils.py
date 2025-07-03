from io import BytesIO
import base64
import requests
import streamlit as st
import pandas as pd
from urllib.parse import quote

# GitHub 配置
GITHUB_TOKEN_KEY = "GITHUB_TOKEN"  # secrets.toml 中的密钥名
REPO_NAME = "TTTriste06/operation_planning-"
BRANCH = "main"

FILENAME_KEYS = {
    "forecast": "预测.xlsx",
    "safety": "安全库存.xlsx",
    "mapping": "新旧料号.xlsx",
    "supplier": "供应商-PC.xlsx",
    "arrival": "到货明细.xlsx",
    "order": "下单明细.xlsx",
    "sales": "销货明细.xlsx",
    "pc": "供应商-PC.xlsx",
}

def upload_to_github(file_obj, filename):
    """
    将 file_obj 文件上传至 GitHub 指定仓库
    """
    token = st.secrets[GITHUB_TOKEN_KEY]
    safe_filename = quote(filename)  # 支持中文

    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{safe_filename}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    file_obj.seek(0)
    content = base64.b64encode(file_obj.read()).decode("utf-8")
    file_obj.seek(0)

    # 检查是否已存在
    sha = None
    get_resp = requests.get(url, headers=headers)
    if get_resp.status_code == 200:
        sha = get_resp.json().get("sha")

    payload = {
        "message": f"upload {filename}",
        "content": content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha
        
    put_resp = requests.put(url, headers=headers, json=payload)
    if put_resp.status_code not in [200, 201]:
        raise Exception(f"❌ 上传失败：{put_resp.status_code} - {put_resp.text}")


def download_from_github(filename):
    """
    从 GitHub 下载文件内容（二进制返回）
    """
    token = st.secrets[GITHUB_TOKEN_KEY]
    safe_filename = quote(filename)

    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{safe_filename}?ref={BRANCH}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        json_resp = response.json()
        return base64.b64decode(json_resp["content"])
    else:
        raise FileNotFoundError(f"❌ GitHub 上找不到文件：{filename} (HTTP {response.status_code})")


def load_file_with_github_fallback(key, uploaded_file):
    """
    加载上传文件或从 GitHub 下载。如果上传了文件，就保存至 GitHub 并返回 DataFrame；
    否则尝试从 GitHub 下载。若失败返回空 DataFrame。

    参数:
        key: str — 文件类型，如 "forecast"
        uploaded_file: 上传文件对象或 None

    返回:
        pd.DataFrame
    """
    filename = FILENAME_KEYS.get(key)
    if not filename:
        st.warning(f"⚠️ 未识别的辅助文件类型：{key}")
        return pd.DataFrame()

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        file_io = BytesIO(file_bytes)
        try:
            upload_to_github(BytesIO(file_bytes), filename)
            st.success(f"✅ 使用上传文件并保存到 GitHub：{filename}")
        except Exception as e:
            st.warning(f"⚠️ 上传失败：{e}")
        return pd.read_excel(file_io)  # ✅ 强制读取 Sheet1

    else:
        try:
            content = download_from_github(filename)
            return pd.read_excel(BytesIO(content))  # ✅ 强制读取 Sheet1
        except FileNotFoundError as e:
            st.warning(str(e))
            return pd.DataFrame()
