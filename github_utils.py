from io import BytesIO
import base64
import requests
import streamlit as st
import pandas as pd
from urllib.parse import quote

# GitHub 配置
GITHUB_TOKEN_KEY = "GITHUB_TOKEN"  # secrets.toml 中的密钥名
REPO_NAME = "TTTriste06/semiexcel"
BRANCH = "main"

FILE_RENAME_MAPPING = {
    "赛卓-新旧料号.xlsx": "mapping_file.xlsx",
    "赛卓-安全库存.xlsx": "safety_file.xlsx",
    "赛卓-预测.xlsx": "pred_file.xlsx"
}

def upload_to_github(file_obj, filename):
    """
    将 file_obj 文件上传至 GitHub 指定仓库
    - file_obj: BytesIO 或类文件对象
    - filename: 仓库中要保存的文件名（含扩展名）
    """
    token = st.secrets[GITHUB_TOKEN_KEY]
    safe_filename = quote(filename)  # 编码支持中文

    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{safe_filename}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 将文件读取并转为 base64
    file_obj.seek(0)
    content = base64.b64encode(file_obj.read()).decode("utf-8")
    file_obj.seek(0)

    # 检查是否已存在（需要获取 SHA）
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
    else:
        print(f"✅ 成功上传文件至 GitHub：{filename}")

def download_from_github(filename):
    """
    从 GitHub 仓库下载指定文件内容（以二进制返回）
    - filename: 仓库中保存的文件名
    - 返回: bytes 内容（可用于 pd.read_excel(BytesIO(...))）
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

def load_or_fallback_from_github(label: str, key: str, filename: str, additional_sheets: dict):
    """优先加载上传文件，否则从 GitHub 加载 fallback 文件"""
    uploaded_file = st.file_uploader(f"📎 上传 {label} 文件", type=["xlsx"], key=key)

    github_filename = FILE_RENAME_MAPPING.get(filename, filename)

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            additional_sheets[filename] = df
            upload_to_github(uploaded_file, github_filename)  # ✅ 上传到 GitHub 使用英文名
            st.success(f"✅ 已上传并缓存：{filename}")
        except Exception as e:
            st.error(f"❌ 解析上传文件失败：{filename} - {e}")
    else:
        try:
            content = download_from_github(github_filename)  # ✅ 下载 GitHub 使用英文名
            if content:
                df = pd.read_excel(BytesIO(content))
                additional_sheets[filename] = df
                st.info(f"ℹ️ 已从 GitHub 加载历史文件：{filename}")
            else:
                st.warning(f"⚠️ 未提供且未在 GitHub 找到历史文件：{filename}")
        except Exception as e:
            st.error(f"❌ 从 GitHub 加载 {filename} 失败: {e}")
