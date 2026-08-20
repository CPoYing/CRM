"""
把 Google 服務帳號的 JSON 金鑰轉成 Streamlit 要的 secrets TOML 格式。

手動複製貼上最容易在 private_key 的換行上出錯，用這支就不會。

用法（在 project_tracker_app 資料夾底下執行）：

    # 只印出來，自己複製去貼到 Streamlit Cloud 的 Secrets
    python make_secrets.py "C:\\Users\\你\\Downloads\\xxxx-123456.json" "試算表網址"

    # 順便寫一份到本機 .streamlit/secrets.toml（本機也要連同一張試算表時用）
    python make_secrets.py "金鑰.json" "試算表網址" --write

    # Windows 直接複製到剪貼簿
    python make_secrets.py "金鑰.json" "試算表網址" | clip

注意：印出來的內容含私鑰，等同密碼。不要貼到聊天室、issue、或任何公開的地方。
"""

import json
import sys
from pathlib import Path

FIELDS = [
    "type", "project_id", "private_key_id", "private_key", "client_email",
    "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url",
    "client_x509_cert_url", "universe_domain",
]


def build_toml(key_data, spreadsheet_url):
    lines = ["[gsheets]", f"spreadsheet = {json.dumps(spreadsheet_url)}", "", "[gcp_service_account]"]
    for field in FIELDS:
        if field in key_data:
            # json.dumps 產生的跳脫字元（\n、\"）剛好符合 TOML 基本字串的規則
            lines.append(f"{field} = {json.dumps(key_data[field], ensure_ascii=False)}")
    return "\n".join(lines) + "\n"


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 1

    key_path = Path(argv[1])
    spreadsheet_url = argv[2].strip()
    write = "--write" in argv[3:]

    if not key_path.exists():
        print(f"❌ 找不到金鑰檔：{key_path}")
        return 1

    try:
        key_data = json.loads(key_path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"❌ 這不是有效的 JSON 檔：{exc}")
        return 1

    missing = [f for f in ("private_key", "client_email", "project_id") if f not in key_data]
    if missing:
        print(f"❌ JSON 裡少了 {missing}，這可能不是服務帳號金鑰檔")
        return 1
    if key_data.get("type") != "service_account":
        print(f"⚠️  type 是 {key_data.get('type')!r}，不是 service_account，可能抓錯檔案")

    if "docs.google.com" not in spreadsheet_url:
        print(f"⚠️  試算表網址看起來怪怪的：{spreadsheet_url}")

    toml_text = build_toml(key_data, spreadsheet_url)

    if write:
        target = Path(__file__).parent / ".streamlit" / "secrets.toml"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(toml_text, encoding="utf-8")
        print(f"✅ 已寫入 {target}")
        print("   （這個路徑已在 .gitignore，不會被推上 GitHub）")
        print()

    print("=" * 60)
    print("以下整段複製，貼到 Streamlit Cloud 的 Settings → Secrets：")
    print("=" * 60)
    print(toml_text)
    print("=" * 60)
    print(f"⚠️  記得把試算表『共用』給這個帳號，權限選「編輯者」：")
    print(f"    {key_data['client_email']}")
    print("⚠️  以上內容含私鑰，等同密碼，不要貼到公開的地方。")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
