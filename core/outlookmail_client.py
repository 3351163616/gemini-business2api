import time
from datetime import datetime
from typing import Any, Dict, List, Optional

import requests

from core.mail_utils import extract_verification_code


class OutlookMailClient:
    """OutlookMail 邮箱管理 API 客户端

    基于已有的 Outlook 账号池获取邮件，用于接收验证码。
    """

    # 类级别索引，用于按顺序选取账号
    _account_index: int = 0
    _accounts_cache: List[Dict[str, Any]] = []

    def __init__(
        self,
        base_url: str = "http://your-outlook-api.com",
        proxy: str = "",
        log_callback=None,
    ) -> None:
        self.base_url = (base_url or "http://your-outlook-api.com").rstrip("/")
        self.proxy_url = (proxy or "").strip()
        self.log_callback = log_callback

        self.email: Optional[str] = None
        self.account_id: Optional[int] = None

    def set_credentials(self, email: str, password: Optional[str] = None) -> None:
        """设置邮箱地址（用于已有账号）"""
        self.email = email

    def _log(self, level: str, message: str) -> None:
        if self.log_callback:
            try:
                self.log_callback(level, message)
            except Exception:
                pass

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        self._log("info", f"📤 发送 {method} 请求: {url}")

        proxies = {"http": self.proxy_url, "https": self.proxy_url} if self.proxy_url else None

        res = requests.request(
            method,
            url,
            proxies=proxies,
            timeout=kwargs.pop("timeout", 120),  # 超时改为 120 秒
            **kwargs,
        )
        self._log("info", f"📥 收到响应: HTTP {res.status_code}")
        return res

    def _fetch_accounts(self) -> List[Dict[str, Any]]:
        """获取所有可用账号"""
        all_accounts = []
        page = 1
        page_size = 100

        while True:
            url = f"{self.base_url}/accounts?page={page}&page_size={page_size}"
            try:
                res = self._request("GET", url)
                if res.status_code != 200:
                    self._log("error", f"❌ 获取账号列表失败: HTTP {res.status_code}")
                    break
                accounts = res.json() if res.content else []
                if not accounts:
                    break
                all_accounts.extend(accounts)
                if len(accounts) < page_size:
                    break
                page += 1
            except Exception as e:
                self._log("error", f"❌ 获取账号列表异常: {e}")
                break

        return all_accounts

    def register_account(self, domain: Optional[str] = None) -> bool:
        """从账号池中按顺序选取下一个邮箱"""
        self._log("info", "🔍 正在从 OutlookMail 账号池获取邮箱...")

        # 刷新账号缓存
        OutlookMailClient._accounts_cache = self._fetch_accounts()

        if not OutlookMailClient._accounts_cache:
            self._log("error", "❌ OutlookMail 账号池为空")
            return False

        total = len(OutlookMailClient._accounts_cache)
        self._log("info", f"📋 账号池共 {total} 个账号")

        # 按顺序选取，循环使用
        index = OutlookMailClient._account_index % total
        account = OutlookMailClient._accounts_cache[index]

        self.email = account.get("email")
        self.account_id = account.get("id")

        # 递增索引供下次使用
        OutlookMailClient._account_index = (index + 1) % total

        self._log("info", f"✅ 选取账号 #{index + 1}/{total}: {self.email}")
        return True

    def _list_emails(self, email: str) -> List[Dict[str, Any]]:
        """获取指定邮箱的邮件列表"""
        url = f"{self.base_url}/emails/by-email/{email}?folder=inbox"
        try:
            res = self._request("GET", url)
            if res.status_code != 200:
                self._log("error", f"❌ 获取邮件列表失败: HTTP {res.status_code}")
                return []
            body = res.json() if res.content else {}
            if not body.get("success"):
                self._log("error", "❌ 获取邮件列表失败: success=false")
                return []
            return list(body.get("emails") or [])
        except Exception as e:
            self._log("error", f"❌ 获取邮件列表异常: {e}")
            return []

    def fetch_verification_code(self, since_time: Optional[datetime] = None) -> Optional[str]:
        """从邮件中提取验证码"""
        if not self.email:
            self._log("error", "❌ 邮箱地址未设置")
            return None

        try:
            self._log("info", f"📬 正在拉取 {self.email} 的邮件...")
            emails = self._list_emails(self.email)

            if not emails:
                self._log("info", "📭 邮箱为空，暂无邮件")
                return None

            # 按接收时间倒序排列（最新的在前）
            emails = sorted(
                emails,
                key=lambda x: x.get("received_time") or "",
                reverse=True
            )
            self._log("info", f"📨 收到 {len(emails)} 封邮件，开始检查验证码...")

            skipped = 0
            for idx, msg in enumerate(emails, 1):
                # 时间过滤
                if since_time:
                    received_time = msg.get("received_time")
                    if received_time:
                        try:
                            # 解析 ISO 格式时间 "2026-01-27T11:49:10Z"
                            msg_time = datetime.fromisoformat(
                                received_time.replace("Z", "+00:00")
                            ).astimezone().replace(tzinfo=None)
                            if msg_time < since_time:
                                skipped += 1
                                continue
                        except Exception:
                            pass

                subject = msg.get("subject") or ""
                self._log("info", f"🔍 检查邮件 {idx}: {subject[:50]}...")

                # 从 body 和 body_preview 中提取验证码
                content = (msg.get("body") or "") + (msg.get("body_preview") or "")
                code = extract_verification_code(content)
                if code:
                    self._log("info", f"✅ 找到验证码: {code}")
                    return code

            if skipped:
                self._log("info", f"⏭️ 跳过 {skipped} 封旧邮件")
            self._log("warning", "⚠️ 所有邮件中均未找到验证码")
            return None

        except Exception as e:
            self._log("error", f"❌ 获取验证码异常: {e}")
            return None

    def poll_for_code(
        self,
        timeout: int = 120,
        interval: int = 4,
        since_time: Optional[datetime] = None,
    ) -> Optional[str]:
        """轮询等待验证码"""
        if not self.email:
            self._log("error", "❌ 邮箱地址未设置")
            return None

        max_retries = max(1, timeout // interval)
        self._log("info", f"⏱️ 开始轮询验证码 (超时 {timeout}秒, 间隔 {interval}秒, 最多 {max_retries} 次)")

        for i in range(1, max_retries + 1):
            self._log("info", f"🔄 第 {i}/{max_retries} 次轮询...")
            code = self.fetch_verification_code(since_time=since_time)
            if code:
                self._log("info", f"🎉 验证码获取成功: {code}")
                return code
            if i < max_retries:
                time.sleep(interval)

        self._log("error", f"⏰ 验证码获取超时 ({timeout}秒)")
        return None
