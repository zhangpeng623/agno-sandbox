import requests
import json
from typing import Optional, List, Dict, Any
from agno.tools import Toolkit
from agno.tools.decorator import tool
from dotenv import load_dotenv
import os

load_dotenv()


class BaiduWebSearch(Toolkit):
    """百度联网搜索工具"""

    def __init__(self):
        super().__init__(name="baidu_web_search")
        self.api_url = os.getenv("BAIDU_SEARCH_API_URL")
        self.api_key = os.getenv("BAIDU_API_KEY")
        self.register(self.search)

    @tool
    def search(self, query: str, max_results: int = 10) -> str:
        """
        使用百度搜索实时信息

        Args:
            query: 要搜索的关键词
            max_results: 返回的最大结果数量（默认10）

        Returns:
            格式化的搜索结果字符串
        """
        if not self.api_key:
            return "错误：未配置百度API Key"

        if not self.api_url:
            return "错误：未配置百度API URL"

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }

        data = {
            "messages": [{"role": "user", "content": query}],
            "max_results": max_results
        }

        try:
            response = requests.post(
                self.api_url,
                headers=headers,
                json=data,
                timeout=30
            )

            if response.status_code == 200:
                return self._format_results(response.json(), query)
            else:
                return f"搜索失败：HTTP {response.status_code} - {response.text}"

        except requests.exceptions.Timeout:
            return "搜索超时，请稍后重试"
        except requests.exceptions.RequestException as e:
            return f"请求异常：{str(e)}"
        except Exception as e:
            return f"未知错误：{str(e)}"

    def _format_results(self, raw_result: Dict[str, Any], query: str) -> str:
        """格式化搜索结果"""
        try:
            # 解析百度API返回结构
            results = []

            # 根据不同返回结构提取数据
            if "result" in raw_result:
                items = raw_result.get("result", {}).get("items", [])
            elif "search_results" in raw_result:
                items = raw_result.get("search_results", [])
            elif "webPages" in raw_result:
                items = raw_result.get("webPages", {}).get("value", [])
            else:
                # 兼容其他返回格式
                items = raw_result.get("results", [])

            if not items:
                return f"未找到关于 '{query}' 的相关信息"

            for idx, item in enumerate(items[:10], 1):
                title = item.get("title") or item.get("name", "无标题")
                url = item.get("url") or item.get("link", "#")
                snippet = item.get("snippet") or item.get("content") or item.get("summary", "暂无摘要")

                results.append(f"{idx}. **{title}**\n   链接：{url}\n   摘要：{snippet}\n")

            return f"🔍 关于 '{query}' 的搜索结果：\n\n" + "\n".join(results)

        except Exception as e:
            return f"解析结果失败：{str(e)}\n原始数据：{json.dumps(raw_result, ensure_ascii=False)[:500]}"

