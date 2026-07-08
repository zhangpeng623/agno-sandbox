import os
import asyncio
from typing import Optional
from pathlib import Path
from agno.tools import Toolkit
from agno.tools.decorator import tool


class BrowserTool(Toolkit):
    """浏览器自动化工具 - 使用 Playwright 控制浏览器"""

    def __init__(self, headless: bool = True, slow_mo: int = 0):
        super().__init__(name="browser_tool")
        self.headless = headless
        self.slow_mo = slow_mo
        self._browser = None
        self._page = None

        self.register(self.open_url)
        self.register(self.get_page_content)
        self.register(self.screenshot)
        self.register(self.click_element)
        self.register(self.fill_input)
        self.register(self.select_option)
        self.register(self.scroll_page)
        self.register(self.wait_for_element)
        self.register(self.get_current_url)
        self.register(self.go_back)
        self.register(self.close_browser)

    def _ensure_browser(self):
        if self._browser is None:
            from playwright.sync_api import sync_playwright
            self._playwright = sync_playwright().start()
            self._browser = self._playwright.chromium.launch(
                headless=self.headless,
                slow_mo=self.slow_mo
            )
            self._page = self._browser.new_page()

    def _run_async(self, coro):
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                import concurrent.futures
                with concurrent.futures.ThreadPoolExecutor() as pool:
                    return pool.submit(asyncio.run, coro).result()
            else:
                return loop.run_until_complete(coro)
        except RuntimeError:
            return asyncio.run(coro)

    @tool
    def open_url(self, url: str, wait_until: str = "load") -> str:
        """
        打开指定URL

        Args:
            url: 要访问的网址
            wait_until: 等待策略，可选 load/domcontentloaded/networkidle

        Returns:
            页面标题
        """
        try:
            self._ensure_browser()
            self._page.goto(url, wait_until=wait_until)
            title = self._page.title()
            return f"已打开: {url}\n页面标题: {title}"
        except Exception as e:
            return f"打开失败: {str(e)}"

    @tool
    def get_page_content(self, selector: str = "body") -> str:
        """
        获取页面内容

        Args:
            selector: CSS选择器，默认获取整个页面

        Returns:
            页面文本内容
        """
        try:
            self._ensure_browser()
            element = self._page.locator(selector)
            text = element.inner_text()
            if len(text) > 5000:
                text = text[:5000] + "\n... (内容过长，已截断)"
            return text
        except Exception as e:
            return f"获取内容失败: {str(e)}"

    @tool
    def screenshot(self, save_path: str = "screenshot.png", full_page: bool = False) -> str:
        """
        截取当前页面截图

        Args:
            save_path: 截图保存路径
            full_page: 是否截取整个页面

        Returns:
            截图保存路径
        """
        try:
            self._ensure_browser()
            path = Path(save_path)
            path.parent.mkdir(parents=True, exist_ok=True)
            self._page.screenshot(path=str(path), full_page=full_page)
            return f"截图已保存: {path.absolute()}"
        except Exception as e:
            return f"截图失败: {str(e)}"

    @tool
    def click_element(self, selector: str) -> str:
        """
        点击页面元素

        Args:
            selector: CSS选择器

        Returns:
            操作结果
        """
        try:
            self._ensure_browser()
            self._page.locator(selector).click()
            return f"已点击: {selector}"
        except Exception as e:
            return f"点击失败: {str(e)}"

    @tool
    def fill_input(self, selector: str, value: str) -> str:
        """
        填写输入框

        Args:
            selector: CSS选择器
            value: 要填入的内容

        Returns:
            操作结果
        """
        try:
            self._ensure_browser()
            self._page.locator(selector).fill(value)
            return f"已填写: {selector}"
        except Exception as e:
            return f"填写失败: {str(e)}"

    @tool
    def select_option(self, selector: str, value: str) -> str:
        """
        选择下拉框选项

        Args:
            selector: CSS选择器
            value: 选项值

        Returns:
            操作结果
        """
        try:
            self._ensure_browser()
            self._page.locator(selector).select_option(value)
            return f"已选择: {value}"
        except Exception as e:
            return f"选择失败: {str(e)}"

    @tool
    def scroll_page(self, direction: str = "down", amount: int = 500) -> str:
        """
        滚动页面

        Args:
            direction: 滚动方向 up/down
            amount: 滚动像素值

        Returns:
            操作结果
        """
        try:
            self._ensure_browser()
            if direction == "down":
                self._page.mouse.wheel(0, amount)
            else:
                self._page.mouse.wheel(0, -amount)
            return f"已向{direction}滚动 {amount}px"
        except Exception as e:
            return f"滚动失败: {str(e)}"

    @tool
    def wait_for_element(self, selector: str, timeout: int = 10000) -> str:
        """
        等待元素出现

        Args:
            selector: CSS选择器
            timeout: 超时时间(毫秒)

        Returns:
            操作结果
        """
        try:
            self._ensure_browser()
            self._page.locator(selector).wait_for(state="visible", timeout=timeout)
            return f"元素已出现: {selector}"
        except Exception as e:
            return f"等待超时: {str(e)}"

    @tool
    def get_current_url(self) -> str:
        """
        获取当前页面URL

        Returns:
            当前URL
        """
        try:
            self._ensure_browser()
            return f"当前URL: {self._page.url}"
        except Exception as e:
            return f"获取失败: {str(e)}"

    @tool
    def go_back(self) -> str:
        """
        返回上一页

        Returns:
            操作结果
        """
        try:
            self._ensure_browser()
            self._page.go_back()
            return f"已返回上一页: {self._page.url}"
        except Exception as e:
            return f"返回失败: {str(e)}"

    @tool
    def close_browser(self) -> str:
        """
        关闭浏览器

        Returns:
            操作结果
        """
        try:
            if self._browser:
                self._browser.close()
                self._playwright.stop()
                self._browser = None
                self._page = None
            return "浏览器已关闭"
        except Exception as e:
            return f"关闭失败: {str(e)}"
