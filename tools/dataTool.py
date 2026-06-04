from datetime import datetime
from agno.tools import Toolkit
from agno.tools.decorator import tool


class DateTool(Toolkit):
    """日期时间工具"""

    def __init__(self):
        super().__init__(name="date_tool")
        self.register(self.get_current_date)
        self.register(self.get_current_time)

    @tool
    def get_current_date(self) -> str:
        """
        获取当前日期

        Returns:
            当前日期，格式：YYYY年MM月DD日
        """
        now = datetime.now()
        return f"今天是{now.year}年{now.month}月{now.day}日"

    @tool
    def get_current_time(self) -> str:
        """
        获取当前时间

        Returns:
            当前时间，格式：HH:MM:SS
        """
        now = datetime.now()
        return f"当前时间是{now.strftime('%H:%M:%S')}"