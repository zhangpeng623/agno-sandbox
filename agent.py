from agno.agent import Agent
from agno.models.dashscope import DashScope
from agno.models.openai import OpenAILike
from dotenv import load_dotenv
from tools.webSearch import BaiduWebSearch
from tools.dataTool import DateTool
from tools.fileTool import FileTool
from agno.tools.baidusearch import BaiduSearchTools
from tools.visionTool import VisionTool
from agno.db.sqlite import SqliteDb
from agno.memory import MemoryManager

import os

load_dotenv()

#使用 MemoryManager（管理用户长期记忆）
db = SqliteDb(db_file="agno.db")

memory_manager = MemoryManager(
    db=db,
    model=OpenAILike(
        id="qwen3.7-max", # 用于生成记忆的模型
        api_key=os.getenv("DASHSCOPE_API_KEY"),
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        ),
    additional_instructions="",
)


agent = Agent(
    tools=[BaiduSearchTools(),DateTool(),FileTool(),VisionTool()],
    db = db,
    memory_manager=memory_manager,
    update_memory_on_run=True,
    add_memories_to_context=True,
    model=OpenAILike(
        id="qwen3.6-plus",
        api_key=os.getenv("DSAHSCOPE_API_KEY"),
        base_url=os.getenv("DSAHSCOPE_API_URL"),
        cache_response=True,
    ),
    instructions="你是一个回答快速的助手，请尽可能的简洁的回答问题，不需要思考太多。",
)
agent.print_response("你还记得我是谁吗", stream=True)
