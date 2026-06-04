# tools/visionTool.py
import os
import json
import base64
from pathlib import Path
from typing import Optional, List, Dict, Any
from agno.tools import Toolkit
from agno.tools.decorator import tool


class VisionTool(Toolkit):
    """视觉大模型工具 - 使用通义千问 VL 分析图片和视频"""

    def __init__(self, api_key: Optional[str] = None, api_url: Optional[str] = None):
        """
        初始化视觉工具

        Args:
            api_key: DashScope API Key
            api_url: DashScope API URL
        """
        super().__init__(name="vision_tool")

        self.api_key = api_key or os.getenv("DASHSCOPE_API_KEY")
        self.api_url = api_url or os.getenv("DASHSCOPE_API_URL",
                                            "https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation")

        # 注册工具方法
        self.register(self.analyze_image)
        self.register(self.analyze_sandbox_image)
        self.register(self.analyze_video_frames)

        print("视觉大模型工具已初始化")

    def _encode_image(self, image_path: str) -> str:
        """将图片转换为 base64"""
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    def _call_vl_model(self, prompt: str, image_paths: List[str]) -> str:
        """
        调用通义千问 VL 模型

        Args:
            prompt: 文本提示词
            image_paths: 图片路径列表（支持1-10张图片）

        Returns:
            模型返回的分析结果
        """
        import requests

        # 构建请求
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        # 构建内容（文本 + 图片）
        content = [{"text": prompt}]
        for img_path in image_paths:
            img_base64 = self._encode_image(img_path)
            content.append({
                "image": f"data:image/jpeg;base64,{img_base64}"
            })

        payload = {
            "model": "qwen-vl-plus",  # 通义千问 VL Plus 模型
            "input": {
                "messages": [
                    {
                        "role": "user",
                        "content": content
                    }
                ]
            },
            "parameters": {
                "result_format": "message",
                "temperature": 0.3,
                "max_tokens": 1024
            }
        }

        try:
            response = requests.post(
                self.api_url,
                headers=headers,
                json=payload,
                timeout=60
            )

            if response.status_code == 200:
                result = response.json()
                # 提取返回的文本
                return result["output"]["choices"][0]["message"]["content"]
            else:
                return f"API 调用失败: {response.status_code} - {response.text}"

        except Exception as e:
            return f"请求异常: {str(e)}"

    @tool
    def analyze_image(self, image_path: str, question: str = "请描述这张图片的内容") -> str:
        """
        分析单张图片

        Args:
            image_path: 图片路径
            question: 要问的问题，默认是描述图片

        Returns:
            图片分析结果
        """
        try:
            path = Path(image_path)
            if not path.exists():
                return f"图片不存在: {image_path}"

            result = self._call_vl_model(question, [str(path)])
            return f"图片分析结果:\n\n{result}"

        except Exception as e:
            return f"分析失败: {str(e)}"

    @tool
    def analyze_sandbox_image(self, image_path: str) -> str:
        """
        专门分析心理沙盘图片，识别沙具和心理特征

        Args:
            image_path: 沙盘图片路径

        Returns:
            沙盘分析结果
        """
        try:
            path = Path(image_path)
            if not path.exists():
                return f"图片不存在: {image_path}"

            # 专业心理沙盘分析提示词
            prompt = """你是一位专业的沙盘游戏治疗师。请分析这张沙盘图片，输出以下信息（用中文，简洁清晰）：

1. 沙具识别：列出图片中能看到的所有沙具类型（如人物、动物、建筑、植物、交通工具等）
2. 空间布局：描述沙具的分布情况（密集/分散、是否有中心区域）
3. 心理特征：
   - 聚集度：沙具之间的紧密程度
   - 秩序性：摆放是否有规律
   - 主题感：是否有明显的故事或主题
4. 建议关注点：指出最值得注意的沙具或区域

只输出分析内容，不要有额外解释。"""

            result = self._call_vl_model(prompt, [str(path)])
            return f"沙盘心理分析:\n\n{result}"

        except Exception as e:
            return f"分析失败: {str(e)}"

    @tool
    def analyze_video_frames(self, video_path: str, frame_count: int = 5) -> str:
        """
        分析视频的关键帧（提取视频中的多帧进行综合分析）

        Args:
            video_path: 视频路径
            frame_count: 要分析的帧数（默认5帧）

        Returns:
            视频分析结果
        """
        try:
            path = Path(video_path)
            if not path.exists():
                return f"视频不存在: {video_path}"

            # 使用 OpenCV 提取视频帧
            import cv2

            cap = cv2.VideoCapture(str(path))
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = cap.get(cv2.CAP_PROP_FPS)
            duration = total_frames / fps

            # 计算要提取的帧位置
            temp_dir = Path("temp_frames")
            temp_dir.mkdir(exist_ok=True)

            frame_paths = []
            interval = max(1, total_frames // frame_count)

            for i, frame_idx in enumerate(range(0, total_frames, interval)[:frame_count]):
                cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
                ret, frame = cap.read()
                if ret:
                    # 保存临时帧
                    temp_path = temp_dir / f"frame_{i}_{frame_idx}.jpg"
                    cv2.imwrite(str(temp_path), frame)
                    frame_paths.append(str(temp_path))

            cap.release()

            # 构建分析提示词
            prompt = f"""这是一个沙盘构建过程的视频（总时长 {duration:.1f} 秒）。这是从视频中提取的 {len(frame_paths)} 个关键帧。
请分析整个沙盘构建过程，输出：

1. 构建顺序：沙具是如何逐步添加的
2. 变化趋势：从第一帧到最后一帧的主要变化
3. 行为观察：用户的可能操作模式（快速/缓慢/犹豫）
4. 心理状态评估：基于沙盘的变化评估可能的心理状态

请综合分析这些帧，不需要逐帧描述。"""

            # 调用视觉模型（支持多图）
            import requests

            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }

            # 构建多图内容
            content = [{"text": prompt}]
            for fp in frame_paths:
                img_base64 = self._encode_image(fp)
                content.append({"image": f"data:image/jpeg;base64,{img_base64}"})

            payload = {
                "model": "qwen-vl-plus",
                "input": {
                    "messages": [
                        {
                            "role": "user",
                            "content": content
                        }
                    ]
                },
                "parameters": {
                    "result_format": "message",
                    "temperature": 0.3,
                    "max_tokens": 1500
                }
            }

            response = requests.post(
                self.api_url,
                headers=headers,
                json=payload,
                timeout=120
            )

            # 清理临时文件
            for fp in frame_paths:
                try:
                    Path(fp).unlink()
                except:
                    pass
            try:
                temp_dir.rmdir()
            except:
                pass

            if response.status_code == 200:
                result = response.json()
                analysis = result["output"]["choices"][0]["message"]["content"]
                return f"视频分析结果（共 {frame_count} 个关键帧）:\n\n{analysis}"
            else:
                return f"API 调用失败: {response.status_code}"

        except ImportError:
            return "需要安装 opencv-python: pip install opencv-python"
        except Exception as e:
            return f"视频分析失败: {str(e)}"