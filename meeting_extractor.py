#!/usr/bin/env python3
"""
会议纪要任务提取工具
从会议记录文本中提取任务信息并生成Excel表格
"""

import os
import sys
import argparse
from pathlib import Path
from typing import List
from enum import Enum
from pydantic import BaseModel, Field
from dotenv import load_dotenv
from openai import OpenAI
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


class TaskType(str, Enum):
    """任务类型枚举"""
    INFO = "信息"
    ACTION = "行动"


class MeetingTask(BaseModel):
    """会议任务数据模型"""
    任务类型: TaskType = Field(description="任务的分类类型：信息（会议中提及的内容）或行动（待办事项）")
    任务描述: str = Field(description="任务的具体描述和要求")
    负责人: str = Field(description="负责执行该任务的人员姓名")
    纳期: str = Field(description="任务的截止日期，格式为YYYY-MM-DD")
    备注: str = Field(description="任务的补充说明或注意事项")


class MeetingResponse(BaseModel):
    """会议响应数据模型"""
    tasks: List[MeetingTask] = Field(description="从会议记录中提取的任务列表")


def load_config() -> dict:
    """加载配置文件"""
    # 加载.env文件
    load_dotenv()
    
    # 获取必需的配置项
    api_key = os.getenv("ARK_API_KEY")
    if not api_key:
        raise ValueError("未找到ARK_API_KEY环境变量，请检查.env文件配置")
    
    model_id = os.getenv("MODEL_ID", "doubao-seed-1.6-250615")
    base_url = os.getenv("BASE_URL", "https://ark.cn-beijing.volces.com/api/v3")
    
    return {
        "api_key": api_key,
        "model_id": model_id,
        "base_url": base_url
    }


def read_meeting_text(file_path: str) -> str:
    """读取会议记录文件"""
    try:
        file_path_obj = Path(file_path)
        if not file_path_obj.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        if not file_path_obj.is_file():
            raise ValueError(f"路径不是文件: {file_path}")
        
        # 读取文件内容，使用UTF-8编码
        with open(file_path_obj, 'r', encoding='utf-8') as f:
            content = f.read().strip()
        
        if not content:
            raise ValueError(f"文件内容为空: {file_path}")
        
        return content
        
    except UnicodeDecodeError:
        raise ValueError(f"文件编码错误，请确保文件是UTF-8编码: {file_path}")
    except Exception as e:
        raise ValueError(f"读取文件失败: {e}")


def extract_tasks(meeting_text: str, config: dict) -> MeetingResponse:
    """使用大模型提取会议任务"""
    try:
        # 初始化OpenAI客户端
        client = OpenAI(
            api_key=config["api_key"],
            base_url=config["base_url"]
        )
        
        # 构建提示词 - 简化为业务逻辑，使用结构化输出
        system_prompt = """你是一个专业的会议纪要分析助手。请仔细分析会议记录内容，提取出所有相关的信息和待办事项。

请将内容分为两种类型：
- 信息：会议中提及的重要内容、决定、讨论结果等信息性内容
- 行动：明确的待办事项、需要执行的任务、后续跟进事项等

对于每项内容，请确定：
- 任务类型：选择"信息"或"行动"
- 任务描述：清晰描述具体内容
- 负责人：相关的责任人姓名，如果没有明确负责人可填"待定"
- 纳期：截止时间或相关时间节点，如果没有明确日期请标注"待定"
- 备注：补充说明、依赖关系或注意事项

请从会议记录中提取所有可识别的信息和行动项。"""

        user_prompt = f"请分析以下会议记录，提取其中的信息和行动项：\n\n{meeting_text}"
        
        # 使用OpenAI结构化输出API
        completion = client.beta.chat.completions.parse(
            model=config["model_id"],
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            response_format=MeetingResponse,  # 指定响应解析模型
            max_tokens=2000,
            temperature=0.1
        )
        
        # 直接获取解析后的结构化响应
        result = completion.choices[0].message.parsed
        
        if not result.tasks:
            raise ValueError("未能从会议记录中提取到信息和行动项")
        
        return result
        
    except Exception as e:
        if "api_key" in str(e).lower() or "unauthorized" in str(e).lower():
            raise ValueError("API密钥验证失败，请检查ARK_API_KEY配置")
        elif "model" in str(e).lower():
            raise ValueError(f"模型调用失败，请检查模型ID配置: {e}")
        else:
            raise ValueError(f"大模型调用失败: {e}")


def generate_excel(tasks: List[MeetingTask], output_path: str) -> None:
    """生成Excel文件"""
    try:
        # 创建工作簿和工作表
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "会议信息行动清单"
        
        # 取消网格线显示
        ws.sheet_view.showGridLines = False
        
        # 定义表头
        headers = ["任务类型", "任务描述", "负责人", "纳期", "备注"]
        
        # 定义边框样式
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 设置表头样式
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        
        # 写入表头
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # 写入任务数据
        for row, task in enumerate(tasks, 2):
            ws.cell(row=row, column=1, value=task.任务类型.value).border = thin_border
            ws.cell(row=row, column=2, value=task.任务描述).border = thin_border
            ws.cell(row=row, column=3, value=task.负责人).border = thin_border
            ws.cell(row=row, column=4, value=task.纳期).border = thin_border
            ws.cell(row=row, column=5, value=task.备注).border = thin_border
        
        # 自动调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            # 设置列宽（最小10，最大50）
            adjusted_width = min(max(max_length + 2, 10), 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # 设置数据行样式
        data_alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = data_alignment
        
        # 保存文件
        output_path_obj = Path(output_path)
        output_path_obj.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        
        print(f"Excel文件已生成: {output_path}")
        print(f"共提取项目 {len(tasks)} 项")
        
    except Exception as e:
        raise ValueError(f"生成Excel文件失败: {e}")


def main():
    """主函数入口"""
    # 设置命令行参数解析
    parser = argparse.ArgumentParser(
        description="会议纪要任务提取工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python meeting_extractor.py meeting_notes.md
  python meeting_extractor.py meeting_notes.md -o tasks.xlsx
        """
    )
    
    parser.add_argument(
        "input_file",
        help="输入的会议记录文件路径（支持.md/.txt等格式）"
    )
    
    parser.add_argument(
        "-o", "--output",
        default="meeting_tasks.xlsx",
        help="输出Excel文件路径 (默认: meeting_tasks.xlsx)"
    )
    
    args = parser.parse_args()
    
    try:
        print("=== 会议纪要任务提取工具 ===")
        print(f"输入文件: {args.input_file}")
        print(f"输出文件: {args.output}")
        print()
        
        # 1. 加载配置
        print("📋 加载配置...")
        config = load_config()
        print("✅ 配置加载成功")
        
        # 2. 读取会议文本
        print("📖 读取会议记录...")
        meeting_text = read_meeting_text(args.input_file)
        print(f"✅ 读取成功，文本长度: {len(meeting_text)} 字符")
        
        # 3. 调用大模型提取任务
        print("🤖 调用AI分析会议内容，提取信息和行动项...")
        print("   (这可能需要几秒钟时间)")
        meeting_response = extract_tasks(meeting_text, config)
        print(f"✅ 信息和行动项提取成功，共识别 {len(meeting_response.tasks)} 项")
        
        # 4. 生成Excel文件
        print("📊 生成Excel文件...")
        generate_excel(meeting_response.tasks, args.output)
        print("✅ 处理完成!")
        
        # 5. 显示提取的任务摘要
        print("\n=== 提取摘要 ===")
        for i, task in enumerate(meeting_response.tasks, 1):
            print(f"{i}. {task.任务类型.value} - {task.负责人} - {task.纳期}")
            print(f"   {task.任务描述[:50]}{'...' if len(task.任务描述) > 50 else ''}")
            print()
            
    except KeyboardInterrupt:
        print("\n❌ 用户中断操作")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 