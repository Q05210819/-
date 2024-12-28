import os
import re
import logging
import yaml
from datetime import datetime
from docx import Document
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('question_processor.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class QuestionConfig:
    """题目处理配置类"""
    def __init__(self, config_path=None):
        # 默认配置
        self.default_config = {
            'separators': ['．', '.'],  # 题号分隔符
            'tags': {
                'answer': ['【答案】', '[答案]', '答案：', '答案:'],
                'difficulty': ['【难度】', '[难度]', '难度：', '难度:'],
                'knowledge': ['【知识点】', '[知识点]', '知识点：', '知识点:'],
                'explanation': ['【详解】', '[详解]', '解析：', '解析:']
            },
            'difficulty_levels': {
                'easy': {'threshold': 0.85, 'name': '易'},
                'medium': {'threshold': 0.85, 'name': '中'},
                'hard': {'name': '难'}
            },
            'options': ['A', 'B', 'C', 'D'],  # 选项
            'judge_answers': ['T', 'F', '对', '错', 'TRUE', 'FALSE', '√', '×', 'true', 'false', 'True', 'False', '正确', '错误'],
            'output': {
                'format': 'excel',
                'columns': ['题型', '题干', '选项', '选项数量', '答案', '解析', '所属知识点', '难度']
            }
        }

        # 加载自定义配置
        self.config = self.default_config.copy()
        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    custom_config = yaml.safe_load(f)
                    self.config.update(custom_config)
            except Exception as e:
                logger.error(f"加载配置文件失败: {str(e)}")
                logger.info("使用默认配置")

    def get_tag_content(self, text, tag_type):
        """从文本中提取标签内容"""
        for tag in self.config['tags'][tag_type]:
            if text.startswith(tag):
                return text[len(tag):].strip()
        return None

    def determine_difficulty(self, value):
        """确定难度等级"""
        try:
            value = float(value)
            if value < self.config['difficulty_levels']['easy']['threshold']:
                return self.config['difficulty_levels']['easy']['name']
            elif value == self.config['difficulty_levels']['medium']['threshold']:
                return self.config['difficulty_levels']['medium']['name']
            else:
                return self.config['difficulty_levels']['hard']['name']
        except ValueError:
            logger.warning(f"无法解析难度值: {value}")
            return self.config['difficulty_levels']['medium']['name']

class QuestionType:
    """题目类型判断类"""
    def __init__(self, config):
        self.config = config

    def determine_type(self, answer):
        """判断题目类型"""
        if not answer:
            logger.info("空答案")
            return '未知类型'

        original_answer = answer
        answer = answer.upper().strip()
        logger.info(f"判断题目类型 - 原始答案: {original_answer}, 处理后: {answer}")

        # 判断题
        judge_answers = self.config.config['judge_answers']
        for judge_ans in judge_answers:
            if judge_ans == original_answer or judge_ans.upper() == answer:
                logger.info(f"识别为判断题")
                return '判断题'

        # 选择题
        valid_options = set(self.config.config['options'])
        answer_chars = set(answer.replace(' ', ''))
        if answer_chars and answer_chars.issubset(valid_options):
            is_multi = len(answer.replace(' ', '')) > 1
            logger.info(f"识别为{'多' if is_multi else '单'}选题")
            return '多选题' if is_multi else '单选题'

        # 填空题
        logger.info(f"识别为填空题")
        return '填空题'

    def count_options(self, choices):
        """计算选项数量"""
        if not choices:
            return 0
        options = self.config.config['options']
        count = 0
        choices = choices.upper()
        for opt in options:
            if f"{opt}." in choices or f"{opt}．" in choices or f"{opt} " in choices:
                count += 1
        return count

class QuestionProcessor:
    """题目处理类"""
    def __init__(self, config):
        self.config = config
        self.question_type = QuestionType(config)
        self.stats = {
            '单选题': 0,
            '多选题': 0,
            '判断题': 0,
            '填空题': 0,
            '未知类型': 0,
            '总题数': 0
        }

    def is_question_start(self, text):
        """检查是否是题目开始"""
        for sep in self.config.config['separators']:
            pattern = r'^\d+' + sep
            if re.match(pattern, text.strip()):
                return True
        return False

    def is_option_start(self, text):
        """检查是否是选项开始"""
        text = text.strip()
        return any(text.startswith(f"{opt}.") or text.startswith(f"{opt}．")
                  for opt in self.config.config['options'])

    def has_image(self, paragraph):
        """检查段落是否包含图片"""
        try:
            for run in paragraph.runs:
                if run._element.findall('.//w:drawing', run._element.nsmap) or \
                   run._element.findall('.//w:pict', run._element.nsmap):
                    return True
        except Exception as e:
            logger.warning(f"检查图片时出错: {str(e)}")
        return False

    def remove_question_number(self, text):
        """移除题目编号"""
        for sep in self.config.config['separators']:
            pattern = r'^\d+' + sep + r'\s*'
            text = re.sub(pattern, '', text)
        return text

    def process_document(self, doc_path):
        """处理文档"""
        try:
            doc = Document(doc_path)
        except Exception as e:
            logger.error(f"打开文档失败: {str(e)}")
            raise

        data = {
            'title_list': [],
            'choice_list': [],
            'answer_list': [],
            'difficulty_list': [],
            'knowledge_list': [],
            'explain_list': [],
            'type_list': [],
            'option_count_list': [],
            'has_image_list': []
        }

        current_question = None
        collecting_title = False
        current_title_lines = []
        current_choices = []
        current_has_image = False

        def save_current_question():
            nonlocal current_question, current_title_lines, current_choices, current_has_image
            if current_question:
                # 保存题干
                full_title = ' '.join(current_title_lines)
                cleaned_title = re.sub(r'（\s*）', '（   ）', full_title)
                cleaned_title = self.remove_question_number(cleaned_title)
                data['title_list'].append(cleaned_title)
                data['has_image_list'].append(current_has_image)

                # 保存选项
                choice_text = '\n'.join(current_choices) if current_choices else ''
                data['choice_list'].append(choice_text)
                data['option_count_list'].append(self.question_type.count_options(choice_text))

                # 保存答案和题型
                answer = current_question.get('answer', '')
                data['answer_list'].append(answer)
                question_type = self.question_type.determine_type(answer)
                data['type_list'].append(question_type)
                self.stats[question_type] = self.stats.get(question_type, 0) + 1

                # 保存其他信息
                data['difficulty_list'].append(current_question.get('difficulty', ''))
                data['knowledge_list'].append(current_question.get('knowledge', ''))
                data['explain_list'].append(current_question.get('explanation', ''))

                # 重置当前题目信息
                current_title_lines.clear()
                current_choices.clear()
                current_has_image = False

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue

            # 如果是新题目开始
            if self.is_question_start(text):
                save_current_question()
                current_question = {'title': text}
                collecting_title = True
                current_title_lines = [text]
                current_has_image = self.has_image(para)
                logger.info(f"发现新题目: {text}")

            # 处理选项
            elif self.is_option_start(text):
                collecting_title = False
                current_choices.append(text)

            # 处理答案
            elif any(text.startswith(tag) for tag in self.config.config['tags']['answer']):
                answer_text = self.config.get_tag_content(text, 'answer')
                if answer_text and current_question:
                    current_question['answer'] = answer_text
                    logger.info(f"处理答案: {answer_text}")

            # 处理其他标签
            elif any(text.startswith(tag) for tag in self.config.config['tags']['difficulty']):
                difficulty_text = self.config.get_tag_content(text, 'difficulty')
                if current_question:
                    current_question['difficulty'] = self.config.determine_difficulty(difficulty_text)

            elif any(text.startswith(tag) for tag in self.config.config['tags']['knowledge']):
                knowledge_text = self.config.get_tag_content(text, 'knowledge')
                if current_question:
                    current_question['knowledge'] = knowledge_text

            elif any(text.startswith(tag) for tag in self.config.config['tags']['explanation']):
                explain_text = self.config.get_tag_content(text, 'explanation')
                if current_question:
                    current_question['explanation'] = explain_text

            # 如果正在收集题干
            elif collecting_title:
                current_title_lines.append(text)
                if not current_has_image:
                    current_has_image = self.has_image(para)

        # 保存最后一个题目
        save_current_question()

        self.stats['总题数'] = len(data['title_list'])
        return data


class QuestionExporter:
    """题目导出类"""

    def __init__(self, config):
        self.config = config

    def prepare_excel_data(self, data):
        """准备Excel数据"""
        excel_data = []
        for i in range(len(data['title_list'])):
            # 处理题干和图片标记
            title_text = data['title_list'][i]
            has_image = data['has_image_list'][i] if i < len(data['has_image_list']) else False

            # 处理选项
            options_text = data['choice_list'][i] if i < len(data['choice_list']) else ''
            option_count = data['option_count_list'][i] if i < len(data['option_count_list']) else 0

            # 创建题目数据字典
            question_data = {
                '题型': data['type_list'][i] if i < len(data['type_list']) else '',
                '题干': title_text,
                '选项': options_text,
                '选项数量': option_count,
                '答案': data['answer_list'][i] if i < len(data['answer_list']) else '',
                '解析': data['explain_list'][i] if i < len(data['explain_list']) else '',
                '所属知识点': data['knowledge_list'][i] if i < len(data['knowledge_list']) else '',
                '难度': data['difficulty_list'][i] if i < len(data['difficulty_list']) else '',
                'has_image': has_image
            }
            excel_data.append(question_data)

        return excel_data

    def _format_excel_worksheet(self, worksheet, data):
        """设置Excel工作表格式"""
        # 设置列宽
        custom_widths = {
            '题型': 12,
            '题干': 60,
            '选项': 40,
            '选项数量': 10,
            '答案': 15,
            '解析': 50,
            '所属知识点': 30,
            '难度': 10
        }

        # 设置浅蓝色填充
        light_blue_fill = PatternFill(start_color='E6F3FF', end_color='E6F3FF', fill_type='solid')

        # 设置列宽
        for idx, column in enumerate(worksheet.columns):
            column_letter = column[0].column_letter
            column_name = self.config.config['output']['columns'][idx]
            worksheet.column_dimensions[column_letter].width = custom_widths.get(column_name, 20)

        # 设置单元格格式
        for row_idx, row in enumerate(worksheet.rows):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')

                if row_idx == 0:  # 表头行
                    cell.font = Font(bold=True)
                elif row_idx <= len(data):  # 数据行
                    if data[row_idx - 1].get('has_image'):  # 检查是否包含图片
                        cell.fill = light_blue_fill

    def export_to_excel(self, data, output_path=None):
        """导出到Excel"""
        try:
            excel_data = self.prepare_excel_data(data)

            # 按题型分组
            grouped_data = []
            type_order = ['单选题', '多选题', '判断题', '填空题', '未知类型']

            for question_type in type_order:
                type_questions = [q for q in excel_data if q['题型'] == question_type]
                if type_questions:
                    if grouped_data:  # 在不同题型之间添加空行
                        empty_row = {key: '' for key in self.config.config['output']['columns']}
                        empty_row['has_image'] = False
                        grouped_data.append(empty_row)
                    grouped_data.extend(type_questions)

            # 创建DataFrame时排除has_image列
            columns = self.config.config['output']['columns']
            df = pd.DataFrame(grouped_data)[columns]

            if output_path is None:
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f'题库导出_{current_time}.xlsx'

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='题库')
                workbook = writer.book
                worksheet = writer.sheets['题库']
                self._format_excel_worksheet(worksheet, grouped_data)

            logger.info(f"Excel文件已生成: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"导出Excel时出错: {str(e)}")
            raise

class QuestionBank:
    """题库管理主类"""
    def __init__(self, config_path=None):
        self.config = QuestionConfig(config_path)
        self.processor = QuestionProcessor(self.config)
        self.exporter = QuestionExporter(self.config)

    def process_file(self, file_path, output_path=None):
        """处理文件的主函数"""
        try:
            logger.info(f"开始处理文件: {file_path}")
            data = self.processor.process_document(file_path)
            output_file = self.exporter.export_to_excel(data, output_path)
            self._print_statistics()
            return output_file
        except Exception as e:
            logger.error(f"处理文件时出错: {str(e)}")
            raise

    def _print_statistics(self):
        """打印统计信息"""
        logger.info("\n=== 题目统计信息 ===")
        logger.info(f"总题目数量：{self.processor.stats['总题数']}题")
        logger.info(f"选择题数量：{self.processor.stats['单选题'] + self.processor.stats['多选题']}题")
        logger.info(f"填空题数量：{self.processor.stats['填空题']}题")
        logger.info(f"判断题数量：{self.processor.stats['判断题']}题")

        logger.info("\n=== 题型分布 ===")
        for q_type, count in self.processor.stats.items():
            if q_type != '总题数':
                logger.info(f"{q_type}: {count}题")


def main():
    """主程序"""
    try:
        # 初始化题库管理器
        question_bank = QuestionBank()

        # 处理Word文档
        doc_path = './2024年12月24日高中信息技术作业 (1).docx'  # 替换为实际的文档路径
        output_path = None  # 可以指定输出路径，默认使用时间戳命名

        output_file = question_bank.process_file(doc_path, output_path)
        logger.info(f"处理完成，输出文件: {output_file}")

    except Exception as e:
        logger.error(f"程序执行出错: {str(e)}")
        raise

if __name__ == "__main__":
    main()