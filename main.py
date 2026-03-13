from src.config.config_manager import ConfigManager
from src.core.formatter import WordFormatter
from src.core.document import Document
from src.core.format_spec import FormatSpecParser

def main():
    try:
        # 加载配置
        config_manager = ConfigManager()

        # 读取测试文档
        doc = Document("./test/test.docx")

        # 加载格式规范
        parser = FormatSpecParser()
        format_spec = parser.get_default_format()

        # 创建格式化器并设置格式规范
        formatter = WordFormatter(doc, config_manager)
        formatter.set_format_spec(format_spec)

        # 应用格式
        formatter.format()

        # 保存格式化后的文档
        doc.save("./test/output.docx")
        print("文档格式化完成，已保存为 output.docx")

    except FileNotFoundError:
        print("错误：找不到测试文档 test.docx")
    except Exception as e:
        print(f"错误：格式化过程中出现异常 - {str(e)}")

if __name__ == "__main__":
    main()
