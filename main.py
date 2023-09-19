import os
import re
import shutil
from method import Cut, PdfToPng
from method import Read_Pdf
from method import PdfProcessor
from concurrent.futures import ProcessPoolExecutor
# from method import BaiDu_OCR
# import json
# import base64
# from urllib.parse import urlencode


class Processor:
    def __init__(self):
        self.cut = Cut()
        self.read_pdf = Read_Pdf()
        self.threads_count = 15
        # 文件绝对路径
        self.absolute_path = os.path.dirname(os.getcwd())
        # 生成成品文件路径
        self.generate_file_dir = self.absolute_path + r"/成品"
        if not os.path.exists(self.generate_file_dir):  # 检查目录是否存在
            os.makedirs(self.generate_file_dir)  # 创建目录
        # 水印文件
        self.watermark_dir = self.absolute_path + r"/watermark.pdf"
        # 临时文件
        self.temp_file = self.absolute_path + r"/temp_pdf"
        if not os.path.exists(self.temp_file):  # 检查目录是否存在
            os.makedirs(self.temp_file)  # 创建目录

        self.temp_cute_file = self.absolute_path + r"/cut_pdf"
        if not os.path.exists(self.temp_cute_file):  # 检查目录是否存在
            os.makedirs(self.temp_cute_file)  # 创建目录

        # self.temp_pdf_file = self.absolute_path + r"/temp_pdf/temp_pdf.pdf"
        # 中文字体文件路径
        self.chinese_font_path = self.absolute_path + r"/fonts/simsun.ttc"
        self.chinese_font_name = r"宋体"
        # 页脚内容
        self.footer_text1 = r"北大法律信息网www.chinalawinfo.com/"
        self.footer_text2 = r"北大法宝www.pkulaw.com/"

    def get_txt_content(self, input_text, TXT_primary_content, input_message_dict):
        try:
            TXT_primary_content["起始页码"] = str(list(input_message_dict.values())[0]["cut_start_page"])
            TXT_primary_content["终止页码"] = str(list(input_message_dict.values())[0]["cut_end_page"])
            TXT_primary_content["标题"] = str(list(input_message_dict.keys())[0])
            TXT_primary_content["自定义中文标题"] = TXT_primary_content["标题"]
        except Exception as e:
            print(f"处理TXT内容，错误信息：\n{e}")
        # 获取第一个元素的键和值
        first_key = next(iter(input_text))
        first_page = input_text[first_key].replace("\u3000 ", "").replace("\u3000", "").replace("", "").replace(" ", "")
        # 获取期刊年份
        re_serial_year = re.search(r"(\d{4})年", first_page, re.DOTALL)
        if re_serial_year:
            serial_year = re_serial_year.group(1)
            TXT_primary_content["期刊年份"] = str(serial_year)
        # 获取期刊号
        re_serial_num = re.search(r"第(\d+)期", first_page, re.DOTALL)
        if re_serial_num:
            serial_num = re_serial_num.group(1)
            TXT_primary_content["期刊号"] = str(serial_num)
        # 获取中文摘要
        re_chinese_abstract = re.search("摘[\s\S]*?要([\s\S]*?)关键词", first_page, re.DOTALL)
        if re_chinese_abstract:
            chinese_abstract = re_chinese_abstract.group(1).strip().replace("\n", "")
            chinese_abstract = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9；,.?!]+', '', chinese_abstract) + '。'
            TXT_primary_content["摘要"] = str(chinese_abstract)
        # 获取中文关键字
        re_chinese_keywords = re.search(r'关键词(.*?)(?=\n)', first_page, re.DOTALL)
        if re_chinese_keywords:
            chinese_keywords = re_chinese_keywords.group(1).strip().replace("\n", "").replace(";", "；")
            chinese_keywords = re.sub(r'[^\u4e00-\u9fa5；]+', '', chinese_keywords)
            TXT_primary_content["关键字"] = str(chinese_keywords)
        # 获取中图分类号
        re_clc_number = re.search(r'中图分类号(.*?)文献标识码', first_page, re.DOTALL)
        if re_clc_number:
            clc_number = re_clc_number.group(1).strip().replace("\n", "").replace(";", "；")
            clc_number = re.sub(r'[^\w\s\u4e00-\u9fa5，。；“”‘’;,:！？]', '', clc_number)
            TXT_primary_content["中图分类号"] = str(clc_number)
        # 获取文献标识码
        re_document_code = re.search(r'文献标识码(.*?)文章编号', first_page, re.DOTALL)
        if re_document_code:
            document_code = re_document_code.group(1).strip().replace("\n", "").replace(";", "；")
            document_code = re.sub(r'[^\w\s\u4e00-\u9fa5，。；“”‘’;,:！？]', '', document_code)
            TXT_primary_content["文献标识码"] = str(document_code)
        # 获取文章编号
        re_article_number = re.search(r'文章编号(.*?)(?=\n)', first_page, re.DOTALL)
        if re_article_number:
            article_number = re_article_number.group(1).strip().replace("\n", "").replace(";", "；")
            article_number = re.sub(r'[\u4e00-\u9fa5【】，。；：“”‘’;,:！？]', '', article_number)
            TXT_primary_content["文章编号"] = str(article_number)

        # 获取最后一个元素的键和值
        last_key = list(input_text.keys())[-1]
        last_page = input_text[last_key].replace("\u3000 ", "").replace("\u3000", "").replace("\n", "").replace("", "")
        # 获取英文摘要
        re_english_abstract1 = re.search(r'Abstract:(.*?)(?=Key.*Words)', last_page, re.DOTALL | re.IGNORECASE)
        re_english_abstract2 = re.search(r'Abstract：(.*?)(?=Key.*Words)', last_page, re.DOTALL | re.IGNORECASE)
        if re_english_abstract1:
            english_abstract1 = re_english_abstract1.group(1).strip()
            TXT_primary_content["英文摘要"] = str(english_abstract1)
        elif re_english_abstract2:
            english_abstract2 = re_english_abstract2.group(1).strip()
            TXT_primary_content["英文摘要"] = str(english_abstract2)
        # 获取英文关键字
        re_english_keywords = re.search(r'Key[\s\S]*?Words([^！？。，；：“”‘’【】\[\]【】\u4e00-\u9fa5]+)', last_page,re.DOTALL | re.IGNORECASE)
        if re_english_keywords:
            english_keywords = re_english_keywords.group(1).replace(": ", "").replace("： ", "").replace(" :", "").replace(" ：", "")
            TXT_primary_content["英文关键字"] = str(english_keywords)

        file_text = ""
        for k, v in TXT_primary_content.items():
            temp_str = k + "＝" + v + '\n'
            file_text += temp_str
        file_text = re.sub(r'＝\n$', '', file_text)
        return file_text

    def generate_txt_file(self, write_txt, filename):
        with open(rf"{self.generate_file_dir}/{filename}.txt", "w", encoding="utf-8") as file:
            file.write(write_txt)

    def generate_pdf_file(self, filename, cut_pdf_dir):
        try:
            finall_pdf = PdfProcessor()
            # 添加水印：cut_pdf_dir切割后的文件，watermark_dir水印文件，temp_pdf_file添加水印的临时文件(添加页脚后删除)
            temp_pdf_file = self.temp_file + rf"/temp_{filename}.pdf"
            finall_pdf.add_watermark(input_pdf=cut_pdf_dir, watermark_pdf=self.watermark_dir, output_pdf=temp_pdf_file)
            # 最终的PDF文件
            finall_pdf_dir = rf"{self.generate_file_dir}/{filename}.pdf"
            # 添加水印后的PDF文件上添加中文页脚
            finall_pdf.add_page_footer(input_pdf=temp_pdf_file, output_pdf=finall_pdf_dir, font_path=self.chinese_font_path,
                                       font_name=self.chinese_font_name, footer_text1=self.footer_text1, footer_text2=self.footer_text2)
            ## print(rf"{filename}的PDF完成")
        except Exception as e:
            print(rf"{filename}的PDF制作失败，错误信息：\n{e}")

    def clear_folder(self, folder_path):
        """清空文件夹"""
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)

    def run_main(self, k, v, TXT_primary_content, source_pdf):
        new_dict = {k: v}
        temp_pdf_file = self.temp_file + rf"/temp_{k}.pdf"
        cut_pdf_dir = self.cut.split_pdf(input_pdf=source_pdf, output_pdf=temp_pdf_file,
                                         start_page=int(v["cut_start_page"]), end_page=int(v["cut_end_page"]))
        # print(rf"切割{k}")
        # 将切割后的PDF，添加水印和页脚，重新生成一个PDF
        self.generate_pdf_file(filename=k, cut_pdf_dir=cut_pdf_dir)

        # 读取PDF，获取一串原始str
        source_str = self.read_pdf.extract_text_from_pdf(pdf_path=cut_pdf_dir)
        # print(source_str)
        # 将str处理为需要的字符串
        write_txt = self.get_txt_content(input_text=source_str, TXT_primary_content=TXT_primary_content,
                                         input_message_dict=new_dict)
        # 写入txt
        self.generate_txt_file(write_txt=write_txt, filename=k)

    def read_txt(self, filename):
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as err:
            print(f"读取txt报错:\n{err}")
        
    def generate_catalogue_txt(self, input_message_dict, name):
        try:
            if input_message_dict.get(name).get("源文件绝对路径") != "":
                temp_catalogue_pdf = self.temp_file + rf"/temp_{name}.pdf"
                cut_pdf_dir = self.cut.split_pdf(input_pdf=input_message_dict[name]["源文件绝对路径"],
                                            output_pdf=temp_catalogue_pdf,
                                            start_page=int(input_message_dict[name]["cut_start_page"]),
                                            end_page=int(input_message_dict[name]["cut_end_page"]))
                # 将切割后的PDF，添加水印和页脚，重新生成一个PDF
                self.generate_pdf_file(filename=name, cut_pdf_dir=cut_pdf_dir)
                source_str = self.read_pdf.chinesedir_extract_text_from_pdf(pdf_path=cut_pdf_dir).replace("…",
                                                                                                               "...")
                self.generate_txt_file(write_txt=f"{name}\n" + source_str, filename=name)
        except Exception as e:
            pass


if __name__ == '__main__':
    TXT_primary_content = {
        "期刊名称": "",
        "期刊年份": "",
        "期刊号": "",
        "起始页码": "",
        "终止页码": "",
        "期刊栏目": "",
        "标题": "",
        "英文标题": "",
        "副标题": "",
        "英文副标题": "",
        "作者": "",
        "作者单位": "",
        "摘要": "",
        "英文摘要": "",
        "关键字": "",
        "英文关键字": "",
        "基金项目": "",
        "中图分类号": "",
        "文献标识码": "",
        "文章编号": "",
        "内容": "",
        "注释": "",
        "参考文献": "",
        "自定义中文标题": "",
        "自定义英文标题": "",
        "文章类型": "",
        "▲": "",
    }

    processor = Processor()
    catalogue_dir = processor.absolute_path + r'/录入目录.txt'
    txt_dir = processor.absolute_path + r'/录入TXT.txt'
    print(processor.read_txt(catalogue_dir))
    print(type(processor.read_txt(catalogue_dir)))
    input_message_dict = eval(processor.read_txt(catalogue_dir))
    source_pdf = input_message_dict["源文件绝对路径"]
    input_message_dict.pop("源文件绝对路径")
    for k, v in input_message_dict.items():
        temp_pdf_file = processor.temp_cute_file + rf"/{k}.pdf"
        cut_pdf_dir = processor.cut.split_pdf(input_pdf=source_pdf, output_pdf=temp_pdf_file,
                                         start_page=int(v["cut_start_page"]), end_page=int(v["cut_end_page"]))
    # # 切割目录
    processor.generate_catalogue_txt(input_message_dict, name="中文目录")
    processor.generate_catalogue_txt(input_message_dict, name="英文目录")

    TXT_input_temp = eval(processor.read_txt(txt_dir))
    TXT_primary_content.update(TXT_input_temp)

    input_message_dict.pop("中文目录")
    input_message_dict.pop("英文目录")
    with ProcessPoolExecutor(max_workers=processor.threads_count) as executor:
        futures = {executor.submit(processor.run_main, k=k, v=v, TXT_primary_content=TXT_primary_content, source_pdf=source_pdf): None for k, v in input_message_dict.items()}
    for future in futures:
        future.result()

    # 关闭进程池
    executor.shutdown()
    processor.clear_folder(folder_path=processor.absolute_path + r"/temp_pdf")
