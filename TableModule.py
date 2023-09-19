import os
import unicodedata
import re


class Generate_Table():
    def __init__(self):
        self.temp_txt_dir = "临时表格"
        self.absolute_path = os.path.dirname(os.getcwd())
        self.generate_file_dir = self.absolute_path + r"\成品"
        if not os.path.exists(self.generate_file_dir):  # 检查目录是否存在
            os.makedirs(self.generate_file_dir)  # 创建目录
        self.filedir = self.absolute_path + r"\生成表格.txt"
        self.cell_width = 16
        # 正则匹配中文字符
        self.pattern = r'[\u4e00-\u9fff，、。：、~《》！？（）{}【】“”‘’]'

    def get_char_width(self, char):
        return 2 if unicodedata.east_asian_width(char) in ('F', 'W') else 1

    def get_text_width(self, text):
        return sum(self.get_char_width(char) for char in text)

    def wrap_text(self, text):
        current_list = ''
        final_list = []
        length = 0
        for character in text:
            if re.search(self.pattern, character):
                length += 2
            else:
                length += 1
            current_list += character

            if length == 17 or length == 18:
                final_list.append(current_list)
                current_list = ''
                length = 0
        # 如果最后的current_list不为空，那么也应该添加到final_list中
        if current_list:
            final_list.append(current_list)
        return final_list

    def build_table(self, data):
        # Wrap text in each cell
        data = [[self.wrap_text(cell) for cell in row] for row in data]
        # Get max rows in each cell
        rows_per_cell = [[len(cell) for cell in row] for row in data]
        max_rows_per_row = [max(cell) for cell in rows_per_cell]
        # Build table title
        total_width = self.cell_width * len(data[0]) + len(data[0]) + 1
        remaining_space = total_width
        half_space = remaining_space // 2
        table = f"{' ' * half_space}{' ' * (remaining_space - half_space)}\n"
        # Build top border
        # table = f"{title}\n"
        table += '┌' + '┬'.join(['─' * (self.cell_width // 2) for _ in range(len(data[0]))]) + '┐\n'
        # Add rows
        for i, row in enumerate(data):
            for j in range(max_rows_per_row[i]):
                table += '│'
                for cell in row:
                    if j < len(cell):
                        text = cell[j]
                        remaining_space = self.cell_width - 2 - self.get_text_width(text)
                        # Fill the remaining space with spaces
                        space = ' ' * remaining_space
                        # If text is full width, use left align; otherwise, center align
                        chinese_str_len = len(re.findall(self.pattern, text)) * 2
                        english_str_len = len([item for item in text if
                                               item not in re.findall(self.pattern, text)])
                        all_len = chinese_str_len + english_str_len
                        if all_len >= self.cell_width - 1:
                            align = '<'
                        else:
                            align = '^'
                        # Fill the remaining space with spaces
                        if align == '^':  # center alignment
                            half_space = remaining_space // 2
                            table += f" {' ' * half_space}{text}{' ' * (remaining_space - half_space)} │"
                        else:  # left alignment
                            text_str = f"{text}{' ' * remaining_space}│"
                            chinese_str_len = len(re.findall(self.pattern, text_str)) * 2
                            english_str_len = len([item for item in text_str if item not in re.findall(self.pattern, text_str)]) -1
                            if chinese_str_len + english_str_len == self.cell_width:
                                text_str = f"{text}{' ' * remaining_space}│"
                                table += text_str
                            else:
                                text_str = f"{text}{' ' * remaining_space} │"
                                table += text_str
                    else:
                        table += ' ' * self.cell_width + '│'
                table += '\n'
            if i < len(data) - 1:
                # Add middle border
                table += '├' + '┼'.join(['─' * (self.cell_width // 2) for _ in range(len(row))]) + '┤\n'
            else:
                # Add bottom border
                table += '└' + '┴'.join(['─' * (self.cell_width // 2) for _ in range(len(row))]) + '┘\n'
        return table

    def generate_txt_file(self, write_txt):
        with open(rf"{self.generate_file_dir}\{self.temp_txt_dir}.txt", "w", encoding="utf-8") as file:
            file.write(write_txt)

    def Table_main(self, str):
        result_dict = {}
        temp_list = []
        for i in str["cells"]:
            temp_dict = {}
            if len(i["layouts"]) == 0:
                value = ""
            else:
                value = i["layouts"][0]["text"]
            temp_dict[i["yec"]] = value
            temp_list.append(temp_dict)
        for item in temp_list:
            # 由于每个元素都是一个只有一个键值对的字典，我们可以直接获取键和值
            for key, value in item.items():
                # 检查键是否已经在结果字典中
                if key in result_dict:
                    # 如果在，我们就添加这个值到对应的值列表中
                    result_dict[key].append(value)
                else:
                    # 如果不在，我们就创建一个新的键和值列表
                    result_dict[key] = [value]
        data = [v for _, v in result_dict.items()]
        txt_text = self.build_table(data=data)
        return txt_text
        # self.generate_txt_file(write_txt=txt_text)
        # print(f"表格制作完成，目录：{self.generate_file_dir}\{self.temp_txt_dir}.txt")
