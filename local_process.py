from TableModule import Generate_Table
import re
import time

#列表字符包含判断
def is_field_in_list(source_data, field_list):
    for field in field_list:
        if field in source_data:
            return True
    return False

#判断是不是中文
def is_chinese_character(char):
    if '\u4e00' <= char <= '\u9fff':
        return True
    else:
        return False

#提取注释前面的字符串
def extract_special_characters(input_string):
    #特殊字符表达式
    special_characters = re.findall(r'①|②|③|④|⑤|⑥|⑦|⑧|⑨|\〔\d+\〕|\[\d+\]', input_string)
    return special_characters

#正则判断内容里的字符串是否存在
def check_regex_match(input_string, regex_list):
    for regex_pattern in regex_list:
        matches = re.findall(r'{}'.format(input_string), regex_pattern)
        if len(matches) > 0:
            return True
    return False

#替换正文里的角标
def replace_string_in_list(input_list, string_to_replace, replacement):
    for i in range(len(input_list)):
        if string_to_replace in input_list[i]:
            input_list[i] = input_list[i].replace(string_to_replace, replacement)

#识别出的错误字符集合
specia_string = ['①','[i0]','[ii]','[m0','[]]','[ng]','[na]','[n]','[',']','〕','〔','[s]','〔1]','[四 ','@','[2r]','[x]']

def data_analysis(datajson):
    data_json = datajson["layouts"]
    #摘要关键字列表
    abstract_keys = ["摘 要","摘要","内容摘要","编者按","[摘  要]","摘  要","内容提要"]
    any_keys = ["中图分类号","文献标识码","文章编号"]
    #内容、注释常量
    count_constant = 0
    #全部内容的列表
    contents_all = []
    #全部注释的列表
    note_all = []
    #摘要内容
    abstracts = []
    #错误信息反馈
    error_message = []
    for data in range(len(data_json)):
        #第一页的内容进行特殊操作
        if data_json[data]["pageNum"][0] == 0:
            # 提取正文的内容
            if data_json[data]["type"] != "foot" \
                and data_json[data]["text"] != "" \
                and data_json[data]["type"] != "foot_pagenum" \
                and data_json[data]["type"] != "corner_note" \
                and data_json[data]["type"] != "head" \
                and data_json[data]["type"] != "note_line" \
                and data_json[data]["subType"] != "doc_title" \
                and data_json[data]["subType"] != "doc_name" \
                and data_json[data]["subType"] != "doc_subtitle" \
                and data_json[data]["type"] != "end_note" \
                and data_json[data]["type"] != "title" \
                and data_json[data]["subType"] != "footer_note":
                if is_field_in_list(data_json[data]["text"], abstract_keys):
                    #提取 摘要 清洗
                    abstract_con = data_json[data]["text"].replace(' ', '')
                    abstracts.append(abstract_con)
                elif data_json[data]["text"].startswith("关键词") or data_json[data]["text"].startswith("关键字") or data_json[data]["text"].startswith("[关键词]"):
                    #提取 关键字 清晰
                    #keywords = data_json[data]["text"].replace(' ', '').split("：")[1]
                    pass
                elif data_json[data]["text"].startswith("[作  者]") or  data_json[data]["text"].startswith("[作者]"):
                    #提取作者相关信息
                    pass
                elif is_field_in_list(data_json[data]["text"], any_keys):
                    #提取中图相关未完成
                    #print(data_json[data]["text"])
                    pass
                else:
                    #提取第一页的内容
                    contents = data_json[data]["text"].replace("?", "？").replace("《:", ":《").replace(":》", ":》").replace(";", "；").replace("(", "（").replace(")", "）").replace(":", "：").replace(":", "：").replace(",", "，")
                    contents_all.append(contents)
            elif data_json[data]["subType"] == "para_title":
                #添加段落title到内容里
                contents = data_json[data]["text"]
                contents_all.append(contents)
            elif data_json[data]["subType"] == "doc_title":
                #提取标题
                #title = data_json[data]["text"].replace(' ', '')
                pass
            else:
                #提取第一页相关注释
                if (data_json[data]["subType"] == "footer_note" or data_json[data]["type"] == "corner_note") and data_json[data]["subType"] != "page":
                    note = data_json[data]["text"]
                    if note.startswith("作者单位") or note.startswith("项目基金") or note.startswith("基金项目") or note.startswith("作者简介") or note.startswith("*") or note.startswith("本文") or note.startswith("收稿") or note.startswith("国家") or note.startswith("定稿"):
                        note = data_json[data]["text"]
                        #添加到注释总表
                        note_all.append(note)
                    else:
                        count_constant += 1
                        note_mark_number = f"[{count_constant}]"
                        #提取注释里前面的角标
                        #print(data_json[data]["text"])
                        note_mark = extract_special_characters(note)
                        if len(note_mark) > 0:
                            #查找原文里的角标是否存在
                            if check_regex_match(note_mark[0], contents_all):
                                replace_string_in_list(contents_all, note_mark[0], note_mark_number)
                            #原文中没有找到角标
                            else:
                                #这里做报错信息提示
                                note_error = "书籍第 " + str(data_json[data]["pageNum"][0]) + " 页 原文应改成 " + note_mark_number + " " + data_json[data]["text"]
                                error_message.append(note_error)
                            #替换注释前角标
                            replace_note = note.replace(note_mark[0], note_mark_number)
                            note_all.append(replace_note)
                        else:
                            #注释前没识别角标到直接拼接字符串
                            replace_note = note_mark_number + note
                            note_all.append(replace_note)
                            note_error = "书籍第 " + str(data_json[data]["pageNum"][0]) + " 页 原文应改成 " + note_mark_number + " " + data_json[data]["text"]
                            error_message.append(note_error)
        #处理不是第一页的内容
        else:
            # 提取正文的内容
            if data_json[data]["type"] != "foot" \
                and data_json[data]["type"] != "foot_pagenum" \
                and data_json[data]["type"] != "corner_note" \
                and data_json[data]["type"] != "head" \
                and data_json[data]["type"] != "note_line" \
                and data_json[data]["subType"] != "footer_note" \
                and data_json[data]["type"] != "end_note" \
                and data_json[data]["subType"] != "page_header" \
                and data_json[data]["subType"] != "none" \
                and data_json[data]["type"] != "table" \
                and data_json[data]["subType"] != "page":
                contents = data_json[data]["text"].replace("?", "？").replace("《:", ":《").replace(":》", ":》").replace(";", "；").replace("(", "（").replace(")", "）").replace(":", "：").replace(",", "，")
                contents_all.append(contents)
            #内容表格提取表格信息
            elif data_json[data]["type"] == "table":
                tables = data_json[data]
                generate_table = Generate_Table()
                contents_all.append(generate_table.Table_main(str=tables))
            else:
                #提取第一页以后得注释
                if (data_json[data]["subType"] == "footer_note" or data_json[data]["type"] == "corner_note") and data_json[data]["subType"] != "page":
                    count_constant += 1
                    note = data_json[data]["text"]
                    note_mark_number = f"[{count_constant}]"
                    #提取注释里前面的角标
                    note_mark = extract_special_characters(note)
                    if len(note_mark) > 0:
                        #查找原文里的角标是否存在
                        if check_regex_match(note_mark[0], contents_all):
                            replace_string_in_list(contents_all, note_mark[0], note_mark_number)
                        #原文中没有找到角标
                        else:
                            #这里做报错信息提示
                            note_error = "书籍第 " + str(data_json[data]["pageNum"][0]) + " 页 原文应改成 " + note_mark_number + " " + data_json[data]["text"]
                            error_message.append(note_error)
                        #替换注释前角标
                        replace_note = note.replace(note_mark[0], note_mark_number)
                        note_all.append(replace_note)
                    else:
                        #注释前没识别角标到直接拼接字符串
                        replace_note = note_mark_number + note
                        note_all.append(replace_note)
                        note_error = "书籍第 " + str(data_json[data]["pageNum"][0]) + " 页 原文应改成 " + note_mark_number + " " + data_json[data]["text"]
                        error_message.append(note_error)
                    
    content_dict = {"期刊名称＝":"","期刊年份＝":"","期刊号＝":"","起始页码＝":"","终止页码＝":"","期刊栏目＝":"","标题＝":"","英文标题＝":"","副标题＝":"", \
                    "英文副标题＝":"","作者＝":"","作者单位＝":"","摘要＝":abstracts,"英文摘要＝":"","关键字＝":"","英文关键字＝":"","基金项目＝":"","中图分类号＝":"", \
                    "文献标识码＝":"","文章编号＝":"","内容＝":contents_all,"注释＝":note_all,"参考文献＝":"","自定义中文标题＝":"","自定义英文标题＝":"","文章类型＝":"",\
                    "▲":"","错误信息":error_message}
    return content_dict
