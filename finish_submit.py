from upload_processing import submit_file
import os
import datetime

def get_directory_name(path):
    #便利目录获取pdf文件名称
    directory_path = rf"{path}"
    pdf_file_name = []
    for filename in os.listdir(directory_path):
        if filename.lower().endswith(".pdf") and os.path.isfile(os.path.join(directory_path, filename)):
            pdf_file_name.append(filename)
    return pdf_file_name

def submit_id(path):
    pdf_file_name = get_directory_name(path)
    execution_flags = [False] * len(pdf_file_name)
    data_id_list = []
    while not all(execution_flags):
        for i, name in enumerate(pdf_file_name):
            if execution_flags[i]:
                continue
            data_id = submit_file(rf"{path}/%s" % (name))
            if len(data_id) > 20:
                data_id_list.append(name +":"+ data_id)
                times = datetime.datetime.now().strftime("%Y-%m-%d")
                seetime= datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(rf"{times}.log", "a", encoding="utf-8") as f:
                    f.write(seetime + "  " + name + ":" + data_id + "\n")
                execution_flags[i] = True
    return data_id_list
