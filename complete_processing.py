from finish_submit import get_directory_name, submit_id
from get_data import query
from local_process import data_analysis

paths = "/Volumes/Datas/data_code/code/cut_pdf"

#data_id_list = submit_id(paths)
data_id_list = ["财经法学2305期正文-1.pdf:docmind-20230917-1f082774"]

execution_flags = [False] * len(data_id_list)
while not all(execution_flags):
    for i, name in enumerate(data_id_list):
        if execution_flags[i]:
                continue
        get_id = query(name.split(":")[1])
        if get_id != None:
            dicts = data_analysis(get_id)
            txt_name = name.split(":")[0].replace(".pdf", "")
            with open(rf"{paths}/{txt_name}.txt", "w", encoding="utf-8") as f:
                for key, value in dicts.items():
                    if isinstance(value, list):
                        if len(value) > 0:
                            f.write(key + value[0] + "\n")
                            for item in value[1:]:
                                f.write("    " + item + "\n")
                    else:
                        f.write(key + value + "\n")
            execution_flags[i] = True
