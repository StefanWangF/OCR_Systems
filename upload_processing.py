from alibabacloud_docmind_api20220711.client import Client as docmind_api20220711Client
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_docmind_api20220711 import models as docmind_api20220711_models
from alibabacloud_tea_util.client import Client as UtilClient
from alibabacloud_tea_util import models as util_models
from alibabacloud_credentials.client import Client as CredClient

def submit_file(filename):
    # 使用默认凭证初始化Credentials Client。
    cred=CredClient()
    config = open_api_models.Config(
        # 通过credentials获取配置中的AccessKey ID
        access_key_id=cred.get_access_key_id(),
        # 通过credentials获取配置中的AccessKey Secret
        access_key_secret=cred.get_access_key_secret()
    )
    #访问的域名
    config.endpoint = f'docmind-api.cn-hangzhou.aliyuncs.com'
    client = docmind_api20220711Client(config)
    request = docmind_api20220711_models.SubmitDocStructureJobAdvanceRequest(
        # file_url_object : 本地文件流
        file_url_object=open(filename, "rb"),
        # file_name ：文件名称。名称必须包含文件类型
        #file_name='3.pdf',
        # file_name_extension : 文件后缀格式。与文件名二选一
        file_name_extension='pdf'
    )
    runtime = util_models.RuntimeOptions()
    try:
        # 复制代码运行请自行打印 API 的返回值
        response = client.submit_doc_structure_job_advance(request, runtime)
        # API返回值格式层级为 body -> data -> 具体属性。可根据业务需要打印相应的结果。如下示例为打印返回的业务id格式
        # 获取属性值均以小写开头，
        data = response.body.data.id
        return data
    except Exception as error:
        # 如有需要，请打印 error
        UtilClient.assert_as_string(error.message)
        #print(error.message)
    