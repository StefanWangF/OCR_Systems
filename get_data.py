from typing import List
from alibabacloud_docmind_api20220711.client import Client as docmind_api20220711Client
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_docmind_api20220711 import models as docmind_api20220711_models
from alibabacloud_tea_util.client import Client as UtilClient
from alibabacloud_credentials.client import Client as CredClient
import simplejson

def query(id_name):
    # 使用默认凭证初始化Credentials Client。
    cred=CredClient()
    config = open_api_models.Config(
        # 通过credentials获取配置中的AccessKey ID
        access_key_id=cred.get_access_key_id(),
        # 通过credentials获取配置中的AccessKey Secret
        access_key_secret=cred.get_access_key_secret()
    )
    # 访问的域名
    config.endpoint = f'docmind-api.cn-hangzhou.aliyuncs.com'
    client = docmind_api20220711Client(config)
    request = docmind_api20220711_models.GetDocStructureResultRequest(
        # id :  任务提交接口返回的id
        id = id_name
    )
    try:
        # 复制代码运行请自行打印 API 的返回值
        response = client.get_doc_structure_result(request)
        # API返回值格式层级为 body -> data -> 具体属性。可根据业务需要打印相应的结果。获取属性值均以小写开头
        # 获取异步任务处理情况,可根据response.body.completed判断是否需要继续轮询结果
        #print(response.body.completed)
        # 获取返回结果。建议先把response.body.data转成json，然后再从json里面取具体需要的值。
        print(response.body)
        data = response.body.data
        return data

    except Exception as error:
        # 如有需要，请打印 error
        UtilClient.assert_as_string(error.message)
