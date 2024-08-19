import sys

import boto3
from botocore import UNSIGNED
from botocore.client import Config

# Yandex object storage (S3)
s3_kwargs = {
    "endpoint_url": "https://storage.yandexcloud.net/",
    "region_name":"ru-central1"
}

# Создаем клиента S3, который использует анонимную сигнатуру (создает анонимные запросы к хранилищу)
s3_client = boto3.client('s3', config=Config(signature_version=UNSIGNED), **s3_kwargs)

# Функция пишет в стандартный вывод url адреса файлов, хранящихся в публичном бакете. Все нужные url, которые будут 
# отбираться в макросе через перехват стандартного вывода, перемешаны с лишними.
def create_urls(client, bucket_name):
    url_for_list = client.list_objects_v2(Bucket=bucket_name)
    for key in url_for_list['Contents']:
        sys.stdout.write(f'https://storage.yandexcloud.net/{bucket_name}/{key["Key"]}\n')
    sys.stdout.flush()
    
if __name__ == "__main__":
    create_urls(s3_client, 'api.hh.ru')
