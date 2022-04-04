from elasticsearch import Elasticsearch
from elasticsearch import helpers
import json
es = Elasticsearch('http://localhost:9200')


def yield_data():
    with open('lotto.json','r', encoding='utf-8') as fd:
        jdat = json.load(fd)['records']
        
    for i in jdat[0:-1]:
        yield {
            "_index": 'lotto-num',
            "_source": i,
        }

helpers.bulk(es, yield_data())