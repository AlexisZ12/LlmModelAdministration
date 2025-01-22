from elasticsearch import Elasticsearch
es = Elasticsearch(["http://localhost:9200"])
index_name = "model"
index_mapping = {
    "mappings": {
        "properties": {
            "ID": {"type": "integer"},
            "品牌": {"type": "keyword"},
            "比例": {"type": "keyword"},
            "赛季": {"type": "keyword"},
            "车队": {"type": "keyword"},
            "型号": {"type": "keyword"},
            "车手": {"type": "keyword"},
            "车号": {"type": "keyword"},
            "分站": {"type": "keyword"},
            "名次": {"type": "keyword"},
            "状态": {"type": "keyword"},
            "入库时间": {"type": "integer"},
            "购入价格": {"type": "integer"},
            "卖出价格": {"type": "integer"},
            "备注": {"type": "text"},
            "SearchText": {"type": "text"},
            "SearchVector": {
                "type": "dense_vector",
                "dims": 1536,
                "index": True,
                "similarity": "cosine"
            }
        }
    }
}
es.indices.create(index=index_name, body=index_mapping, ignore=400)