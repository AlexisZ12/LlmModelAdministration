from elasticsearch import Elasticsearch
from openai import OpenAI
from pywebio.input import *
from pywebio.output import *
import openpyxl
from datetime import datetime
import math
import json



def AssenblyData(EsResponse, size):
    data = []
    for i in range(0, size):
        source = EsResponse['hits']['hits'][i]['_source']
        row = [source['ID'], source['品牌'], source['比例'], source['赛季'], source['车队'], source['型号'], source['车手'], source['车号'], source['分站'], source['名次'], source['状态'], source['入库时间'], source['购入价格'], source['卖出价格'], source['备注']]
        data.append(row)
    for i in range(0, 10):
        data.append(["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])
    
    return data



def SearchAll(es, index_name):
    search = {
        "query": {
            "match_all": {}
        },
        "sort": [{"ID": {"order": "asc"}}],
        "from": 0,
        "size": 10000,
    }
    response = es.search(index = index_name, body = search)
    size = response['hits']['total']['value']
    data = AssenblyData(response, size)
    
    return data, size



def SearchInStock(es, index_name):
    search = {
        "query": {
            "term": {
                "状态": "入库"
            }
        },
        "sort": [{"ID": {"order": "asc"}}],
        "from": 0,
        "size": 10000,
    }
    response = es.search(index = index_name, body = search)
    size = response['hits']['total']['value']
    data = AssenblyData(response, size)
    
    return data, size



def SearchAdvanceSale(es, index_name):
    search = {
        "query": {
            "term": {
                "状态": "预购"
            }
        },
        "sort": [{"ID": {"order": "asc"}}],
        "from": 0,
        "size": 10000,
    }
    response = es.search(index = index_name, body = search)
    size = response['hits']['total']['value']
    data = AssenblyData(response, size)
    
    return data, size



def SearchUnsubscribe(es, index_name):
    search = {
        "query": {
            "term": {
                "状态": "退订"
            }
        },
        "sort": [{"ID": {"order": "asc"}}],
        "from": 0,
        "size": 10000,
    }
    response = es.search(index = index_name, body = search)
    size = response['hits']['total']['value']
    data = AssenblyData(response, size)
    
    return data, size



def SearchSaled(es, index_name):
    search = {
        "query": {
            "term": {
                "状态": "已售"
            }
        },
        "sort": [{"ID": {"order": "asc"}}],
        "from": 0,
        "size": 10000,
    }
    response = es.search(index = index_name, body = search)
    size = response['hits']['total']['value']
    data = AssenblyData(response, size)
    
    return data, size



def main():

    with open('config.json', 'r', encoding='utf-8') as file:
        config = json.load(file)

    setting = input_group(
        "初始化",
        [
            input("ElasticSearch地址", value=config['ESURL'], name="EsUrl"),
            input("索引名称", value=config['INDEX'], name="Index"),
            input("Api Key", value=config['KEY'], name="Key"),
            input("Base Url", value=config['BASE'], name="Base"),
            input("导出地址", value=config['PATH'], name="path")
        ])
    
    es = Elasticsearch([setting['EsUrl']])
    index_name = setting['Index']
    
    
    client = OpenAI(api_key=setting['Key'], base_url=setting['Base'])
    if not es.indices.exists(index=index_name):
        popup('提示', '索引不存在')
        exit()
    head = ["ID", "品牌", "比例", "赛季", "车队", "型号", "车手", "车号", "分站", "名次", "状态", "入库时间", "购入价格", "卖出价格", "备注"]
    total = SearchAll(es, index_name)[1]
    
    while True:
        menu = actions("", [{'label': '批量查找', 'value':1}, {'label': '定向查找', 'value':2}, {'label': '修改条目', 'value':3}, {'label': '增加条目', 'value':4}, {'label': '导出文件', 'value':5}])

        match menu:
            case 1:
                while True:
                    BatchSearch = actions("", [{'label': '查看全部', 'value':1}, {'label': '查看入库', 'value':2}, {'label': '查看预售', 'value':3}, {'label': '查看退订', 'value':4}, {'label': '查看已售', 'value':5}, {'label': '返回主菜单', 'value': 0, 'color': 'warning'}])

                    match BatchSearch:
                        case 0:
                            clear(scope="ROOT")
                            break
                        
                        case 1:
                            data, size = SearchAll(es, index_name)
                            page = 0
                            pagemax = math.ceil(size/10)
                            
                            while True:
                                showlist = []
                                showlist.append(head)
                                showlist = showlist + data[page*10: page*10+10]
                                
                                put_table(showlist)
                                action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回批量查找菜单', 'value': 0, 'color': 'warning'}])
                                
                                match action:
                                    case 0:
                                        clear(scope="ROOT")
                                        break
                                    case 1:
                                        page = page - 1
                                        if page < 0:
                                            page = 0
                                        clear(scope="ROOT")
                                    case 2:
                                        page = page + 1
                                        if page >= pagemax:
                                            page = pagemax - 1
                                        clear(scope="ROOT")
                    
                        case 2:
                            data, size = SearchInStock(es, index_name)
                            page = 0
                            pagemax = math.ceil(size/10)
                            
                            while True:
                                showlist = []
                                showlist.append(head)
                                showlist = showlist + data[page*10: page*10+10]
                                
                                put_table(showlist)
                                action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回批量查找菜单', 'value': 0, 'color': 'warning'}])
                                
                                match action:
                                    case 0:
                                        clear(scope="ROOT")
                                        break
                                    case 1:
                                        page = page - 1
                                        if page < 0:
                                            page = 0
                                        clear(scope="ROOT")
                                    case 2:
                                        page = page + 1
                                        if page >= pagemax:
                                            page = pagemax - 1
                                        clear(scope="ROOT")
                    
                        case 3:
                            data, size = SearchAdvanceSale(es, index_name)
                            page = 0
                            pagemax = math.ceil(size/10)
                            
                            while True:
                                showlist = []
                                showlist.append(head)
                                showlist = showlist + data[page*10: page*10+10]
                                
                                put_table(showlist)
                                action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回批量查找菜单', 'value': 0, 'color': 'warning'}])
                                
                                match action:
                                    case 0:
                                        clear(scope="ROOT")
                                        break
                                    case 1:
                                        page = page - 1
                                        if page < 0:
                                            page = 0
                                        clear(scope="ROOT")
                                    case 2:
                                        page = page + 1
                                        if page >= pagemax:
                                            page = pagemax - 1
                                        clear(scope="ROOT")
                    
                        case 4:
                            data, size = SearchUnsubscribe(es, index_name)
                            page = 0
                            pagemax = math.ceil(size/10)
                            
                            while True:
                                showlist = []
                                showlist.append(head)
                                showlist = showlist + data[page*10: page*10+10]
                                
                                put_table(showlist)
                                action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回批量查找菜单', 'value': 0, 'color': 'warning'}])
                                
                                match action:
                                    case 0:
                                        clear(scope="ROOT")
                                        break
                                    case 1:
                                        page = page - 1
                                        if page < 0:
                                            page = 0
                                        clear(scope="ROOT")
                                    case 2:
                                        page = page + 1
                                        if page >= pagemax:
                                            page = pagemax - 1
                                        clear(scope="ROOT")
                    
                        case 5:
                            data, size = SearchSaled(es, index_name)
                            page = 0
                            pagemax = math.ceil(size/10)
                            
                            while True:
                                showlist = []
                                showlist.append(head)
                                showlist = showlist + data[page*10: page*10+10]
                                
                                put_table(showlist)
                                action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回主菜单', 'value': 0, 'color': 'warning'}])
                                
                                match action:
                                    case 0:
                                        clear(scope="ROOT")
                                        break
                                    case 1:
                                        page = page - 1
                                        if page < 0:
                                            page = 0
                                        clear(scope="ROOT")
                                    case 2:
                                        page = page + 1
                                        if page >= pagemax:
                                            page = pagemax - 1
                                        clear(scope="ROOT")
            
            case 2:
                while True:
                    search = input_group(
                        "选择检索条件",
                        [
                            select("搜索条目", options = ["品牌", "比例", "赛季", "车队", "型号", "车手", "车号", "分站", "名次", "状态", "入库时间", "购入价格", "卖出价格", "备注", "全部", "语义查询"], name = "field"),
                            input("搜索字段", name="condition"),
                            actions("", [{'label': '确认', 'value': 1}, {'label': '返回主菜单', 'value': 0, 'color': 'warning'}], name = "action")
                        ])

                    match search['action']:
                        case 0:
                            clear(scope="ROOT")
                            break
                        
                        case 1:
                            if search['field'] == '语义查询':
                                vector = client.embeddings.create(input=search['condition'], model="text-embedding-3-small")
                                search = {
                                    "query": {
                                        "script_score": {
                                            "query": {
                                                "match_all": {}
                                            },
                                            "script": {
                                                "source": "cosineSimilarity(params.query_vector, 'SearchVector') + 1.0",
                                                "params": {
                                                    "query_vector": vector.data[0].embedding
                                                }
                                            }
                                        }
                                    },
                                    "from": 0,
                                    "size": 10000,
                                }
                                response = es.search(index = index_name, body = search)
                                size = response['hits']['total']['value']
                                data = AssenblyData(response, size)
                                page = 0
                                pagemax = math.ceil(size/10)

                                while True:                    
                                    showlist = []
                                    showlist.append(head)
                                    showlist = showlist + data[page*10: page*10+10]

                                    put_table(showlist)
                                    action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回检索', 'value': 0, 'color': 'warning'}])
                                    
                                    match action:
                                        case 0:
                                            clear(scope="ROOT")
                                            break
                                        case 1:
                                            page = page - 1
                                            if page < 0:
                                                page = 0
                                            clear(scope="ROOT")
                                        case 2:
                                            page = page + 1
                                            if page >= pagemax:
                                                page = pagemax - 1
                                            clear(scope="ROOT")

                            elif search['field'] == '全部':
                                search = {
                                    "query": {
                                        "multi_match": {
                                            "query": search['condition'],
                                            "fields": ["品牌", "比例", "赛季", "车队", "型号", "车手", "车号", "分站", "名次", "状态", "备注"]
                                        }
                                    },
                                    "sort": [{"ID": {"order": "asc"}}],
                                    "from": 0,
                                    "size": 10000,
                                }

                                response = es.search(index = index_name, body = search)
                                size = response['hits']['total']['value']
                                data = AssenblyData(response, size)
                                page = 0
                                pagemax = math.ceil(size/10)
                                
                                while True:
                                    showlist = []
                                    showlist.append(head)
                                    showlist = showlist + data[page*10: page*10+10]
                                    
                                    put_table(showlist)
                                    action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回检索', 'value': 0, 'color': 'warning'}])
                                    
                                    match action:
                                        case 0:
                                            clear(scope="ROOT")
                                            break
                                        case 1:
                                            page = page - 1
                                            if page < 0:
                                                page = 0
                                            clear(scope="ROOT")
                                        case 2:
                                            page = page + 1
                                            if page >= pagemax:
                                                page = pagemax - 1
                                            clear(scope="ROOT")

                            else:
                                search = {
                                    "query": {
                                        "match": {
                                            search['field']: search['condition']
                                        }
                                    },
                                    "sort": [{"ID": {"order": "asc"}}],
                                    "from": 0,
                                    "size": 10000,
                                }

                                response = es.search(index = index_name, body = search)
                                size = response['hits']['total']['value']
                                data = AssenblyData(response, size)
                                page = 0
                                pagemax = math.ceil(size/10)
                                
                                while True:
                                    showlist = []
                                    showlist.append(head)
                                    showlist = showlist + data[page*10: page*10+10]
                                    
                                    put_table(showlist)
                                    action = actions("", [{'label': '上一页', 'value': 1}, {'label': '下一页', 'value': 2}, {'label': '返回检索', 'value': 0, 'color': 'warning'}])
                                    
                                    match action:
                                        case 0:
                                            clear(scope="ROOT")
                                            break
                                        case 1:
                                            page = page - 1
                                            if page < 0:
                                                page = 0
                                            clear(scope="ROOT")
                                        case 2:
                                            page = page + 1
                                            if page >= pagemax:
                                                page = pagemax - 1
                                            clear(scope="ROOT")
            
            case 3:
                while True:
                    search = input_group(
                        "修改条目",
                        [
                            input("请输入id:", name = "id"),
                            actions(name="action", buttons=[{'label': '确认', 'value': 1}, {'label': '返回主菜单', 'value': 0, 'color': 'warning'}])
                        ])

                    match search['action']:
                        case 0:
                            clear(scope="ROOT")
                            break

                        case 1:
                            while True:
                                response = es.get(index = index_name, id = search['id'])
                                data = input_group(
                                    "ID: {0}".format(response['_source']['ID']),
                                    [
                                        input("品牌", value = response['_source']['品牌'], name = "brand"),
                                        input("比例", value = response['_source']['比例'], name = "scale"),
                                        input("赛季", value = response['_source']['赛季'], name = "season"),
                                        input("车队", value = response['_source']['车队'], name = "team"),
                                        input("型号", value = response['_source']['型号'], name = "type"),
                                        input("车手", value = response['_source']['车手'], name = "driver"),
                                        input("车号", value = response['_source']['车号'], name = "number"),
                                        input("分站", value = response['_source']['分站'], name = "stage"),
                                        input("名次", value = response['_source']['名次'], name = "rank"),
                                        input("状态", value = response['_source']['状态'], name = "state"),
                                        input("入库时间", value = response['_source']['入库时间'], name = "date"),
                                        input("购入价格", value = response['_source']['购入价格'], name = "price_in"),
                                        input("卖出价格", value = response['_source']['卖出价格'], name = "price_out"),
                                        input("备注", value = response['_source']['备注'], name = "remark"),
                                        actions(name="action", buttons=[{'label': '确认', 'value': 1}, {'label': '返回搜索菜单', 'value': 0, 'color': 'warning'}])
                                    ])
                                
                                match data['action']:
                                    case 0:
                                        clear(scope="ROOT")
                                        break

                                    case 1:
                                        ModelDescribe = "这是一台品牌为{0}的{1}比例的赛车模型，这台模型的原型车为{2}赛季{3}车队{4}号车手{5}驾驶的{6}。".format(data['brand'], data['scale'], data['season'], data['team'], data['number'], data['driver'], data['type'])
        
                                        if data['rank'] == "DNF":
                                            RaceDescribe = "{0}在{1}的正赛中未完赛。".format(data['driver'], data['stage'])
                                        elif data['rank'] == "Pole":
                                            RaceDescribe = "{0}在{1}的排位赛中取得杆位。".format(data['driver'], data['stage'])
                                        elif data['rank'] == "测试":
                                            RaceDescribe = "{0}在{1}参与了测试".format(data['driver'], data['stage'])
                                        else:
                                            RaceDescribe = "{0}在{1}的正赛中取得了第{2}名。".format(data['driver'], data['stage'], data['rank'])
                                        
                                        if data['state'] == "退订":
                                            StateDescribe = "这台模型已经退订"
                                        elif data['state'] == "预购":
                                            if(data['price_in'] == '0'):
                                                StateDescribe = "这台模型已经预订，预定价格未知"
                                            else:
                                                StateDescribe = "这台模型已经预订，预定价格为{0}".format(data['price_in'])
                                        elif data['state'] == "入库":
                                            StateDescribe = "这台模型已经入库，购入价格为{0}，入库时间为{1}".format(data['price_in'], data['date'])
                                        elif data['state'] == "已售":
                                            StateDescribe = "这台模型已经出售，购入价格为{0}，卖出价格为{1}，入库时间为{2}".format(data['price_in'], data['price_out'], data['date'])

                                        TextDescribe = ModelDescribe + RaceDescribe + StateDescribe + "这台模型还有如下特点：" + data['remark']
                                        vector = client.embeddings.create(
                                            input = TextDescribe,
                                            model = "text-embedding-3-small"
                                        )

                                        document = {
                                            "doc": {
                                                "品牌": data['brand'],
                                                "比例": data['scale'],
                                                "赛季": data['season'],
                                                "车队": data['team'],
                                                "型号": data['type'],
                                                "车手": data['driver'],
                                                "车号": data['number'],
                                                "分站": data['stage'],
                                                "名次": data['rank'],
                                                "状态": data['state'],
                                                "入库时间": int(data['date']),
                                                "购入价格": int(data['price_in']),
                                                "卖出价格": int(data['price_out']),
                                                "备注": data['remark'],
                                                "SearchText": TextDescribe,
                                                "SearchVector": vector.data[0].embedding
                                            }
                                        }
                                        response = es.update(index = index_name, id = search['id'], body = document)
                                        popup('提示', '修改成功')

                                        clear(scope="ROOT")
                                        break
            
            case 4:
                data = input_group(
                    "ID: {0}".format(total+1),
                    [
                        input("品牌", name = "brand"),
                        input("比例", name = "scale"),
                        input("赛季", name = "season"),
                        input("车队", name = "team"),
                        input("型号", name = "type"),
                        input("车手", name = "driver"),
                        input("车号", name = "number"),
                        input("分站", name = "stage"),
                        input("名次", name = "rank"),
                        input("状态", name = "state"),
                        input("入库时间", name = "date"),
                        input("购入价格", name = "price_in"),
                        input("卖出价格", name = "price_out"),
                        input("备注", name = "remark"),
                        actions(name="action", buttons=[{'label': '确认', 'value': 1}, {'label': '返回搜索菜单', 'value': 0, 'color': 'warning'}])
                    ])
                
                if(data['date'] == ''):
                    data['date'] = '0'
                if(data['price_out'] == ''):
                    data['price_out'] = '0'
                
                match data['action']:
                    case 0:
                        clear(scope="ROOT")
                        continue

                    case 1:
                        ModelDescribe = "这是一台品牌为{0}的{1}比例的赛车模型，这台模型的原型车为{2}赛季{3}车队{4}号车手{5}驾驶的{6}。".format(data['brand'], data['scale'], data['season'], data['team'], data['number'], data['driver'], data['type'])

                        if data['rank'] == "DNF":
                            RaceDescribe = "{0}在{1}的正赛中未完赛。".format(data['driver'], data['stage'])
                        elif data['rank'] == "Pole":
                            RaceDescribe = "{0}在{1}的排位赛中取得杆位。".format(data['driver'], data['stage'])
                        elif data['rank'] == "测试":
                            RaceDescribe = "{0}在{1}参与了测试".format(data['driver'], data['stage'])
                        else:
                            RaceDescribe = "{0}在{1}的正赛中取得了第{2}名。".format(data['driver'], data['stage'], data['rank'])
                        
                        if data['state'] == "退订":
                            StateDescribe = "这台模型已经退订"
                        elif data['state'] == "预购":
                            if(data['price_in'] == '0'):
                                StateDescribe = "这台模型已经预订，预定价格未知"
                            else:
                                StateDescribe = "这台模型已经预订，预定价格为{0}".format(data['price_in'])
                        elif data['state'] == "入库":
                            StateDescribe = "这台模型已经入库，购入价格为{0}，入库时间为{1}".format(data['price_in'], data['date'])
                        elif data['state'] == "已售":
                            StateDescribe = "这台模型已经出售，购入价格为{0}，卖出价格为{1}，入库时间为{2}".format(data['price_in'], data['price_out'], data['date'])

                        TextDescribe = ModelDescribe + RaceDescribe + StateDescribe + "这台模型还有如下特点：" + data['remark']
                        vector = client.embeddings.create(
                            input = TextDescribe,
                            model = "text-embedding-3-small"
                        )

                        data = {
                            "ID": total+1,
                            "品牌": data['brand'],
                            "比例": data['scale'],
                            "赛季": data['season'],
                            "车队": data['team'],
                            "型号": data['type'],
                            "车手": data['driver'],
                            "车号": data['number'],
                            "分站": data['stage'],
                            "名次": data['rank'],
                            "状态": data['state'],
                            "入库时间": int(data['date']),
                            "购入价格": int(data['price_in']),
                            "卖出价格": int(data['price_out']),
                            "备注": data['remark'],
                            "SearchText": TextDescribe,
                            "SearchVector": vector.data[0].embedding
                        }
                        es.index(index = index_name, id = total+1, document = data)
                        total = total + 1
                        
                        popup('提示', '录入成功')

            case 5:
                search = {
                    "query": {
                        "match_all": {}
                    },
                    "sort": [{"ID": {"order": "asc"}}],
                    "from": 0,
                    "size": 10000,
                }
                response = es.search(index = index_name, body = search)
                size = response['hits']['total']['value']
                data = []
                for i in range(0, size):
                    source = response['hits']['hits'][i]['_source']
                    row = [source['ID'], source['品牌'], source['比例'], source['赛季'], source['车队'], source['型号'], source['车手'], source['车号'], source['分站'], source['名次'], source['状态'], source['入库时间'], source['购入价格'], source['卖出价格'], source['备注']]
                    data.append(row)

                newworkbook = openpyxl.Workbook()
                newsheet = newworkbook.active
                newsheet.title = "ModelList"

                newsheet.append(head)
                for each in data:
                    newsheet.append(each)
                
                newworkbook.save('{0}\{1}.xlsx'.format(setting['path'], datetime.now().strftime('%Y%m%d_%H%M')))

                popup('提示', '导出完成，导出{0}条数据'.format(size))



if __name__ == '__main__':
    main()