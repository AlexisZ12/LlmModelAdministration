# 使用教程

## 1. 解压数据库
解压elasticsearch-8.15.0-windows-x86_64.zip到当前位置

![pic1](pic\pic1.png)

## 2. 修改数据库参数
编辑文件：elasticsearch-8.15.0\config\elasticsearch.yml

![pic2](pic\pic2.png)

按照以下模板更改参数，将四个含enabled的设置后的true改为false

```
# Enable security features
xpack.security.enabled: false

xpack.security.enrollment.enabled: false

# Enable encryption for HTTP API client connections, such as Kibana, Logstash, and Agents
xpack.security.http.ssl:
  enabled: false
  keystore.path: certs/http.p12

# Enable encryption and mutual authentication between cluster nodes
xpack.security.transport.ssl:
  enabled: false
  verification_mode: certificate
  keystore.path: certs/transport.p12
  truststore.path: certs/transport.p12
```

## 3. 启动数据库

鼠标双击运行：elasticsearch-8.15.0\bin\elasticsearch.bat

## 4. 初始化数据库

鼠标双击运行：Initialize.exe

## 5. 修改个性化参数

编辑文件：config.json

```
{
    "KEY":"",
    "BASE":"",
    "ESURL":"http://localhost:9200",
    "INDEX":"model",
    "PATH": "log"
}
```

如无特殊需求只需在KEY和BASE后的引号中填入你的openai的key和base，openai的中转服务可以在网上购买

## 6. 运行程序

鼠标双击运行：LlmModelAdministrationClient.exe