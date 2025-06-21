### 合同模块

```text
本版本发布时间为 2021-05-01  适配jdk版本为 1.8
```

#### 1 配置
##### 1.1 添加依赖
```
<dependency>
    <groupId>com.na</groupId>
    <artifactId>na-contract</artifactId>
    <version>1.0.0</version>
</dependency>
        
或者

<dependency>
    <groupId>com.na</groupId>
    <artifactId>na-contract</artifactId>
    <version>1.0.0</version>
    <scope>system</scope>
    <systemPath>${project.basedir}/../lib/na-contract-1.0.0.jar</systemPath>
</dependency>

相关依赖

    <dependency>
        <groupId>com.openhtmltopdf</groupId>
        <artifactId>openhtmltopdf-pdfbox</artifactId>
        <version>1.0.10</version> <!-- 推荐在JDK 1.8下使用的稳定版本 -->
    </dependency>
    <dependency>
        <groupId>com.openhtmltopdf</groupId>
        <artifactId>openhtmltopdf-slf4j</artifactId>
        <version>1.0.10</version> <!-- 支持中文字体 -->
    </dependency>
```

##### 1.2 配置
```yaml
na:
  contract:
    key: 秘钥需要申请
```

##### 1.3 使用
```java
@Autowired
private NaHtmlUtil naHtmlUtil;


@GetMapping("/test")
@AnonymousAccess
public void test(){
    String htmlTempFilePath = "D:\\hetong.html";
    Map<String, Object> sourceMap = new HashMap<>();
    String targetFilePath = "D:\\hetong.pdf";
    sourceMap.put("userIdB","1111");
    sourceMap.put("nameB","2222");
    System.out.println(naHtmlUtil.renderHtmlToFile(htmlTempFilePath,sourceMap,targetFilePath));
    // 打印为false，则生成失败或没有权限
}
```


# 【注意】启动类配置
```
如果你的包名不是以com.na开头的，需要配置
@ComponentScan(basePackages = {"com.na", "com.ziji.baoming"}) // 扫描多个包路径
```
