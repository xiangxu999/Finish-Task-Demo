## 业务场景

在平日系统维护过程中，经常会涉及到数据导入的工作。由于对接的人员一般是老师或者政府人员，所给的数据资料一般为excel，而往往需要导入的是某一张表的主键。

比如现在需要根据一份课题评审分组Excel表，将相关课题分配到对应评审分组下。这时候你所得到的Excel表往往只有课题的相关信息，并没有课题的唯一标识。这时候就需要根据年份、状态、身份证等信息去课题表中查询到对应的课题，然后把课题的唯一标识插入到评审分组下。

## 实现思路

1、读取excel。

2、根据读取到的内容拼接适当的sql语句到对应表中查询。

3、根据查询到的数据拼接适当的sql语句插入到对应表中。

## 相关代码

**工具选取**

easyexcel：excel读取

jdbc：数据库操作

hutool：相关工具

**数据源配置**

```yaml
spring:
  datasource:
    primary:
      driverClassName: com.mysql.cj.jdbc.Driver
      username: root
      password: root
      jdbc-url: jdbc:mysql://localhost:3306/test?useUnicode=true&characterEncoding=UTF-8&useSSL=false&autoReconnect=true&failOverReadOnly=false&serverTimezone=GMT%2B8
```

```java
@Configuration
public class DataSourceConfig {
    @Bean(name = "primaryDataSource")
    @Qualifier("primaryDataSource")
    @ConfigurationProperties(prefix = "spring.datasource.primary")
    public DataSource primaryDataSource() {
        return DataSourceBuilder.create().build();
    }

    @Bean(name = "primaryJdbcTemplate")
    public JdbcTemplate primaryJdbcTemplate(@Qualifier("primaryDataSource") DataSource dataSource) {
        return new JdbcTemplate(dataSource);
    }
}
```

**excel监听器**

```java
public class NoModelDataListener extends AnalysisEventListener<Map<Integer, String>> {
    private static final Logger LOGGER = LoggerFactory.getLogger(NoModelDataListener.class);
    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 5;
    List<Map<Integer, String>> list = new ArrayList<Map<Integer, String>>();

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
        //LOGGER.info("解析到一条数据:{}", JSON.toJSONString(data));
        //list.add(data);
        //if (list.size() >= BATCH_COUNT) {
        //    saveData();
        //    list.clear();
        //}
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        //saveData();
        //LOGGER.info("所有数据解析完成！");
    }

    /**
     * 加上存储数据库
     */
    private void saveData() {
        //LOGGER.info("{}条数据，开始存储数据库！", list.size());
        //LOGGER.info("存储数据库成功！");
    }
}
```

**任务实现**

```java
@SpringBootTest
class Task01ReadExcelAndInsertDbApplicationTests {

    @Autowired
    @Qualifier("primaryJdbcTemplate")
    protected JdbcTemplate jdbcTemplate;

    @Test
    void contextLoads() {
        // 要读的excel文件目录
        String fileName = "src/main/resources/templates/学生分组名单.xlsx";

        // excel的表格（从0开始）
        Integer sheetNum = 0;

        // 分组的id
        String gid = "D217050C8B444D65B97C98A51C4E3667";

        // 开始读excel
        List<Map<Integer, String>> listMap = EasyExcel.read(fileName, new NoModelDataListener()).sheet(sheetNum).doReadSync();

        for (Map<Integer, String> data : listMap) {
            // 返回每条数据的键值对 表示所在的列 和所在列的值
            String uid = data.get(2);
            String sql = "insert into student_group_person values (?,?,?)";
            jdbcTemplate.update(sql, IdUtil.simpleUUID(), gid, uid);
        }
    }
}
```

> Excel中有多张表的时候，要注意列的对应。