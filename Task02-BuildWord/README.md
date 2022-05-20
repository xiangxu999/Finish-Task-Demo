## 业务场景

现在由于信息化普及，很多报告采用了电子报告形式。我们往往需要根据当前的用户去获取相关信息再生成对应的报告

## 实现思路

1、根据需求的模板去制作符合poi的word模板。

2、根据唯一信息去查询到报告中需要的信息。

3、按照POI的格式去编辑数据，生成报告。

## 相关代码

**测试代码：**

```java
    @Test
    void buildWord() throws IOException {
        // 表格数据
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        Configure config = Configure.builder()
                .bind("courseList", policy).build();
        List<JSONObject> courseList = new ArrayList<>();
        JSONObject course1 = new JSONObject();
        course1.put("name", "数学");
        course1.put("teacher", "数学老师");
        course1.put("finish", "授课中");
        JSONObject course2 = new JSONObject();
        course2.put("name", "英语");
        course2.put("teacher", "英语老师");
        course2.put("finish", "授课中");
        JSONObject course3 = new JSONObject();
        course3.put("name", "语文");
        course3.put("teacher", "语文老师");
        course3.put("finish", "已结课");
        courseList.add(course1);
        courseList.add(course2);
        courseList.add(course3);
        // 图表数据
        ChartMultiSeriesRenderData chart = Charts
                .ofMultiSeries("学生成绩", new String[]{"", "", ""})
                .addSeries("数学", new Double[]{100.0, 120.5, 145.0})
                .addSeries("英语", new Double[]{121.5, 135.5, 147.0})
                .addSeries("英语", new Double[]{127.5, 136.5, 148.0})
                .create();
        XWPFTemplate template = XWPFTemplate.compile("src/main/resources/templates/学生报告模板.docx", config).render(
                new HashMap<String, Object>() {{
                    // 普通数据，根据需求要从数据库中去拿
                    put("year", "2021-2022下半");
                    put("name", "wxx");
                    put("school", "第一中学");
                    put("class", "高三二班");
                    // 图表数据
                    put("lineChart", chart);
                    // 表格数据
                    put("courseList", courseList);
                }});
        template.writeAndClose(Files.newOutputStream(Paths.get("src/main/resources/templates/out.docx")));
    }
```

> 这里相关信息是直接构造的，实际业务上需要根据条件去数据库中进行查询。