package me.xu.task01;

import cn.hutool.core.util.IdUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.fastjson.JSON;
import me.xu.task01.listener.NoModelDataListener;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.jdbc.core.JdbcTemplate;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

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
