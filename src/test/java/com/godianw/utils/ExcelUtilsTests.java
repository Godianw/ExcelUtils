package com.godianw.utils;

import com.godianw.po.ExamplePo;
import org.apache.commons.beanutils.BeanUtils;
import org.junit.Test;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 示例用测试类
 * @author : Godianw
 * @version : v0.1
 */
public class ExcelUtilsTests {

    @Test
    public void testRead() {
        File file = new File("example.xlsx");
        try {
            InputStream in = new FileInputStream(file);
            Map<String, String> header2TableMap = new HashMap<>();
            header2TableMap.put("ID", "id");
            header2TableMap.put("NAME", "name");
            header2TableMap.put("COUNT", "count");
            header2TableMap.put("PRICE", "price");
            header2TableMap.put("REAL", "isReal");
            header2TableMap.put("CREATE_TIME", "createTimeStr");
            List<Map<String, String>> resultList = ExcelUtils.read(in, header2TableMap);
            for (Map<String, String> map : resultList) {
                ExamplePo examplePo = new ExamplePo();
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                examplePo.setCreateTime(sdf.parse(map.get("createTimeStr")));
                BeanUtils.populate(examplePo, map);
                System.out.println(examplePo);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testWrite() throws ParseException {
        List<ExamplePo> examplePos = new ArrayList<>();
        for (int i = 0; i < 1000; ++ i) {
            ExamplePo examplePo = new ExamplePo();
            examplePo.setId(String.valueOf(i + 1));
            examplePo.setName("name" + i);
            examplePo.setCount(i);
            examplePo.setPrice(3.25);
            examplePo.setReal(i % 2 == 0);
            examplePo.setCreateTime(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss")
                            .parse("2018-11-11 11:11:11"));
            examplePos.add(examplePo);
        }
        Map<String, String> model2HeaderMap = new HashMap<>();
        model2HeaderMap.put("id", "ID");
        model2HeaderMap.put("name", "NAME");
        model2HeaderMap.put("count", "COUNT");
        model2HeaderMap.put("price", "PRICE");
        try {
            String fileName = "E:\\exportExample.xlsx";
            File file = new File(fileName);
            if (file.exists() && file.isFile()) {
                file.delete();
            }
            OutputStream out = new FileOutputStream(file);
            ExcelUtils.write(out, fileName, model2HeaderMap, examplePos);
        } catch (IllegalAccessException | IOException e) {
            e.printStackTrace();
        }
    }
}
