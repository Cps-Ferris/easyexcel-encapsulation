package cn.cps.easyexcel.test;

import cn.cps.easyexcel.excel.CustomCellWriteHandler;
import cn.cps.easyexcel.excel.ExcelUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created with IntelliJ IDEA
 *
 * @Author yuanhaoyue swithaoy@gmail.com
 * @Description
 * @Date 2018-06-05
 * @Time 16:56
 */
@RestController
public class ExcelController {
    /**
     * 读取 Excel（允许多个 sheet）
     */
    @RequestMapping(value = "readExcelWithSheets", method = RequestMethod.POST)
    public Object readExcelWithSheets(MultipartFile excel) {
        return ExcelUtil.readExcel(excel, new ImportInfo());
    }

    /**
     * 读取 Excel（指定某个 sheet）
     */
    @RequestMapping(value = "readExcel", method = RequestMethod.POST)
    public Object readExcel(MultipartFile excel, int sheetNo,
                            @RequestParam(defaultValue = "1") int headLineNum) {
        return ExcelUtil.readExcel(excel, new ImportInfo(), sheetNo, headLineNum);
    }

    /**
     * 导出 Excel（一个 sheet）
     */
    @RequestMapping(value = "writeExcel", method = RequestMethod.GET)
    public void writeExcel(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName = "第一个 sheet";

        ExcelUtil.writeExcel(response, list, fileName, sheetName, new ExportInfo());
    }

    /**
     * 导出 Excel（多个 sheet）
     */
    @RequestMapping(value = "writeExcelWithSheets", method = RequestMethod.GET)
    public void writeExcelWithSheets(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName1 = "第一个 sheet";
        String sheetName2 = "第二个 sheet";
        String sheetName3 = "第三个 sheet";

        ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
    }


    /**
     * 填充 Excel（根据上下行是否相同合并单元格）
     */
    @RequestMapping(value = "fillExcel", method = RequestMethod.GET)
    public void fillExcel(HttpServletResponse response) throws IOException {

        try {
            //设置输入流，设置响应域
            response.setContentType("application/ms-excel");
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode("填充Excel导出.xlsx","utf-8");
            response.setHeader("Content-disposition","attachment;filename="+fileName);

            // 需要合并的sheet
            int[] mergeSheetIndex = {0};

            // 需要合并的列
            int[] mergeColumnIndex = {0, 1, 2, 10};

            // 根据需要合并的数据做一个结尾（目的让合并在我们想结束位置 就结束合并）
            long mergeBeginRowIndex = 7;
            long mergeEndRowIndex = mergeBeginRowIndex + getFillList().size();

            String templatePath = this.getClass().getResource("/template/fill.xlsx").getPath();

            ExcelWriter excelWriter =
                    EasyExcel.write(new BufferedOutputStream(response.getOutputStream()))
                            //07的excel版本,节省内存
                            .excelType(ExcelTypeEnum.XLSX)
                            //是否自动关闭输入流
                            .autoCloseStream(Boolean.TRUE)
                            .registerWriteHandler(new CustomCellWriteHandler(mergeSheetIndex, mergeColumnIndex, mergeBeginRowIndex, mergeEndRowIndex))
                            .withTemplate("D://fill.xlsx").build();
//              // 自定义列宽度，有数字会
//                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
            // 设置excel保护密码
//                .password("123456")
            WriteSheet writeSheet = EasyExcel.writerSheet(0).build();
            WriteSheet writeSheet1 = EasyExcel.writerSheet(1).build();
            excelWriter.fill(getSummary(), writeSheet);
            excelWriter.fill(getFillList(), writeSheet);
            excelWriter.fill(getList(), writeSheet1); //这里只是演示 可以填充多个sheet
            // 千万别忘记关闭流
            excelWriter.finish();
        }catch (Exception e) {
            e.printStackTrace();
        }

    }


    private List<ExportInfo> getList() {
        List<ExportInfo> list = new ArrayList<>();
        ExportInfo model1 = new ExportInfo();
        model1.setName("howie");
        model1.setAge("19");
        model1.setAddress("123456789");
        model1.setEmail("123456789@gmail.com");
        list.add(model1);
        ExportInfo model2 = new ExportInfo();
        model2.setName("harry");
        model2.setAge("20");
        model2.setAddress("198752233");
        model2.setEmail("198752233@gmail.com");
        list.add(model2);
        return list;
    }

    private Map<String,String> getSummary(){
        Map<String,String> map = new HashMap<>();
        map.put("keYongMoney", "10000");
        map.put("orderMoney", "9884");
        map.put("ensureMoney", "1000");
        map.put("date", "2020-12-12");
        return map;
    }

    private List<FillInfo> getFillList() {
        List<FillInfo> list = new ArrayList<>();
        FillInfo model1 = new FillInfo();
        model1.setOrderNo("NX001");
        model1.setPeiMoney("500");
        model1.setYuMoney("100");
        model1.setGoodsName("商品1");
        model1.setHeTongMoney("100000");
        model1.setHeTongCount("100");
        model1.setFaHuoMoney("910");
        model1.setFaHuoCount("90");
        model1.setJieSuanMoney("510");
        model1.setJieSuanCount("50");
        model1.setStatus("进行中");
        list.add(model1);
        FillInfo model2 = new FillInfo();
        model2.setOrderNo("NX001");
        model2.setPeiMoney("500");
        model2.setYuMoney("100");
        model2.setGoodsName("商品2");
        model2.setHeTongMoney("665800");
        model2.setHeTongCount("120");
        model2.setFaHuoMoney("300");
        model2.setFaHuoCount("60");
        model2.setJieSuanMoney("200");
        model2.setJieSuanCount("60");
        model2.setStatus("进行中");
        list.add(model2);
        FillInfo model3 = new FillInfo();
        model3.setOrderNo("NX002");
        model3.setPeiMoney("1000");
        model3.setYuMoney("900");
        model3.setGoodsName("商品1");
        model3.setHeTongMoney("486518");
        model3.setHeTongCount("3060");
        model3.setFaHuoMoney("300");
        model3.setFaHuoCount("60");
        model3.setJieSuanMoney("200");
        model3.setJieSuanCount("60");
        model3.setStatus("订单完成");
        list.add(model3);
        FillInfo model4 = new FillInfo();
        model4.setOrderNo("NX002");
        model4.setPeiMoney("1000");
        model4.setYuMoney("900");
        model4.setGoodsName("商品2");
        model4.setHeTongMoney("486518");
        model4.setHeTongCount("3060");
        model4.setFaHuoMoney("300");
        model4.setFaHuoCount("60");
        model4.setJieSuanMoney("200");
        model4.setJieSuanCount("60");
        model4.setStatus("订单完成");
        list.add(model4);
        return list;
    }
}
