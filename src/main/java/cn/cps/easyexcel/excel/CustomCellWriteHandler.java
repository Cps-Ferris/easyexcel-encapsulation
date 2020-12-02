package cn.cps.easyexcel.excel;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.ehcache.impl.internal.classes.commonslang.ArrayUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import java.util.Arrays;
import java.util.List;

/**
 * @Author: Cai Peishen
 * @Date: 2020/11/20 15:01
 * @Description: EasyExcel写入拦截器
 */
public class CustomCellWriteHandler implements CellWriteHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(CustomCellWriteHandler.class);

    /**
     * 合并sheet的下标
     */
    private int[] mergeSheetIndex;

    /**
     * 合并字段的下标
     */
    private int[] mergeColumnIndex;

    /**
     * 起始结束合并行
     */
    private long mergeBeginRowIndex;
    private long mergeEndRowIndex;


    public CustomCellWriteHandler() {
    }

    public CustomCellWriteHandler(int[] mergeSheetIndex, int[] mergeColumnIndex, long mergeBeginRowIndex, long mergeEndRowIndex) {
        this.mergeSheetIndex = mergeSheetIndex;
        this.mergeColumnIndex = mergeColumnIndex;
        this.mergeBeginRowIndex = mergeBeginRowIndex;
        this.mergeEndRowIndex = mergeEndRowIndex;
    }

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {
    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell,
                                Head head, Integer relativeRowIndex, Boolean isHead) {
    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {

        Object curData = cell.getCellTypeEnum() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue();

        // 获取当前sheet第几页
        Integer sheetNo = writeSheetHolder.getSheetNo();

        // 这里可以对cell进行任何操作
        //LOGGER.info("{} sheet - {}行 - {}列 - {}", sheetNo, cell.getRowIndex(), cell.getColumnIndex(), curData);

        boolean isContainSheet = false;
        for(int curSheet : mergeSheetIndex) {
            if(curSheet == sheetNo) {
                isContainSheet = true;
                break;
            }
        }

        // Arrays.asList(mergeSheetIndex).contains(sheetNo.intValue()) 这个方式有问题
        if(isContainSheet){
            //当前行 当前列
            int curRowIndex = cell.getRowIndex();
            int curColIndex = cell.getColumnIndex();
            if(curRowIndex >= mergeBeginRowIndex && curRowIndex <= mergeEndRowIndex){
                // 哪几个sheet进行合并
                if (!StringUtils.isEmpty(curData)) {
                    for (int i = 0; i < mergeColumnIndex.length; i++) {
                        if (curColIndex == mergeColumnIndex[i]) {
                            mergeWithPrevRow(writeSheetHolder, cell, curRowIndex, curColIndex);
                            break;
                        }
                    }
                }
            }
        }

    }

    /**
     * 当前单元格向上合并
     *
     * @param writeSheetHolder
     * @param cell             当前单元格
     * @param curRowIndex      当前行
     * @param curColIndex      当前列
     */
    private void mergeWithPrevRow(WriteSheetHolder writeSheetHolder, Cell cell, int curRowIndex, int curColIndex) {
        //获取当前行的当前列的数据和上一行的当前列列数据，通过上一行数据是否相同进行合并
        Object curData = cell.getCellTypeEnum() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue();
        Cell preCell = null;

        // 获取当前sheet第几页
        Integer sheetNo = writeSheetHolder.getSheetNo().intValue();

        try {
            Row row = cell.getSheet().getRow(curRowIndex - 1);
            if(row==null){
                //LOGGER.info("《《《《《《《《《《《   合并异常... {}sheet - {}行 - {}列 - {}", sheetNo, cell.getRowIndex(), cell.getColumnIndex(), curData);
                return;
            }
            preCell = row.getCell(curColIndex);

            Object preData = preCell.getCellTypeEnum() == CellType.STRING ? preCell.getStringCellValue() : preCell.getNumericCellValue();
            // 比较当前行的第一列的单元格与上一行是否相同，相同合并当前单元格与上一行
            if (curData.equals(preData)) {

                //LOGGER.info("合并数据... {} sheet - {}行 - {}列 - {}", sheetNo, cell.getRowIndex(), cell.getColumnIndex(), curData);

                Sheet sheet = writeSheetHolder.getSheet();
                List<CellRangeAddress> mergeRegions = sheet.getMergedRegions();

                boolean isMergedBefore = false;

                // 如果是第一列，上下单元格内容相同则合并
                // 如果非第一列，上下单元格内容相同并且第一列是合并的情况则合并

                // 以第一列为准，当前行的第一列合并再合并
                int colNorm = 0;

                if(curColIndex == colNorm){
                    isMergedBefore = true;
                }else {
                    for(int i = 0; i < mergeRegions.size(); i++){
                        CellRangeAddress cellRangeAddr = mergeRegions.get(i);
                        if(cellRangeAddr.isInRange(curRowIndex, colNorm)) {
                            isMergedBefore = true;
                            break;
                        }
                    }
                }

                if(isMergedBefore) {
                    boolean isMerged = false;
                    for (int i = 0; i < mergeRegions.size() && !isMerged; i++) {
                        CellRangeAddress cellRangeAddr = mergeRegions.get(i);
                            // 若上一个单元格已经被合并，则先移出原有的合并单元，再重新添加合并单元
                            if (cellRangeAddr.isInRange(curRowIndex - 1, curColIndex)) {
                                sheet.removeMergedRegion(i);
                                cellRangeAddr.setLastRow(curRowIndex);
                                sheet.addMergedRegion(cellRangeAddr);
                                isMerged = true;
                            }
                        }
                    // 若上一个单元格未被合并，则新增合并单元
                    if (!isMerged) {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(curRowIndex - 1, curRowIndex, curColIndex,
                                curColIndex);
                        sheet.addMergedRegion(cellRangeAddress);
                    }
                }
            }
        }catch (Exception e) {
            LOGGER.info("EasyExcel合并单元格出现异常:{}", e.getStackTrace());
        }
    }

}