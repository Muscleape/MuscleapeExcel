package com.muscleape.excel.context;

import com.muscleape.excel.event.WriteHandler;
import com.muscleape.excel.metadata.*;
import com.muscleape.excel.support.ExcelTypeEnum;
import com.muscleape.excel.util.StyleUtil;
import com.muscleape.excel.util.WorkBookUtil;
import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.CollectionUtils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * A context is the main anchorage point of a excel writer.
 *
 * @author Muscleape
 */
@Data
public class WriteContext {

    /***
     * The sheet currently written
     */
    private Sheet currentSheet;

    /**
     * current param
     */
    private MuscleapeSheet currentSheetParam;

    /**
     * The sheet currently written's name
     */
    private String currentSheetName;

    /**
     *
     */
    private MuscleapeTable currentTable;

    /**
     * Excel type
     */
    private ExcelTypeEnum excelType;

    /**
     * POI Workbook
     */
    private Workbook workbook;

    /**
     * Final output stream
     */
    private OutputStream outputStream;

    /**
     * Written form collection
     */
    private Map<Integer, MuscleapeTable> tableMap = new ConcurrentHashMap<>();

    /**
     * Cell default style
     */
    private CellStyle defaultCellStyle;

    /**
     * Current table head  style
     */
    private CellStyle currentHeadCellStyle;

    /**
     * Current table content  style
     */
    private CellStyle currentContentCellStyle;

    /**
     * the header attribute of excel
     */
    private ExcelHeadProperty excelHeadProperty;

    private boolean needHead;

    private WriteHandler afterWriteHandler;

    public WriteHandler getAfterWriteHandler() {
        return afterWriteHandler;
    }

    public WriteContext(InputStream templateInputStream, OutputStream out, ExcelTypeEnum excelType,
                        boolean needHead, WriteHandler afterWriteHandler) throws IOException {
        this.needHead = needHead;
        this.outputStream = out;
        this.afterWriteHandler = afterWriteHandler;
        this.workbook = WorkBookUtil.createWorkBook(templateInputStream, excelType);
        this.defaultCellStyle = StyleUtil.buildDefaultCellStyle(this.workbook);

    }

    /**
     * @param sheet
     */
    public void currentSheet(MuscleapeSheet sheet) {
        if (null == currentSheetParam || currentSheetParam.getSheetNo() != sheet.getSheetNo()) {
            cleanCurrentSheet();
            currentSheetParam = sheet;
            try {
                this.currentSheet = workbook.getSheetAt(sheet.getSheetNo() - 1);
            } catch (Exception e) {
                this.currentSheet = WorkBookUtil.createSheet(workbook, sheet);
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.sheet(sheet.getSheetNo(), currentSheet);
                }
            }
            StyleUtil.buildSheetStyle(currentSheet, sheet.getColumnWidthMap());
            /** **/
            this.initCurrentSheet(sheet);
        }

    }

    private void initCurrentSheet(MuscleapeSheet sheet) {

        /** **/
        initExcelHeadProperty(sheet.getHead(), sheet.getClazz());

        initTableStyle(sheet.getTableStyle());

        initTableHead();

    }

    private void cleanCurrentSheet() {
        this.currentSheet = null;
        this.currentSheetParam = null;
        this.excelHeadProperty = null;
        this.currentHeadCellStyle = null;
        this.currentContentCellStyle = null;
        this.currentTable = null;

    }

    /**
     * init excel header
     *
     * @param head
     * @param clazz
     */
    private void initExcelHeadProperty(List<List<String>> head, Class<? extends BaseRowModel> clazz) {
        if (head != null || clazz != null) {
            this.excelHeadProperty = new ExcelHeadProperty(clazz, head);
        }
    }

    public void initTableHead() {
        if (needHead && null != excelHeadProperty && !CollectionUtils.isEmpty(excelHeadProperty.getHead())) {
            int startRow = currentSheet.getLastRowNum();
            if (startRow > 0) {
                startRow += 4;
            } else {
                startRow = currentSheetParam.getStartRow();
            }
            addMergedRegionToCurrentSheet(startRow);
            int i = startRow;
            for (; i < this.excelHeadProperty.getRowNum() + startRow; i++) {
                Row row = WorkBookUtil.createRow(currentSheet, i);
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.row(i, row);
                }
                addOneRowOfHeadDataToExcel(row, this.excelHeadProperty.getHeadByRowNum(i - startRow));
            }
        }
    }

    private void addMergedRegionToCurrentSheet(int startRow) {
        for (MuscleapeCellRange cellRangeModel : excelHeadProperty.getCellRangeModels()) {
            currentSheet.addMergedRegion(
                    new CellRangeAddress(
                            cellRangeModel.getFirstRow() + startRow,
                            cellRangeModel.getLastRow() + startRow,
                            cellRangeModel.getFirstCol(),
                            cellRangeModel.getLastCol()
                    )
            );
        }
    }

    private void addOneRowOfHeadDataToExcel(Row row, List<String> headByRowNum) {
        if (headByRowNum != null && headByRowNum.size() > 0) {
            for (int i = 0; i < headByRowNum.size(); i++) {
                Cell cell = WorkBookUtil.createCell(row, i, getCurrentHeadCellStyle(), headByRowNum.get(i));
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.cell(i, cell);
                }
            }
        }
    }

    private void initTableStyle(MuscleapeTableStyle tableStyle) {
        if (tableStyle != null) {
            this.currentHeadCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableHeadFont(),
                    tableStyle.getTableHeadBackGroundColor());
            this.currentContentCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableContentFont(),
                    tableStyle.getTableContentBackGroundColor());
        }
    }

    private void cleanCurrentTable() {
        this.excelHeadProperty = null;
        this.currentHeadCellStyle = null;
        this.currentContentCellStyle = null;
        this.currentTable = null;

    }

    public void currentTable(MuscleapeTable table) {
        if (null == currentTable || currentTable.getTableNo() != table.getTableNo()) {
            cleanCurrentTable();
            this.currentTable = table;
            this.initExcelHeadProperty(table.getHead(), table.getClazz());
            this.initTableStyle(table.getTableStyle());
            this.initTableHead();
        }

    }

    public CellStyle getCurrentHeadCellStyle() {
        return this.currentHeadCellStyle == null ? defaultCellStyle : this.currentHeadCellStyle;
    }
}


