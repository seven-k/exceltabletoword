import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

/**
 * @author yin.cl
 * @since 2019/3/9 11:44
 */
public class App {
    public static void main(String[] args) {
        App app = new App();
        List<User> userList = app.readExcel("/Users/yiyezhiqiu/Desktop/df/副本花名册.xlsx", 2, 0);
        app.readWord("/Users/yiyezhiqiu/Desktop/df/个人备案.docx", userList);
    }

    /**
     * 读写world
     * @param filePath word文件路径
     */
    public void readWord(String filePath, List<User> userList) {
        if (userList == null || userList.size() == 0) {
            return;
        }
        try {
            //载入文档 //如果是office2007  docx格式
            for (User user : userList) {
                String userName = user.getUserName();
                System.out.println("正常处理..." + userName);
                FileInputStream in = new FileInputStream(filePath);
                int dotAt = filePath.lastIndexOf(".");
                StringBuilder sb = new StringBuilder(filePath);
                sb.replace(dotAt, dotAt, "_" + userName + "_" + Utils.getCurrentTimeStr2());
                File outFile = new File(sb.toString());
                OutputStream outputStream = new FileOutputStream(outFile);
                if (filePath.toLowerCase().endsWith("docx")) {
                    //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
                    //得到word文档的信息
                    XWPFDocument xwpf = new XWPFDocument(in);
                    //得到word中的表格
                    XWPFTable table = xwpf.getTables().get(0);
                    //Iterator<XWPFTable> it = xwpf.getTablesIterator(); 另一种方式获取word中的所有表格
                    List<XWPFTableRow> rows = table.getRows();
                    // 第二行
                    XWPFTableRow table2 = rows.get(2);
                    List<XWPFTableCell> table2Cells = table2.getTableCells();
                    table2Cells.get(1).setText(user.getUserName());
                    table2Cells.get(3).setText(user.getSex());
                    table2Cells.get(5).setText(user.getBirthday());

                    XWPFTableRow table3 = rows.get(3);
                    List<XWPFTableCell> table3Cells = table3.getTableCells();
                    table3Cells.get(1).setText(user.getCultureLevel());
                    table3Cells.get(3).setText(user.getPhone());

                    XWPFTableRow table5 = rows.get(5);
                    List<XWPFTableCell> table5Cells = table5.getTableCells();
                    table5Cells.get(1).setText(user.getAddress());

                    xwpf.write(outputStream);
                    outputStream.close();
                    xwpf.close();
                    in.close();
/*
                    //读取每一行数据
                    for (int i = 1; i < rows.size(); i++) {
                        XWPFTableRow row = rows.get(i);
                        //读取每一列数据
                        List<XWPFTableCell> cells = row.getTableCells();
                        for (XWPFTableCell cell : cells) {
                            //输出当前的单元格的数据
                            //cell.setText();
                            System.out.println(cell.getText());
                        }
                    }*/

                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取Excel
     *
     * @param filePath
     * @param sheetAt
     * @param columnAt
     * @return
     */
    private List<User> readExcel(String filePath, int sheetAt, int columnAt) {
        int length = filePath.split("\\.").length;
        if (length < 2) {
            System.out.println("文件格式错误");
            return Lists.newArrayList();
        }

        //新建输出文件
        //int dotAt = filePath.lastIndexOf(".");
        //StringBuilder sb = new StringBuilder(filePath);
        //sb.replace(dotAt, dotAt, "_" + Utils.getCurrentTimeStr2());
        //File outFile = new File(sb.toString());
        List<User> userList = Lists.newArrayList();
        try {
            FileInputStream fileInputStream = new FileInputStream(filePath);
            //OutputStream outputStream = new FileOutputStream(outFile);
            Workbook wb = null;
            if (filePath.toLowerCase().endsWith("xls")) {
                wb = new HSSFWorkbook(fileInputStream);
            } else if (filePath.toLowerCase().endsWith("xlsx")) {
                wb = new XSSFWorkbook(fileInputStream);
            }
            Sheet sheet = wb.getSheetAt(sheetAt - 1);
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 6; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                System.out.println("current row:" + i);
                if (Utils.isNotEmpty(getStringValueFromCell(row.getCell(1)))) {
                    User user = User.builder()
                            .userName(getStringValueFromCell(row.getCell(1)))
                            .sex(getStringValueFromCell(row.getCell(2)))
                            .birthday(getStringValueFromCell(row.getCell(3)))
                            .cultureLevel(getStringValueFromCell(row.getCell(4)))
                            .phone(getStringValueFromCell(row.getCell(7)))
                            .address(getStringValueFromCell(row.getCell(8))).build();
                    userList.add(user);
                }

            }
            //System.out.println(userList);
            //wb.write(outputStream);
            //outputStream.close();
            wb.close();
            fileInputStream.close();
            //jTextAreaInfo.append("[" + Utils.getCurrentTimeStr() + "]完成...");
            //jTextAreaInfo.setCaretPosition(jTextAreaInfo.getLineCount());


        } catch (Exception ex) {
            System.out.println("出错了：" + ex.getMessage());
            //jTextAreaInfo.append("readExcel_Exception:" + ex.getMessage() + "\n");
        }
        System.out.println("Total:"+userList.size());
        return userList;

    }

    /**
     * Excel单元格值转String
     *
     * @param cell 单元格
     * @return 单元格值
     */
    public String getStringValueFromCell(Cell cell) {
        SimpleDateFormat sFormat = new SimpleDateFormat("yyyy-MM-dd");
        DecimalFormat decimalFormat = new DecimalFormat("#.#");
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            cellValue = cell.getStringCellValue();
        } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                double d = cell.getNumericCellValue();
                Date date = HSSFDateUtil.getJavaDate(d);
                cellValue = sFormat.format(date);
            } else {
                cellValue = decimalFormat.format((cell.getNumericCellValue()));
            }
        } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            cellValue = "";
        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
            cellValue = "";
        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
            cellValue = cell.getCellFormula().toString();
        }
        return cellValue;
    }


}
