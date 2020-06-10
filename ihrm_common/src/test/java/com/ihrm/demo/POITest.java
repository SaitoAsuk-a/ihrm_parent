package com.ihrm.demo;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;

/**
 * @author liyu
 * @date 2020/6/9 11:10
 * @description
 */
public class POITest {

    @Test
    public void simpleWrite() throws IOException {

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("test");

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("test");

        FileOutputStream fos = new FileOutputStream("E:\\test\\test.xlsx");
        workbook.write(fos);
        fos.close();
    }

    /**
     * 样式处理
     *
     * @throws IOException
     */
    @Test
    public void simpleWriteFormat() throws IOException {

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("test");

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("test");

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);

        Font font = workbook.createFont();
        font.setFontName("华文行楷");

        //行高
        row.setHeightInPoints(50);
        //列宽(字符宽度)
        sheet.setColumnWidth(2, 31 * 256);

        //剧中显示
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cell.setCellStyle(cellStyle);

        FileOutputStream fos = new FileOutputStream("E:\\test\\test2.xlsx");
        workbook.write(fos);
        fos.close();
    }

    /**
     * 设置图片
     *
     * @throws IOException
     */
    @Test
    public void simpleWritePic() throws IOException {

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("test");

        //读取图片流
        FileInputStream fileInputStream = new FileInputStream("E:\\test\\rei-amagai-img14252.jpg");
        //转化为二进制数组
        byte[] bytes = IOUtils.toByteArray(fileInputStream);
        fileInputStream.read(bytes);
        //向poi内存添加一张图片，返回索引
        int i = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        //绘制图片工具类
        CreationHelper creationHelper = workbook.getCreationHelper();
        //创建一个绘图对象
        Drawing<?> patriarch = sheet.createDrawingPatriarch();
        //创建锚点，设置图片坐标
        ClientAnchor clientAnchor = creationHelper.createClientAnchor();
        clientAnchor.setRow1(1);
        clientAnchor.setCol1(1);
        //绘制图片
        Picture picture = patriarch.createPicture(clientAnchor, i);
        //自适应渲染图片
        picture.resize();

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("test");

        FileOutputStream fos = new FileOutputStream("E:\\test\\test3.xlsx");
        workbook.write(fos);
        fos.close();
    }

    /**
     * 简单读取
     *
     * @throws IOException
     */
    @Test
    public void simpleRead() throws IOException {

        //1.根据Excel文件创建工作簿
        Workbook wb = new XSSFWorkbook("H:\\heima\\26-传统行业SaaS解决方案\\08-员工管理及POI\\08-员工管理及POI\\01-员工管理及POI入门\\资源\\资源\\Excel相关\\demo.xlsx");

        //2.获取Sheet
        Sheet sheet = wb.getSheetAt(0);//参数：索引

        //3.获取Sheet中的每一行，和每一个单元格
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {

            Row row = sheet.getRow(rowNum);//根据索引获取每一个行

            StringBuilder sb = new StringBuilder();
            for (int cellNum = 2; cellNum < row.getLastCellNum(); cellNum++) {

                //根据索引获取每一个单元格
                Cell cell = row.getCell(cellNum);

                //获取每一个单元格的内容
                Object value = getCellValue(cell);

                sb.append(value).append("-");
            }

            System.out.println(sb.toString());
        }
    }

    /**
     * 百万级数据基于事件驱动读取
     * @throws IOException
     */
    @Test
    public void simpleReadEvent() throws IOException, OpenXML4JException, SAXException {

        String path = "C:\\Users\\ThinkPad\\Desktop\\ihrm\\day8\\资源\\百万数据报表\\demo.xlsx";

        //1.根据excel报表获取OPCPackage
        OPCPackage opcPackage = OPCPackage.open(path, PackageAccess.READ);
        //2.创建XSSFReader
        XSSFReader reader = new XSSFReader(opcPackage);
        //3.获取SharedStringTable对象
        SharedStringsTable table = reader.getSharedStringsTable();
        //4.获取styleTable对象
        StylesTable stylesTable = reader.getStylesTable();
        //5.创建Sax的xmlReader对象
        XMLReader xmlReader = XMLReaderFactory.createXMLReader();
        //6.注册事件处理器
        XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(stylesTable,table,new SheetHandler(),false);
        xmlReader.setContentHandler(xmlHandler);
        //7.逐行读取
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();
        while (sheetIterator.hasNext()) {
            InputStream stream = sheetIterator.next(); //每一个sheet的流数据
            InputSource is = new InputSource(stream);
            xmlReader.parse(is);
        }
    }

    public static Object getCellValue(Cell cell) {
        //1.获取到单元格的属性类型
        CellStyle cellStyle = cell.getCellStyle();
        String dataFormatString = cellStyle.getDataFormatString();
        CellType cellType = cell.getCellType();
        //2.根据单元格数据类型获取数据
        Object value = null;
        switch (cellType) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    //日期格式
                    value = cell.getDateCellValue();
                } else {
                    //数字
                    value = cell.getNumericCellValue();
                }
                break;
            case FORMULA: //公式
                value = cell.getCellFormula();
                break;
            default:
                break;
        }
        return value;
    }
}
