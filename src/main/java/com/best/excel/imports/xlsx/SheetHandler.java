package com.best.excel.imports.xlsx;

import com.best.excel.enmus.CellValueType;
import com.best.excel.imports.parse.ISaxRowRead;
import com.best.excel.imports.sax.SaxReadCellEntity;
import com.google.common.collect.Lists;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.io.IOException;
import java.util.List;

/**
 * 回调接口
 *
 * @author zhangbo
 * @date 2019年3月30日12:04:26
 */
public class SheetHandler extends DefaultHandler {

    private ISaxRowRead read;
    /**
     * 共享字符串表
     */
    private SharedStringsTable sst;
    /**
     * 上一次的索引值
     */
    private String lastIndex;
    /**
     * 总行数
     */
    private int totalRows = 0;
    /**
     * 判断整行是否为空行的标记
     */
    private boolean flag = false;
    /**
     * 当前行
     */
    private int curRow = 0;
    /**
     * 当前列
     */
    private int curCol = 0;
    /**
     * T元素标识
     */
    private boolean isTElement;
    /**
     * 单元格数据类型，默认为字符串类型
     */
    private CellValueType nextDataType = CellValueType.String;

    private final DataFormatter formatter = new DataFormatter();

    /**
     * 单元格日期格式的索引
     */
    private short formatIndex;

    /**
     * 日期格式字符串
     */
    private String formatString;

    /**
     * 定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
     */
    private String preRef = null, ref = null;

    /**
     * 定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
     */
    private String maxRef = null;

    /**
     * 单元格
     */
    private StylesTable stylesTable;

    /**
     * 存储行记录的容器
     */
    private List<SaxReadCellEntity> rowList = Lists.newArrayList();

    /**
     * 单元格标识
     */
    private static final String CELLS_IDENTIFIED = "c";
    /**
     * 单元格值标识
     */
    private static final String CELLS_VALUE_IDENTIFIED = "v";
    /**
     * 行尾标识
     */
    private static final String ROW_END_IDENTIFIED = "row";
    /**
     * t标识
     */
    private static final String T_IDENTIFIED = "t";
    /**
     * 布尔类型标识
     */
    private static final String BOOLEAN_IDENTIFIED = "b";
    /**
     * 字符串类型标识
     */
    private static final String STRING_IDENTIFIED = "s";

    /**
     * 字符串类型标识
     */
    private static final String DATE_M_D_YY = "m/d/yy";


    public SheetHandler(SharedStringsTable sst, ISaxRowRead rowRead, XSSFReader xssfReader) throws IOException, InvalidFormatException {
        this.sst = sst;
        this.read = rowRead;
        this.stylesTable = xssfReader.getStylesTable();
        this.sst = xssfReader.getSharedStringsTable();
    }


    /**
     * 第一个执行
     *
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     * @throws SAXException
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        //c => 单元格
        if (CELLS_IDENTIFIED.equals(name)) {
            //前一个单元格的位置
            if (preRef == null) {
                preRef = attributes.getValue("r");
            } else {
                preRef = ref;
            }

            //当前单元格的位置
            ref = attributes.getValue("r");
            //设定单元格类型
            this.setNextDataType(attributes);
        }

        //当元素为t时
        if (T_IDENTIFIED.equals(name)) {
            isTElement = true;
        } else {
            isTElement = false;
        }
        //置空
        lastIndex = "";
    }


    /**
     * 第三个执行
     *
     * @param uri
     * @param localName
     * @param name
     * @throws SAXException
     */
    @Override
    public void endElement(String uri, String localName, String name) {

        //t元素也包含字符串
        //将单元格内容加入rowList中，在这之前先去掉字符串前后的空白符
        if (isTElement) {
            String value = lastIndex.trim();
            rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, value));
            curCol++;
            isTElement = false;
            //如果里面某个单元格含有值，则标识该行不为空行
            if (value != null && !"".equals(value)) {
                flag = true;
            }
        } else if (CELLS_VALUE_IDENTIFIED.equals(name)) {
            //v => 单元格的值，如果单元格是字符串，则v标签的值为该字符串在SST中的索引
            //根据索引值获取对应的单元格值
            String value = this.getDataValue(lastIndex.trim(), "");
            //补全单元格之间的空单元格
            if (!ref.equals(preRef)) {
                int len = countNullCell(ref, preRef);
                for (int i = 0; i < len; i++) {
                    rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, ""));
                    curCol++;
                }
            }
            curCol++;
            //如果里面某个单元格含有值，则标识该行不为空行
            if (value != null && !"".equals(value)) {
                flag = true;
            }
        } else {
            //如果标签名称为row，这说明已到行尾，调用optRows()方法
            if (ROW_END_IDENTIFIED.equals(name)) {
                //默认第一行为表头，以该行单元格数目为最大数目
                if (curRow == 0) {
                    maxRef = ref;
                }
                //补全一行尾部可能缺失的单元格
                if (maxRef != null) {
                    int len = countNullCell(maxRef, ref);
                    for (int i = 0; i <= len; i++) {
                        rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, ""));
                        curCol++;
                    }
                }
                read.parse(curRow, rowList);
                totalRows++;
                rowList.clear();
                curRow++;
                curCol = 0;
                preRef = null;
                ref = null;
                flag = false;
            }
        }
    }

    /**
     * 处理数据类型
     *
     * @param attributes
     */
    public void setNextDataType(Attributes attributes) {
        //cellType为空，则表示该单元格类型为数字
        nextDataType = CellValueType.Number;
        formatIndex = -1;
        formatString = null;
        //单元格类型
        String cellType = attributes.getValue("t");
        String cellStyleStr = attributes.getValue("s");

        //处理布尔值
        if (BOOLEAN_IDENTIFIED.equals(cellType)) {
            nextDataType = CellValueType.Boolean;
            //处理字符串
        } else if (STRING_IDENTIFIED.equals(cellType)) {
            nextDataType = CellValueType.String;
        }
        //处理日期
        if (cellStyleStr != null) {
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();

            if (formatString.contains(DATE_M_D_YY)) {
                nextDataType = CellValueType.Date;
            }

            if (formatString == null) {
                nextDataType = CellValueType.Null;
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
        }
    }


    /**
     * 对解析出来的数据进行类型处理
     *
     * @param value   单元格的值，
     *                value代表解析：BOOL的为0或1， ERROR的为内容值，FORMULA的为内容值，INLINESTR的为索引值需转换为内容值，
     *                SSTINDEX的为索引值需转换为内容值， NUMBER为内容值，DATE为内容值
     * @param thisStr 一个空字符串
     * @return
     */
    @SuppressWarnings("deprecation")
    public String getDataValue(String value, String thisStr) {
        switch (nextDataType) {
            // 这几个的顺序不能随便交换，交换了很可能会导致数据错误
            //布尔值
            case Boolean:
                char first = value.charAt(0);
                thisStr = first == '0' ? "FALSE" : "TRUE";
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.Boolean, thisStr));
                break;
            //字符串
            case String:
                String sstIndex = value;
                try {
                    int idx = Integer.parseInt(sstIndex);
                    //根据idx索引值获取内容值
                    XSSFRichTextString xss = new XSSFRichTextString(sst.getEntryAt(idx));
                    thisStr = xss.toString();
                    //clear
                    xss = null;
                } catch (NumberFormatException ex) {
                    thisStr = value;
                }
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.String, thisStr));
                break;
            //数字
            case Number:
                if (formatString != null) {
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                } else {
                    thisStr = value;
                }
                thisStr = thisStr.replace("_", "").trim();
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.Number, thisStr));
                break;
            //日期
            case Date:
                thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                // 对日期字符串作特殊处理，去掉T
                thisStr = thisStr.replace("T", " ");
                rowList.add(curCol, new SaxReadCellEntity(CellValueType.Date, thisStr));
                break;
            default:
                thisStr = " ";
                break;
        }
        return thisStr;
    }

    public int countNullCell(String ref, String preRef) {
        //excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
        String xfd = ref.replaceAll("\\d+", "");
        String xfd_1 = preRef.replaceAll("\\d+", "");

        xfd = fillChar(xfd, 3, '@', true);
        xfd_1 = fillChar(xfd_1, 3, '@', true);

        char[] letter = xfd.toCharArray();
        char[] letter_1 = xfd_1.toCharArray();
        int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
        return res - 1;
    }

    public String fillChar(String str, int len, char let, boolean isPre) {
        int len_1 = str.length();
        if (len_1 < len) {
            if (isPre) {
                for (int i = 0; i < (len - len_1); i++) {
                    str = let + str;
                }
            } else {
                for (int i = 0; i < (len - len_1); i++) {
                    str = str + let;
                }
            }
        }
        return str;
    }


    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        //得到单元格内容的值
        lastIndex += new String(ch, start, length);
    }

}
