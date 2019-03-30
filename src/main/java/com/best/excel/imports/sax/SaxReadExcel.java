package com.best.excel.imports.sax;

import com.best.excel.imports.config.ImportParams;
import com.best.excel.imports.handler.ReadRowHandler;
import com.best.excel.imports.parse.ISaxRowRead;
import com.best.excel.imports.parse.SaxRowRead;
import com.best.excel.imports.xlsx.SheetHandler;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

/**
 * sax解析处理类
 *
 * @author BG317957
 * @date 2019年3月29日19:50:48
 */
public class SaxReadExcel {

    public <T> List<T> readExcel(InputStream inputstream, Class<?> pojoClass, ImportParams params,
                                 ISaxRowRead rowRead, ReadRowHandler handler) throws Exception {
        try {
            OPCPackage opcPackage = OPCPackage.open(inputstream);
            return readExcel(opcPackage, pojoClass, params, rowRead, handler);
        } catch (Exception e) {
            throw e;
        }
    }

    private <T> List<T> readExcel(OPCPackage opcPackage, Class<?> pojoClass, ImportParams params,
                                  ISaxRowRead rowRead, ReadRowHandler hanlder) throws Exception {
        try {
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            SharedStringsTable sst = xssfReader.getSharedStringsTable();
            if (rowRead == null) {
                rowRead = new SaxRowRead(pojoClass, params, hanlder);
            }
            XMLReader parser = fetchSheetParser(sst, rowRead, xssfReader);
            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext() && sheetIndex < params.getSheetNum()) {
                sheetIndex++;
                InputStream sheet = sheets.next();
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource);
                sheet.close();
            }
            return rowRead.getList();
        } catch (Exception e) {
            throw e;
        }
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst,
                                       ISaxRowRead rowRead, XSSFReader xssfReader) throws SAXException, IOException, InvalidFormatException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst, rowRead, xssfReader);
        parser.setContentHandler(handler);
        return parser;
    }
}
