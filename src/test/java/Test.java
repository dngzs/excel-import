import com.best.excel.imports.expose.DefaultImportDataUtils;
import com.google.common.collect.Lists;

import java.io.FileInputStream;
import java.util.List;

public class Test {

    public static void main(String[] args) {
        String book1 = "C:\\Users\\Administrator\\Desktop\\Book1.xls";

        String book2 = "C:\\Users\\Administrator\\Desktop\\1.xlsx";


        List<PayOrderDetailEo> List = Lists.newArrayList();
        try {
            DefaultImportDataUtils.defaultImportExcelBySax("Book1.xls"
                    , new FileInputStream(book2), PayOrderDetailEo.class,List);
            System.out.println(List.get(0));
        }catch (Exception e){
            e.printStackTrace();
        }

    }
}
