import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.util.*;

public class ExcelCreation {
    public static void main(String[] args) throws IOException {

//        String companyName = "AAPL";
//        String apiKey = "7759164af885a77ae927b986d5762b49";
//        URL url = new URL("https://financialmodelingprep.com/api/v3/balance-sheet-statement/" + companyName + "?limit=120&apikey=" + apiKey);
//        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(url.openStream()));

        File file = new File("test.json");
        Parcer parcer = new Parcer(file);
        List listParser = parcer.parceFile();
        LinkedHashMap listFinal = parcer.parce(listParser, 2);
        System.out.println(listFinal.get("calendarYear"));
        Object[] listFinalArr = listFinal.keySet().toArray();
        for (int i = 0; i < listFinalArr.length; i++) {
            System.out.println(i + " : " + listFinalArr[i]);
        }



        for (int i = 0; i < listFinal.size(); i++) {
            LinkedHashMap listFinal2 = parcer.parce(listParser, 0);
            var k = listFinal2.get(listFinalArr[i]);
            System.out.println(k.getClass());
        }


    XSSFWorkbook workbook = new XSSFWorkbook();

    XSSFSheet sheet = workbook.createSheet();

//        Row header = sheet.createRow(0);
//        header.createCell(0).setCellValue("Name");
//        header.createCell(1).setCellValue("Year");
//        for (int i = 0; i < listFinal.size(); i++) {
//            int s = 0;
//            LinkedHashMap listFinal2 = parcer.parce(listParser, 0);
//            var k = listFinalArr[i];
//            if(!(listFinal2.get(listFinalArr[i]) instanceof String)){
//                Row temp = sheet.createRow(s+1);
//                temp.createCell(0).setCellValue(k.toString());
//                s++;
//            }
//
//
//
//        }
//
    Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Year");


        for (int i = 0; i < listFinal.size(); i++) {
            Row temp = sheet.createRow(i+5);
            var k = listFinalArr[i];
            temp.createCell(0).setCellValue(k.toString()); }


            for (int j = 0; j < listParser.size(); j++) {
                for (int i = 0; i < listFinal.size(); i++) {
                    LinkedHashMap listFinal2 = parcer.parce(listParser, j);
                    Row temp = sheet.getRow(i+1);
                    var k = listFinal2.get(listFinalArr[i]);
                    if (k instanceof Long){
                        k = (long)k/1000000;
                        temp.createCell(j+1).setCellValue((long) k);
                    } else if (k instanceof Double) temp.createCell(j+1).setCellValue(k.toString());
                }
   }


        FileOutputStream out = new FileOutputStream(
                new File("Test.xlsx"));
        workbook.write(out);
        out.close();

}
}
