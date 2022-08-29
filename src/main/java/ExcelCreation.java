import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.util.*;

public class ExcelCreation {
    public static void main(String[] args) throws IOException {

//        String companyName = "AAPL";
//        String apiKey = "7759164af885a77ae927b986d5762b49";
//        URL url = new URL("https://financialmodelingprep.com/api/v3/balance-sheet-statement/" + companyName + "?limit=120&apikey=" + apiKey);
//        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(url.openStream()));

        File file = new File("test.json");
        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode objYear = objectMapper.readTree(file);

        System.out.println(objYear.get(0).get("date").asText());


        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet();

        DataFormat format = workbook.createDataFormat();
//
//        XSSFFont font = workbook.createFont();
//        font.setFontHeightInPoints((short)10);
//        font.setFontName("Arial");
//        font.setColor(IndexedColors.WHITE.getIndex());
//        font.setBold(true);
//        font.setItalic(false);

        XSSFCellStyle numbStyle = workbook.createCellStyle();
        numbStyle.setDataFormat(format.getFormat("# ##0"));


        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont headerStyleFont = workbook.createFont();
        headerStyleFont.setBold(true);
        headerStyleFont.setFontHeight(13);
        headerStyle.setFont(headerStyleFont);

        XSSFCellStyle aggregateNumStyle = workbook.createCellStyle();
        XSSFFont aggregateNumStyleFont = workbook.createFont();
        aggregateNumStyleFont.setBold(true);
        aggregateNumStyle.setFont(aggregateNumStyleFont);

        XSSFCellStyle aggregateNumStyleBig = workbook.createCellStyle();
        XSSFFont aggregateNumStyleFontBig = workbook.createFont();
        aggregateNumStyleFontBig.setFontHeight(12);
        aggregateNumStyleFontBig.setBold(true);
        aggregateNumStyleBig.setFont(aggregateNumStyleFontBig);

        XSSFCellStyle numbStyleBold = workbook.createCellStyle();
        numbStyleBold.setDataFormat(format.getFormat("# ##0"));
        numbStyleBold.setFont(aggregateNumStyleFont);

            int referenceHeight = 1;

            sheet.addMergedRegion(new CellRangeAddress(referenceHeight, referenceHeight, 0, objYear.size()+1));
//
            Row header = sheet.createRow(referenceHeight);
            Cell cell= header.createCell(0);
            cell.setCellValue("Balance sheet");
            cell.setCellStyle(headerStyle);

            referenceHeight++;

            Row years = sheet.createRow(referenceHeight);
            referenceHeight++;


            Row assets = sheet.createRow(referenceHeight);
            int assetsInt = referenceHeight;
            Cell assetsCell = assets.createCell(0);
                assetsCell.setCellValue("Assets");
                assetsCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;

            Row currentAssets = sheet.createRow(referenceHeight);
            Cell currentAssetsCell = currentAssets.createCell(0);
                    currentAssetsCell.setCellValue("Current Assets");
                    currentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row cash = sheet.createRow(referenceHeight);
            cash.createCell(0).setCellValue("Cash and cash equivalents");
            int cashInt = referenceHeight;
            referenceHeight++;

            Row shortTermInv = sheet.createRow(referenceHeight);
            shortTermInv.createCell(0).setCellValue("Short-term investments");
            int shortTermInvInt = referenceHeight;
            referenceHeight++;

            Row accountsReceivable = sheet.createRow(referenceHeight);
            accountsReceivable.createCell(0).setCellValue("Accounts receivable, net");
            int accountsReceivableInt = referenceHeight;
            referenceHeight++;

            Row inventory = sheet.createRow(referenceHeight);
            inventory.createCell(0).setCellValue("Inventory");
            int inventoryInt = referenceHeight;
            referenceHeight++;

            Row otherCurrentAssets = sheet.createRow(referenceHeight);
            Cell otherCurrentAssetsCell = otherCurrentAssets.createCell(0);
                otherCurrentAssetsCell.setCellValue("Other current assets");
                currentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row totalCurrentAssets = sheet.createRow(referenceHeight);
            Cell totalCurrentAssetsCell = totalCurrentAssets.createCell(0);
                    totalCurrentAssetsCell.setCellValue("Total current assets");
                    totalCurrentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row nonCurrentAssets = sheet.createRow(referenceHeight);
            Cell nonCurrentAssetsCell = nonCurrentAssets.createCell(0);
                nonCurrentAssetsCell.setCellValue("Non-current assets");
                nonCurrentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row propertyPlantEquipmentNet = sheet.createRow(referenceHeight);
            propertyPlantEquipmentNet.createCell(0).setCellValue("PP&E");
            referenceHeight++;

            Row longTermInvestments = sheet.createRow(referenceHeight);
            longTermInvestments.createCell(0).setCellValue("Long-term investments");
            referenceHeight++;

            Row taxAssets = sheet.createRow(referenceHeight);
            taxAssets.createCell(0).setCellValue("Deferred tax asset");
            referenceHeight++;

            Row intangibleAssets = sheet.createRow(referenceHeight);
            intangibleAssets.createCell(0).setCellValue("Intangible assets");
            referenceHeight++;

            Row goodwill = sheet.createRow(referenceHeight);
            goodwill.createCell(0).setCellValue("Goodwill");
            referenceHeight++;

            Row otherNonCurrentAssets = sheet.createRow(referenceHeight);
            otherNonCurrentAssets.createCell(0).setCellValue("Other non-current assets");
            referenceHeight++;

            Row totalNonCurrentAssets = sheet.createRow(referenceHeight);
            Cell totalNonCurrentAssetsCell = totalNonCurrentAssets.createCell(0);
            totalNonCurrentAssetsCell.setCellValue("Total non-current assets");
            totalNonCurrentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row otherAssets = sheet.createRow(referenceHeight);
            otherAssets.createCell(0).setCellValue("Other assets");
            referenceHeight++;

            Row totalAssets = sheet.createRow(referenceHeight);
            Cell totalAssetsCell = totalAssets.createCell(0);
            totalAssetsCell.setCellValue("Total assets");
            totalAssetsCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;

            Row liabilitiesAndEquity = sheet.createRow(referenceHeight);
            Cell liabilitiesAndEquityCell = liabilitiesAndEquity.createCell(0);
            liabilitiesAndEquityCell.setCellValue("Liabilities and Equity");
            liabilitiesAndEquityCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;

            Row currentLiabilities = sheet.createRow(referenceHeight);
            Cell currentLiabilitiesCell = currentLiabilities.createCell(0);
            currentLiabilitiesCell.setCellValue("Current liabilities");
            currentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row accountPayables = sheet.createRow(referenceHeight);
            accountPayables.createCell(0).setCellValue("Accounts payable");
            referenceHeight++;

            Row shortTermDebt = sheet.createRow(referenceHeight);
            shortTermDebt.createCell(0).setCellValue("Short-term debt");
            referenceHeight++;

            Row taxPayables = sheet.createRow(referenceHeight);
            taxPayables.createCell(0).setCellValue("Taxes payable");
            referenceHeight++;

            Row deferredRevenue = sheet.createRow(referenceHeight);
            deferredRevenue.createCell(0).setCellValue("Deferred revenue");
            referenceHeight++;

            Row otherCurrentLiabilities = sheet.createRow(referenceHeight);
            otherCurrentLiabilities.createCell(0).setCellValue("Other current liabilities");
            referenceHeight++;

            Row totalCurrentLiabilities = sheet.createRow(referenceHeight);
            Cell totalCurrentLiabilitiesCell = totalCurrentLiabilities.createCell(0);
            totalCurrentLiabilitiesCell.setCellValue("Total non-current assets");
            totalCurrentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row nonCurrentLiabilities = sheet.createRow(referenceHeight);
            Cell nonCurrentLiabilitiesCell = nonCurrentLiabilities.createCell(0);
            nonCurrentLiabilitiesCell.setCellValue("Non-current liabilities");
            nonCurrentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row longTermDebt = sheet.createRow(referenceHeight);
            longTermDebt.createCell(0).setCellValue("Long-term liablities");
            referenceHeight++;

            Row deferredTaxLiabilitiesNonCurrent = sheet.createRow(referenceHeight);
            deferredTaxLiabilitiesNonCurrent.createCell(0).setCellValue("Deferred tax liability");
            referenceHeight++;

            Row deferredRevenueNonCurrent = sheet.createRow(referenceHeight);
            deferredRevenueNonCurrent.createCell(0).setCellValue("Deferred revenue");
            referenceHeight++;

            Row otherNonCurrentLiabilities = sheet.createRow(referenceHeight);
            otherNonCurrentLiabilities.createCell(0).setCellValue("Other long-term liabilities");
            referenceHeight++;

            Row totalNonCurrentLiabilities = sheet.createRow(referenceHeight);
            Cell totalNonCurrentLiabilitiesCell = totalNonCurrentLiabilities.createCell(0);
            totalNonCurrentLiabilitiesCell.setCellValue("Total non-current liabilities");
            totalNonCurrentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row otherLiabilities = sheet.createRow(referenceHeight);
            otherLiabilities.createCell(0).setCellValue("Other liabilities");
            referenceHeight++;

            Row capitalLeaseObligations = sheet.createRow(referenceHeight);
            capitalLeaseObligations.createCell(0).setCellValue("Capital lease obligations");
            referenceHeight++;


            Row equity = sheet.createRow(referenceHeight);
            Cell equityCell = equity.createCell(0);
            equityCell.setCellValue("Equity");
            equityCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row preferredEquity = sheet.createRow(referenceHeight);
            Cell preferredEquityCell = preferredEquity.createCell(0);
            preferredEquityCell.setCellValue("Total non-current liabilities");
            preferredEquityCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;


            sheet.autoSizeColumn(0);


///////////////////////////////////
///////////////////////////////////
///////////////////////////////////
///////////////////////////////////


        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            years.createCell(i + 2).setCellValue(d);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("cashAndCashEquivalents").asLong();
            Cell tempCell = cash.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("shortTermInvestments").asLong();
            Cell tempCell = shortTermInv.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("netReceivables").asLong();
            Cell tempCell = accountsReceivable.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("inventory").asLong();
            Cell tempCell = inventory.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherCurrentAssets").asLong();
            Cell tempCell = otherCurrentAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalCurrentAssets").asLong();
            Cell tempCell = totalCurrentAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("propertyPlantEquipmentNet").asLong();
            Cell tempCell = propertyPlantEquipmentNet.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("longTermInvestments").asLong();
            Cell tempCell = longTermInvestments.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("taxAssets").asLong();
            Cell tempCell = taxAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("intangibleAssets").asLong();
            Cell tempCell = intangibleAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("goodwill").asLong();
            Cell tempCell = goodwill.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherNonCurrentAssets").asLong();
            Cell tempCell = otherNonCurrentAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalNonCurrentAssets").asLong();
            Cell tempCell = totalNonCurrentAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherAssets").asLong();
            Cell tempCell = otherAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalAssets").asLong();
            Cell tempCell = totalAssets.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("accountPayables").asLong();
            Cell tempCell = accountPayables.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("shortTermDebt").asLong();
            Cell tempCell = shortTermDebt.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("taxPayables").asLong();
            Cell tempCell = taxPayables.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredRevenue").asLong();
            Cell tempCell = deferredRevenue.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherCurrentLiabilities").asLong();
            Cell tempCell = otherCurrentLiabilities.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalCurrentLiabilities").asLong();
            Cell tempCell = totalCurrentLiabilities.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("longTermDebt").asLong();
            Cell tempCell = longTermDebt.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredTaxLiabilitiesNonCurrent").asLong();
            Cell tempCell = deferredTaxLiabilitiesNonCurrent.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredRevenueNonCurrent").asLong();
            Cell tempCell = deferredRevenueNonCurrent.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherNonCurrentLiabilities").asLong();
            Cell tempCell = otherNonCurrentLiabilities.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalNonCurrentLiabilities").asLong();
            Cell tempCell = totalNonCurrentLiabilities.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherLiabilities").asLong();
            Cell tempCell = otherLiabilities.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("capitalLeaseObligations").asLong();
            Cell tempCell = capitalLeaseObligations.createCell(i + 2);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }






//        for (int i = 0; i < listFinal.size(); i++) {
//            Row temp = sheet.createRow(i+5);
//            var k = listFinalArr[i];
//            temp.createCell(0).setCellValue(k.toString()); }
//
//
//            for (int j = 0; j < listParser.size(); j++) {
//                for (int i = 0; i < listFinal.size(); i++) {
//                    LinkedHashMap listFinal2 = parcer.parce(listParser, j);
//                    Row temp = sheet.getRow(i+1);
//                    var k = listFinal2.get(listFinalArr[i]);
//                    if (k instanceof Long){
//                        k = (long)k/1000000;
//                        temp.createCell(j+1).setCellValue((long) k);
//                    } else if (k instanceof Double) temp.createCell(j+1).setCellValue(k.toString());
//                }
//   }


            FileOutputStream out = new FileOutputStream(
                    new File("Test.xlsx"));
            workbook.write(out);
            out.close();
    }
    public static int divider(long longNumber){
        return (int)(longNumber/1_000_000);
    }

    }

