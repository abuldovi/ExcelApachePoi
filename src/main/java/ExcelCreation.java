import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.net.URL;
import java.util.Date;

public class ExcelCreation {
    public static void main(String[] args) throws Exception {

        Date current = new Date();
        String companyName = JOptionPane.showInputDialog ("Type company ticker");
        int showNull = Integer.parseInt(JOptionPane.showInputDialog ("Type 0 to show lines with no information, type 1 otherwise"));


        String apiKey = "7759164af885a77ae927b986d5762b49";
        URL url = new URL("https://financialmodelingprep.com/api/v3/balance-sheet-statement/" + companyName + "?limit=120&apikey=" + apiKey);
        URL incomeURL = new URL("https://financialmodelingprep.com/api/v3/income-statement/" + companyName + "?limit=120&apikey=" + apiKey);
        URL cfURL = new URL("https://financialmodelingprep.com/api/v3/cash-flow-statement/" + companyName + "?limit=120&apikey=" + apiKey);

        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode objYear = objectMapper.readTree(url);
        JsonNode incomeObjYear = objectMapper.readTree(incomeURL);
        JsonNode cfObjYear = objectMapper.readTree(cfURL);



        try{
            String cik = objYear.get(0).get("cik").asText();
            System.out.println(objYear.get(0).get("cik").asText());
        } catch (NullPointerException e){
            JOptionPane.showMessageDialog(new JFrame("Error"), "Company Name or apiKey is incorrect, please restart the program");
            System.exit(0);
        }
        String cik = objYear.get(0).get("cik").asText();


        JsonNode NetIncomeLossAttributableToNoncontrollingInterestSecObj = SecParcer.getSecJsonNode(cik, "NetIncomeLossAttributableToNoncontrollingInterest");
        JsonNode PreferredStockSharesAuthorizedSecObj = SecParcer.getSecJsonNode(cik, "PreferredStockSharesAuthorized");
        JsonNode PreferredStockSharesOutstandingSecObj = SecParcer.getSecJsonNode(cik, "PreferredStockSharesOutstanding");
        JsonNode PreferredStockSharesIssuedSecObj = SecParcer.getSecJsonNode(cik, "PreferredStockSharesIssued");
        JsonNode CommonStockSharesAuthorizedSecObj = SecParcer.getSecJsonNode(cik, "CommonStockSharesAuthorized");
        JsonNode CommonStockSharesIssuedSecObj = SecParcer.getSecJsonNode(cik, "CommonStockSharesIssued");
        JsonNode CommonStockSharesOutstandingSecObj = SecParcer.getSecJsonNode(cik, "CommonStockSharesOutstanding");
        JsonNode NetIncomeLossAvailableToCommonStockholdersBasicSecObj = SecParcer.getSecJsonNode(cik, "NetIncomeLossAvailableToCommonStockholdersBasic");
        JsonNode NetIncomeLossAvailableToCommonStockholdersDilutedSecObj = SecParcer.getSecJsonNode(cik, "NetIncomeLossAvailableToCommonStockholdersDiluted");
        JsonNode DividendsCommonStockSecObj = SecParcer.getSecJsonNode(cik, "DividendsCommonStock");




        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet();

        DataFormat format = workbook.createDataFormat();

     // Styles -------------------------------------------------------------------------------------------------

        XSSFCellStyle numbStyle = workbook.createCellStyle();
        String numbFormat = "# ##0.00";
        numbStyle.setDataFormat(format.getFormat(numbFormat));
    //    numbStyle.setDataFormat(format.getFormat("# ##0"));


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
        numbStyleBold.setDataFormat(format.getFormat(numbFormat));
        numbStyleBold.setFont(aggregateNumStyleFont);

        // Left column -------------------------------------------------------------------------------------------------

            int referenceHeight = 1; // Padding from top
            int referenceHeightInit = referenceHeight;
            int paddingLeft = 1;
            int zeroCounter = 0;


            sheet.addMergedRegion(new CellRangeAddress(referenceHeight, referenceHeight, 0, objYear.size()));

//----------------------------------------
            Row header = sheet.createRow(referenceHeight);
            Cell cell= header.createCell(0);
            cell.setCellValue("Balance sheet");
            cell.setCellStyle(headerStyle);

            referenceHeight++;

//----------------------------------------
            Row years = sheet.createRow(referenceHeight);
            referenceHeight++;


//----------------------------------------
            Row assets = sheet.createRow(referenceHeight);
            Cell assetsCell = assets.createCell(0);
                assetsCell.setCellValue("Assets");
                assetsCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;

//----------------------------------------
            Row currentAssets = sheet.createRow(referenceHeight);
            Cell currentAssetsCell = currentAssets.createCell(0);
                    currentAssetsCell.setCellValue("Current Assets");
                    currentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

//----------------------------------------
            Row cash = sheet.createRow(referenceHeight);
            cash.createCell(0).setCellValue("Cash and cash equivalents");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("cashAndCashEquivalents").asLong();
            Cell tempCell = cash.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if(d==0){
                zeroCounter++;
            }

        }
            if (zeroCounter==5){
                referenceHeight--;
            }
            zeroCounter = 0;

//----------------------------------------
            Row shortTermInv = sheet.createRow(referenceHeight);
            shortTermInv.createCell(0).setCellValue("Short-term investments");
            referenceHeight++;


        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("shortTermInvestments").asLong();
            Cell tempCell = shortTermInv.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row accountsReceivable = sheet.createRow(referenceHeight);
            accountsReceivable.createCell(0).setCellValue("Accounts receivable, net");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("netReceivables").asLong();
            Cell tempCell = accountsReceivable.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row inventory = sheet.createRow(referenceHeight);
            inventory.createCell(0).setCellValue("Inventory");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("inventory").asLong();
            Cell tempCell = inventory.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row otherCurrentAssets = sheet.createRow(referenceHeight);
            Cell otherCurrentAssetsCell = otherCurrentAssets.createCell(0);
                otherCurrentAssetsCell.setCellValue("Other current assets");
                currentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherCurrentAssets").asLong();
            Cell tempCell = otherCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row totalCurrentAssets = sheet.createRow(referenceHeight);
            Cell totalCurrentAssetsCell = totalCurrentAssets.createCell(0);
                    totalCurrentAssetsCell.setCellValue("Total current assets");
                    totalCurrentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalCurrentAssets").asLong();
            Cell tempCell = totalCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row nonCurrentAssets = sheet.createRow(referenceHeight);
            Cell nonCurrentAssetsCell = nonCurrentAssets.createCell(0);
                nonCurrentAssetsCell.setCellValue("Non-current assets");
                nonCurrentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

//----------------------------------------
            Row propertyPlantEquipmentNet = sheet.createRow(referenceHeight);
            propertyPlantEquipmentNet.createCell(0).setCellValue("PP&E");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("propertyPlantEquipmentNet").asLong();
            Cell tempCell = propertyPlantEquipmentNet.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row longTermInvestments = sheet.createRow(referenceHeight);
            longTermInvestments.createCell(0).setCellValue("Long-term investments");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("longTermInvestments").asLong();
            Cell tempCell = longTermInvestments.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row taxAssets = sheet.createRow(referenceHeight);
            taxAssets.createCell(0).setCellValue("Deferred tax asset");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("taxAssets").asLong();
            Cell tempCell = taxAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row intangibleAssets = sheet.createRow(referenceHeight);
            intangibleAssets.createCell(0).setCellValue("Intangible assets");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("intangibleAssets").asLong();
            Cell tempCell = intangibleAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row goodwill = sheet.createRow(referenceHeight);
            goodwill.createCell(0).setCellValue("Goodwill");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("goodwill").asLong();
            Cell tempCell = goodwill.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row otherNonCurrentAssets = sheet.createRow(referenceHeight);
            otherNonCurrentAssets.createCell(0).setCellValue("Other non-current assets");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherNonCurrentAssets").asLong();
            Cell tempCell = otherNonCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row totalNonCurrentAssets = sheet.createRow(referenceHeight);
            Cell totalNonCurrentAssetsCell = totalNonCurrentAssets.createCell(0);
            totalNonCurrentAssetsCell.setCellValue("Total non-current assets");
            totalNonCurrentAssetsCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalNonCurrentAssets").asLong();
            Cell tempCell = totalNonCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;


        Row otherAssets = sheet.createRow(referenceHeight);
            otherAssets.createCell(0).setCellValue("Other assets");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherAssets").asLong();
            Cell tempCell = otherAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row totalAssets = sheet.createRow(referenceHeight);
            Cell totalAssetsCell = totalAssets.createCell(0);
            totalAssetsCell.setCellValue("Total assets");
            totalAssetsCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalAssets").asLong();
            Cell tempCell = totalAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row liabilitiesAndEquity = sheet.createRow(referenceHeight);
            Cell liabilitiesAndEquityCell = liabilitiesAndEquity.createCell(0);
            liabilitiesAndEquityCell.setCellValue("Liabilities and Equity");
            liabilitiesAndEquityCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;

//----------------------------------------
            Row currentLiabilities = sheet.createRow(referenceHeight);
            Cell currentLiabilitiesCell = currentLiabilities.createCell(0);
            currentLiabilitiesCell.setCellValue("Current liabilities");
            currentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

//----------------------------------------
            Row accountPayables = sheet.createRow(referenceHeight);
            accountPayables.createCell(0).setCellValue("Accounts payable");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("accountPayables").asLong();
            Cell tempCell = accountPayables.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row shortTermDebt = sheet.createRow(referenceHeight);
            shortTermDebt.createCell(0).setCellValue("Short-term debt");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("shortTermDebt").asLong();
            Cell tempCell = shortTermDebt.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;


//----------------------------------------
            Row taxPayables = sheet.createRow(referenceHeight);
            taxPayables.createCell(0).setCellValue("Taxes payable");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("taxPayables").asLong();
            Cell tempCell = taxPayables.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row deferredRevenue = sheet.createRow(referenceHeight);
            deferredRevenue.createCell(0).setCellValue("Deferred revenue");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredRevenue").asLong();
            Cell tempCell = deferredRevenue.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row otherCurrentLiabilities = sheet.createRow(referenceHeight);
            otherCurrentLiabilities.createCell(0).setCellValue("Other current liabilities");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherCurrentLiabilities").asLong();
            Cell tempCell = otherCurrentLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row totalCurrentLiabilities = sheet.createRow(referenceHeight);
            Cell totalCurrentLiabilitiesCell = totalCurrentLiabilities.createCell(0);
            totalCurrentLiabilitiesCell.setCellValue("Total non-current assets");
            totalCurrentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;





//----------------------------------------
            Row nonCurrentLiabilities = sheet.createRow(referenceHeight);
            Cell nonCurrentLiabilitiesCell = nonCurrentLiabilities.createCell(0);
            nonCurrentLiabilitiesCell.setCellValue("Non-current liabilities");
            nonCurrentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

//----------------------------------------
            Row longTermDebt = sheet.createRow(referenceHeight);
            longTermDebt.createCell(0).setCellValue("Long-term liablities");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("longTermDebt").asLong();
            Cell tempCell = longTermDebt.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row deferredTaxLiabilitiesNonCurrent = sheet.createRow(referenceHeight);
            deferredTaxLiabilitiesNonCurrent.createCell(0).setCellValue("Deferred tax liability");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredTaxLiabilitiesNonCurrent").asLong();
            Cell tempCell = deferredTaxLiabilitiesNonCurrent.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row deferredRevenueNonCurrent = sheet.createRow(referenceHeight);
            deferredRevenueNonCurrent.createCell(0).setCellValue("Deferred revenue");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredRevenueNonCurrent").asLong();
            Cell tempCell = deferredRevenueNonCurrent.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row otherNonCurrentLiabilities = sheet.createRow(referenceHeight);
            otherNonCurrentLiabilities.createCell(0).setCellValue("Other long-term liabilities");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherNonCurrentLiabilities").asLong();
            Cell tempCell = otherNonCurrentLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row totalNonCurrentLiabilities = sheet.createRow(referenceHeight);
            Cell totalNonCurrentLiabilitiesCell = totalNonCurrentLiabilities.createCell(0);
            totalNonCurrentLiabilitiesCell.setCellValue("Total non-current liabilities");
            totalNonCurrentLiabilitiesCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalNonCurrentLiabilities").asLong();
            Cell tempCell = totalNonCurrentLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row otherLiabilities = sheet.createRow(referenceHeight);
            otherLiabilities.createCell(0).setCellValue("Other liabilities");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherLiabilities").asLong();
            Cell tempCell = otherLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row capitalLeaseObligations = sheet.createRow(referenceHeight);
            capitalLeaseObligations.createCell(0).setCellValue("Capital lease obligations");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("capitalLeaseObligations").asLong();
            Cell tempCell = capitalLeaseObligations.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;


//----------------------------------------
            Row equity = sheet.createRow(referenceHeight);
            Cell equityCell = equity.createCell(0);
            equityCell.setCellValue("Equity");
            equityCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

//----------------------------------------
            Row preferredEquity = sheet.createRow(referenceHeight);
            Cell preferredEquityCell = preferredEquity.createCell(0);
            preferredEquityCell.setCellValue("Preferred equity");
            preferredEquityCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

//----------------------------------------
            Row preferredStockSharesAuthorized = sheet.createRow(referenceHeight);
            preferredStockSharesAuthorized.createCell(0).setCellValue("Authorized shares");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4 - i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, PreferredStockSharesAuthorizedSecObj);
            Cell tempCell = preferredStockSharesAuthorized.createCell(i + paddingLeft);
            if (k < 0) {
                tempCell.setCellValue("N/A");
            } else {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if (showNull == 1) {
                if (k < 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;


//----------------------------------------
            Row preferredStockSharesOutstanding = sheet.createRow(referenceHeight);
            preferredStockSharesOutstanding.createCell(0).setCellValue("Outstanding shares");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, PreferredStockSharesOutstandingSecObj);
            Cell tempCell = preferredStockSharesOutstanding.createCell(i + paddingLeft);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if (showNull == 1) {
                if (k < 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row preferredStockSharesIssued = sheet.createRow(referenceHeight);
            preferredStockSharesIssued.createCell(0).setCellValue("Issued shares");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, PreferredStockSharesIssuedSecObj);
            Cell tempCell = preferredStockSharesIssued.createCell(i + paddingLeft);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if (showNull == 1) {
                if (k < 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row commonStockHeader = sheet.createRow(referenceHeight);
            Cell commonStockHeaderCell = commonStockHeader.createCell(0);
            commonStockHeaderCell.setCellValue("Common Stock");
            commonStockHeaderCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;
            

//----------------------------------------
            Row commonStockSharesAuthorized = sheet.createRow(referenceHeight);
            commonStockSharesAuthorized.createCell(0).setCellValue("Authorized shares");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, CommonStockSharesAuthorizedSecObj);
            Cell tempCell = commonStockSharesAuthorized.createCell(i + paddingLeft);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if (showNull == 1) {
                if (k < 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row commonStockSharesIssued = sheet.createRow(referenceHeight);
            commonStockSharesIssued.createCell(0).setCellValue("Issued shares");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, CommonStockSharesIssuedSecObj);
            Cell tempCell = commonStockSharesIssued.createCell(i + paddingLeft);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if (showNull == 1) {
                if (k < 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row commonStockSharesOutstanding = sheet.createRow(referenceHeight);
            commonStockSharesOutstanding.createCell(0).setCellValue("Outstanding shares");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, CommonStockSharesOutstandingSecObj);
            Cell tempCell = commonStockSharesOutstanding.createCell(i + paddingLeft);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if (showNull == 1) {
                if (k < 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;


        Row commonStock = sheet.createRow(referenceHeight);
            commonStock.createCell(0).setCellValue("Common stock");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("commonStock").asLong();
            Cell tempCell = commonStock.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row othertotalStockholdersEquity = sheet.createRow(referenceHeight);
            othertotalStockholdersEquity.createCell(0).setCellValue("Other stockholders equity");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("othertotalStockholdersEquity").asLong();
            Cell tempCell = othertotalStockholdersEquity.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row retainedEarnings = sheet.createRow(referenceHeight);
            retainedEarnings.createCell(0).setCellValue("Retained earnings ");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("retainedEarnings").asLong();
            Cell tempCell = retainedEarnings.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row accumulatedOtherComprehensiveIncomeLoss = sheet.createRow(referenceHeight);
            accumulatedOtherComprehensiveIncomeLoss.createCell(0).setCellValue("Accumulated other comprehensive (loss) income (income/(loss))");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("accumulatedOtherComprehensiveIncomeLoss").asLong();
            Cell tempCell = accumulatedOtherComprehensiveIncomeLoss.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;
//----------------------------------------
//----------------------------------------
            Row minorityInterest = sheet.createRow(referenceHeight);
            minorityInterest.createCell(0).setCellValue("Noncontrolling interests (in subsidiaries)");
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("minorityInterest").asLong();
            Cell tempCell = minorityInterest.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

//----------------------------------------
            Row totalStockholdersEquity = sheet.createRow(referenceHeight);
            Cell totalStockholdersEquityCell = totalStockholdersEquity.createCell(0);
            totalStockholdersEquityCell.setCellValue("Total equity");
            totalStockholdersEquityCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalStockholdersEquity").asLong();
            Cell tempCell = totalStockholdersEquity.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

            
//----------------------------------------
            Row totalLiabilitiesAndTotalEquity = sheet.createRow(referenceHeight);
            Cell totalLiabilitiesAndTotalEquityCell = totalLiabilitiesAndTotalEquity.createCell(0);
            totalLiabilitiesAndTotalEquityCell.setCellValue("Total liabilities and equity");
            totalLiabilitiesAndTotalEquityCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalLiabilitiesAndTotalEquity").asLong();
            Cell tempCell = totalLiabilitiesAndTotalEquity.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        
///////////////////////////////////
///////////////////////////////////
///////////////////////////////////
///////////////////////////////////


        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            years.createCell(i + paddingLeft).setCellValue(d);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;


        CellRangeAddress borderBS = new CellRangeAddress(referenceHeightInit, referenceHeight-1, paddingLeft-1, objYear.size());

        RegionUtil.setBorderTop(BorderStyle.MEDIUM, borderBS, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, borderBS, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, borderBS, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, borderBS, sheet);

        CellRangeAddress borderBSHeader = new CellRangeAddress(referenceHeightInit, referenceHeightInit, paddingLeft-1, objYear.size());

        RegionUtil.setBorderTop(BorderStyle.MEDIUM, borderBSHeader, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, borderBSHeader, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, borderBSHeader, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, borderBSHeader, sheet);

        ////////////////////////
        ////////////////////////
        ///Income Statement////
        ///////////////////////
        ///////////////////////


        int paddingIncome = paddingLeft + objYear.size()+2;
        referenceHeight = 1;

        sheet.addMergedRegion(new CellRangeAddress(referenceHeight, referenceHeight, paddingIncome, paddingIncome+incomeObjYear.size()));

        //----------------------------------------

        Row headerIncome = sheet.getRow(referenceHeight);
        Cell cellIncome= headerIncome.createCell(paddingIncome);
        cellIncome.setCellValue("IncomeStatement");
        cellIncome.setCellStyle(headerStyle);
        referenceHeight++;

        //----------------------------------------

        Row incomeYears = sheet.getRow(referenceHeight);
        referenceHeight++;
        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            int d = incomeObjYear.get(4-i).get("calendarYear").asInt();
            incomeYears.createCell(i + paddingIncome+1).setCellValue(d);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row incomeRevenue = sheet.getRow(referenceHeight);
        Cell incomeRevenueCell = incomeRevenue.createCell(paddingIncome);
        incomeRevenueCell.setCellValue("Revenue");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("revenue").asLong();
            Cell tempCell = incomeRevenue.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



//----------------------------------------

        Row costOfRevenue = sheet.getRow(referenceHeight);
        Cell costOfRevenueCell = costOfRevenue.createCell(paddingIncome);
        costOfRevenueCell .setCellValue("Cost of sales");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("costOfRevenue").asLong();
            Cell tempCell = costOfRevenue.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row grossProfit = sheet.getRow(referenceHeight);
        Cell grossProfitCell = grossProfit.createCell(paddingIncome);
        grossProfitCell.setCellValue("Gross profit");
        grossProfitCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("grossProfit").asLong();
            Cell tempCell = grossProfit.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row operatingExpensesHeader = sheet.getRow(referenceHeight);
        Cell operatingExpensesHeaderCell = operatingExpensesHeader.createCell(paddingIncome);
        operatingExpensesHeaderCell.setCellValue("Operating expenses");
        operatingExpensesHeaderCell.setCellStyle(aggregateNumStyleBig);
        referenceHeight++;

        //----------------------------------------

        Row generalAndAdministrativeExpenses = sheet.getRow(referenceHeight);
        Cell generalAndAdministrativeExpensesCell = generalAndAdministrativeExpenses.createCell(paddingIncome);
        generalAndAdministrativeExpensesCell.setCellValue("General and administrative");


        referenceHeight++;


        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("generalAndAdministrativeExpenses").asLong();
            Cell tempCell = generalAndAdministrativeExpenses.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row researchAndDevelopmentExpenses = sheet.getRow(referenceHeight);
        Cell researchAndDevelopmentExpensesCell = researchAndDevelopmentExpenses.createCell(paddingIncome);
        researchAndDevelopmentExpensesCell.setCellValue("Research and development");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("researchAndDevelopmentExpenses").asLong();
            Cell tempCell = researchAndDevelopmentExpenses.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;




        //----------------------------------------

        Row sellingAndMarketingExpenses = sheet.getRow(referenceHeight);
        Cell sellingAndMarketingExpensesCell = sellingAndMarketingExpenses.createCell(paddingIncome);
        sellingAndMarketingExpensesCell.setCellValue("Selling and marketing");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("sellingAndMarketingExpenses").asLong();
            Cell tempCell = sellingAndMarketingExpenses.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;




        //----------------------------------------

        Row sellingGeneralAndAdministrativeExpenses = sheet.getRow(referenceHeight);
        Cell sellingGeneralAndAdministrativeExpensesCell = sellingGeneralAndAdministrativeExpenses.createCell(paddingIncome);
        sellingGeneralAndAdministrativeExpensesCell.setCellValue("Selling, General and Administrative expense");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("sellingGeneralAndAdministrativeExpenses").asLong();
            Cell tempCell = sellingGeneralAndAdministrativeExpenses.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row otherExpenses = sheet.getRow(referenceHeight);
        Cell otherExpensesCell = otherExpenses.createCell(paddingIncome);
        otherExpensesCell.setCellValue("Other expenses");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("otherExpenses").asLong();
            Cell tempCell = otherExpenses.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row operatingExpenses = sheet.getRow(referenceHeight);
        Cell operatingExpensesCell = operatingExpenses.createCell(paddingIncome);
        operatingExpensesCell.setCellStyle(aggregateNumStyle);
        operatingExpensesCell.setCellValue("Total operating expenses");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("operatingExpenses").asLong();
            Cell tempCell = operatingExpenses.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row ebitda = sheet.getRow(referenceHeight);
        Cell ebitdaCell = ebitda.createCell(paddingIncome);
        ebitdaCell.setCellValue("EBITDA");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("ebitda").asLong();
            Cell tempCell = ebitda.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row interestIncome = sheet.getRow(referenceHeight);
        Cell interestIncomeCell = interestIncome.createCell(paddingIncome);
        interestIncomeCell.setCellValue("Interest income ");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("interestIncome").asLong();
            Cell tempCell = interestIncome.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row interestExpense = sheet.getRow(referenceHeight);
        Cell interestExpenseCell = interestExpense.createCell(paddingIncome);
        interestExpenseCell.setCellValue("Interest expense");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("interestExpense").asLong();
            Cell tempCell = interestExpense.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row depreciationAndAmortization = sheet.getRow(referenceHeight);
        Cell depreciationAndAmortizationCell = depreciationAndAmortization.createCell(paddingIncome);
        depreciationAndAmortizationCell.setCellValue("Depreciation and amortization");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("depreciationAndAmortization").asLong();
            Cell tempCell = depreciationAndAmortization.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row operatingIncome = sheet.getRow(referenceHeight);
        Cell operatingIncomeCell = operatingIncome.createCell(paddingIncome);
        operatingIncomeCell.setCellValue("Operating income");
        operatingIncomeCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("operatingIncome").asLong();
            Cell tempCell = operatingIncome.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row totalOtherIncomeExpensesNet = sheet.getRow(referenceHeight);
        Cell totalOtherIncomeExpensesNetCell = totalOtherIncomeExpensesNet.createCell(paddingIncome);
        totalOtherIncomeExpensesNetCell.setCellValue("Other income and expense");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("totalOtherIncomeExpensesNet").asLong();
            Cell tempCell = totalOtherIncomeExpensesNet.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row incomeBeforeTax = sheet.getRow(referenceHeight);
        Cell incomeBeforeTaxCell = incomeBeforeTax.createCell(paddingIncome);
        incomeBeforeTaxCell.setCellValue("Income before taxes");
        incomeBeforeTaxCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("incomeBeforeTax").asLong();
            Cell tempCell = incomeBeforeTax.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row incomeTaxExpense = sheet.getRow(referenceHeight);
        Cell incomeTaxExpenseCell = incomeTaxExpense.createCell(paddingIncome);
        incomeTaxExpenseCell.setCellValue("Provision for taxes");

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("incomeTaxExpense").asLong();
            Cell tempCell = incomeTaxExpense.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row netIncome = sheet.getRow(referenceHeight);
        Cell netIncomeCell = netIncome.createCell(paddingIncome);
        netIncomeCell.setCellValue("Net income (loss)");
        netIncomeCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            long d = incomeObjYear.get(4-i).get("netIncome").asLong();
            Cell tempCell = netIncome.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;



        //----------------------------------------

        Row netIncomeLossAttributableToNoncontrollingInterest = sheet.getRow(referenceHeight);
        netIncomeLossAttributableToNoncontrollingInterest.createCell(paddingIncome).setCellValue("Net income attributable to non-controlling interests");
        referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValueIncome(d, NetIncomeLossAttributableToNoncontrollingInterestSecObj);
            Cell tempCell = netIncomeLossAttributableToNoncontrollingInterest.createCell(i + paddingLeft + paddingIncome);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if(k<0){
                zeroCounter++;
            }

        }
        if (zeroCounter==5){
            referenceHeight--;
        }
        zeroCounter = 0;

        //----------------------------------------

        Row netIncomeLossAvailableToCommonStockholdersBasic = sheet.getRow(referenceHeight);
        netIncomeLossAvailableToCommonStockholdersBasic.createCell(paddingIncome).setCellValue("Net income attributable to common shareholders - basic");
        referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValueIncome(d, NetIncomeLossAvailableToCommonStockholdersBasicSecObj);
            Cell tempCell = netIncomeLossAvailableToCommonStockholdersBasic.createCell(i + paddingLeft+paddingIncome);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if(k<0){
                zeroCounter++;
            }

        }
        if (zeroCounter==5){
            referenceHeight--;
        }
        zeroCounter = 0;

        //----------------------------------------

        Row netIncomeLossAvailableToCommonStockholdersDiluted = sheet.getRow(referenceHeight);
        netIncomeLossAvailableToCommonStockholdersDiluted.createCell(paddingIncome).setCellValue("Net income attributable to common shareholders - diluted");
        referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValueIncome(d, NetIncomeLossAvailableToCommonStockholdersDilutedSecObj);
            Cell tempCell = netIncomeLossAvailableToCommonStockholdersDiluted.createCell(i + paddingLeft+paddingIncome);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(divider(k));
                tempCell.setCellStyle(numbStyle);
            }

            if(k<0){
                zeroCounter++;
            }

        }
        if (zeroCounter==5){
            referenceHeight--;
        }
        zeroCounter = 0;

        //----------------------------------------

        Row eps = sheet.getRow(referenceHeight);
        Cell epsCell = eps.createCell(paddingIncome);
        epsCell.setCellValue("Earnings per share - basic");

        referenceHeight++;


        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            double d = incomeObjYear.get(4-i).get("eps").asDouble();
            Cell tempCell = eps.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(d);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row weightedAverageShsOut = sheet.getRow(referenceHeight);
        Cell weightedAverageShsOutCell = weightedAverageShsOut.createCell(paddingIncome);
        weightedAverageShsOutCell.setCellValue("Shares used in computing earning per share - basic");

        referenceHeight++;


        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            double d = incomeObjYear.get(4-i).get("weightedAverageShsOut").asDouble();
            Cell tempCell = weightedAverageShsOut.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(d);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row epsdiluted = sheet.getRow(referenceHeight);
        Cell epsdilutedCell = epsdiluted.createCell(paddingIncome);
        epsdilutedCell.setCellValue("Earnings per share - diluted");

        referenceHeight++;


        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            double d = incomeObjYear.get(4-i).get("epsdiluted").asDouble();
            Cell tempCell = epsdiluted.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(d);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row weightedAverageShsOutDil = sheet.getRow(referenceHeight);
        Cell weightedAverageShsOutDilCell = weightedAverageShsOutDil.createCell(paddingIncome);
        weightedAverageShsOutDilCell.setCellValue("Shares used in computing earning per share - diluted");

        referenceHeight++;


        for (int i = incomeObjYear.size()-1; i >= 0 ; i--) {
            double d = incomeObjYear.get(4-i).get("weightedAverageShsOutDil").asDouble();
            Cell tempCell = weightedAverageShsOutDil.createCell(i + paddingLeft + paddingIncome);
            tempCell.setCellValue(d);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row dividendsCommonStock = sheet.getRow(referenceHeight);
        dividendsCommonStock.createCell(paddingIncome).setCellValue("Dividends per share");
        referenceHeight++;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            double k = SecParcer.secValueIncomeDouble(d, DividendsCommonStockSecObj);
            Cell tempCell = dividendsCommonStock.createCell(i + paddingLeft+paddingIncome);
            if(k<0){tempCell.setCellValue("N/A");}
            else
            {
                tempCell.setCellValue(k);
                tempCell.setCellStyle(numbStyle);
            }

            if(k<0){
                zeroCounter++;
            }

        }
        if (zeroCounter==5){
            for (int i = objYear.size(); i >= 0 ; i--) {
                Cell tempCell = dividendsCommonStock.createCell(i + paddingLeft+paddingIncome-1);
                tempCell.setCellValue("");}

            referenceHeight--;
        }
        zeroCounter = 0;


        CellRangeAddress borderIncome = new CellRangeAddress(referenceHeightInit, referenceHeight-1, paddingIncome, paddingIncome+incomeObjYear.size());

        RegionUtil.setBorderTop(BorderStyle.MEDIUM, borderIncome, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, borderIncome, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, borderIncome, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, borderIncome, sheet);

        CellRangeAddress borderIncomeHeader = new CellRangeAddress(referenceHeightInit, referenceHeightInit, paddingIncome, paddingIncome+incomeObjYear.size());

        RegionUtil.setBorderTop(BorderStyle.MEDIUM, borderIncomeHeader, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, borderIncomeHeader, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, borderIncomeHeader, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, borderIncomeHeader, sheet);
        
        
        ////////////////////////
        ////////////////////////
        ///Cash flow statement////
        ///////////////////////
        ///////////////////////


        int paddingCF = paddingLeft + objYear.size()+ 2 + paddingLeft + incomeObjYear.size() + 2;
        referenceHeight = 1;

        sheet.addMergedRegion(new CellRangeAddress(referenceHeight, referenceHeight, paddingCF, paddingCF+cfObjYear.size()));

        //----------------------------------------

        Row cfHeader = sheet.getRow(referenceHeight);
        Cell cfHeaderCell= headerIncome.createCell(paddingCF);
        cfHeaderCell.setCellValue("Cash flow statement");
        cfHeaderCell.setCellStyle(headerStyle);
        referenceHeight++;

        //----------------------------------------

        Row cfYears = sheet.getRow(referenceHeight);
        referenceHeight++;
        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            int d = cfObjYear.get(4-i).get("calendarYear").asInt();
            cfYears.createCell(i + paddingCF+1).setCellValue(d);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row operatingActivities = sheet.getRow(referenceHeight);
        Cell operatingActivitiesCell = operatingActivities.createCell(paddingCF);
        operatingActivitiesCell.setCellValue("Operating activities");
        operatingActivitiesCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        //----------------------------------------

        Row netIncomeCF = sheet.getRow(referenceHeight);
        Cell netIncomeCFCell = netIncomeCF.createCell(paddingCF);
        netIncomeCFCell.setCellValue("Net income (loss)");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("netIncome").asLong();
            Cell tempCell = netIncomeCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row adjustmentToReconcile = sheet.getRow(referenceHeight);
        Cell adjustmentToReconcileCell= adjustmentToReconcile.createCell(paddingCF);
        adjustmentToReconcileCell.setCellValue("Adjustments to reconcile net income (loss)");
        adjustmentToReconcileCell.setCellStyle(aggregateNumStyle);
        referenceHeight++;

        //----------------------------------------

        Row depreciationAndAmortizationCF = sheet.getRow(referenceHeight);
        Cell depreciationAndAmortizationCFCell = depreciationAndAmortizationCF.createCell(paddingCF);
        depreciationAndAmortizationCFCell.setCellValue("Depreciation and amortization");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("depreciationAndAmortization").asLong();
            Cell tempCell = depreciationAndAmortizationCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row stockBasedCompensationCF = sheet.getRow(referenceHeight);
        Cell stockBasedCompensationCFCell = stockBasedCompensationCF.createCell(paddingCF);
        stockBasedCompensationCFCell.setCellValue("Share-based compensation");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("stockBasedCompensation").asLong();
            Cell tempCell = stockBasedCompensationCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row otherNonCashItemsCF = sheet.getRow(referenceHeight);
        Cell otherNonCashItemsCFCell = otherNonCashItemsCF.createCell(paddingCF);
        otherNonCashItemsCFCell.setCellValue("Other non-cash items");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("otherNonCashItems").asLong();
            Cell tempCell = otherNonCashItemsCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row changeInWorkingCapitalCF = sheet.getRow(referenceHeight);
        Cell changeInWorkingCapitalCFCell = changeInWorkingCapitalCF.createCell(paddingCF);
        changeInWorkingCapitalCFCell.setCellValue("Changes in working capital");
        changeInWorkingCapitalCFCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("changeInWorkingCapital").asLong();
            Cell tempCell = changeInWorkingCapitalCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row accountsReceivablesCF = sheet.getRow(referenceHeight);
        Cell accountsReceivablesCFCell = accountsReceivablesCF.createCell(paddingCF);
        accountsReceivablesCFCell.setCellValue("Accounts receivable");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("accountsReceivables").asLong();
            Cell tempCell = accountsReceivablesCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row accountsPayablesCF = sheet.getRow(referenceHeight);
        Cell accountsPayablesCFCell = accountsPayablesCF.createCell(paddingCF);
        accountsPayablesCFCell.setCellValue("Accounts payables");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("accountsPayables").asLong();
            Cell tempCell = accountsPayablesCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row inventoryCF = sheet.getRow(referenceHeight);
        Cell inventoryCFCell = inventoryCF.createCell(paddingCF);
        inventoryCFCell.setCellValue("Inventory");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("inventory").asLong();
            Cell tempCell = inventoryCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row deferredIncomeTax = sheet.getRow(referenceHeight);
        Cell deferredIncomeTaxCell = deferredIncomeTax.createCell(paddingCF);
        deferredIncomeTaxCell.setCellValue("Deferred income taxes");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("deferredIncomeTax").asLong();
            Cell tempCell = deferredIncomeTax.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row otherWorkingCapitalCF = sheet.getRow(referenceHeight);
        Cell otherWorkingCapitalCFCell = otherWorkingCapitalCF.createCell(paddingCF);
        otherWorkingCapitalCFCell.setCellValue("Other working capital");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long changeInWorkingCapitalLong = cfObjYear.get(4-i).get("changeInWorkingCapital").asLong();
            long accountsReceivablesLong = cfObjYear.get(4-i).get("accountsReceivables").asLong();
            long accountsPayablesLong = cfObjYear.get(4-i).get("accountsPayables").asLong();
            long inventoryLong = cfObjYear.get(4-i).get("inventory").asLong();
            long deferredIncomeTaxLong = cfObjYear.get(4-i).get("deferredIncomeTax").asLong();
            long d = changeInWorkingCapitalLong - accountsPayablesLong - accountsReceivablesLong - inventoryLong - deferredIncomeTaxLong;
            Cell tempCell = otherWorkingCapitalCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;


        //----------------------------------------

        Row netCashProvidedByOperatingActivitiesCF = sheet.getRow(referenceHeight);
        Cell netCashProvidedByOperatingActivitiesCFCell = netCashProvidedByOperatingActivitiesCF.createCell(paddingCF);
        netCashProvidedByOperatingActivitiesCFCell.setCellValue("Net cash provided by (used in) operating activities ");
        netCashProvidedByOperatingActivitiesCFCell.setCellStyle(aggregateNumStyle);
        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("netCashProvidedByOperatingActivities").asLong();
            Cell tempCell = netCashProvidedByOperatingActivitiesCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row investingActivitiesCF = sheet.getRow(referenceHeight);
        Cell investingActivitiesCFCell = investingActivitiesCF.createCell(paddingCF);
        investingActivitiesCFCell.setCellValue("Investing activities");
        investingActivitiesCFCell.setCellStyle(aggregateNumStyle);
        referenceHeight++;

        //----------------------------------------

        Row investmentsInPropertyPlantAndEquipmentCF = sheet.getRow(referenceHeight);
        Cell investmentsInPropertyPlantAndEquipmentCFCell = investmentsInPropertyPlantAndEquipmentCF.createCell(paddingCF);
        investmentsInPropertyPlantAndEquipmentCFCell.setCellValue("Acquisition of fixed assets");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("investmentsInPropertyPlantAndEquipment").asLong();
            Cell tempCell = investmentsInPropertyPlantAndEquipmentCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row purchasesOfInvestmentsCF = sheet.getRow(referenceHeight);
        Cell purchasesOfInvestmentsCFCell = purchasesOfInvestmentsCF.createCell(paddingCF);
        purchasesOfInvestmentsCFCell.setCellValue("Acquisition of debt and equity investments");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("purchasesOfInvestments").asLong();
            Cell tempCell = purchasesOfInvestmentsCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row salesMaturitiesOfInvestmentsCF = sheet.getRow(referenceHeight);
        Cell salesMaturitiesOfInvestmentsCFCell = salesMaturitiesOfInvestmentsCF.createCell(paddingCF);
        salesMaturitiesOfInvestmentsCFCell.setCellValue("Sale proceeds from debt and equity investments");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("salesMaturitiesOfInvestments").asLong();
            Cell tempCell = salesMaturitiesOfInvestmentsCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row otherInvestingActivitesCF = sheet.getRow(referenceHeight);
        Cell otherInvestingActivitesCFCell = otherInvestingActivitesCF.createCell(paddingCF);
        otherInvestingActivitesCFCell.setCellValue("Other investing activities");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("otherInvestingActivites").asLong();
            Cell tempCell = otherInvestingActivitesCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row acquisitionsNetCF = sheet.getRow(referenceHeight);
        Cell acquisitionsNetCFCell = acquisitionsNetCF.createCell(paddingCF);
        acquisitionsNetCFCell.setCellValue("Acquisitions");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("acquisitionsNet").asLong();
            Cell tempCell = acquisitionsNetCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row netCashUsedForInvestingActivitesCF = sheet.getRow(referenceHeight);
        Cell netCashUsedForInvestingActivitesCFCell = netCashUsedForInvestingActivitesCF.createCell(paddingCF);
        netCashUsedForInvestingActivitesCFCell.setCellValue("Net cash used in investing activities");
        netCashUsedForInvestingActivitesCFCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("netCashUsedForInvestingActivites").asLong();
            Cell tempCell = netCashUsedForInvestingActivitesCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row financingActivitiesCF = sheet.getRow(referenceHeight);
        Cell financingActivitiesCFCell = financingActivitiesCF.createCell(paddingCF);
        financingActivitiesCFCell.setCellValue("Financing activities");
        financingActivitiesCFCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        //----------------------------------------

        Row commonStockIssuedCF = sheet.getRow(referenceHeight);
        Cell commonStockIssuedCFCell = commonStockIssuedCF.createCell(paddingCF);
        commonStockIssuedCFCell.setCellValue("Proceeds from issuing securities");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("commonStockIssued").asLong();
            Cell tempCell = commonStockIssuedCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row commonStockRepurchasedCF = sheet.getRow(referenceHeight);
        Cell commonStockRepurchasedCFCell = commonStockRepurchasedCF.createCell(paddingCF);
        commonStockRepurchasedCFCell.setCellValue("Payments to reacquire stock");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("commonStockRepurchased").asLong();
            Cell tempCell = commonStockRepurchasedCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row dividendsPaidCF = sheet.getRow(referenceHeight);
        Cell dividendsPaidCFCell = dividendsPaidCF.createCell(paddingCF);
        dividendsPaidCFCell.setCellValue("Dividends paid to shareholders");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("dividendsPaid").asLong();
            Cell tempCell = dividendsPaidCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row debtRepaymentCF = sheet.getRow(referenceHeight);
        Cell debtRepaymentCFCell = debtRepaymentCF.createCell(paddingCF);
        debtRepaymentCFCell.setCellValue("Repayment of borrowings");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("debtRepayment").asLong();
            Cell tempCell = debtRepaymentCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row otherFinancingActivitesCF = sheet.getRow(referenceHeight);
        Cell otherFinancingActivitesCFCell = otherFinancingActivitesCF.createCell(paddingCF);
        otherFinancingActivitesCFCell.setCellValue("Other financing activities");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("otherFinancingActivites").asLong();
            Cell tempCell = otherFinancingActivitesCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row netCashUsedProvidedByFinancingActivitiesCF = sheet.getRow(referenceHeight);
        Cell netCashUsedProvidedByFinancingActivitiesCFCell = netCashUsedProvidedByFinancingActivitiesCF.createCell(paddingCF);
        netCashUsedProvidedByFinancingActivitiesCFCell.setCellValue("Net cash (used in) provided by financing activities ");
        netCashUsedProvidedByFinancingActivitiesCFCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("netCashUsedProvidedByFinancingActivities").asLong();
            Cell tempCell = netCashUsedProvidedByFinancingActivitiesCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row effectOfForexChangesOnCashCF = sheet.getRow(referenceHeight);
        Cell effectOfForexChangesOnCashCFCell = effectOfForexChangesOnCashCF.createCell(paddingCF);
        effectOfForexChangesOnCashCFCell.setCellValue("Foreign currency effect on cash, cash equivalents, and restricted cash");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("effectOfForexChangesOnCash").asLong();
            Cell tempCell = effectOfForexChangesOnCashCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row netChangeInCashCF = sheet.getRow(referenceHeight);
        Cell netChangeInCashCFCell = netChangeInCashCF.createCell(paddingCF);
        netChangeInCashCFCell.setCellValue("Deferred income taxes");
        netChangeInCashCFCell.setCellStyle(aggregateNumStyle);

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("netChangeInCash").asLong();
            Cell tempCell = netChangeInCashCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row cashAtBeginningOfPeriodCF = sheet.getRow(referenceHeight);
        Cell cashAtBeginningOfPeriodCFCell = cashAtBeginningOfPeriodCF.createCell(paddingCF);
        cashAtBeginningOfPeriodCFCell.setCellValue("Starting cash balance");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("cashAtBeginningOfPeriod").asLong();
            Cell tempCell = cashAtBeginningOfPeriodCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------

        Row cashAtEndOfPeriodCF = sheet.getRow(referenceHeight);
        Cell cashAtEndOfPeriodCFCell = cashAtEndOfPeriodCF.createCell(paddingCF);
        cashAtEndOfPeriodCFCell.setCellValue("Cash at the end of period");

        referenceHeight++;

        for (int i = cfObjYear.size()-1; i >= 0 ; i--) {
            long d = cfObjYear.get(4-i).get("cashAtEndOfPeriod").asLong();
            Cell tempCell = cashAtEndOfPeriodCF.createCell(i + paddingCF + 1);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);
            if (showNull == 1) {
                if (d == 0) {
                    zeroCounter++;
                }
                if (zeroCounter == 5) {
                referenceHeight--;
            }
            }
            }
        zeroCounter = 0;

        //----------------------------------------


        CellRangeAddress borderCF = new CellRangeAddress(referenceHeightInit, referenceHeight-1, paddingCF, paddingCF+cfObjYear.size());

        RegionUtil.setBorderTop(BorderStyle.MEDIUM, borderCF, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, borderCF, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, borderCF, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, borderCF, sheet);

        CellRangeAddress borderCFHeader = new CellRangeAddress(referenceHeightInit, referenceHeightInit, paddingCF, paddingCF+cfObjYear.size());

        RegionUtil.setBorderTop(BorderStyle.MEDIUM, borderCFHeader, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, borderCFHeader, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, borderCFHeader, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, borderCFHeader, sheet);


        for (int i = 0; i < paddingCF + cfObjYear.size() + 4; i++) {
            sheet.autoSizeColumn(i);
        }


            FileOutputStream out = new FileOutputStream(
                    new File(companyName + "_Report.xlsx"));
            workbook.write(out);
            out.close();

        Date current2 = new Date();
        System.out.println(current2.getTime()-current.getTime());


    }


    public static double divider(long longNumber){
        return (double)(longNumber/1_000_000);
    }


}


