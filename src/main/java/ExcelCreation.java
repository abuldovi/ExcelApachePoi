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

public class ExcelCreation {
    public static void main(String[] args) throws Exception {

        String companyName = "GM";
        String apiKey = "7759164af885a77ae927b986d5762b49";
        URL url = new URL("https://financialmodelingprep.com/api/v3/balance-sheet-statement/" + companyName + "?limit=120&apikey=" + apiKey);
//        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(url.openStream()));

//        File file = new File("test.json");
        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode objYear = objectMapper.readTree(url);

        try{
            String cik = objYear.get(0).get("cik").asText();
            System.out.println(objYear.get(0).get("cik").asText());
        } catch (NullPointerException e){
            System.out.println("Company Name or apiKey is incorrect");
            throw e;
        }
        String cik = objYear.get(0).get("cik").asText();



        JsonNode PreferredStockSharesAuthorizedSecObj = SecParcer.getSecJsonNode(cik, "PreferredStockSharesAuthorized");
        JsonNode PreferredStockSharesOutstandingSecObj = SecParcer.getSecJsonNode(cik, "PreferredStockSharesOutstanding");
        JsonNode PreferredStockSharesIssuedSecObj = SecParcer.getSecJsonNode(cik, "PreferredStockSharesIssued");
        JsonNode CommonStockSharesAuthorizedSecObj = SecParcer.getSecJsonNode(cik, "CommonStockSharesAuthorized");
        JsonNode CommonStockSharesIssuedSecObj = SecParcer.getSecJsonNode(cik, "CommonStockSharesIssued");
        JsonNode CommonStockSharesOutstandingSecObj = SecParcer.getSecJsonNode(cik, "CommonStockSharesOutstanding");




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

            sheet.addMergedRegion(new CellRangeAddress(referenceHeight, referenceHeight, 0, objYear.size()));

            Row header = sheet.createRow(referenceHeight);
            Cell cell= header.createCell(0);
            cell.setCellValue("Balance sheet");
            cell.setCellStyle(headerStyle);

            referenceHeight++;

            Row years = sheet.createRow(referenceHeight);
            referenceHeight++;


            Row assets = sheet.createRow(referenceHeight);
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
            referenceHeight++;

            Row shortTermInv = sheet.createRow(referenceHeight);
            shortTermInv.createCell(0).setCellValue("Short-term investments");
            referenceHeight++;

            Row accountsReceivable = sheet.createRow(referenceHeight);
            accountsReceivable.createCell(0).setCellValue("Accounts receivable, net");
            referenceHeight++;

            Row inventory = sheet.createRow(referenceHeight);
            inventory.createCell(0).setCellValue("Inventory");
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
            preferredEquityCell.setCellValue("Preferred equity");
            preferredEquityCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row preferredStockSharesAuthorized = sheet.createRow(referenceHeight);
            preferredStockSharesAuthorized.createCell(0).setCellValue("Authorized shares");
            referenceHeight++;

            Row preferredStockSharesOutstanding = sheet.createRow(referenceHeight);
            preferredStockSharesOutstanding.createCell(0).setCellValue("Outstanding shares");
            referenceHeight++;

            Row preferredStockSharesIssued = sheet.createRow(referenceHeight);
            preferredStockSharesIssued.createCell(0).setCellValue("Issued shares");
            referenceHeight++;

            Row commonStockHeader = sheet.createRow(referenceHeight);
            Cell commonStockHeaderCell = commonStockHeader.createCell(0);
            commonStockHeaderCell.setCellValue("Common Stock");
            commonStockHeaderCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row commonStockSharesAuthorized = sheet.createRow(referenceHeight);
            commonStockSharesAuthorized.createCell(0).setCellValue("Authorized shares");
            referenceHeight++;

            Row commonStockSharesIssued = sheet.createRow(referenceHeight);
            commonStockSharesIssued.createCell(0).setCellValue("Issued shares");
            referenceHeight++;

            Row commonStockSharesOutstanding = sheet.createRow(referenceHeight);
            commonStockSharesOutstanding.createCell(0).setCellValue("Outstanding shares");
            referenceHeight++;

            Row commonStock = sheet.createRow(referenceHeight);
            commonStock.createCell(0).setCellValue("Common stock");
            referenceHeight++;

            Row othertotalStockholdersEquity = sheet.createRow(referenceHeight);
            othertotalStockholdersEquity.createCell(0).setCellValue("Other stockholders equity");
            referenceHeight++;

            Row retainedEarnings = sheet.createRow(referenceHeight);
            retainedEarnings.createCell(0).setCellValue("Retained earnings ");
            referenceHeight++;

            Row accumulatedOtherComprehensiveIncomeLoss = sheet.createRow(referenceHeight);
            accumulatedOtherComprehensiveIncomeLoss.createCell(0).setCellValue("Accumulated other comprehensive (loss) income (income/(loss))");
            referenceHeight++;

            Row minorityInterest = sheet.createRow(referenceHeight);
            minorityInterest.createCell(0).setCellValue("Noncontrolling interests (in subsidiaries)");
            referenceHeight++;

            Row totalStockholdersEquity = sheet.createRow(referenceHeight);
            Cell totalStockholdersEquityCell = totalStockholdersEquity.createCell(0);
            totalStockholdersEquityCell.setCellValue("Total equity");
            totalStockholdersEquityCell.setCellStyle(aggregateNumStyle);
            referenceHeight++;

            Row totalLiabilitiesAndTotalEquity = sheet.createRow(referenceHeight);
            Cell totalLiabilitiesAndTotalEquityCell = totalLiabilitiesAndTotalEquity.createCell(0);
            totalLiabilitiesAndTotalEquityCell.setCellValue("Total liabilities and equity");
            totalLiabilitiesAndTotalEquityCell.setCellStyle(aggregateNumStyleBig);
            referenceHeight++;







///////////////////////////////////
///////////////////////////////////
///////////////////////////////////
///////////////////////////////////
        int paddingLeft = 1;

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            years.createCell(i + paddingLeft).setCellValue(d);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("cashAndCashEquivalents").asLong();
            Cell tempCell = cash.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("shortTermInvestments").asLong();
            Cell tempCell = shortTermInv.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("netReceivables").asLong();
            Cell tempCell = accountsReceivable.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("inventory").asLong();
            Cell tempCell = inventory.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherCurrentAssets").asLong();
            Cell tempCell = otherCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalCurrentAssets").asLong();
            Cell tempCell = totalCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("propertyPlantEquipmentNet").asLong();
            Cell tempCell = propertyPlantEquipmentNet.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("longTermInvestments").asLong();
            Cell tempCell = longTermInvestments.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("taxAssets").asLong();
            Cell tempCell = taxAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("intangibleAssets").asLong();
            Cell tempCell = intangibleAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("goodwill").asLong();
            Cell tempCell = goodwill.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherNonCurrentAssets").asLong();
            Cell tempCell = otherNonCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalNonCurrentAssets").asLong();
            Cell tempCell = totalNonCurrentAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherAssets").asLong();
            Cell tempCell = otherAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalAssets").asLong();
            Cell tempCell = totalAssets.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("accountPayables").asLong();
            Cell tempCell = accountPayables.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("shortTermDebt").asLong();
            Cell tempCell = shortTermDebt.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("taxPayables").asLong();
            Cell tempCell = taxPayables.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredRevenue").asLong();
            Cell tempCell = deferredRevenue.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherCurrentLiabilities").asLong();
            Cell tempCell = otherCurrentLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalCurrentLiabilities").asLong();
            Cell tempCell = totalCurrentLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("longTermDebt").asLong();
            Cell tempCell = longTermDebt.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredTaxLiabilitiesNonCurrent").asLong();
            Cell tempCell = deferredTaxLiabilitiesNonCurrent.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("deferredRevenueNonCurrent").asLong();
            Cell tempCell = deferredRevenueNonCurrent.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherNonCurrentLiabilities").asLong();
            Cell tempCell = otherNonCurrentLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalNonCurrentLiabilities").asLong();
            Cell tempCell = totalNonCurrentLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("otherLiabilities").asLong();
            Cell tempCell = otherLiabilities.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("capitalLeaseObligations").asLong();
            Cell tempCell = capitalLeaseObligations.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, PreferredStockSharesAuthorizedSecObj);
            Cell tempCell = preferredStockSharesAuthorized.createCell(i + paddingLeft);
                    if(k<0){tempCell.setCellValue("None");}
                    else
                    {
                        tempCell.setCellValue(divider(k));
                        tempCell.setCellStyle(numbStyle);
                    }


        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, PreferredStockSharesOutstandingSecObj);
            Cell tempCell = preferredStockSharesOutstanding.createCell(i + paddingLeft);
                    if(k<0){tempCell.setCellValue("None");}
                    else
                    {
                        tempCell.setCellValue(divider(k));
                        tempCell.setCellStyle(numbStyle);
                    }

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, PreferredStockSharesIssuedSecObj);
            Cell tempCell = preferredStockSharesIssued.createCell(i + paddingLeft);
                    if(k<0){tempCell.setCellValue("None");}
                    else
                    {
                        tempCell.setCellValue(divider(k));
                        tempCell.setCellStyle(numbStyle);
                    }

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, CommonStockSharesAuthorizedSecObj);
            Cell tempCell = commonStockSharesAuthorized.createCell(i + paddingLeft);
                    if(k<0){tempCell.setCellValue("None");}
                    else
                    {
                        tempCell.setCellValue(divider(k));
                        tempCell.setCellStyle(numbStyle);
                    }

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, CommonStockSharesIssuedSecObj);
            Cell tempCell = commonStockSharesIssued.createCell(i + paddingLeft);
                    if(k<0){tempCell.setCellValue("None");}
                    else
                    {
                        tempCell.setCellValue(divider(k));
                        tempCell.setCellStyle(numbStyle);
                    }

        }

        for (int i = objYear.size()-1; i >= 0 ; i--) {
            int d = objYear.get(4-i).get("calendarYear").asInt();
            long k = SecParcer.secValue(d, CommonStockSharesOutstandingSecObj);
            Cell tempCell = commonStockSharesOutstanding.createCell(i + paddingLeft);
                    if(k<0){tempCell.setCellValue("None");}
                    else
                    {
                        tempCell.setCellValue(divider(k));
                        tempCell.setCellStyle(numbStyle);
                    }

        }


        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("commonStock").asLong();
            Cell tempCell = commonStock.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }
        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("othertotalStockholdersEquity").asLong();
            Cell tempCell = othertotalStockholdersEquity.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }
        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("retainedEarnings").asLong();
            Cell tempCell = retainedEarnings.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }
        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("accumulatedOtherComprehensiveIncomeLoss").asLong();
            Cell tempCell = accumulatedOtherComprehensiveIncomeLoss.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }
        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("minorityInterest").asLong();
            Cell tempCell = minorityInterest.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyle);

        }
        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalStockholdersEquity").asLong();
            Cell tempCell = totalStockholdersEquity.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }
        for (int i = objYear.size()-1; i >= 0 ; i--) {
            long d = objYear.get(4-i).get("totalLiabilitiesAndTotalEquity").asLong();
            Cell tempCell = totalLiabilitiesAndTotalEquity.createCell(i + paddingLeft);
            tempCell.setCellValue(divider(d));
            tempCell.setCellStyle(numbStyleBold);

        }




        for (int i = 0; i < objYear.size()+2; i++) {
            sheet.autoSizeColumn(i);
        }


            FileOutputStream out = new FileOutputStream(
                    new File(companyName + "_Report.xlsx"));
            workbook.write(out);
            out.close();
    }
    public static double divider(long longNumber){
        return (double)(longNumber/1_000_000);
    }

    }

