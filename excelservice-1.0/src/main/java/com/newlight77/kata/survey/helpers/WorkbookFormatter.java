package com.newlight77.kata.survey.helpers;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.newlight77.kata.survey.model.Address;
import com.newlight77.kata.survey.model.AddressStatus;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;

public class WorkbookFormatter {

  static final int workbookHeaderColumnWidth = 10500;
  static final String workbookSheetName = "Survey";
  static final int workbookColumnWidth = 6000;
  static final int workbookMinColumnIndexToSetColumnWidth = 1;
  static final int workbookMaxColumnIndexToSetColumnWidth = 18;
  static final String workbookHeaderFont = "Arial";
  static final short workbookHeaderFontSize = 14;
  static final boolean workbookHeaderFontIsBold = true;
  static final boolean workbookHeaderIsTextWrapped = false;
  static final String workbookTitleFont = "Arial";
  static final short workbookTitleFontSize = 12;
  static final String workbookHeaderCol0Label = "Survey";
  static final boolean workbookDataIsTextWrapped = true;

  static final String workbookRow2Col0Label = "Client";

  static final String workbookRow6Col0Label = "Number of surveys";

  static final String workbookRow8Col0Label = "N° street";
  static final String workbookRow8Col1Label = "streee";
  static final String workbookRow8Col2Label = "Postal code";
  static final String workbookRow8Col3Label = "City";
  static final String workbookRow8Col4Label = "Status";

  static final int workbookSurveyDataStartRowIndex = 9;

  public static Workbook formatCampaignToWorkbook(Campaign campaign, Survey survey) {

    Workbook campaignResultsWorkbook = createWorkbook();

    CellStyle headerStyle = getHeaderStyle(campaignResultsWorkbook);
    CellStyle titleStyle = getTitleStyle(campaignResultsWorkbook);
    CellStyle defaultStyle = getDefaultStyle(campaignResultsWorkbook);

    Sheet surveySheet = createSurveySheet(campaignResultsWorkbook);

    Row surveySheetHeader = createRow(surveySheet, 0);
    Cell headerTitleLabel = createCell(surveySheetHeader, 0, workbookHeaderCol0Label, headerStyle);

    Row surveySheetClientLabel = createRow(surveySheet, 2);
    Cell clientLabel = createCell(surveySheetClientLabel, 0, workbookRow2Col0Label, titleStyle);
    Row surveySheetClientInfo = createRow(surveySheet, 3);
    Cell clientInfo = createCell(surveySheetClientInfo, 0, survey.getClient(), defaultStyle);
    
    Row surveySheetClientAddressInfo = createRow(surveySheet, 4);
    Cell clientAddress = createCell(surveySheetClientAddressInfo, 0, formatClientAddress(survey.getClientAddress()), defaultStyle);

    Row surveySheetSurveyDataInfo = createRow(surveySheet, 6);
    Cell numberOfSurveysLabel = createCell(surveySheetSurveyDataInfo, 0, workbookRow6Col0Label);
    Cell numberOfSurveys = createCell(surveySheetSurveyDataInfo, 1, campaign.getAddressStatuses().size());

    Row surveySheetSurveyDataHeader = createRow(surveySheet, 8);
    Cell streetNumberLabel = createCell(surveySheetSurveyDataHeader, 0, workbookRow8Col0Label);
    Cell streetLabel = createCell(surveySheetSurveyDataHeader, 1, workbookRow8Col1Label);
    Cell postalCodeLabel = createCell(surveySheetSurveyDataHeader, 2, workbookRow8Col2Label);
    Cell cityLabel = createCell(surveySheetSurveyDataHeader, 3, workbookRow8Col3Label);
    Cell statusLabel = createCell(surveySheetSurveyDataHeader, 4, workbookRow8Col4Label);

    int currentIndex = workbookSurveyDataStartRowIndex;

    for (AddressStatus addressStatus : campaign.getAddressStatuses()) {

      Row surveySheetSurveyData = createRow(surveySheet, currentIndex);
      Cell streetNumber = createCell(surveySheetSurveyData, 0, addressStatus.getAddress().getStreetNumber(), defaultStyle);
      Cell streetName = createCell(surveySheetSurveyData, 1, addressStatus.getAddress().getStreetName(), defaultStyle);
      Cell streetPostalCode = createCell(surveySheetSurveyData, 2, addressStatus.getAddress().getPostalCode(), defaultStyle);
      Cell streetCity = createCell(surveySheetSurveyData, 3, addressStatus.getAddress().getCity(), defaultStyle);
      Cell status = createCell(surveySheetSurveyData, 4, addressStatus.getStatus().toString(), defaultStyle);

      currentIndex++;

    }

    return campaignResultsWorkbook;

  }

  private static Workbook createWorkbook(){
    return new XSSFWorkbook();
  }

  private static Sheet createSurveySheet(Workbook workbook){

    Sheet sheet = workbook.createSheet(workbookSheetName);
    sheet.setColumnWidth(0, workbookHeaderColumnWidth);

    for (int i = workbookMinColumnIndexToSetColumnWidth; i <= workbookMaxColumnIndexToSetColumnWidth; i++) {
      sheet.setColumnWidth(i, workbookColumnWidth);
    }

    return sheet;
  }

  private static Row createRow(Sheet sheet, int rowIndex){

    return sheet.createRow(rowIndex);

  }

  private static Cell createCell(Row row, int colIndex, String value, CellStyle style){

    Cell cell = row.createCell(colIndex);
    cell.setCellValue(value);
    cell.setCellStyle(style);

    return cell;

  }

  private static Cell createCell(Row row, int colIndex, String value){

    Cell cell = row.createCell(colIndex);
    cell.setCellValue(value);

    return cell;

  }

  private static Cell createCell(Row row, int colIndex, int value){

    Cell cell = row.createCell(colIndex);
    cell.setCellValue(value);

    return cell;

  }

  private static CellStyle getHeaderStyle(Workbook workbook){

    CellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    headerStyle.setFont(getHeaderFont(workbook));
    headerStyle.setWrapText(workbookHeaderIsTextWrapped);

    return headerStyle;

  }

  private static CellStyle getTitleStyle(Workbook workbook){

    CellStyle titleStyle = workbook.createCellStyle();
    titleStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
    titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    
    titleStyle.setFont(getTitleFont(workbook));

    return titleStyle;

  }

  private static CellStyle getDefaultStyle(Workbook workbook){

    CellStyle defaulSstyle = workbook.createCellStyle();
    defaulSstyle.setWrapText(workbookDataIsTextWrapped);

    return defaulSstyle;

  }

  private static XSSFFont getTitleFont(Workbook workbook){

    XSSFFont titleFont = ((XSSFWorkbook) workbook).createFont();
    titleFont.setFontName(workbookTitleFont);
    titleFont.setFontHeightInPoints(workbookTitleFontSize);
    titleFont.setUnderline(FontUnderline.SINGLE);

    return titleFont;
    
  }

  private static XSSFFont getHeaderFont(Workbook workbook){

    XSSFFont headerFont = ((XSSFWorkbook) workbook).createFont();
    headerFont.setFontName(workbookHeaderFont);
    headerFont.setFontHeightInPoints(workbookHeaderFontSize);
    headerFont.setBold(workbookHeaderFontIsBold);

    return headerFont;
    
  }

  private static String formatClientAddress(Address clientAddress){
    String formattedClientAddress = String.format("%s %s %s %s"
      , clientAddress.getStreetNumber()
      , clientAddress.getStreetName()
      , clientAddress.getPostalCode()
      , clientAddress.getCity());

        return formattedClientAddress;
  }

}
