package com.newlight77.kata.survey.constants;

import lombok.Data;

@Data
public final class WokbookFormatterConstants {

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
    
}
