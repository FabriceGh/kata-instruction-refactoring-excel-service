package com.newlight77.kata.survey.helpers;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.assertj.core.api.Assertions;
import org.junit.Before;
import org.junit.Test;

import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import com.newlight77.kata.survey.util.JsonUtil;

public class WorkbookFormatterTest {

    private Survey surveyMock;
    private Campaign campaignMock;
    private Workbook workbookMock;

    @Before
    public void init() {
        surveyMock = JsonUtil.instance().fromJsonFile("../jsonMocks/survey.json", Survey.class);
        campaignMock = JsonUtil.instance().fromJsonFile("../jsonMocks/campaign.json", Campaign.class);
    
        workbookMock = WorkbookFormatter.formatCampaignToWorkbook(campaignMock, surveyMock);
    }

    @Test
    public void workbook_should_be_not_null() {

        Assertions.assertThat(workbookMock).isNotNull();

    }

    @Test
    public void workbook_should_have_a_Survey_sheet() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Assertions.assertThat(sheet).isNotNull();

    }

    @Test
    public void workbook_Survey_sheet_row_0_should_be_not_null() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row headerRow = sheet.getRow(0);
        Assertions.assertThat(headerRow).isNotNull();

    }

    @Test
    public void workbook_Survey_sheet_row_2_should_be_not_null() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row clientRow = sheet.getRow(2);
        Assertions.assertThat(clientRow).isNotNull();

    }

    @Test
    public void workbook_Survey_sheet_row_3_should_be_not_null() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row clientNameRow = sheet.getRow(3);
        Assertions.assertThat(clientNameRow).isNotNull();

    }

    @Test
    public void workbook_Survey_sheet_row_4_should_be_not_null() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row clientAddressRow = sheet.getRow(4);
        Assertions.assertThat(clientAddressRow).isNotNull();

    }

    @Test
    public void workbook_Survey_sheet_row_6_should_be_not_null() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row numberOfSurveysRow = sheet.getRow(6);
        Assertions.assertThat(numberOfSurveysRow).isNotNull();

    }

    @Test
    public void workbook_Survey_sheet_row_8_should_be_not_null() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row surveysHeaderRow = sheet.getRow(8);
        Assertions.assertThat(surveysHeaderRow).isNotNull();

    }

    @Test
    public void workbook_Row0_Cell0_should_contain_Survey_text() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row headerRow = sheet.getRow(0);
        Assertions.assertThat(headerRow.getCell(0).getStringCellValue()).isEqualTo("Survey");

    }

    @Test
    public void workbook_Row2_Cell0_should_contain_Client_text() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row clientRow = sheet.getRow(2);
        Assertions.assertThat(clientRow.getCell(0).getStringCellValue()).isEqualTo("Client");

    }

    @Test
    public void workbook_Row3_Cell0_should_contain_Client_mocked_name_value() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row clientNameRow = sheet.getRow(3);
        Assertions.assertThat(clientNameRow.getCell(0)).isNotNull();
        Assertions.assertThat(clientNameRow.getCell(0).getStringCellValue()).isEqualTo(surveyMock.getClient());

    }

    @Test
    public void workbook_Row4_Cell0_should_contain_Client_mocked_address_value() {

        String clientAddress = surveyMock.getClientAddress().getStreetNumber() 
                                + " " + surveyMock.getClientAddress().getStreetName() 
                                + surveyMock.getClientAddress().getPostalCode() 
                                + " " + surveyMock.getClientAddress().getCity();

        Sheet sheet = workbookMock.getSheet("Survey");
        Row clientAddressRow = sheet.getRow(4);
        Assertions.assertThat(clientAddressRow.getCell(0)).isNotNull();
        Assertions.assertThat(clientAddressRow.getCell(0).getStringCellValue()).isEqualTo(clientAddress);

    }

    @Test
    public void workbook_Row6_Cell0_should_contain_NumberOfSurvey_text() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row numberOfSurveysRow = sheet.getRow(6);
        Assertions.assertThat(numberOfSurveysRow.getCell(0).getStringCellValue()).isEqualTo("Number of surveys");

    }

    @Test
    public void workbook_Row6_Cell1_should_contain_mocked_NumberOfSurvey_value() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row numberOfSurveysRow = sheet.getRow(6);
        Assertions.assertThat(numberOfSurveysRow.getCell(1)).isNotNull();
        Assertions.assertThat(numberOfSurveysRow.getCell(1).getNumericCellValue()).isEqualTo(campaignMock.getAddressStatuses().size());

    }

    @Test
    public void workbook_Row8_Cell1_should_contain_Street_text() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row surveysHeaderRow = sheet.getRow(8);
        Assertions.assertThat(surveysHeaderRow.getCell(1).getStringCellValue()).isEqualTo("streee");

    }

    @Test
    public void workbook_Row8_Cell2_should_contain_PostalCode_text() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row surveysHeaderRow = sheet.getRow(8);
        Assertions.assertThat(surveysHeaderRow.getCell(2).getStringCellValue()).isEqualTo("Postal code");

    }

    @Test
    public void workbook_Row8_Cell3_should_contain_City_text() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row surveysHeaderRow = sheet.getRow(8);
        Assertions.assertThat(surveysHeaderRow.getCell(3).getStringCellValue()).isEqualTo("City");

    }

    @Test
    public void workbook_Row8_Cell4_should_contain_Status_text() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row surveysHeaderRow = sheet.getRow(8);
        Assertions.assertThat(surveysHeaderRow.getCell(4).getStringCellValue()).isEqualTo("Status");

    }

    @Test
    public void workbook_Row9_Cell0_To_4_should_not_be_Null() {

        Sheet sheet = workbookMock.getSheet("Survey");
        Row campaignResultsRow = sheet.getRow(9);
        Assertions.assertThat(campaignResultsRow.getCell(0)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(1)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(2)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(3)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(4)).isNotNull();

    }

    @Test
    public void workbook_CampaignLastRow_Cell0_To_4_should_not_be_Null() {

        int campaignSize = campaignMock.getAddressStatuses().size();
        int lastResultRow = 9 + campaignSize - 1;

        Sheet sheet = workbookMock.getSheet("Survey");
        Row campaignLastResultsRow = sheet.getRow(lastResultRow);
        Assertions.assertThat(campaignLastResultsRow.getCell(0)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(1)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(2)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(3)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(4)).isNotNull();
    }

}
