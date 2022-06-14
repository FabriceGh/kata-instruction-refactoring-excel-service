package com.newlight77.kata.survey.helpers;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.assertj.core.api.Assertions;
import org.junit.Test;

import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import com.newlight77.kata.survey.util.JsonUtil;

public class WorkbookFormatterTest {

    @Test
    public void shouldFormatCampaignToWorkbook() {

        // Arrange
        String surveyJson = "{\n" +
                "    \"id\" : \"surveyId\",\n" +
                "    \"sommary\" : \"sommary\",\n" +
                "    \"client\" : \"client's name\",\n" +
                "    \"clientAddress\" : {\n" +
                "        \"id\" : \"addressId1\",\n" +
                "        \"streetNumber\" : \"10\",\n" +
                "        \"streetName\" : \"rue de Rivoli\",\n" +
                "        \"postalCode\" : \"75001\",\n" +
                "        \"city\" : \"Paris\"\n" +
                "    },\n" +
                "    \"questions\" : [{\n" +
                "        \"id\" : \"questionId1\",\n" +
                "        \"question\" : \"question1\"\n" +
                "    }, \n" +
                "    {\n" +
                "       \"id\" : \"questionId2\",\n" +
                "       \"question\" : \"question2\"\n" +
                "   }] \n" +
                "}";

        Survey survey = JsonUtil.instance().fromJson(surveyJson, Survey.class);

        String campaignJson = "{\n" +
                "    \"id\" : \"campaignId\",\n" +
                "    \"surveyId\" : \"surveyId\",\n" +
                "    \"addressStatuses\" : [ {\n" +
                "        \"id\" : \"addressStatusesId1\",\n" +
                "        \"address\" : {\n" +
                "          \"id\" : \"addressId1\",\n" +
                "          \"streetNumber\" : \"10\",\n" +
                "          \"streetName\" : \"rue de Rivoli\",\n" +
                "          \"postalCode\" : \"75001\",\n" +
                "          \"city\" : \"Paris\"\n" +
                "        },\n" +
                "        \"status\" : \"DONE\"\n" +
                "    }, {\n" +
                "        \"id\" : \"addressStatusesId2\",\n" +
                "        \"address\" : {\n" +
                "          \"id\" : \"addressId2\",\n" +
                "          \"streetNumber\" : \"40\",\n" +
                "          \"streetName\" : \"rue de Louvre\",\n" +
                "          \"postalCode\" : \"75001\",\n" +
                "          \"city\" : \"Paris\"\n" +
                "        },\n" +
                "        \"status\" : \"TODO\"\n" +
                "    }] \n" +
                "}";
        Campaign campaign = JsonUtil.instance().fromJson(campaignJson, Campaign.class);

        Workbook workbook = WorkbookFormatter.formatCampaignToWorkbook(campaign, survey);

        Assertions.assertThat(workbook).isNotNull();

        Sheet sheet = workbook.getSheet("Survey");
        Assertions.assertThat(sheet).isNotNull();

        Row headerRow = sheet.getRow(0);
        Assertions.assertThat(headerRow).isNotNull();
        Assertions.assertThat(headerRow.getCell(0)).isNotNull();
        Assertions.assertThat(headerRow.getCell(0).getStringCellValue()).isEqualTo("Survey");

        Row clientRow = sheet.getRow(2);
        Assertions.assertThat(clientRow).isNotNull();
        Assertions.assertThat(clientRow.getCell(0)).isNotNull();
        Assertions.assertThat(clientRow.getCell(0).getStringCellValue()).isEqualTo("Client");

        Row clientNameRow = sheet.getRow(3);
        Assertions.assertThat(clientNameRow).isNotNull();
        Assertions.assertThat(clientNameRow.getCell(0)).isNotNull();
        Assertions.assertThat(clientNameRow.getCell(0).getStringCellValue()).isEqualTo(survey.getClient());

        Row clientAddressRow = sheet.getRow(4);
        Assertions.assertThat(clientAddressRow).isNotNull();
        Assertions.assertThat(clientAddressRow.getCell(0)).isNotNull();

        Row SurveyNumberRow = sheet.getRow(6);
        Assertions.assertThat(SurveyNumberRow).isNotNull();
        Assertions.assertThat(SurveyNumberRow.getCell(0)).isNotNull();
        Assertions.assertThat(SurveyNumberRow.getCell(0).getStringCellValue()).isEqualTo("Number of surveys");
        Assertions.assertThat(SurveyNumberRow.getCell(1)).isNotNull();
        Assertions.assertThat(SurveyNumberRow.getCell(1).getNumericCellValue()).isEqualTo(campaign.getAddressStatuses().size());

        Row surveysHeaderRow = sheet.getRow(8);
        Assertions.assertThat(surveysHeaderRow).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(1)).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(1).getStringCellValue()).isEqualTo("streee");

        Assertions.assertThat(surveysHeaderRow).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(2)).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(2).getStringCellValue()).isEqualTo("Postal code");

        Assertions.assertThat(surveysHeaderRow).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(3)).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(3).getStringCellValue()).isEqualTo("City");

        Assertions.assertThat(surveysHeaderRow).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(4)).isNotNull();
        Assertions.assertThat(surveysHeaderRow.getCell(4).getStringCellValue()).isEqualTo("Status");

        Row campaignResultsRow = sheet.getRow(9);
        Assertions.assertThat(campaignResultsRow.getCell(0)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(1)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(2)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(3)).isNotNull();
        Assertions.assertThat(campaignResultsRow.getCell(4)).isNotNull();

        int campaignSize = campaign.getAddressStatuses().size();
        int lastResultRow = 9 + campaignSize - 1;

        Row campaignLastResultsRow = sheet.getRow(lastResultRow);
        Assertions.assertThat(campaignLastResultsRow.getCell(0)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(1)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(2)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(3)).isNotNull();
        Assertions.assertThat(campaignLastResultsRow.getCell(4)).isNotNull();
    }
}