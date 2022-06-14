package com.newlight77.kata.survey.service;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import com.newlight77.kata.survey.client.CampaignClient;
import com.newlight77.kata.survey.helpers.WorkbookFormatter;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;

@Component
public class ExportCampaignService {

  private CampaignClient campaignWebService;
  private MailService mailService;
  private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
  
  public ExportCampaignService(final CampaignClient campaignWebService, MailService mailService) {
    this.campaignWebService = campaignWebService;
    this.mailService = mailService;
  }

  public void creerSurvey(Survey survey) {
    campaignWebService.createSurvey(survey);
  }

  public Survey getSurvey(String id) {
    return campaignWebService.getSurvey(id);
  }

  public void createCampaign(Campaign campaign) {
    campaignWebService.createCampaign(campaign);
  }

  public Campaign getCampaign(String id) {
    return campaignWebService.getCampaign(id);
  }

  public void sendResults(Campaign campaign, Survey survey) {
    Workbook workbook = WorkbookFormatter.formatCampaignToWorkbook(campaign, survey);

    writeFileAndSend(survey, workbook);

  }

  protected void writeFileAndSend(Survey survey, Workbook workbook) {
    try {
      File resultFile = new File(System.getProperty("java.io.tmpdir"), "survey-" + survey.getId() + "-" + dateTimeFormatter.format(LocalDate.now()) + ".xlsx");
      FileOutputStream outputStream = new FileOutputStream(resultFile);
      workbook.write(outputStream);

      mailService.send(resultFile);
      resultFile.deleteOnExit();
    } catch(Exception ex) {
        throw new RuntimeException("Errorr while trying to send email", ex);
    } finally {
      try {
        workbook.close();
      } catch(Exception e) {
        // CANT HAPPEN
      }
    }
  }

}
