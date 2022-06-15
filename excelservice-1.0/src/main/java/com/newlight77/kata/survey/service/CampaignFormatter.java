package com.newlight77.kata.survey.service;

import org.springframework.boot.autoconfigure.condition.ConditionalOnCloudPlatform;
import org.springframework.stereotype.Component;

import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;

public interface CampaignFormatter<T> {

    public T format(Campaign campaign, Survey survey);
    
}
