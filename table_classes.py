#import statements
import os
import datetime


class Table(object):
    date = datetime.datetime.now()
    testing_JIRA = ""

    def __init__(self, date, testing_JIRA):
        self.date = date
        self.testing = testing_JIRA
    
class bundle_rows(Table):
    jira_code = ""
    summary = ""
    team = ""
    bundle_status = ""
    jira_status = ""
    assignee = ""
    reporter = ""
    priority = ""
    sub_priority = ""
    category = ""
    priortization_date = datetime.datetime.now()
    is_HRS = None
    hrqa_epqas_date = datetime.datetime.now()
    phire_cr = ""
    cr_type = ""
    hrs_epm_date = datetime.datetime.now()
    off_reason = ""
    project_association = ""

    def __init__(self, jira_code, summary, team, bundle_status, jira_status, assignee, reporter,
     priority, sub_priority, category, priortization_date, is_HRS, hrqa_epqas_date, phire_cr,
     cr_type, hrs_epm_date, off_reason, project_association):
        self.jira_code = jira_code
        self.summary = summary
        self.team = team
        self.bundle_status = bundle_status
        self.jira_status = jira_status
        self.assignee = assignee
        self.reporter = reporter
        self.priority = priority
        self.sub_priority = sub_priority
        self.category = category
        self.priortization_date = priortization_date
        self.is_HRS = is_HRS
        self.hrqa_epqas_date = hrqa_epqas_date
        self.phire_cr = phire_cr
        self.cr_type = cr_type
        self.hrs_epm_date = hrs_epm_date
        self.off_reason = off_reason
        self.project_association = project_association
    
    def toList(self):
        return [self.jira_code, self.summary, self.team, self.bundle_status, self.jira_status,
         self.assignee, self.reporter, self.priority, self.sub_priority, self.category,
         self.priortization_date, self.is_HRS, self.hrqa_epqas_date, self.phire_cr, self.cr_type,
         self.hrs_epm_date, self.off_reason, self.project_association]

class audit_rows():
    CR = ""
    jiraTracking = ""
    targetDB = ""
    migratedDate = ""
    migratedBy = ""
    crType = ""
    srcQuery = ""

    def __init__(self, CR, jiraTracking, targetDB, migratedDate, migratedBy, crType, srcQuery):
        self.CR = CR
        self.jiraTracking = jiraTracking
        self.targetDB = targetDB
        self.migratedDate = migratedDate
        self.migratedBy = migratedBy
        self.crType = crType
        self.srcQuery = srcQuery
    
    def toList(self):
        return [self.CR, self.jiraTracking, self.targetDB, self.migratedDate,
         self.migratedBy, self.crType, self.srcQuery]
