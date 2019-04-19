"""This script creates the monthly member detail excel reports for the QBH program
Clinical Quality Metrics"""

import sys
import xlsxwriter
import pandas
from qbfunctions import qbfunctions

def rate(den_col, num_col, den, startrow=9):
    """This function takes column number of a denominator, column number of a numerator
    and returns a string of an excel equation that will recalculate with autofilters"""
    den_range = "{0}{1}:{0}{2}".format(qbfunctions.colnum_string(den_col), int(startrow+1), int(startrow+den))
    num_range = "{0}{1}:{0}{2}".format(qbfunctions.colnum_string(num_col), int(startrow+1), int(startrow+den))
    return_str="""=concatenate("Denominator: ", SUBTOTAL(2, {0}), " Numerator: ", SUBTOTAL(2, {1}), " Rate: ", ROUND((SUBTOTAL(2, {1})/SUBTOTAL(2, {0}))*100, 2), "%")""".format(den_range, num_range)
    print(return_str)
    return return_str

# Current reporting period
try:
    YEAR = int(sys.argv[1])
    MONTH = int(sys.argv[2])
except BaseException:
    print("ERROR: Incorrect Number of Arguments Supplied." +
          " cqm_detail.py needs valid year and month:" +
          "./cqm_detail.py 2019 2")
    quit()

DATES = qbfunctions.dates(MONTH, YEAR)

PROVIDER = "/n04/data/p4vrept/Data/Provider/qb_hospitals_2019.sas7bdat"

QBH_DIR = "/n04/data/p4vrept/Programs/Quality_Blue_Hospital_2019/"
DATA_DIR = QBH_DIR + "DATA/mastheads/"
HOSP03_DIR = DATA_DIR + "hosp03/"
HOSP04_DIR = DATA_DIR + "hosp04/"
HOSP19_DIR = DATA_DIR + "hosp1920/"
HOSP21_DIR = DATA_DIR + "hosp21/"
READM_DIR = DATA_DIR + "readm/"
HOSP22_DIR = DATA_DIR + "hosp222324/"
HOSP23_DIR = DATA_DIR + "hosp222324/"
HOSP24_DIR = DATA_DIR + "hosp222324/"

HOSP03_DETAIL = HOSP03_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_HOSP03 = HOSP03_DETAIL + "member_hosp03_scored.sas7bdat"
HOSP03_SCORES = DATA_DIR + "hosp03_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

HOSP04_DETAIL = HOSP04_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_HOSP04 = HOSP04_DETAIL + "member_hosp04_scored.sas7bdat"
HOSP04_SCORES = DATA_DIR + "hosp04_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

HOSP19_DETAIL = HOSP19_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_HOSP19 = HOSP19_DETAIL + "member_hosp19.sas7bdat"
HOSP19_SCORES = DATA_DIR + "hosp19_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

MEMBER_HOSP20 = HOSP19_DETAIL + "member_hosp20.sas7bdat"
HOSP20_SCORES = DATA_DIR + "hosp20_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

HOSP21_DETAIL = HOSP21_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_HOSP21 = HOSP21_DETAIL + "member_hosp21_scored.sas7bdat"
HOSP21_SCORES = DATA_DIR + "hosp21_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

RRAMA_DETAIL = READM_DIR + str(YEAR) + "_" + \
    str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_RRAMA = RRAMA_DETAIL + "member_rrama.sas7bdat"
RRAMA_SCORES = DATA_DIR + "rrama_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

RRACOMM_DETAIL = READM_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_RRACOMM = RRACOMM_DETAIL + "member_rracomm.sas7bdat"
RRACOMM_SCORES = DATA_DIR + "rracomm_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

HOSP22_DETAIL = HOSP22_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_HOSP22 = HOSP22_DETAIL + "member_hosp22.sas7bdat"
HOSP22_SCORES = DATA_DIR + "hosp22_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

HOSP23_DETAIL = HOSP23_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_HOSP23 = HOSP23_DETAIL + "member_hosp23.sas7bdat"
HOSP23_SCORES = DATA_DIR + "hosp23_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

HOSP24_DETAIL = HOSP24_DIR + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp/"
MEMBER_HOSP24 = HOSP24_DETAIL + "member_hosp24.sas7bdat"
HOSP24_SCORES = DATA_DIR + "hosp24_scores_" + str(YEAR) + \
    str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"

CQM_MBR_DETAIL = DATA_DIR + "cqm_mbr_detail_" + \
    str(YEAR) + "_" + str('{0:0>2}'.format(MONTH)) + "_prosp.sas7bdat"


PROVIDER = qbfunctions.loadsas(PROVIDER)
PROVIDER = qbfunctions.decoder(PROVIDER)

# Load the SAS datasets into pandas dataframes
MEMBER_HOSP03 = qbfunctions.loadsas(MEMBER_HOSP03)
HOSP03_SCORES = qbfunctions.loadsas(HOSP03_SCORES)
MEMBER_HOSP04 = qbfunctions.loadsas(MEMBER_HOSP04)
HOSP04_SCORES = qbfunctions.loadsas(HOSP04_SCORES)
MEMBER_HOSP19 = qbfunctions.loadsas(MEMBER_HOSP19)
HOSP19_SCORES = qbfunctions.loadsas(HOSP19_SCORES)
MEMBER_HOSP20 = qbfunctions.loadsas(MEMBER_HOSP20)
HOSP20_SCORES = qbfunctions.loadsas(HOSP20_SCORES)
MEMBER_HOSP21 = qbfunctions.loadsas(MEMBER_HOSP21)
HOSP21_SCORES = qbfunctions.loadsas(HOSP21_SCORES)
MEMBER_RRAMA = qbfunctions.loadsas(MEMBER_RRAMA)
RRAMA_SCORES = qbfunctions.loadsas(RRAMA_SCORES)
MEMBER_RRACOMM = qbfunctions.loadsas(MEMBER_RRACOMM)
RRACOMM_SCORES = qbfunctions.loadsas(RRACOMM_SCORES)
MEMBER_HOSP22 = qbfunctions.loadsas(MEMBER_HOSP22)
HOSP22_SCORES = qbfunctions.loadsas(HOSP22_SCORES)
MEMBER_HOSP23 = qbfunctions.loadsas(MEMBER_HOSP23)
HOSP23_SCORES = qbfunctions.loadsas(HOSP23_SCORES)
MEMBER_HOSP24 = qbfunctions.loadsas(MEMBER_HOSP24)
HOSP24_SCORES = qbfunctions.loadsas(HOSP24_SCORES)
CQM_MBR_DETAIL = qbfunctions.loadsas(CQM_MBR_DETAIL)

# Decode the byte objects into strings
MEMBER_HOSP03 = qbfunctions.decoder(MEMBER_HOSP03)
HOSP03_SCORES = qbfunctions.decoder(HOSP03_SCORES)
MEMBER_HOSP04 = qbfunctions.decoder(MEMBER_HOSP04)
HOSP04_SCORES = qbfunctions.decoder(HOSP04_SCORES)
MEMBER_HOSP19 = qbfunctions.decoder(MEMBER_HOSP19)
HOSP19_SCORES = qbfunctions.decoder(HOSP19_SCORES)
MEMBER_HOSP20 = qbfunctions.decoder(MEMBER_HOSP20)
HOSP20_SCORES = qbfunctions.decoder(HOSP20_SCORES)
MEMBER_HOSP21 = qbfunctions.decoder(MEMBER_HOSP21)
HOSP21_SCORES = qbfunctions.decoder(HOSP21_SCORES)
MEMBER_RRAMA = qbfunctions.decoder(MEMBER_RRAMA)
RRAMA_SCORES = qbfunctions.decoder(RRAMA_SCORES)
MEMBER_RRACOMM = qbfunctions.decoder(MEMBER_RRACOMM)
RRACOMM_SCORES = qbfunctions.decoder(RRACOMM_SCORES)
MEMBER_HOSP22 = qbfunctions.decoder(MEMBER_HOSP22)
HOSP22_SCORES = qbfunctions.decoder(HOSP22_SCORES)
MEMBER_HOSP23 = qbfunctions.decoder(MEMBER_HOSP23)
HOSP23_SCORES = qbfunctions.decoder(HOSP23_SCORES)
MEMBER_HOSP24 = qbfunctions.decoder(MEMBER_HOSP24)
HOSP24_SCORES = qbfunctions.decoder(HOSP24_SCORES)
CQM_MBR_DETAIL = qbfunctions.decoder(CQM_MBR_DETAIL)

# Fix the dates into excel format


def sasdate2xl(dfcol):
    newcol = []
    for i in newcol:
        if i != "":
            i = i - 9342
        newcol.append(i)
    return newcol


MEMBER_HOSP03["EACM_BIR_DT"] = MEMBER_HOSP03["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP03["EAC_ADMM_DT"] = MEMBER_HOSP03["EAC_ADMM_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP03["EAC_DCG_DT"] = MEMBER_HOSP03["EAC_DCG_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP03["dx_svce_dt"] = MEMBER_HOSP03["dx_svce_dt"].where(MEMBER_HOSP03["dx_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP03["dx2_svce_dt"] = MEMBER_HOSP03["dx2_svce_dt"].where(MEMBER_HOSP03["dx2_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP03["dx3_svce_dt"] = MEMBER_HOSP03["dx3_svce_dt"].where(MEMBER_HOSP03["dx3_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP03["proc_svce_dt"] = MEMBER_HOSP03["proc_svce_dt"].where(MEMBER_HOSP03["proc_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP03["pall_care_svce_dt"] = MEMBER_HOSP03["pall_care_svce_dt"].where(MEMBER_HOSP03["pall_care_svce_dt"] != "").apply(lambda x: x + 21916)

MEMBER_HOSP04["EACM_BIR_DT"] = MEMBER_HOSP04["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP04["EAC_ADMM_DT"] = MEMBER_HOSP04["EAC_ADMM_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP04["EAC_DCG_DT"] = MEMBER_HOSP04["EAC_DCG_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP04["dx_svce_dt"] = MEMBER_HOSP04["dx_svce_dt"].where(MEMBER_HOSP04["dx_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP04["dx2_svce_dt"] = MEMBER_HOSP04["dx2_svce_dt"].where(MEMBER_HOSP04["dx2_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP04["dx3_svce_dt"] = MEMBER_HOSP04["dx3_svce_dt"].where(MEMBER_HOSP04["dx3_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP04["proc_svce_dt"] = MEMBER_HOSP04["proc_svce_dt"].where(MEMBER_HOSP04["proc_svce_dt"] != "").apply(lambda x: x + 21916)
MEMBER_HOSP04["pall_care_svce_dt"] = MEMBER_HOSP04["pall_care_svce_dt"].where(MEMBER_HOSP04["pall_care_svce_dt"] != "").apply(lambda x: x + 21916)

MEMBER_HOSP19["EACM_BIR_DT"] = MEMBER_HOSP19["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP19["SVCE_DT"] = MEMBER_HOSP19["SVCE_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP19["num_svce_dt"] = MEMBER_HOSP19["num_svce_dt"].where(MEMBER_HOSP19["num_svce_dt"] != "").apply(lambda x: x + 21916)

MEMBER_HOSP20["EACM_BIR_DT"] = MEMBER_HOSP20["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP20["SVCE_DT"] = MEMBER_HOSP20["SVCE_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP20["num_svce_dt"] = MEMBER_HOSP20["num_svce_dt"].where(MEMBER_HOSP20["num_svce_dt"] != "").apply(lambda x: x + 21916)

MEMBER_HOSP21["EACM_BIR_DT"] = MEMBER_HOSP21["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP21["EAC_ADMM_DT"] = MEMBER_HOSP21["EAC_ADMM_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP21["EAC_DCG_DT"] = MEMBER_HOSP21["EAC_DCG_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP21["follow_up_svce_dt"] = MEMBER_HOSP21["follow_up_svce_dt"].where(MEMBER_HOSP21["follow_up_svce_dt"] != "").apply(lambda x: x + 21916)

MEMBER_RRAMA["eacm_bir_dt"] = MEMBER_RRAMA["eacm_bir_dt"].apply(lambda x: x + 21916)
MEMBER_RRAMA["ADM_DT2"] = MEMBER_RRAMA["ADM_DT2"].apply(lambda x: x + 21916)
MEMBER_RRAMA["IESD2"] = MEMBER_RRAMA["IESD2"].apply(lambda x: x + 21916)
MEMBER_RRAMA["READMITDATE30"] = MEMBER_RRAMA["READMITDATE30"].where(MEMBER_RRAMA["READMITDATE30"] != "").apply(lambda x: x + 21916)

MEMBER_RRACOMM["eacm_bir_dt"] = MEMBER_RRACOMM["eacm_bir_dt"].apply(lambda x: x + 21916)
MEMBER_RRACOMM["ADM_DT2"] = MEMBER_RRACOMM["ADM_DT2"].apply(lambda x: x + 21916)
MEMBER_RRACOMM["IESD2"] = MEMBER_RRACOMM["IESD2"].apply(lambda x: x + 21916)
MEMBER_RRACOMM["READMITDATE30"] = MEMBER_RRACOMM["READMITDATE30"].where(MEMBER_RRACOMM["READMITDATE30"] != "").apply(lambda x: x + 21916)

MEMBER_HOSP22["EACM_BIR_DT"] = MEMBER_HOSP22["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP22["svce_dt"] = MEMBER_HOSP22["svce_dt"].apply(lambda x: x + 21916)
MEMBER_HOSP23["EACM_BIR_DT"] = MEMBER_HOSP23["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP23["svce_dt"] = MEMBER_HOSP23["svce_dt"].apply(lambda x: x + 21916)
MEMBER_HOSP24["EACM_BIR_DT"] = MEMBER_HOSP24["EACM_BIR_DT"].apply(lambda x: x + 21916)
MEMBER_HOSP24["svce_dt"] = MEMBER_HOSP24["svce_dt"].apply(lambda x: x + 21916)

CQM_MBR_DETAIL["mbr_bir_dt"] = CQM_MBR_DETAIL["mbr_bir_dt"].apply(lambda x: x + 21916)
CQM_MBR_DETAIL["LAST_PCP_VISIT_DATE"] = CQM_MBR_DETAIL["LAST_PCP_VISIT_DATE"].where(CQM_MBR_DETAIL["LAST_PCP_VISIT_DATE"] != "").apply(lambda x: x + 21916)


#print("NO FILTERING", member_hosp03)
# print(member_hosp03.columns.values)

DEFINITIONS = []
DEFINITIONS.append(['Worksheet', 'Category', 'Column', 'Definition'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Index Admission', 'Patient Number',
                    'The hospital ID for the member'])
DEFINITIONS.append(['Hosp03: Pall Care',
                    'Index Admission',
                    'Admission Date',
                    'The date that the member was admitted to an inpatient facility'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Index Admission', 'Discharge Date',
                    'The date that the member was discharged from an inpatient facility'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Index Admission', 'Claim Number',
                    'The claim number submitted for the member stay at an inpatient facility'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Index Admission', 'Type of Bill',
                    'The type of bill code that identified the claim as a stay at an inpatient facility'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Denominator Criteria', 'Description',
                    'The description of the diagnosis or procedure code that qualified the mem' +
                    'ber for the Palliative Care measure'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Denominator Criteria', 'Code',
                    'The diagnosis or procedure code that qualified the member for the Palliative Care measure'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Denominator Criteria', 'Service Date',
                    'The service date of the claim that qualified the member for the Palliative Care measure'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Numerator Criteria', 'AIS',
                    'An indicator of whether the member participates in the Advanced' +
                    ' Illness Service program (MA Only).'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Numerator Criteria', 'Code',
                    'The Palliative Care consult Diagnosis code that made the member compliant.'])
DEFINITIONS.append(['Hosp03: Pall Care', 'Numerator Criteria', 'Service Date',
                    'The date of the Palliative Care Consult'])
DEFINITIONS.append(['Hosp19 & Hosp20: 3 Day Return to ED', 'Index ED Visit',
                    'Patient Number', 'The hospital ID for the member'])
DEFINITIONS.append(['Hosp19 & Hosp20: 3 Day Return to ED', 'Index ED Visit',
                    'Service Date', 'The service date of the ED visit for the member. If the' +
                    ' ED visit spanned multiple days, this will be the most recent service date'])
DEFINITIONS.append(['Hosp19 & Hosp20: 3 Day Return to ED',
                    'Index ED Visit',
                    'Diagnosis',
                    'The primary diagnosis code for the ED visit'])
DEFINITIONS.append(['Hosp19 & Hosp20: 3 Day Return to ED',
                    'Index ED Visit',
                    'Diagnosis Description',
                    'The primary diagnosis code description for the ED visit'])
DEFINITIONS.append(
    [
        'Hosp19 & Hosp20: 3 Day Return to ED',
        'Return ED Visit',
        'Service Date',
        'The service date of the ED visit for the member. If the ED visit ' +
        'spanned multiple days, this will be the earliest service date'])
DEFINITIONS.append(['Hosp19 & Hosp20: 3 Day Return to ED',
                    'Return ED Visit',
                    'Diagnosis',
                    'The primary diagnosis code for the ED visit'])
DEFINITIONS.append(['Hosp19 & Hosp20: 3 Day Return to ED',
                    'Return ED Visit',
                    'Diagnosis Description',
                    'The primary diagnosis code description for the ED visit'])
DEFINITIONS.append(['Hosp19 & Hosp20: 3 Day Return to ED',
                    'Return ED Visit',
                    'Place of Capture',
                    'When there is a qualifying return to ED then this will display the name of the facility' +
                    ' when the return occurs to the facility of the Index ED or will show "Other"'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up',
                    'Index Admission',
                    'Patient Number',
                    'The hospital ID for the member'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up',
                    'Index Admission',
                    'Admission Date',
                    'The date that the member was admitted to an inpatient facility'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up',
                    'Index Admission',
                    'Discharge Date',
                    'The date that the member was discharged from an inpatient facility'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up', 'Index Admission', 'Claim Number',
                    'The claim number submitted for the member stay at an inpatient facility'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up', 'Index Admission', 'Type of Bill',
                    'The type of bill code that identified the claim as a stay at an inpatient facility'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up', 'Index Admission', 'Description',
                    'The description of the DRG that qualified the member for the measure'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up', 'Index Admission', 'DRG',
                    'The Diagnosis-Related Group Code'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up',
                    'Index Admission',
                    'Discharge Status',
                    'The Discharge Status Code'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up',
                    'Follow-up Visit',
                    'Service Date',
                    'The date that a follow-up visit occurred'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up',
                    'Follow-up Visit',
                    'Description',
                    'The description for the type of follow-up visit'])
DEFINITIONS.append(['Hosp21: 7 Day Follow-up', 'Follow-up Visit', 'Code',
                    'The procedure code used to identify the follow-up visit'])
DEFINITIONS.append(['Readm - MA and Comm', 'Index Admission', 'Patient Number',
                    'The hospital ID for the member'])
DEFINITIONS.append(['Readm - MA and Comm', 'Index Admission', 'Admission Date',
                    'The date the admission occurred'])
DEFINITIONS.append(['Readm - MA and Comm', 'Index Admission', 'IESD',
                    'The date that the 30 day readmission period began. This date ' +
                    'represents the discharge date from the hospital, or discharge from ' +
                    'an SNF if the member was transferred after the original admission.'])
DEFINITIONS.append(['Readm - MA and Comm', 'Index Admission', 'Diagnosis',
                    'The primary diagnosis for the index admission.'])
DEFINITIONS.append(['Readm - MA and Comm', 'Index Admission', 'Description',
                    'The primary diagnosis description for the index admission.'])
DEFINITIONS.append(['Readm - MA and Comm', 'Index Admission', 'MDC Description',
                    'The Major Diagnostic Category description for the index admission.'])
DEFINITIONS.append(['Readm - MA and Comm', 'Index Admission', 'DRG Description',
                    'The DRG description for the index admission.'])
DEFINITIONS.append(['Readm - MA and Comm', 'Readmission', 'Date',
                    'The date a readmission occurred.'])
DEFINITIONS.append(['Readm - MA and Comm', 'Readmission', 'Diagnosis',
                    'The primary diagnosis of the readmission.'])
DEFINITIONS.append(['Readm - MA and Comm',
                    'Readmission',
                    'Place of Capture',
                    'When there is a qualifying Readmission then this will display the name of the facility' +
                    ' when the return occurs to the facility of the Index Admission or will show "Other"'])


def filterdf(df, col, var):
    df = df.copy(deep=True).where(df[col] == var).dropna(how='all')
    return df


def cqm_detail_report(hospital_id, id_var, name_var, outdir):

    prov = filterdf(PROVIDER, id_var, hospital_id)
    hosp03 = filterdf(MEMBER_HOSP03, id_var, hospital_id)
    try:
        hosp03_den = hosp03['hosp03_den'].sum()
        hosp03_num = hosp03['hosp03_num'].sum()
    except BaseException:
        hosp03_den = 0
        hosp03_num = 0

    #print(hosp03_den, hosp03_num)

    hosp04 = filterdf(MEMBER_HOSP04, id_var, hospital_id)
    try:
        hosp04_den = hosp04['hosp04_den'].sum()
        hosp04_num = hosp04['hosp04_num'].sum()
    except BaseException:
        hosp04_den = 0
        hosp04_num = 0

    #print(hosp04_den, hosp04_num)

    #print("First dataset load", hosp03)
    hosp19 = filterdf(MEMBER_HOSP19, id_var, hospital_id)
    try:
        hosp19_den = hosp19['hosp19_den'].sum()
        hosp19_num = hosp19['hosp19_num'].sum()
    except BaseException:
        hosp19_den = 0
        hosp19_num = 0

    hosp20 = filterdf(MEMBER_HOSP20, id_var, hospital_id)
    try:
        hosp20_den = hosp20['hosp20_den'].sum()
        hosp20_num = hosp20['hosp20_num'].sum()
    except BaseException:
        hosp20_den = 0
        hosp20_num = 0

    hosp21 = filterdf(MEMBER_HOSP21, id_var, hospital_id)
    try:
        hosp21_den = hosp21['hosp21_den'].sum()
        hosp21_num = hosp21['hosp21_num'].sum()
    except BaseException:
        hosp21_den = 0
        hosp21_num = 0

    rrama = filterdf(MEMBER_RRAMA, id_var, hospital_id)
    try:
        rrama_den = rrama['rrama_den'].sum()
        rrama_num = rrama['rrama_num'].sum()
    except BaseException:
        rrama_den = 0
        rrama_num = 0

    rracomm = filterdf(MEMBER_RRACOMM, id_var, hospital_id)
    try:
        rracomm_den = rracomm['rracomm_den'].sum()
        rracomm_num = rracomm['rracomm_num'].sum()
    except BaseException:
        rracomm_den = 0
        rracomm_num = 0

    hosp22 = filterdf(MEMBER_HOSP22, id_var, hospital_id)
    try:
        hosp22_den = hosp22['hosp22_den'].sum()
        hosp22_num = hosp22['hosp22_num'].sum()
    except BaseException:
        hosp22_den = 0
        hosp22_num = 0

    hosp23 = filterdf(MEMBER_HOSP23, id_var, hospital_id)
    try:
        hosp23_den = hosp23['hosp23_den'].sum()
        hosp23_num = hosp23['hosp23_num'].sum()
    except BaseException:
        hosp23_den = 0
        hosp23_num = 0

    hosp24 = filterdf(MEMBER_HOSP24, id_var, hospital_id)
    try:
        hosp24_den = hosp24['hosp24_den'].sum()
        hosp24_num = hosp24['hosp24_num'].sum()
    except BaseException:
        hosp24_den = 0
        hosp24_num = 0

    mbr_detail = filterdf(CQM_MBR_DETAIL, id_var, hospital_id)

    # Create the xlsx file
    workbook = xlsxwriter.Workbook(
        outdir +
        "H_" +
        hospital_id +
        "_MO_NH_Quality_Blue_CQM_Member_Detail.xlsx",
        {
            'nan_inf_to_errors': True})

    qbfunctions.highmark_styles(workbook)

    worksheet = workbook.add_worksheet('Definitions')
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(3, 0, "CLAIMS INCURRED THROUGH: " +
                    DATES["Claims Incurred"], qbfunctions.title2)

    startrow = 5
    worksheet.write_row(startrow, 0, DEFINITIONS[0], qbfunctions.header)
    currow = startrow + 1
    worksheet.set_column(0, 2, 20)
    worksheet.set_column(3, 3, 150)
    for i in range(1, len(DEFINITIONS)):
        worksheet.write_row(currow, 0, DEFINITIONS[i], qbfunctions.table_body)
        currow = currow + 1

    worksheet = workbook.add_worksheet('Member Summary')
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' + prov[name_var].values[0],
                    qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " + DATES["Claims Paid"],
                    qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)

    qbfunctions.cqm_mbr_detail(worksheet, mbr_detail, startrow=5, startcol=0)

    worksheet = workbook.add_worksheet('Hosp03 - Pall Care MA')
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(12, 21, hosp03_den, startrow=9), qbfunctions.title2) 

    qbfunctions.hosp03_detail(worksheet, hosp03, startrow=6, startcol=0)


    worksheet = workbook.add_worksheet('Hosp04 - Pall Care COMM')
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(12, 21, hosp04_den, startrow=9), qbfunctions.title2) 
    qbfunctions.hosp04_detail(worksheet, hosp04, startrow=6, startcol=0)



    worksheet = workbook.add_worksheet("Hosp19 - 3 Day ED (MA)")
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(6, 10, hosp19_den, startrow=9), qbfunctions.title2)

    qbfunctions.hosp19_detail(
        worksheet,
        hosp19,
        startrow=6,
        startcol=0,
        title=qbfunctions.cqmbenchmarks['hosp19']['title'])

    worksheet = workbook.add_worksheet("Hosp20 - 3 Day ED (Comm)")
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(6, 10, hosp20_den, startrow=9), qbfunctions.title2)

    qbfunctions.hosp19_detail(
        worksheet,
        hosp20,
        startrow=6,
        startcol=0,
        title=qbfunctions.cqmbenchmarks['hosp20']['title'])

    worksheet = workbook.add_worksheet("Readm - MA")
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(7, 12, rrama_den, startrow=9), qbfunctions.title2)

    qbfunctions.reads_detail(worksheet, rrama, startrow=6, startcol=0,
                             title=qbfunctions.cqmbenchmarks['rrama']['title'])

    worksheet = workbook.add_worksheet("Readm - Comm")
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(7, 12, rracomm_den, startrow=9), qbfunctions.title2)

    qbfunctions.reads_detail(
        worksheet,
        rracomm,
        startrow=6,
        startcol=0,
        title=qbfunctions.cqmbenchmarks['rracomm']['title'])
    worksheet = workbook.add_worksheet("Hosp21 - 7 Day Follow-up")
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' + prov[name_var].values[0],
                    qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " + DATES["Claims Paid"],
                    qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(7, 13, hosp21_den, startrow=9), qbfunctions.title2)

    qbfunctions.hosp21_detail(worksheet, hosp21, startrow=6, startcol=0)
    
    worksheet = workbook.add_worksheet('Hosp22 - Preop Lab')
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(6, 8, hosp22_den, startrow=9), qbfunctions.title2) 

    qbfunctions.hosp22_detail(worksheet, hosp22, startrow=6, startcol=0)

    worksheet = workbook.add_worksheet('Hosp23 - Preop Cardiac')
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(6, 8, hosp23_den, startrow=9), qbfunctions.title2) 

    qbfunctions.hosp23_detail(worksheet, hosp23, startrow=6, startcol=0)
    
    worksheet = workbook.add_worksheet('Hosp24 - Preop EKG')
    worksheet.write(
        0,
        0,
        str(YEAR) +
        ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
        qbfunctions.title1)
    worksheet.write(1, 0, hospital_id + ' - ' +
                    prov[name_var].values[0], qbfunctions.title2)
    worksheet.write(2, 0, "CLAIMS PAID THROUGH: " +
                    DATES["Claims Paid"], qbfunctions.title2)
    worksheet.write(
        3,
        0,
        "CLAIMS INCURRED THROUGH: " +
        DATES["Claims Incurred"],
        qbfunctions.title2)
    worksheet.write(4, 0, rate(6, 8, hosp24_den, startrow=9), qbfunctions.title2) 

    qbfunctions.hosp24_detail(worksheet, hosp24, startrow=6, startcol=0)
    
    
    workbook.close()

#cqm_detail_report("000390050", id_var="hospital_id2", name_var="hospital_name2", outdir=QBH_DIR+"OUT/CQM_Member_Detail/")
#cqm_detail_report("000390050", id_var="quality_blue_id", name_var="quality_blue_name", outdir=QBH_DIR+"OUT/CQM_Member_Detail/")
# cqm_detail_report("003900929")
# cqm_detail_report("002098676")


def create_reports():
    hospitals = CQM_MBR_DETAIL.reset_index(drop=True).copy(deep=True)
    hospitals = hospitals.drop_duplicates(subset=['hospital_id2'])

    for i in hospitals['hospital_id2']:
        try:
            outdir = QBH_DIR + "OUT/CQM_Member_Detail/"
            cqm_detail_report(i, id_var='hospital_id2',
                              name_var='hospital_name2', outdir=outdir)
        except BaseException:
            print("ERROR: Hospital report " + i + " was unable to be created")
    for i in hospitals['quality_blue_id']:
        try:
            outdir = QBH_DIR + "OUT/CQM_Member_Detail/Aggregate"
            cqm_detail_report(i, id_var='quality_blue_id',
                              name_var='quality_blue_name', outdir=outdir)
        except BaseException:
            print("ERROR: Hospital report " + i + " was unable to be created")


create_reports()


# Create the xlsx file
WORKBOOK = xlsxwriter.Workbook(
    "/n04/data/p4vrept/Programs/Quality_Blue_Hospital_2019/OUT/" +
    "Quality_Blue_" +
    str(YEAR) +
    "_" +
    str(MONTH) +
    "_Hosp03_Detail.xlsx",
    {
        'nan_inf_to_errors': True})
qbfunctions.highmark_styles(WORKBOOK)

WORKSHEET = WORKBOOK.add_worksheet("Hosp03")
WORKSHEET.write(
    0,
    0,
    str(YEAR) +
    ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
    qbfunctions.title1)
#worksheet.write(1, 0, hospital_id, qbfunctions.title2)
WORKSHEET.write(2, 0, "Hosp03")

qbfunctions.hosp03_detail(WORKSHEET, MEMBER_HOSP03,
                          startrow=4, startcol=0, hospital=True)

WORKBOOK.close()


# Create the xlsx file
WORKBOOK = xlsxwriter.Workbook(
    "/n04/data/p4vrept/Programs/Quality_Blue_Hospital_2019/OUT/" +
    "Quality_Blue_" +
    str(YEAR) +
    "_" +
    str(MONTH) +
    "_Hosp19_Detail.xlsx",
    {
        'nan_inf_to_errors': True})

qbfunctions.highmark_styles(WORKBOOK)

WORKSHEET = WORKBOOK.add_worksheet("Hosp19")
WORKSHEET.write(
    0,
    0,
    str(YEAR) +
    ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
    qbfunctions.title1)
#worksheet.write(1, 0, hospital_id, qbfunctions.title2)
WORKSHEET.write(2, 0, "Hosp19")

qbfunctions.hosp19_detail(WORKSHEET, MEMBER_HOSP19,
                          startrow=4, startcol=0, hospital=True)

WORKBOOK.close()

# Create the xlsx file
WORKBOOK = xlsxwriter.Workbook(
    "/n04/data/p4vrept/Programs/Quality_Blue_Hospital_2019/OUT/" +
    "Quality_Blue_" +
    str(YEAR) +
    "_" +
    str(MONTH) +
    "_Hosp20_Detail.xlsx",
    {
        'nan_inf_to_errors': True})

qbfunctions.highmark_styles(WORKBOOK)

WORKSHEET = WORKBOOK.add_worksheet("Hosp20")
WORKSHEET.write(
    0,
    0,
    str(YEAR) +
    ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
    qbfunctions.title1)
#worksheet.write(1, 0, hospital_id, qbfunctions.title2)
WORKSHEET.write(2, 0, "Hosp20")

qbfunctions.hosp19_detail(WORKSHEET, MEMBER_HOSP20,
                          startrow=4, startcol=0, hospital=True)

WORKBOOK.close()


# Create the xlsx file
WORKBOOK = xlsxwriter.Workbook(
    "/n04/data/p4vrept/Programs/Quality_Blue_Hospital_2019/OUT/" +
    "Quality_Blue_" +
    str(YEAR) +
    "_" +
    str(MONTH) +
    "_Hosp21_Detail.xlsx",
    {
        'nan_inf_to_errors': True})

qbfunctions.highmark_styles(WORKBOOK)

WORKSHEET = WORKBOOK.add_worksheet("Hosp21")
WORKSHEET.write(
    0,
    0,
    str(YEAR) +
    ' QUALITY BLUE HOSPITAL - CLINICAL QUALITY METRICS DETAIL',
    qbfunctions.title1)
#WORKSHEET.write(1, 0, hospital_id, qbfunctions.title2)
WORKSHEET.write(2, 0, "Hosp21")

qbfunctions.hosp21_detail(WORKSHEET, MEMBER_HOSP21,
                          startrow=4, startcol=0, hospital=True)

WORKBOOK.close()
