import math
import os
import sys
import pandas as pd
import numpy as np
from datetime import date
import datetime

# Stars Footnotes
footnotes = []
footnotes.append("Please Note the Following")
footnotes.append(
    "- True Performance providers will NOT be paid under the Stars Incentive Program.")
footnotes.append(
    "- Aggregate Star Rating: Includes only measures with 10 or more eligible members")
footnotes.append("- Measurement Period: The reports are produced to reflect claims adjudicated at a point in time; the reports reflect claims incurred and paid by the report measurement end date.")
footnotes.append("- Measures are divided into two classes: 'Dynamic' measures cannot be 'closed' prior to year's end due to last visit or continuous care monitoring requirements; and 'Static' measures that have a one time compliance fulfillment requirement.")
footnotes.append("- Provider Issues: Providers can be terminated from Highmark networks or their identification information could be reassigned.  Providers can experience significant change in the number of members that are attributed to them.")
footnotes.append("- Various factors can affect which members appear in the denominators for each measure and can cause the denominators to change from one report to the next:")
footnotes.append("               Factors such as new diagnosis, membership attribution, members moving into or leaving a practice, measure exclusion events, continuous enrollment requirements, member age, member deaths, etc.")
footnotes.append(
    "- Gaps are defined as 'member-measure' gaps because one member might have gaps for more than one measure. ")
footnotes.append(
    "- Most Recent Encounter:  A number of measures require the use of the most recent date of service as the encounter to be used for compliance determination.  Because of this constraint a compliance status for a member ~{newline} can change from one measurement period to a subsequent measurement period. Impacted measures include Comprehensive Diabetes Care (CDC) test for HbA1c.")
footnotes.append(
    "- Beyond Remediation: This value represents gaps that cannot be 'closed' based on the remaining time in the measurement period.  ")
footnotes.append(
    "(See also the 'Stars - Column Definitions' tab for more detail on each of the columns)")


# Quality Bundle Benchmarks
qbbenchmarks = {
    'title': "MA Stars Quality Bundle",
    'bonus5': {
        'points': 25,
        'score': 4.51,
        'title': '+5 Bonus'
    },
    'bonus2': {
        'points': 22,
        'score': 4.01,
        'title': '+2 Bonus'
    },
    'max': {
        'points': 20,
        'score': 4.0,
        'title': 'Max'
    },
    'mid': {
        'points': 13,
        'score': 3.75,
        'title': 'Mid'
    },
    'min': {
        'points': 7,
        'score': 3.5,
        'title': 'Min'
    }
}


# Dictionaries of CQM benchmarks and parameters
cqmbenchmarks = {
    'title': "Clinical Quality Metrics",
    'rracomm': {
        # Display title of measure
        'title': 'Risk-Adjusted All-Cause Readmissions - Commercial',
        'region': 'PROV_REGION2',  # Region variable name of input dataset
        'hospital_id': 'PROV_NBR',  # Hospital Blue Shield # of input dataset
        'num': 'numerator',
        'den': 'denominator',
        'exp': 'expected',
        'market': 'market_expected',
        'points': 7,  # Maximum Points Available
        'pointsmid': 5,  # Mid Points Available
        'pointsmin': 3.5,  # Min Points Available
        'max': .0590,  # Maximum Benchmark (100% Points)
        'mid': .0683,  # Mid Benchmark (75% Points
        'min': .0774},  # Minimum Benchmark (50% Points)
    'rrama': {
        'title': 'Risk-Adjusted All-Cause Readmissions - MA',  # Display title of measure
        'region': 'PROV_REGION2',  # Region variable name of input dataset
        'hospital_id': 'PROV_NBR',  # Hospital Blue Shield # of input dataset
        'num': 'numerator',
        'den': 'denominator',
        'exp': 'expected',
        'market': 'market_expected',
        'points': 7,
        'pointsmid': 5,  # Mid Points Available
        'pointsmin': 3.5,  # Min Points Available
        'max': .1292,
        'mid': .1432,
        'min': .1558
    },
    'hosp03': {
        'title': 'HOSP03: Palliative Care for Complex Patients - MA',
        'sheet': 'QB_NU_SUMMARY_3',
        'hospital_id': 'obillprv',
        'num': 'NUCOMB',
        'den': 'DECOMB',
        'points': 3,
        'pointsmid': 2.25,  # Mid Points Available
        'pointsmin': 1,  # Min Points Available
        'max': .3480,
        'mid': .3225,
        'min': .2889
    },
    'hosp04': {
        'title': 'HOSP03: Palliative Care for Complex Patients - COM',
        'sheet': 'QB_NU_SUMMARY_3',
        'hospital_id': 'obillprv',
        'num': 'NUCOMB',
        'den': 'DECOMB',
        'points': 3,
        'pointsmid': 2.25,  # Mid Points Available
        'pointsmin': 1,  # Min Points Available
        'max': .1154,
        'mid': .0802,
        'min': .0530
    },
    'hosp19': {
        'title': 'HOSP19: 3 Day Return Visits to the ED - MA',
        'sheet': 'QB_NU_SUMMARY_19',
        'hospital_id': 'OBILLPRV',
        'num': 'NU',
        'den': 'DE',
        'points': 3,
        'pointsmid': 2.5,  # Mid Points Available
        'pointsmin': 2,  # Min Points Available
        'max': .0650,
        'mid': .0706,
        'min': .0812
    },
    'hosp20': {
        'title': 'HOSP20: 3 Day Return Visits to the ED - Commercial',
        'sheet': 'QB_NU_SUMMARY_20',
        'hospital_id': 'OBILLPRV',
        'num': 'NU',
        'den': 'DE',
        'points': 3,
        'pointsmid': 2.5,  # Mid Points Available
        'pointsmin': 2,  # Min Points Available
        'max': .0450,
        'mid': .0489,
        'min': .0523
    },
    'hosp21': {
        'title': 'HOSP21: 7 Day Follow-Up',
        'sheet': 'QB_NU_SUMMARY_21',
        'hospital_id': 'OBILLPRV',
        'num': 'NU',
        'den': 'DE',
        'points': 4,
        'pointsmid': 3,  # Mid Points Available
        'pointsmin': 2,  # Min Points Available
        'max': .6391,
        'mid': .6062,
        'min': .5677
    },
    'hosp22': {
        'title': 'HOSP22: Preop Laboratory Studies [Profiled]',
        'sheet': 'QB_NU_SUMMARY_22',
        'hospital_id': 'OBILLPRV',
        'num': 'NU',
        'den': 'DE',
        'points': 0,
        'pointsmid': 0,  # Mid Points Available
        'pointsmin': 0,  # Min Points Available
        'max': .0000,
        'mid': .0000,
        'min': .0000
    },
    'hosp23': {
        'title': 'HOSP23: Preop Cardiac Echo or Stress Testing [Profiled]',
        'sheet': 'QB_NU_SUMMARY_23',
        'hospital_id': 'OBILLPRV',
        'num': 'NU',
        'den': 'DE',
        'points': 0,
        'pointsmid': 0,  # Mid Points Available
        'pointsmin': 0,  # Min Points Available
        'max': .0000,
        'mid': .0000,
        'min': .0000
    },
    'hosp24': {
        'title': 'HOSP24: Preop EKG, Chest X-Ray, Pulm Function Test [Profiled]',
        'sheet': 'QB_NU_SUMMARY_24',
        'hospital_id': 'OBILLPRV',
        'num': 'NU',
        'den': 'DE',
        'points': 0,
        'pointsmid': 0,  # Mid Points Available
        'pointsmin': 0,  # Min Points Available
        'max': .0000,
        'mid': .0000,
        'min': .0000
    }
}


overallbenchmarks = {
#   'max': .63,
#   'mid': .51,
#   'min': .45
    'Large': {
        'max': .63,
        'mid': .51,
        'min': .45
    },
    'Medium': {
        'max': .62,
        'mid': .50,
        'min': .38
    },
    'Small': {
        'max': .71,
        'mid': .585,
        'min': .43
    },
    'Very-Small': {
        'max': .86,
        'mid': .60,
        'min': .41
    },
    'Specialty': {
        'max': .80,
        'mid': .60,
        'min': .34
    }
}

epibenchmarks = {
    'max': .85,
    'mid': .04,
    'min': .005,
    'MJR': {
        'title': 'Major Joint Replacement of the Lower Extremity',
        'points': 30,
    },
    'COPD': {
        'title': 'Chronic Obstructive Pulmonary Disease',
        'points': 5
    },
    'PNEU': {
        'title': 'Pneumonia',
        'points': 5,
    },
    'SEPSIS': {
        'title': 'Sepsis',
        'points': 5,
    },
    'CARD': {
        'title': 'Cardiac arrhythmia',
        'points': 5,
    },
    'PCI': {
        'title': 'Percutaneous coronary intervention',
        'points': 5,
    },
    'EGD': {
        'title': 'Esophagitis, gastroenteritis and other digestive disorders',
        'points': 5,
    },
    'MBP': {
        'title': 'Major bowel procedure',
        'points': 5,
    },
    'SFUS': {
        'title': 'Spinal fusion (non-cervical)',
        'points': 5,
    },
    'STROKE': {
        'title': 'Stroke',
        'points': 5,
    }
    
}

# This is a function used to set the score to 0 if the points available are 0


def pointsavailable(df):
    df.loc[df['Points Earned'] > df['Points Available'], 'Points Earned'] = 0
    return df

# This function processes readmissions sas datasets from James Trang


def readm(filename, measure_cd):
    # Load Commercial Readmissions
    df = loadsas(filename)

    # Clean Commercial Readmissions
    # When SAS files are read turned into pandas, character variables are loaded as a UTF8 encoded byte string.
    # Decode the byte strings into normal strings with the lambda function
    df['Region'] = df[cqmbenchmarks[measure_cd]
                      ['region']].apply(lambda x: x.decode("utf-8"))
    df['Hospital ID'] = df[cqmbenchmarks[measure_cd]
                           ['hospital_id']].apply(lambda x: x.decode("utf-8"))
    df['Denominator'] = df[cqmbenchmarks[measure_cd]['den']]
    df['Numerator'] = df[cqmbenchmarks[measure_cd]['num']]
    df['Expected Numerator'] = df[cqmbenchmarks[measure_cd]
                                  ['exp']] * df['Denominator']
    df['market_expected'] = df[cqmbenchmarks[measure_cd]['market']]

    # Set the index to Hospital ID and then join the Provider Crossreference
    # dataset.
    df = df.set_index('Hospital ID').join(
        xref.set_index('Hospital ID'), rsuffix="_XREF")
    # Reset to the default index
    df.reset_index(inplace=True)

    # If the Xref ID is blank, then set the "Keep Hospital ID" indicator to
    # True
    df['Keep Hospital ID'] = pd.isnull(df['XRef ID'])
    # Now , if Keep Hospital ID is True, use the Hospital ID from the original
    # dataset, else use the XRef ID from the Crossreference Dataset
    df['Hospital ID'] = np.where(
        df['Keep Hospital ID'],
        df['Hospital ID'],
        df['XRef ID'])

    # Aggregate the totals for hospitals that were combined in the crossreference
    # FIXME there may be a problem with using the mean market_expected, if two
    # hospitals are not in the same market
    readm_comm_groups = df.groupby('Hospital ID')
    df = readm_comm_groups.agg({
        'Denominator': np.sum,
        'Numerator': np.sum,
        'Expected Numerator': np.sum,
        'market_expected': np.mean
    })
    # Make all the NaN numbers = 0
    df = df.fillna(0)
    df['expected'] = df['Expected Numerator'] / df['Denominator']
    df['Rate'] = df['Numerator'] / df['Denominator']
    df['Risk_Adjusted_Rate'] = (
        df['Rate'] / df['expected']) * df['market_expected']
    df.reset_index(inplace=True)
    # print(df)
    # Add Cutpoints
    df['Points Available'] = df["Denominator"].apply(
        lambda x: cqmbenchmarks[measure_cd]['points'] if x >= 25 else 0)
    df['100% Points'] = cqmbenchmarks[measure_cd]['max']
    df['75% Points'] = cqmbenchmarks[measure_cd]['mid']
    df['50% Points'] = cqmbenchmarks[measure_cd]['min']
    df['Clinical Quality Metric'] = cqmbenchmarks[measure_cd]['title']
    df['Rate'] = df['Risk_Adjusted_Rate']
    # Determine Points Earned
    df['Points Earned'] = 0
    maxpoints = cqmbenchmarks[measure_cd]['points']
    df.loc[df['Risk_Adjusted_Rate'] <= df['50% Points'],
           'Points Earned'] = cqmbenchmarks[measure_cd]['pointsmin']
    df.loc[df['Risk_Adjusted_Rate'] <= df['75% Points'],
           'Points Earned'] = cqmbenchmarks[measure_cd]['pointsmid']
    df.loc[df['Risk_Adjusted_Rate'] <=
           df['100% Points'], 'Points Earned'] = maxpoints
    # Filter the pandas dataframe to only the necessary columns
    df = df[['Hospital ID',
             'Clinical Quality Metric',
             'Points Available',
             'Points Earned',
             'Rate',
             'Denominator',
             'Numerator',
             '100% Points',
             '75% Points',
             '50% Points',
             'expected',
             'market_expected']]
    df = pointsavailable(df)
    return df.copy(deep=True)


def cqm(filename, measure_cd):
    xlsx = loadxl(filename)

    # df
    df = sheet(xlsx, cqmbenchmarks[measure_cd]['sheet'])
    # print(df)
    df['Hospital ID'] = df[cqmbenchmarks[measure_cd][
        'hospital_id']].apply(lambda x: '{0:0>9}'.format(x))
    df['Denominator'] = df[cqmbenchmarks[measure_cd]['den']]
    df['Numerator'] = df[cqmbenchmarks[measure_cd]['num']]

    df = df.set_index('Hospital ID').join(
        xref.set_index('Hospital ID'), rsuffix="_XREF")
    df.reset_index(inplace=True)

    df['Keep Hospital ID'] = pd.isnull(df['XRef ID'])
    df['Hospital ID'] = np.where(
        df['Keep Hospital ID'],
        df['Hospital ID'],
        df['XRef ID'])

    df_groups = df.groupby('Hospital ID')
    df = df_groups.agg({
        'Denominator': np.sum,
        'Numerator': np.sum
    })
    df = df.reset_index()
    df = df.fillna(0)

    df['Rate'] = df['Numerator'] / df['Denominator']
    df['Points Available'] = df['Denominator'].apply(
        lambda x: cqmbenchmarks[measure_cd]['points'] if x >= 25 else 0)
    df['100% Points'] = cqmbenchmarks[measure_cd]['max']
    df['75% Points'] = cqmbenchmarks[measure_cd]['mid']
    df['50% Points'] = cqmbenchmarks[measure_cd]['min']

    if cqmbenchmarks[measure_cd]['max'] > cqmbenchmarks[measure_cd]['min']:
        df.loc[df['Rate'] < df['50% Points'], 'Points Earned'] = 0
        df.loc[df['Rate'] >= df['50% Points'],
               'Points Earned'] = cqmbenchmarks[measure_cd]['pointsmin']
        df.loc[df['Rate'] >= df['75% Points'],
               'Points Earned'] = cqmbenchmarks[measure_cd]['pointsmid']
        df.loc[df['Rate'] >= df['100% Points'],
               'Points Earned'] = cqmbenchmarks[measure_cd]['points']
    elif cqmbenchmarks[measure_cd]['max'] < cqmbenchmarks[measure_cd]['min']:
        df.loc[df['Rate'] > df['50% Points'], 'Points Earned'] = 0
        df.loc[df['Rate'] <= df['50% Points'],
               'Points Earned'] = cqmbenchmarks[measure_cd]['pointsmin']
        df.loc[df['Rate'] <= df['75% Points'],
               'Points Earned'] = cqmbenchmarks[measure_cd]['pointsmid']
        df.loc[df['Rate'] <= df['100% Points'],
               'Points Earned'] = cqmbenchmarks[measure_cd]['points']
    else:
        print("ERROR: The benchmarks for", measure_cd, "are not right")
        quit()

    df.fillna(0)

    df['Clinical Quality Metric'] = cqmbenchmarks[measure_cd]['title']

    df = df[['Hospital ID',
             'Clinical Quality Metric',
             'Points Available',
             'Points Earned',
             'Rate',
             'Denominator',
             'Numerator',
             '100% Points',
             '75% Points',
             '50% Points']]
    df = pointsavailable(df)

    return df.copy(deep=True)


def decoder(df):
    columns = list(df.columns.values)
    for i in columns:
        for j in range(0, len(df[i])):
            try:
                df[i].values[j] = df[i].values[j].decode("utf-8")
                #df[i] = df[i].apply(lambda x: x.decode("utf-8"))
            except BaseException:
                continue
    return df

# Load SAS datasets
# This function loads SAS files as a pandas dataframe


def loadsas(filename):
    # Try to load the SAS dataset. If that doesn't work, print an error and
    # quit
    try:
        sas7bdat = pd.read_sas(filename).fillna("")
        print(
            "The SAS dataset",
            filename,
            "was loaded successfully. The dataset contains:")
        print(sas7bdat.head())
        return sas7bdat
    except BaseException:
        print(
            "ERROR:",
            filename,
            "is not a valid SAS dataset. Please pass a valid SAS7BDAT file to this program")
        tracebackerror()
        quit()

        # This function loads excel files


def loadxl(filename):
    xlsx = pd.ExcelFile(filename)
    print("This Excel workbook contains the sheets", xlsx.sheet_names)
    return xlsx


# This functions loads worksheets from excel files as a pandas dataframe
def sheet(xlsx, sheetname):
    try:
        sheet = xlsx.parse(sheetname)
        print("Successfully created dataframe from Excel worksheet", sheetname)
        print("Worksheet contains:")
        print(sheet.head())
        return sheet
    except BaseException:
        print("ERROR: Unable to create dataframe from the specified worksheet.")
        tracebackerror()
        quit()

# Calculate the points earned


def score_bundle(df, star_rating, dictionary):
    df2 = df.copy(deep=True)
    df2["Points Available"] = df2[star_rating].apply(
        lambda x: dictionary['max']['points'] if x >= 1 else 0)
    df2["Points Earned"] = 0
    df2["Points Earned"] = df2[star_rating].apply(lambda x: dictionary['bonus5']['points'] if x >= dictionary['bonus5']['score']
                                                  else dictionary['bonus2']['points'] if x >= dictionary['bonus2']['score'] else dictionary['max']['points'] if x == dictionary['max']['score']
                                                  else dictionary['mid']['points'] if x >= dictionary['mid']['score'] else dictionary['min']['points'] if x >= dictionary['min']['score'] else 0)
    return df2.copy(deep=True)


# Web Safe Colors taken from https://highwire.highmark.com/sites/iwov/hwt259/docs/Highmark_BrandGuide.pdf #4.3.1
# color numbers go from 1 (darkest) to higher (lightest)
black = '#000000'  # Title2 Color
white = '#ffffff'
grey1 = '#39454e'  # Text color
grey2 = '#7a8b97'
grey3 = '#c9cfd4'
grey4 = '#e8ebec'
blue1 = '#004362'
blue2 = '#0071a9'
blue3 = '#00a4e4'  # TITLE1 Color
blue4 = '#98c5dc'  # Header Color
blue5 = '#e1eff6'
green1 = '#4c691f'
green2 = '#8dc63f'
green3 = '#e8f4d9'
orange1 = '#c15119'
orange2 = '#f09628'
orange3 = '#fcead4'
purple1 = '#4c2358'
purple2 = '#a54399'
purple3 = '#f2e3f0'
red1 = '#781419'
red2 = '#e31b23'
yellow1 = '#ffd100'


# This function binds styles to an xlsxwriter workbook object. The styles must be defined before they can be used.
# Because of the limitations of xlsxwriter, number formats have to be
# applied in style.
def highmark_styles(workbook):

    workbook.formats[0].set_font_size(11)
    workbook.formats[0].set_bg_color(white)
    workbook.formats[0].set_border(0)
    workbook.formats[0].set_font_color(grey1)
    workbook.formats[0].set_text_wrap()

    global title1
    title1 = workbook.add_format()
    title1.set_bold(True)
    title1.set_font_color(blue3)
    title1.set_border(0)
    title1.set_top(0)
    title1.set_bottom(0)
    title1.set_left(0)
    title1.set_right(0)
    title1.set_bg_color(white)

    global title2
    title2 = workbook.add_format()
    title2.set_bold(True)
    title2.set_font_color(black)
    title2.set_border(0)
    title2.set_top(0)
    title2.set_bottom(0)
    title2.set_left(0)
    title2.set_right(0)
    title2.set_bg_color(white)

    global table_title
    table_title = workbook.add_format()
    table_title.set_bold(True)
    table_title.set_bg_color(blue3)
    table_title.set_font_color(white)
    table_title.set_border(0)
    table_title.set_text_wrap(True)

    global table_title_dec
    table_title_dec = workbook.add_format()
    table_title_dec.set_bold(True)
    table_title_dec.set_bg_color(blue3)
    table_title_dec.set_font_color(white)
    table_title_dec.set_border(0)
    table_title_dec.set_text_wrap(True)
    table_title_dec.set_num_format('#0.##')

    global table_title_whole
    table_title_whole = workbook.add_format()
    table_title_whole.set_bold(True)
    table_title_whole.set_bg_color(blue3)
    table_title_whole.set_font_color(white)
    table_title_whole.set_border(0)
    table_title_whole.set_text_wrap(True)
    table_title_whole.set_num_format('#0')

    global header
    header = workbook.add_format()
    header.set_font_color(blue3)
    header.set_bg_color(white)
    header.set_bold(True)
    header.set_border(0)
    header.set_bottom(1)
    header.set_bottom_color(blue3)

    global headerwrap
    headerwrap = workbook.add_format()
    headerwrap.set_font_color(blue3)
    headerwrap.set_bg_color(white)
    headerwrap.set_bold(True)
    headerwrap.set_border(0)
    headerwrap.set_bottom(1)
    headerwrap.set_bottom_color(blue3)
    headerwrap.set_text_wrap(True)

    global headerwrap_num
    headerwrap_num = workbook.add_format()
    headerwrap_num.set_font_color(blue3)
    headerwrap_num.set_bg_color(white)
    headerwrap_num.set_bold(True)
    headerwrap_num.set_border(0)
    headerwrap_num.set_bottom(1)
    headerwrap_num.set_bottom_color(blue3)
    headerwrap_num.set_text_wrap(True)
    headerwrap_num.set_align('right')

    global headerwrap_group
    headerwrap_group = workbook.add_format()
    headerwrap_group.set_font_color(blue3)
    headerwrap_group.set_bg_color(white)
    headerwrap_group.set_bold(True)
    headerwrap_group.set_border(1)
    headerwrap_group.set_bottom(1)
    headerwrap_group.set_border_color(blue3)
    headerwrap_group.set_text_wrap(True)
    headerwrap_group.set_align('center')

    global header_num
    header_num = workbook.add_format()
    header_num.set_font_color(blue3)
    header_num.set_bg_color(white)
    header_num.set_bold(True)
    header_num.set_border(0)
    header_num.set_bottom(1)
    header_num.set_bottom_color(blue3)
    header_num.set_align('right')

    global header_center
    header_center = workbook.add_format()
    header_center.set_font_color(blue3)
    header_center.set_bg_color(white)
    header_center.set_bold(True)
    header_center.set_border(0)
    header_center.set_bottom(1)
    header_center.set_bottom_color(blue3)
    header_center.set_align('center')

    global table_body
    table_body = workbook.add_format()
    table_body.set_bold(False)
    table_body.set_font_color(black)
    table_body.set_bg_color(grey4)
    table_body.set_border(0)
    table_body.set_bottom(1)
    table_body.set_top(1)
    table_body.set_bottom_color(grey2)
    table_body.set_top_color(grey2)
    table_body.set_num_format('#,###')

    global table_body_center
    table_body_center = workbook.add_format()
    table_body_center.set_bold(False)
    table_body_center.set_font_color(black)
    table_body_center.set_bg_color(grey4)
    table_body_center.set_border(0)
    table_body_center.set_bottom(1)
    table_body_center.set_top(1)
    table_body_center.set_bottom_color(grey2)
    table_body_center.set_top_color(grey2)
    table_body_center.set_align('center')

    global table_body_na
    table_body_na = workbook.add_format()
    table_body_na.set_bold(False)
    table_body_na.set_font_color(black)
    table_body_na.set_bg_color(grey3)
    table_body_na.set_border(1)
    table_body_na.set_bottom(1)
    table_body_na.set_top(1)
    table_body_na.set_left(1)
    table_body_na.set_right(1)
    table_body_na.set_bottom_color(grey2)
    table_body_na.set_top_color(grey2)
    table_body_na.set_left_color(grey2)
    table_body_na.set_right_color(grey2)

    global table_body_date
    table_body_date = workbook.add_format()
    table_body_date.set_bold(False)
    table_body_date.set_font_color(black)
    table_body_date.set_bg_color(grey4)
    table_body_date.set_border(0)
    table_body_date.set_bottom(1)
    table_body_date.set_top(1)
    table_body_date.set_bottom_color(grey2)
    table_body_date.set_top_color(grey2)
    table_body_date.set_num_format('mmmm-yy')

    global table_body_date2
    table_body_date2 = workbook.add_format()
    table_body_date2.set_bold(False)
    table_body_date2.set_font_color(black)
    table_body_date2.set_bg_color(grey4)
    table_body_date2.set_border(0)
    table_body_date2.set_bottom(1)
    table_body_date2.set_top(1)
    table_body_date2.set_bottom_color(grey2)
    table_body_date2.set_top_color(grey2)
    table_body_date2.set_num_format('mm/dd/yyyy')
    table_body_date2.set_align('center')

    global current_score
    current_score = workbook.add_format()
    current_score.set_bold(False)
    current_score.set_font_color(black)
    current_score.set_bg_color(blue4)
    current_score.set_border(1)
    current_score.set_border_color(blue3)
    current_score.set_num_format('#0%')

    global table_body_green
    table_body_green = workbook.add_format()
    table_body_green.set_bold(False)
    table_body_green.set_font_color(black)
    table_body_green.set_bg_color(green2)
    table_body_green.set_border(0)
    table_body_green.set_bottom(1)
    table_body_green.set_top(1)
    table_body_green.set_bottom_color(grey2)
    table_body_green.set_top_color(grey2)

    global table_body_yellow
    table_body_yellow = workbook.add_format()
    table_body_yellow.set_bold(False)
    table_body_yellow.set_font_color(black)
    table_body_yellow.set_bg_color(yellow1)
    table_body_yellow.set_border(0)
    table_body_yellow.set_bottom(1)
    table_body_yellow.set_top(1)
    table_body_yellow.set_bottom_color(grey2)
    table_body_yellow.set_top_color(grey2)

    global table_body_orange
    table_body_orange = workbook.add_format()
    table_body_orange.set_bold(False)
    table_body_orange.set_font_color(black)
    table_body_orange.set_bg_color(orange2)
    table_body_orange.set_border(0)
    table_body_orange.set_bottom(1)
    table_body_orange.set_top(1)
    table_body_orange.set_bottom_color(grey2)
    table_body_orange.set_top_color(grey2)

    global table_body_red
    table_body_red = workbook.add_format()
    table_body_red.set_bold(False)
    table_body_red.set_font_color(black)
    table_body_red.set_bg_color(red2)
    table_body_red.set_border(0)
    table_body_red.set_bottom(1)
    table_body_red.set_top(1)
    table_body_red.set_bottom_color(grey2)
    table_body_red.set_top_color(grey2)

    global table_body_num
    table_body_num = workbook.add_format()
    table_body_num.set_bold(False)
    table_body_num.set_font_color(black)
    table_body_num.set_bg_color(grey4)
    table_body_num.set_border(0)
    table_body_num.set_bottom(1)
    table_body_num.set_top(1)
    table_body_num.set_bottom_color(grey2)
    table_body_num.set_top_color(grey2)
    table_body_num.set_align('right')

    global table_body_num2
    table_body_num2 = workbook.add_format()
    table_body_num2.set_bold(False)
    table_body_num2.set_font_color(black)
    table_body_num2.set_bg_color(grey4)
    table_body_num2.set_border(0)
    table_body_num2.set_bottom(1)
    table_body_num2.set_top(1)
    table_body_num2.set_bottom_color(grey2)
    table_body_num2.set_top_color(grey2)
    table_body_num2.set_align('right')
    table_body_num2.set_num_format('#0.##')

    global table_body_pct
    table_body_pct = workbook.add_format()
    table_body_pct.set_bold(False)
    table_body_pct.set_font_color(black)
    table_body_pct.set_bg_color(grey4)
    table_body_pct.set_border(0)
    table_body_pct.set_bottom(1)
    table_body_pct.set_top(1)
    table_body_pct.set_bottom_color(grey2)
    table_body_pct.set_top_color(grey2)
    table_body_pct.set_num_format('#0%')
    table_body_pct.set_align('right')

    global table_body_pct_red
    table_body_pct_red = workbook.add_format()
    table_body_pct_red.set_bold(False)
    table_body_pct_red.set_font_color(black)
    table_body_pct_red.set_bg_color(red2)
    table_body_pct_red.set_border(0)
    table_body_pct_red.set_bottom(1)
    table_body_pct_red.set_top(1)
    table_body_pct_red.set_bottom_color(grey2)
    table_body_pct_red.set_top_color(grey2)
    table_body_pct_red.set_num_format('#0%')
    table_body_pct_red.set_align('right')

    global table_body_pct_yellow
    table_body_pct_yellow = workbook.add_format()
    table_body_pct_yellow.set_bold(False)
    table_body_pct_yellow.set_font_color(black)
    table_body_pct_yellow.set_bg_color(yellow1)
    table_body_pct_yellow.set_border(0)
    table_body_pct_yellow.set_bottom(1)
    table_body_pct_yellow.set_top(1)
    table_body_pct_yellow.set_bottom_color(grey2)
    table_body_pct_yellow.set_top_color(grey2)
    table_body_pct_yellow.set_num_format('#0%')
    table_body_pct_yellow.set_align('right')

    global table_body_pct_orange
    table_body_pct_orange = workbook.add_format()
    table_body_pct_orange.set_bold(False)
    table_body_pct_orange.set_font_color(black)
    table_body_pct_orange.set_bg_color(orange2)
    table_body_pct_orange.set_border(0)
    table_body_pct_orange.set_bottom(1)
    table_body_pct_orange.set_top(1)
    table_body_pct_orange.set_bottom_color(grey2)
    table_body_pct_orange.set_top_color(grey2)
    table_body_pct_orange.set_num_format('#0%')
    table_body_pct_orange.set_align('right')

    global table_body_pct_green
    table_body_pct_green = workbook.add_format()
    table_body_pct_green.set_bold(False)
    table_body_pct_green.set_font_color(black)
    table_body_pct_green.set_bg_color(green2)
    table_body_pct_green.set_border(0)
    table_body_pct_green.set_bottom(1)
    table_body_pct_green.set_top(1)
    table_body_pct_green.set_bottom_color(grey2)
    table_body_pct_green.set_top_color(grey2)
    table_body_pct_green.set_num_format('#0%')
    table_body_pct_green.set_align('right')

    global table_body_dollar
    table_body_dollar = workbook.add_format()
    table_body_dollar.set_bold(False)
    table_body_dollar.set_font_color(black)
    table_body_dollar.set_bg_color(grey4)
    table_body_dollar.set_border(0)
    table_body_dollar.set_bottom(1)
    table_body_dollar.set_top(1)
    table_body_dollar.set_bottom_color(grey2)
    table_body_dollar.set_top_color(grey2)
    table_body_dollar.set_num_format('$#,##0.00')

    global table_body_dollar_red
    table_body_dollar_red = workbook.add_format()
    table_body_dollar_red.set_bold(False)
    table_body_dollar_red.set_font_color(black)
    table_body_dollar_red.set_bg_color(red2)
    table_body_dollar_red.set_border(0)
    table_body_dollar_red.set_bottom(1)
    table_body_dollar_red.set_top(1)
    table_body_dollar_red.set_bottom_color(grey2)
    table_body_dollar_red.set_top_color(grey2)
    table_body_dollar_red.set_num_format('$#,##0.00')

    global table_body_dollar_orange
    table_body_dollar_orange = workbook.add_format()
    table_body_dollar_orange.set_bold(False)
    table_body_dollar_orange.set_font_color(black)
    table_body_dollar_orange.set_bg_color(orange2)
    table_body_dollar_orange.set_border(0)
    table_body_dollar_orange.set_bottom(1)
    table_body_dollar_orange.set_top(1)
    table_body_dollar_orange.set_bottom_color(grey2)
    table_body_dollar_orange.set_top_color(grey2)
    table_body_dollar_orange.set_num_format('$#,##0.00')

    global table_body_dollar_green
    table_body_dollar_green = workbook.add_format()
    table_body_dollar_green.set_bold(False)
    table_body_dollar_green.set_font_color(black)
    table_body_dollar_green.set_bg_color(green2)
    table_body_dollar_green.set_border(0)
    table_body_dollar_green.set_bottom(1)
    table_body_dollar_green.set_top(1)
    table_body_dollar_green.set_bottom_color(grey2)
    table_body_dollar_green.set_top_color(grey2)
    table_body_dollar_green.set_num_format('$#,##0.00')

    global table_body_dollar_yellow
    table_body_dollar_yellow = workbook.add_format()
    table_body_dollar_yellow.set_bold(False)
    table_body_dollar_yellow.set_font_color(black)
    table_body_dollar_yellow.set_bg_color(yellow1)
    table_body_dollar_yellow.set_border(0)
    table_body_dollar_yellow.set_bottom(1)
    table_body_dollar_yellow.set_top(1)
    table_body_dollar_yellow.set_bottom_color(grey2)
    table_body_dollar_yellow.set_top_color(grey2)
    table_body_dollar_yellow.set_num_format('$#,##0.00')

    global table_body_pct2
    table_body_pct2 = workbook.add_format()
    table_body_pct2.set_bold(False)
    table_body_pct2.set_font_color(black)
    table_body_pct2.set_bg_color(grey4)
    table_body_pct2.set_border(0)
    table_body_pct2.set_bottom(1)
    table_body_pct2.set_top(1)
    table_body_pct2.set_bottom_color(grey2)
    table_body_pct2.set_top_color(grey2)
    table_body_pct2.set_num_format('#0.##%')
    table_body_pct2.set_align('right')

    global table_body_pct2_red
    table_body_pct2_red = workbook.add_format()
    table_body_pct2_red.set_bold(False)
    table_body_pct2_red.set_font_color(black)
    table_body_pct2_red.set_bg_color(red2)
    table_body_pct2_red.set_border(0)
    table_body_pct2_red.set_bottom(1)
    table_body_pct2_red.set_top(1)
    table_body_pct2_red.set_bottom_color(grey2)
    table_body_pct2_red.set_top_color(grey2)
    table_body_pct2_red.set_num_format('#0.##%')
    table_body_pct2_red.set_align('right')

    global table_body_pct2_yellow
    table_body_pct2_yellow = workbook.add_format()
    table_body_pct2_yellow.set_bold(False)
    table_body_pct2_yellow.set_font_color(black)
    table_body_pct2_yellow.set_bg_color(yellow1)
    table_body_pct2_yellow.set_border(0)
    table_body_pct2_yellow.set_bottom(1)
    table_body_pct2_yellow.set_top(1)
    table_body_pct2_yellow.set_bottom_color(grey2)
    table_body_pct2_yellow.set_top_color(grey2)
    table_body_pct2_yellow.set_num_format('#0.##%')
    table_body_pct2_yellow.set_align('right')

    global table_body_pct2_orange
    table_body_pct2_orange = workbook.add_format()
    table_body_pct2_orange.set_bold(False)
    table_body_pct2_orange.set_font_color(black)
    table_body_pct2_orange.set_bg_color(orange2)
    table_body_pct2_orange.set_border(0)
    table_body_pct2_orange.set_bottom(1)
    table_body_pct2_orange.set_top(1)
    table_body_pct2_orange.set_bottom_color(grey2)
    table_body_pct2_orange.set_top_color(grey2)
    table_body_pct2_orange.set_num_format('#0.##%')
    table_body_pct2_orange.set_align('right')

    global table_body_pct2_green
    table_body_pct2_green = workbook.add_format()
    table_body_pct2_green.set_bold(False)
    table_body_pct2_green.set_font_color(black)
    table_body_pct2_green.set_bg_color(green2)
    table_body_pct2_green.set_border(0)
    table_body_pct2_green.set_bottom(1)
    table_body_pct2_green.set_top(1)
    table_body_pct2_green.set_bottom_color(grey2)
    table_body_pct2_green.set_top_color(grey2)
    table_body_pct2_green.set_num_format('#0.##%')
    table_body_pct2_green.set_align('right')


def legend(worksheet, startrow=1, startcol=9):
    worksheet.merge_range(
        startrow,
        startcol,
        startrow,
        startcol + 1,
        "Legend:",
        table_title)
    worksheet.write(startrow + 1, startcol, " ", table_body_green)
    worksheet.write(startrow + 2, startcol, " ", table_body_yellow)
    worksheet.write(startrow + 3, startcol, " ", table_body_orange)
    worksheet.write(startrow + 4, startcol, " ", table_body_red)
    worksheet.write(startrow + 1, startcol + 1, "Max", table_body)
    worksheet.write(startrow + 2, startcol + 1, "Mid", table_body)
    worksheet.write(startrow + 3, startcol + 1, "Min", table_body)
    worksheet.write(startrow + 4, startcol + 1, "Zero", table_body)


def epitables(
        worksheet,
        startrow,
        hospital_size,
        epi_available,
        epi_earned,
        MJRcount,
        MJRcost,
        MJRpoints,
        MJRearned,
        MJRweight,
        MJRpercspd,
        MJRpctl,
        MJRimprovement,
        COPDcount,
        COPDcost,
        COPDpoints,
        COPDearned,
        COPDweight,
        COPDpercspd,
        COPDpctl,
        COPDimprovement,
        EGDcount,
        EGDcost,
        EGDpoints,
        EGDearned,
        EGDweight,
        EGDpercspd,
        EGDpctl,
        EGDimprovement,
        SEPSIScount,
        SEPSIScost,
        SEPSISpoints,
        SEPSISearned,
        SEPSISweight,
        SEPSISpercspd,
        SEPSISpctl,
        SEPSISimprovement,
        CARDcount,
        CARDcost,
        CARDpoints,
        CARDearned,
        CARDweight,
        CARDpercspd,
        CARDpctl,
        CARDimprovement,
        PCIcount,
        PCIcost,
        PCIpoints,
        PCIearned,
        PCIweight,
        PCIpercspd,
        PCIpctl,
        PCIimprovement,
        PNEUcount,
        PNEUcost,
        PNEUpoints,
        PNEUearned,
        PNEUweight,
        PNEUpercspd,
        PNEUpctl,
        PNEUimprovement,
        MBPcount,
        MBPcost,
        MBPpoints,
        MBPearned,
        MBPweight,
        MBPpercspd,
        MBPpctl,
        MBPimprovement,
        SFUScount,
        SFUScost,
        SFUSpoints,
        SFUSearned,
        SFUSweight,
        SFUSpercspd,
        SFUSpctl,
        SFUSimprovement,
        STROKEcount,
        STROKEcost,
        STROKEpoints,
        STROKEearned,
        STROKEweight,
        STROKEpercspd,
        STROKEpctl,
        STROKEimprovement,
        MJRScoring_Flag,
        COPDScoring_Flag,
        EGDScoring_Flag,
        SEPSISScoring_Flag,
        CARDScoring_Flag,
        PCIScoring_Flag,
        PNEUScoring_Flag,
        MBPScoring_Flag,
        SFUSScoring_Flag,
        STROKEScoring_Flag):
    print("************************worksheet.merge_range( IN DEF EPITABLES********************************************")
    print("************************worksheet.merge_range( IN DEF EPITABLES********************************************")
    print("************************worksheet.merge_range( IN DEF EPITABLES********************************************")
    worksheet.merge_range(
        startrow,
        0,
        startrow,
        14,
        'EPISODES OF CARE',
        table_title)
    worksheet.merge_range(startrow + 1, 0, startrow + 1, 3, "Episode", header)
    worksheet.write(startrow + 1, 4, 'Available', header_num)
    worksheet.write(startrow + 1, 5, 'Earned', header_center)
    worksheet.write(startrow + 1, 6, 'Percentile', header_center)
    worksheet.merge_range(startrow + 1, 7, startrow + 1, 8, 'Improvement', header_num)
    worksheet.merge_range(startrow + 1,9,startrow + 1,10,'Point Weight',header_num)
    worksheet.merge_range(startrow + 1,11,startrow + 1,12,'Tot Qualified $',header_num)
    worksheet.merge_range(startrow + 1,13,startrow + 1,14,'Percent Spend',header_num)
    worksheet.merge_range(startrow, 16,startrow,25,'THRESHOLDS',table_title)
    worksheet.merge_range(startrow + 1,16,startrow + 1,17,'100% Points (Max)',header_num)
    worksheet.merge_range(startrow + 1,18,startrow + 1,21,'75% Points (Mid)',header_num)
    worksheet.merge_range(startrow + 1,22,startrow + 1,25,'25-75% Points (Min)',header_num)
    #worksheet.merge_range(startrow + 2,16,startrow + 2,17,'=75th Percentile',header_num)
    #worksheet.merge_range(startrow + 2,18,startrow + 2,19,'=4% Improvement',header_num)
    #worksheet.merge_range(startrow + 2,20,startrow + 2,21,'=0.5% Improvement',header_num)


    def epirow(
            row,
            measure_cd,
            epicount,
            cost,
            points,
            earned,
            pctl,
            improvement,
            ptweight,
            pctspend,
            scoring_flag):
        worksheet.merge_range(
            row,
            0,
            row,
            3,
            epibenchmarks[measure_cd]['title'],
            table_body)
        worksheet.write(row, 4, points, table_body)
        if points != 0 and earned == 0:
            worksheet.write(row, 5, earned, table_body_red)
        if earned != 0 or points == 0:
            worksheet.write(row, 5, earned, table_body)
        worksheet.write(row, 6, math.floor(pctl * 100), table_body)
        if scoring_flag == 'Not Qualified':
            worksheet.merge_range(row, 4, row, 6, 'Not Qualified', table_body)
        if scoring_flag == ' ':
            worksheet.merge_range(row, 4, row, 6, 'Not Qualified', table_body)
        if scoring_flag == 'Profiled':
            worksheet.merge_range(row, 4, row, 6, 'Profiled', table_body)
        worksheet.merge_range(row, 7, row, 8, improvement, table_body_pct2)
        worksheet.merge_range(row, 9, row, 10, ptweight, table_body_pct)
        worksheet.merge_range(row, 11, row, 12, cost, table_body_dollar)
        worksheet.merge_range(row, 13, row, 14, pctspend, table_body_pct2)
        if earned == points and points != 0:
            worksheet.merge_range(row,16,row,17,u'\u226575th Percentile',table_body_pct_green)
            #worksheet.merge_range(row, 18, row, 19, str(points) + " Points", table_body_pct_green)
            worksheet.write(row, 10, math.floor(pctl * 100), table_body_green)
        else:
            worksheet.merge_range(row,16,row,17,u'\u226575th Percentile',table_body_pct)
            #worksheet.merge_range(row,18,row,19,str(points) + " Points",table_body_pct)
        if earned == points * .75 and points != 0:
            if hospital_size == 'Large':
                worksheet.merge_range(row,18,row,21,u'\u22652% Improvement',table_body_pct_yellow)
            elif hospital_size == 'Medium':
                worksheet.merge_range(row,18,row,21,u'\u22653% Improvement',table_body_pct_yellow)
            else:
                worksheet.merge_range(row,18,row,21,u'\u22654% Improvement',table_body_pct_yellow)
            #worksheet.merge_range(row, 22, row, 23, str(points * .75) + " Points", table_body_pct_yellow)
            worksheet.write(row, 12, improvement, table_body_pct2_yellow)
        else:
            if hospital_size == 'Large':
        	  	  worksheet.merge_range(row,18,row,21,u'\u22652% Improvement',table_body_pct)
            elif hospital_size == 'Medium':
                worksheet.merge_range(row,18,row,21,u'\u22653% Improvement',table_body_pct_yellow)
            else:
        	  	  worksheet.merge_range(row,18,row,21,u'\u22654% Improvement',table_body_pct)
            #worksheet.merge_range(row, 22, row, 23, str(points * .75) + " Points", table_body_pct)
        if earned == points * .5 and points != 0:
            worksheet.merge_range(row,22,row,25,u'\u2265.5% Improvement',table_body_pct_orange)
            #worksheet.merge_range(row, 26, row, 27, str(points * .5) + " Points", table_body_pct_orange)
            worksheet.write(row, 12, improvement, table_body_pct2_orange)
        else:
            worksheet.merge_range(row,22,row,25,u'\u2265.5% Improvement',table_body_pct)
            #worksheet.merge_range(row, 26, row, 27, str(points * .5) + " Points", table_body_pct)

    epirow(
        startrow + 2,
        'MJR',
        MJRcount,
        MJRcost,
        MJRpoints,
        MJRearned,
        MJRpctl,
        MJRimprovement,
        MJRweight,
        MJRpercspd,
        MJRScoring_Flag)
    epirow(
        startrow + 3,
        'COPD',
        COPDcount,
        COPDcost,
        COPDpoints,
        COPDearned,
        COPDpctl,
        COPDimprovement,
        COPDweight,
        COPDpercspd,
        COPDScoring_Flag)
    epirow(
        startrow + 4,
        'EGD',
        EGDcount,
        EGDcost,
        EGDpoints,
        EGDearned,
        EGDpctl,
        EGDimprovement,
        EGDweight,
        EGDpercspd,
        EGDScoring_Flag)
    epirow(
        startrow + 5,
        'SEPSIS',
        SEPSIScount,
        SEPSIScost,
        SEPSISpoints,
        SEPSISearned,
        SEPSISpctl,
        SEPSISimprovement,
        SEPSISweight,
        SEPSISpercspd,
        SEPSISScoring_Flag)
    epirow(
        startrow + 6,
        'CARD',
        CARDcount,
        CARDcost,
        CARDpoints,
        CARDearned,
        CARDpctl,
        CARDimprovement,
        CARDweight,
        CARDpercspd,
        CARDScoring_Flag)
    epirow(
        startrow + 7,
        'PCI',
        PCIcount,
        PCIcost,
        PCIpoints,
        PCIearned,
        PCIpctl,
        PCIimprovement,
        PCIweight,
        PCIpercspd,
        PCIScoring_Flag)
    epirow(
        startrow + 8,
        'PNEU',
        PNEUcount,
        PNEUcost,
        PNEUpoints,
        PNEUearned,
        PNEUpctl,
        PNEUimprovement,
        PNEUweight,
        PNEUpercspd,
        PNEUScoring_Flag)
    epirow(
        startrow + 9,
        'MBP',
        MBPcount,
        MBPcost,
        MBPpoints,
        MBPearned,
        MBPpctl,
        MBPimprovement,
        MBPweight,
        MBPpercspd,
        MBPScoring_Flag)
    epirow(
        startrow + 10,
        'SFUS',
        SFUScount,
        SFUScost,
        SFUSpoints,
        SFUSearned,
        SFUSpctl,
        SFUSimprovement,
        SFUSweight,
        SFUSpercspd,
        SFUSScoring_Flag)
    epirow(
        startrow + 11,
        'STROKE',
        STROKEcount,
        STROKEcost,
        STROKEpoints,
        STROKEearned,
        STROKEpctl,
        STROKEimprovement,
        STROKEweight,
        STROKEpercspd,
        STROKEScoring_Flag)

    worksheet.merge_range(
        startrow + 12,
        0,
        startrow + 12,
        3,
        'TOTAL',
        table_title)
    worksheet.write(startrow + 12, 4, epi_available, table_title)
    worksheet.write(startrow + 12, 5, epi_earned, table_title_dec)
    worksheet.merge_range(startrow + 12, 6, startrow + 12, 14, '', table_title)
    worksheet.merge_range(startrow + 12, 16, startrow + 12, 25, '', table_title)


# FIXME - Rewrite to use itercols function

def qbtables(
        worksheet,
        startrow,
        startcol=0,
        qb_star_rating=0,
        qb_points=0,
        qb_earned=0,
        score=True,
        thresholds=True):

    # If score
    print("************************IF SCORE IN DEF QBTABLES********************************************")
    print("************************IF SCORE IN DEF QBTABLES********************************************")
    print("************************IF SCORE IN DEF QBTABLES********************************************")
    if score:
        worksheet.merge_range(
            startrow,
            startcol,
            startrow,
            startcol + 3,
            'QUALITY BUNDLE',
            table_title)
        worksheet.merge_range(
            startrow + 1,
            startcol,
            startrow + 1,
            2,
            "Measure",
            header)
        worksheet.write(startrow + 1, startcol + 3, 'Score', header_num)
        worksheet.merge_range(
            startrow + 2,
            startcol,
            startrow + 2,
            startcol + 2,
            'Star Rating',
            table_body)
        worksheet.write(
            startrow + 2,
            startcol + 3,
            qb_star_rating,
            table_body_num2)
        worksheet.merge_range(
            startrow + 3,
            startcol,
            startrow + 3,
            startcol + 2,
            'POINTS EARNED',
            table_title)
        worksheet.write(startrow + 3, startcol + 3, qb_earned, table_title)
    elif score == False:
        startcol = startcol - 5

    if thresholds:
        worksheet.merge_range(
            startrow,
            startcol + 5,
            startrow,
            startcol + 8,
            "THRESHOLDS",
            table_title)
        worksheet.merge_range(
            startrow + 1,
            startcol + 5,
            startrow + 1,
            startcol + 6,
            "Points",
            header)
        worksheet.merge_range(
            startrow + 1,
            startcol + 7,
            startrow + 1,
            startcol + 8,
            "Criteria",
            header)
        if qb_earned == qbbenchmarks['bonus5']['points']:
            if score:
                worksheet.write(
                    startrow + 2,
                    startcol + 3,
                    qb_star_rating,
                    table_body_green)
            worksheet.merge_range(
                startrow + 2,
                startcol + 5,
                startrow + 2,
                startcol + 6,
                qbbenchmarks['bonus5']['title'],
                table_body_green)
            worksheet.merge_range(startrow +
                                  2, startcol +
                                  7, startrow +
                                  2, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['bonus5']['score']) +
                                  " Stars", table_body_green)
        else:
            worksheet.merge_range(
                startrow + 2,
                startcol + 5,
                startrow + 2,
                startcol + 6,
                qbbenchmarks['bonus5']['title'],
                table_body)
            worksheet.merge_range(startrow +
                                  2, startcol +
                                  7, startrow +
                                  2, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['bonus5']['score']) +
                                  " Stars", table_body)
        if qb_earned == qbbenchmarks['bonus2']['points']:
            if score:
                worksheet.write(
                    startrow + 2,
                    startcol + 3,
                    qb_star_rating,
                    table_body_green)
            worksheet.merge_range(
                startrow + 3,
                startcol + 5,
                startrow + 3,
                startcol + 6,
                qbbenchmarks['bonus2']['title'],
                table_body_green)
            worksheet.merge_range(startrow +
                                  3, startcol +
                                  7, startrow +
                                  3, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['bonus2']['score']) +
                                  " Stars", table_body_green)
        else:
            worksheet.merge_range(
                startrow + 3,
                startcol + 5,
                startrow + 3,
                startcol + 6,
                qbbenchmarks['bonus2']['title'],
                table_body)
            worksheet.merge_range(startrow +
                                  3, startcol +
                                  7, startrow +
                                  3, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['bonus2']['score']) +
                                  " Stars", table_body)
        if qb_earned == qbbenchmarks['max']['points']:
            if score:
                worksheet.write(
                    startrow + 2,
                    3,
                    qb_star_rating,
                    table_body_green)
            worksheet.merge_range(startrow +
                                  4, startcol +
                                  5, startrow +
                                  4, startcol +
                                  6, str(qbbenchmarks['max']['points']) +
                                  " Points", table_body_green)
            worksheet.merge_range(startrow +
                                  4, startcol +
                                  7, startrow +
                                  4, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['max']['score']) +
                                  " Stars", table_body_green)
        else:
            worksheet.merge_range(startrow +
                                  4, startcol +
                                  5, startrow +
                                  4, startcol +
                                  6, str(qbbenchmarks['max']['points']) +
                                  " Points", table_body)
            worksheet.merge_range(startrow +
                                  4, startcol +
                                  7, startrow +
                                  4, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['max']['score']) +
                                  " Stars", table_body)
        if qb_earned == qbbenchmarks['mid']['points']:
            if score:
                worksheet.write(
                    startrow + 2,
                    startcol + 3,
                    qb_star_rating,
                    table_body_yellow)
            worksheet.merge_range(startrow +
                                  5, startcol +
                                  5, startrow +
                                  5, startcol +
                                  6, str(qbbenchmarks['mid']['points']) +
                                  " Points", table_body_yellow)
            worksheet.merge_range(startrow +
                                  5, startcol +
                                  7, startrow +
                                  5, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['mid']['score']) +
                                  " Stars", table_body_yellow)
        else:
            worksheet.merge_range(startrow +
                                  5, startcol +
                                  5, startrow +
                                  5, startcol +
                                  6, str(qbbenchmarks['mid']['points']) +
                                  " Points", table_body)
            worksheet.merge_range(startrow +
                                  5, startcol +
                                  7, startrow +
                                  5, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['mid']['score']) +
                                  " Stars", table_body)
        if qb_earned == qbbenchmarks['min']['points']:
            if score:
                worksheet.write(
                    startrow + 2,
                    startcol + 3,
                    qb_star_rating,
                    table_body_orange)
            worksheet.merge_range(startrow +
                                  6, startcol +
                                  5, startrow +
                                  6, startcol +
                                  6, str(qbbenchmarks['min']['points']) +
                                  " Points", table_body_orange)
            worksheet.merge_range(startrow +
                                  6, startcol +
                                  7, startrow +
                                  6, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['min']['score']) +
                                  " Stars", table_body_orange)
        else:
            worksheet.merge_range(startrow +
                                  6, startcol +
                                  5, startrow +
                                  6, startcol +
                                  6, str(qbbenchmarks['min']['points']) +
                                  " Points", table_body)
            worksheet.merge_range(startrow +
                                  6, startcol +
                                  7, startrow +
                                  6, startcol +
                                  8, u"\u2265" +
                                  str(qbbenchmarks['min']['score']) +
                                  " Stars", table_body)
        if qb_earned == 0 and qb_points != 0:
            if score:
                worksheet.write(
                    startrow + 2,
                    startcol + 3,
                    qb_star_rating,
                    table_body_red)
        worksheet.merge_range(
            startrow + 7,
            startcol + 5,
            startrow + 7,
            startcol + 8,
            "",
            table_title)


def itercol(
        ws,
        startrow=0,
        startcol=0,
        length=0,
        text=" ",
        style=None,
        endrow=False):
    if not endrow:
        endrow = startrow
    if length < 0:
        print("Length must be 1 or longer")
    elif length == 0:
        try:
            ws.write(startrow, startcol, text, style)
        except BaseException:
            ws.write(startrow, startcol, '', style)
        return startcol + 1
    else:
        try:
            ws.merge_range(
                startrow,
                startcol,
                endrow,
                startcol +
                length,
                text,
                style)
        except BaseException:
            ws.merge_range(
                startrow,
                startcol,
                endrow,
                startcol +
                length,
                '',
                style)
        return startcol + length + 1


def colnum_string(n):
    string = ""
    n = n + 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

# output:AB


def episodes(filename):

    xlsx = loadxl(filename)
    df = sheet(xlsx, 'Sheet1')

    df['quality_blue_id'] = df['BSID'].apply(lambda x: '{0:0>9}'.format(x))

    df['measure_cd'] = df['Episode'].apply(lambda x: 'MJR' if x == 'Major joint replacement of the lower extremity'
                                           else 'COPD' if x == 'Chronic obstructive pulmonary disease, bronchitis, asthma'
                                           else 'CARD' if x == 'Cardiac arrhythmia'
                                           else 'EGD' if x == 'Esophagitis, gastroenteritis and other digestive disorders'
                                           else 'MBP' if x == 'Major bowel procedure'
                                           else 'PCI' if x == 'Percutaneous coronary intervention'
                                           else 'SEPSIS' if x == 'Sepsis'
                                           else 'SFUS' if x == 'Spinal fusion (non-cervical)'
                                           else 'STROKE' if x == 'Stroke'
                                           else 'PNEU' if x == 'Simple pneumonia and respiratory infections' else '')
    print("LINE 1636 QBFUNCTIONS.PY LINE 1636 QBFUNCTIONS.PY LINE 1636 QBFUNCTIONS.PY LINE 1636 QBFUNCTIONS.PY LINE 1636 QBFUNCTIONS.PY")
    print("TRYING TO FILTER OUT")
    df = df[df["measure_cd"] != ''].copy(deep=True)
    print("LINE 1638 QBFUNCTIONS.PY LINE 1638 QBFUNCTIONS.PY LINE 1638 QBFUNCTIONS.PY LINE 1638 QBFUNCTIONS.PY LINE 1638 QBFUNCTIONS.PY")
    print(df)
    #VPF - 11/272018 NO LONGER USING SETTING THE POINTS FOR EPISODES BUT WILL BE USING THE INPUT FROM THE EPISODE SCORING SINCE IT IS NOT PROPORTIONAL
    #try:
    #    df['Points Available'] = df['measure_cd'].apply(
    #        lambda x: epibenchmarks[x]['points'])
    #except BaseException:
    #    df['Points Available'] = df['Points_Available']
    df['Minimum Number'] = df['Episode_Count'].apply(
        lambda x: 0 if x < 10 else 1)
    # df['Points Available'] = df['Points Available'] * df['Minimum Number']
    df['max_dist'] = df['Avg_Cost_75P'] - df['Hospital_Avg_Cost']
    df['mid_dist'] = df['Prev_Hospital_Avg_Cost'] - \
        (.04 * df['Prev_Hospital_Avg_Cost']) - df['Hospital_Avg_Cost']
    df['min_dist'] = df['Prev_Hospital_Avg_Cost'] - \
        (.005 * df['Prev_Hospital_Avg_Cost']) - df['Hospital_Avg_Cost']

    df['Points Earned %'] = 0
    # df.loc[df['Improvement'] >= epibenchmarks['min'], 'Points Earned %'] = .5
    # df.loc[df['Improvement'] >= epibenchmarks['mid'], 'Points Earned %'] = .75
    # df.loc[df['Percentile'] >= epibenchmarks['max'], 'Points Earned %'] = 1
    # df['Points Earned'] = df['Points Available'] * df['Points Earned %']
    df['Points Earned'] = df['Points_Earned']

    return df.copy(deep=True)


def comment(number, neg_string, pos_string):
    if number > 0:
        comment = str(abs(number)) + ' ' + pos_string
    elif number < 0:
        comment = str(abs(number)) + ' ' + neg_string
    else:
        comment = "No Change"
    return comment


def matrix(
        worksheet,
        startrow,
        epi_earned,
        cqm_earned,
        qb_points,
        qb_earned,
        qb_star_rating,
        hosp03_points,
        hosp03_earned,
        hosp03_rate,
        hosp03_denominator,
        hosp03_numerator,
        hosp04_points,
        hosp04_earned,
        hosp04_rate,
        hosp04_denominator,
        hosp04_numerator,
        hosp19_points,
        hosp19_earned,
        hosp19_rate,
        hosp19_denominator,
        hosp19_numerator,
        hosp20_points,
        hosp20_earned,
        hosp20_rate,
        hosp20_denominator,
        hosp20_numerator,
        hosp21_points,
        hosp21_earned,
        hosp21_rate,
        hosp21_denominator,
        hosp21_numerator,
        rrama_points,
        rrama_earned,
        rrama_rate,
        rrama_denominator,
        rrama_numerator,
        rrama_expected,
        rrama_mkt,
        rracomm_points,
        rracomm_earned,
        rracomm_rate,
        rracomm_denominator,
        rracomm_numerator,
        rracomm_expected,
        rracomm_mkt,
        MJRpoints,
        MJRearned,
        MJRmaxdist,
        MJRmiddist,
        MJRmindist,
        COPDpoints,
        COPDearned,
        COPDmaxdist,
        COPDmiddist,
        COPDmindist,
        PNEUpoints,
        PNEUearned,
        PNEUmaxdist,
        PNEUmiddist,
        PNEUmindist,
        EPI_available_TOTAL1,
        epi_points_earned_total1,
        comment_desc_TOTAL1,
        EPI_available_TOTAL2,
        epi_points_earned_total2,
        comment_desc_TOTAL2,
        EPI_available_TOTAL3,
        epi_points_earned_total3,
        comment_desc_TOTAL3,
        EPI_available_TOTAL4,
        epi_points_earned_total4,
        comment_desc_TOTAL4,
        hospsize):

    print("********************************PRINT START MATRIX****************************************************")
    print("********************************PRINT START MATRIX****************************************************")
    print("********************************PRINT START MATRIX****************************************************")
    print("hosp21_denominator")
    print(hosp21_denominator)
    # Lists of possible scores for Episodes, and distance from those scores

    if EPI_available_TOTAL1 > 0:
        theepilist = [[4,
                    EPI_available_TOTAL4,
                    epi_points_earned_total4,
                    comment_desc_TOTAL4],
                   [3,
                    EPI_available_TOTAL3,
                    epi_points_earned_total3,
                    comment_desc_TOTAL3],
                   [2,
                    EPI_available_TOTAL2,
                    epi_points_earned_total2,
                    comment_desc_TOTAL2],
                   [1,
                    EPI_available_TOTAL1,
                    epi_points_earned_total1,
                    comment_desc_TOTAL1]]
    else:
        theepilist = [[0, 0, 0, ""]]
    print("********************************PRINT theepilist****************************************************")
    print("********************************PRINT theepilist****************************************************")
    print("********************************PRINT theepilist****************************************************")
    print(theepilist)
    
    print("********************************PRINT if qb_points > 0:****************************************************")
    print("********************************PRINT if qb_points > 0:****************************************************")
    print("********************************PRINT if qb_points > 0:****************************************************")
    if qb_points > 0:
        qblist = [[6,
                   qbbenchmarks['bonus5']['points'],
                   qbbenchmarks['bonus5']['points'] - qb_earned,
                   qbbenchmarks['bonus5']['score'] - qb_star_rating],
                  [5,
                   qbbenchmarks['bonus2']['points'],
                   qbbenchmarks['bonus2']['points'] - qb_earned,
                   qbbenchmarks['bonus2']['score'] - qb_star_rating],
                  [4,
                   qbbenchmarks['max']['points'],
                   qbbenchmarks['max']['points'] - qb_earned,
                   qbbenchmarks['max']['score'] - qb_star_rating],
                  [3,
                   qbbenchmarks['mid']['points'],
                   qbbenchmarks['mid']['points'] - qb_earned,
                   qbbenchmarks['mid']['score'] - qb_star_rating],
                  [2,
                   qbbenchmarks['min']['points'],
                   qbbenchmarks['min']['points'] - qb_earned,
                   qbbenchmarks['min']['score'] - qb_star_rating],
                  [1,
                   0,
                   0 - qb_earned,
                   3.49 - qb_star_rating]]
    else:
        qblist = [[0, 0, 0, 0]]

    print("********************************PRINT if hosp03_points > 0:****************************************************")
    print("********************************PRINT if hosp03_points > 0:****************************************************")
    print("********************************PRINT if hosp03_points > 0:****************************************************")
    print("hosp03_points")
    print(hosp03_points)
    print("hosp04_points")
    print(hosp04_points)
    print("hosp19_points")
    print(hosp19_points)
    print("hosp20_points")
    print(hosp20_points)
    print("hosp21_points")
    print(hosp21_points)
    print("rrama_points")
    print(rrama_points)
    print("rracomm_points")
    print(rracomm_points)
    if hosp03_points > 0:
        print("IN hosp03_points > 0")
        hosp03list = [[4,
                       3,
                       cqmbenchmarks['hosp03']['points'] - hosp03_earned,
                       cqmbenchmarks['hosp03']['max'] - hosp03_rate],
                      [3,
                       cqmbenchmarks['hosp03']['pointsmid'],
                       cqmbenchmarks['hosp03']['pointsmid'] - hosp03_earned,
                       cqmbenchmarks['hosp03']['mid'] - hosp03_rate],
                      [2,
                       cqmbenchmarks['hosp03']['pointsmin'],
                       cqmbenchmarks['hosp03']['pointsmin'] - hosp03_earned,
                       cqmbenchmarks['hosp03']['min'] - hosp03_rate],
                      [1,
                       0,
                       0 - hosp03_earned,
                       cqmbenchmarks['hosp03']['min'] - .0001 - hosp03_rate]]
    else:
        hosp03list = [[0, 0, 0, 0]]

    print("********************************PRINT if hosp04_points > 0:****************************************************")
    print("********************************PRINT if hosp04_points > 0:****************************************************")
    print("********************************PRINT if hosp04_points > 0:****************************************************")
    print(hosp04_points)
    if hosp04_points > 0:
        print("IN hosp04_points > 0")
        hosp04list = [[4,
                       3,
                       cqmbenchmarks['hosp04']['points'] - hosp04_earned,
                       cqmbenchmarks['hosp04']['max'] - hosp04_rate],
                      [3,
                       cqmbenchmarks['hosp04']['pointsmid'],
                       cqmbenchmarks['hosp04']['pointsmid'] - hosp04_earned,
                       cqmbenchmarks['hosp04']['mid'] - hosp04_rate],
                      [2,
                       cqmbenchmarks['hosp04']['pointsmin'],
                       cqmbenchmarks['hosp04']['pointsmin'] - hosp04_earned,
                       cqmbenchmarks['hosp04']['min'] - hosp04_rate],
                      [1,
                       0,
                       0 - hosp04_earned,
                       cqmbenchmarks['hosp04']['min'] - .0001 - hosp04_rate]]
    else:
        hosp04list = [[0, 0, 0, 0]]

    print("********************************PRINT if hosp19_points > 0:****************************************************")
    print("********************************PRINT if hosp19_points > 0:****************************************************")
    print("********************************PRINT if hosp19_points > 0:****************************************************")
    if hosp19_points > 0:
        print("IN hosp19_points > 0")
        hosp19list = [[4,
                       3,
                       cqmbenchmarks['hosp19']['points'] - hosp19_earned,
                       cqmbenchmarks['hosp19']['max'] - hosp19_rate],
                      [3,
                       cqmbenchmarks['hosp19']['pointsmid'],
                       cqmbenchmarks['hosp19']['pointsmid'] - hosp19_earned,
                       cqmbenchmarks['hosp19']['mid'] - hosp19_rate],
                      [2,
                       cqmbenchmarks['hosp19']['pointsmin'],
                       cqmbenchmarks['hosp19']['pointsmin'] - hosp19_earned,
                       cqmbenchmarks['hosp19']['min'] - hosp19_rate],
                      [1,
                       0,
                       0 - hosp19_earned,
                       cqmbenchmarks['hosp19']['min'] + .0001 - hosp19_rate]]
    else:
        hosp19list = [[0, 0, 0, 0]]

    print("********************************PRINT if hosp20_points > 0:****************************************************")
    print("********************************PRINT if hosp20_points > 0:****************************************************")
    print("********************************PRINT if hosp20_points > 0:****************************************************")
    if hosp20_points > 0:
        print("IN hosp20_points > 0")
        hosp20list = [[4,
                       3,
                       cqmbenchmarks['hosp20']['points'] - hosp20_earned,
                       cqmbenchmarks['hosp20']['max'] - hosp20_rate],
                      [3,
                       cqmbenchmarks['hosp20']['pointsmid'],
                       cqmbenchmarks['hosp20']['pointsmid'] - hosp20_earned,
                       cqmbenchmarks['hosp20']['mid'] - hosp20_rate],
                      [2,
                       cqmbenchmarks['hosp20']['pointsmin'],
                       cqmbenchmarks['hosp20']['pointsmin'] - hosp20_earned,
                       cqmbenchmarks['hosp20']['min'] - hosp20_rate],
                      [1,
                       0,
                       0 - hosp20_earned,
                       cqmbenchmarks['hosp20']['min'] + .0001 - hosp20_rate]]
    else:
        hosp20list = [[0, 0, 0, 0]]

    print("********************************PRINT if hosp21_points > 0:****************************************************")
    print("********************************PRINT if hosp21_points > 0:****************************************************")
    print("********************************PRINT if hosp21_points > 0:****************************************************")
    print(hosp21_points)
    if hosp21_points > 0:
        print("IN hosp21_points > 0")
        hosp21list = [[4,
                       4,
                       cqmbenchmarks['hosp21']['points'] - hosp21_earned,
                       cqmbenchmarks['hosp21']['max'] - hosp21_rate],
                      [3,
                       cqmbenchmarks['hosp21']['pointsmid'],
                       cqmbenchmarks['hosp21']['pointsmid'] - hosp21_earned,
                       cqmbenchmarks['hosp21']['mid'] - hosp21_rate],
                      [2,
                       cqmbenchmarks['hosp21']['pointsmin'],
                       cqmbenchmarks['hosp21']['pointsmin'] - hosp21_earned,
                       cqmbenchmarks['hosp21']['min'] - hosp21_rate],
                      [1,
                       0,
                       0 - hosp21_earned,
                       cqmbenchmarks['hosp21']['min'] - .0001 - hosp21_rate]]
    else:
        hosp21list = [[0, 0, 0, 0]]

    print("********************************PRINT if rrama_points > 0:****************************************************")
    print("********************************PRINT if rrama_points > 0:****************************************************")
    print("********************************PRINT if rrama_points > 0:****************************************************")
    if rrama_points > 0:
        print("IN rrama_points > 0")
        rramalist = [[4,
                      7,
                      cqmbenchmarks['rrama']['points'] - rrama_earned,
                      cqmbenchmarks['rrama']['max'] - rrama_rate],
                     [3,
                      cqmbenchmarks['rrama']['pointsmid'],
                      cqmbenchmarks['rrama']['pointsmid'] - rrama_earned,
                      cqmbenchmarks['rrama']['mid'] - rrama_rate],
                     [2,
                      cqmbenchmarks['rrama']['pointsmin'],
                      cqmbenchmarks['rrama']['pointsmin'] - rrama_earned,
                      cqmbenchmarks['rrama']['min'] - rrama_rate],
                     [1,
                      0,
                      0 - rrama_earned,
                      cqmbenchmarks['rrama']['min'] + .0001 - rrama_rate]]
    else:
        rramalist = [[0, 0, 0, 0]]

    print("********************************PRINT if rracomm_points > 0:****************************************************")
    print("********************************PRINT if rracomm_points > 0:****************************************************")
    print("********************************PRINT if rracomm_points > 0:****************************************************")
    if rracomm_points > 0:
        print("IN rracomm_points > 0")
        rracommlist = [[4,
                        7,
                        cqmbenchmarks['rracomm']['points'] - rracomm_earned,
                        cqmbenchmarks['rracomm']['max'] - rracomm_rate],
                       [3,
                        cqmbenchmarks['rracomm']['pointsmid'],
                        cqmbenchmarks['rracomm']['pointsmid'] - rracomm_earned,
                        cqmbenchmarks['rracomm']['mid'] - rracomm_rate],
                       [2,
                        cqmbenchmarks['rracomm']['pointsmin'],
                        cqmbenchmarks['rracomm']['pointsmin'] - rracomm_earned,
                        cqmbenchmarks['rracomm']['min'] - rracomm_rate],
                       [1,
                        0,
                        0 - rracomm_earned,
                        cqmbenchmarks['rracomm']['min'] + .0001 - rracomm_rate]]
    else:
        rracommlist = [[0, 0, 0, 0]]

    print("********************************PRINT if cqmheader = [****************************************************")
    print("********************************PRINT if cqmheader = [****************************************************")
    print("********************************PRINT if cqmheader = [****************************************************")
    cqmheader = [
        "Points Earned",
        "Total Numerator Change",
        "Hosp03 Change",
        "Hosp19 Change",
        "Hosp20 Change",
        "Readm MA Change",
        "Readm Comm Change",
        "Hosp04 Change",
        "Hosp21 Change"]
    cqmlist = []
    print("hosp03list[0]")
    print(hosp03list[0])
    print("hosp04list[0]")
    print(hosp04list[0])
    print("hosp19list[0]")
    print(hosp19list[0])
    print("hosp20list[0]")
    print(hosp20list[0])
    print("hosp21list[0]")
    print(hosp21list[0])
    print("rramalist[0]")
    print(rramalist[0])
    print("rracommlist[0]")
    print(rracommlist[0])
    print("********************************PRINT BEFORE THE LOOPER FOR CQMLIST****************************************************")
    print("********************************PRINT BEFORE THE LOOPER FOR CQMLIST****************************************************")
    print("********************************PRINT BEFORE THE LOOPER FOR CQMLIST****************************************************")
    for i in hosp03list:
        for j in hosp19list:
            for k in hosp20list:
                for l in rramalist:
                    for m in rracommlist:
                        for n in hosp04list:
                            for o in hosp21list:
                                if i[2] != 0:
                                    if i[3] < 0:
                                        hosp03_den = math.floor(i[3] * hosp03_denominator)
                                    if i[3] > 0:
                                        hosp03_den = math.ceil(i[3] * hosp03_denominator)
                                else:
                                    hosp03_den = 0
                                if j[2] != 0:
                                    if j[3] < 0:
                                        hosp19_den = math.floor(j[3] * hosp19_denominator)
                                    if j[3] > 0:
                                        hosp19_den = math.ceil(j[3] * hosp19_denominator)
                                else:
                                    hosp19_den = 0
                                if k[2] != 0:
                                    if k[3] < 0:
                                        hosp20_den = math.floor(k[3] * hosp20_denominator)
                                    if k[3] > 0:
                                        hosp20_den = math.ceil(k[3] * hosp20_denominator)
                                else:
                                    hosp20_den = 0
                                if l[2] != 0:
                                    if l[3] < 0:
                                        rrama_den = math.floor((rrama_expected * l[3] / rrama_mkt) * rrama_denominator)
                                    if l[3] > 0:
                                        rrama_den = math.ceil((rrama_expected * l[3] / rrama_mkt) * rrama_denominator)
                                else:
                                    rrama_den = 0
                                if m[2] != 0:
                                    if m[3] < 0:
                                        rracomm_den = math.floor((rracomm_expected * m[3] / rracomm_mkt) * rracomm_denominator)
                                if m[3] > 0:
                                    rracomm_den = math.ceil((rracomm_expected * m[3] / rracomm_mkt) * rracomm_denominator)
                                else:
                                    rracomm_den = 0
                                if n[2] != 0:
                                    if n[3] < 0:
                                        hosp04_den = math.floor(n[3] * hosp04_denominator)
                                    if n[3] > 0:
                                        hosp04_den = math.ceil(n[3] * hosp04_denominator)
                                else:
                                    hosp04_den = 0
                                if o[2] != 0:
                                    print("+++++O[2]+++++")
                                    if o[3] < 0:
                                        hosp21_den = math.floor(o[3] * hosp21_denominator)
                                    if o[3] > 0:
                                        hosp21_den = math.ceil(o[3] * hosp21_denominator)
                                else:
                                    print("-----O[2]-----")
                                    hosp21_den = 0
                                print("IN THE LOOPER FOR CQMLIST")
                                print("IN THE LOOPER FOR CQMLIST")
                                print("IN THE LOOPER FOR CQMLIST")
                                print("i[1]")
                                print(i[1])
                                print("j[1]")
                                print(j[1])
                                print("k[1]")
                                print(k[1])
                                print("l[1]")
                                print(l[1])
                                print("m[1]")
                                print(m[1])
                                print("n[1]")
                                print(n[1])
                                print("o[1]")
                                print(o[1])
                                print("hosp03_den")
                                print(hosp03_den)
                                print("hosp04_den")
                                print(hosp04_den)
                                print("hosp19_den")
                                print(hosp19_den)
                                print("hosp20_den")
                                print(hosp20_den)
                                print("hosp21_den")
                                print(hosp21_den)
                                print("rrama_den")
                                print(rrama_den)
                                print("rracomm_den")
                                print(rracomm_den)
                                cqmlist.append(
                                    [i[1] + j[1] + k[1] + l[1] + m[1] + n[1] + o[1],
                                    abs(hosp03_den) +
                                    abs(hosp19_den) +
                                    abs(hosp20_den) +
                                    abs(rrama_den) +
                                    abs(rracomm_den) +
                                    abs(hosp04_den) +
                                    abs(hosp21_den),
                                    hosp03_den,
                                    hosp19_den,
                                    hosp20_den,
                                    rrama_den,
                                    rracomm_den,
                                    hosp04_den,
                                    hosp21_den])
                                print("END THE LOOPER FOR CQMLIST")
                                print("END THE LOOPER FOR CQMLIST")
                                print("END THE LOOPER FOR CQMLIST")
        # Sort to determine the minimum change over all episodes
    print("OUT OF THE LOOPER FOR CQMLIST")
    print("OUT OF THE LOOPER FOR CQMLIST")
    print("OUT OF THE LOOPER FOR CQMLIST")
    cqmlist.sort(key=lambda x: abs(x[1]))
    # Now sort by the score at each minimum change
    cqmlist.sort(key=lambda x: x[0])
    print("********************************PRINT CQMLIST****************************************************")
    print("********************************PRINT CQMLIST****************************************************")
    print("********************************PRINT CQMLIST****************************************************")
    #print(cqmlist)
    print(cqmlist[0])
    #worksheet2.write_row(1, 0, cqmheader)
    # for i in range(len(cqmlist)):
    #worksheet2.write_row(1+i+1, 0, cqmlist[i])

    qblist.sort(key=lambda x: x[1])
    print("********************************PRINT qblist****************************************************")
    print("********************************PRINT qblist****************************************************")
    print("********************************PRINT qblist****************************************************")
    #print(qblist)

    if MJRpoints > 0:
        MJRlist = [[4,
                    MJRpoints,
                    MJRpoints - MJRearned,
                    MJRmaxdist],
                   [3,
                    MJRpoints * .75,
                    MJRpoints * .75 - MJRearned,
                    MJRmiddist],
                   [2,
                    MJRpoints * .5,
                    MJRpoints * .5 - MJRearned,
                    MJRmindist],
                   [1,
                    0,
                    0 - MJRearned,
                    MJRmindist - 1]]
    else:
        MJRlist = [[0, 0, 0, 0]]
    if COPDpoints > 0:
        COPDlist = [[4,
                     COPDpoints,
                     COPDpoints - COPDearned,
                     COPDmaxdist],
                    [3,
                     COPDpoints * .75,
                     COPDpoints * .75 - COPDearned,
                     COPDmiddist],
                    [2,
                     COPDpoints * .5,
                     COPDpoints * .5 - COPDearned,
                     COPDmindist],
                    [1,
                     0,
                     0 - COPDearned,
                     COPDmindist - 1]]
    else:
        COPDlist = [[0, 0, 0, 0]]
    if PNEUpoints > 0:
        PNEUlist = [[4,
                     PNEUpoints,
                     PNEUpoints - PNEUearned,
                     PNEUmaxdist],
                    [3,
                     PNEUpoints * .75,
                     PNEUpoints * .75 - PNEUearned,
                     PNEUmiddist],
                    [2,
                     PNEUpoints * .5,
                     PNEUpoints * .5 - PNEUearned,
                     PNEUmindist],
                    [1,
                     0,
                     0 - PNEUearned,
                     PNEUmindist - 1]]
    else:
        PNEUlist = [[0, 0, 0, 0]]

    # Calculate all possible episode scores by iterating through each of the
    # lists above
    epi_scores = []
    for i in MJRlist:
        for j in COPDlist:
            for k in PNEUlist:
                epi_scores.append(
                    [i[1] + j[1] + k[1], i[3] + j[3] + k[3], i[3], j[3], k[3]])
    print("********************************PRINT epi_scores****************************************************")
    print("********************************PRINT epi_scores****************************************************")
    print("********************************PRINT epi_scores****************************************************")
    print(epi_scores)
    epi_scores2 = []
    for i in theepilist:
        epi_scores2.append([i[2], i[2], i[2], i[2], i[2], i[3]])
    print("********************************PRINT epi_scores2****************************************************")
    print("********************************PRINT epi_scores2****************************************************")
    print("********************************PRINT epi_scores2****************************************************")
    print(epi_scores2)

    # Sort to determine the minimum change over all episodes
    epi_scores.sort(key=lambda x: abs(x[1]))
    # Now sort by the score at each minimum change
    epi_scores.sort(key=lambda x: x[0])

    # Sort to determine the minimum change over all episodes
    epi_scores2.sort(key=lambda x: abs(x[1]))
    # Now sort by the score at each minimum change
    epi_scores2.sort(key=lambda x: x[0])
    # This next block of code down to cqmlist_detail = newlist
    # will take the first value of cqmlist_detail and then compare it
    # with the next value. Since the cqmlist_detail is sorted by points and total_den,
    # if the next of cqmlist_detail is higher than the first, it will add it to newlist
    # and use that newlist value as the comparisons value. This is essentially
    # a "NODUPKEY" proc sort in SAS
    newlist = []
    newlist.append(epi_scores[0])
    newlisti = 0
    for i in range(len(epi_scores) - 1):

        if epi_scores[i + 1][0] > newlist[newlisti][0]:
            newlisti = newlisti + 1
            newlist.append(epi_scores[i + 1])

    epi_scores = newlist
    print("********************************PRINT epi_scores = newlist****************************************************")
    print("********************************PRINT epi_scores = newlist****************************************************")
    print("********************************PRINT epi_scores = newlist****************************************************")
    print(epi_scores)

    newlist = []
    newlist.append(epi_scores2[0])
    newlisti = 0
    for i in range(len(epi_scores2) - 1):

        if epi_scores2[i + 1][0] > newlist[newlisti][0]:
            newlisti = newlisti + 1
            newlist.append(epi_scores2[i + 1])

    epi_scores2 = newlist
    print("********************************PRINT epi_scores2 = newlist****************************************************")
    print("********************************PRINT epi_scores2 = newlist****************************************************")
    print("********************************PRINT epi_scores2 = newlist****************************************************")
    print(epi_scores2)
    
    newlist = []
    newlist.append(cqmlist[0])
    newlisti = 0
    for i in range(len(cqmlist) - 1):

        if cqmlist[i + 1][0] > newlist[newlisti][0]:
            newlisti = newlisti + 1
            newlist.append(cqmlist[i + 1])
            print(cqmlist[i + 1][0])

    cqmlist = newlist
    #print(epi_scores)
    #print(cqmlist)

    worksheet.write(startrow, 0, "Legend:", header)
    worksheet.write(startrow + 1, 0, " ", table_body_pct_red)
    worksheet.write(startrow + 1, 1, "Zero Payout")
    worksheet.write(startrow + 2, 0, " ", table_body_pct_orange)
    worksheet.write(startrow + 2, 1, "Min Payout")
    worksheet.write(startrow + 1, 4, " ", table_body_pct_yellow)
    worksheet.write(startrow + 1, 5, "Mid Payout")
    worksheet.write(startrow + 2, 4, " ", table_body_pct_green)
    worksheet.write(startrow + 2, 5, "Max Payout")
    worksheet.write(startrow + 1, 8, " ", current_score)
    worksheet.write(startrow + 1, 9, "Current Score")

    # This is the row and column where we will start our matrix.
    startrow = startrow + 3
    startcol = 0
    worksheet.write(startrow, startcol, '', table_title)
    worksheet.write(startrow + 1, startcol, '', table_title)
    worksheet.write(startrow + 2, startcol, '', table_title)
    worksheet.write(startrow + 3, startcol, "Episodes", table_title)
    curcol = startcol + 1
    worksheet.write(startrow, curcol, '', table_title)
    worksheet.write(startrow + 1, curcol, '', table_title)
    worksheet.write(startrow + 2, curcol, 'Quality', table_title)
    worksheet.write(startrow + 3, curcol, "Bundle", table_title)
    curcol = curcol + 1
    currow = startrow
    worksheet.merge_range(
        currow,
        curcol,
        currow,
        curcol +
        len(cqmlist) -
        1,
        "SCORING MATRIX",
        table_title)
    currow = currow + 1
    worksheet.merge_range(
        currow,
        curcol,
        currow,
        curcol +
        len(cqmlist) -
        1,
        "",
        table_title)
    currow = currow + 1
    worksheet.merge_range(
        currow,
        curcol,
        currow,
        curcol + len(cqmlist) - 1,
        "Clinical Quality Metrics",
        table_title)
    currow = currow + 1

    for k in range(len(cqmlist)):
        worksheet.write(currow, curcol, cqmlist[k][0], header)
        cqmcomment = "Points:" + str(cqmlist[k][0]) + "\n Hosp03 Target: " + str(cqmlist[k][2] + hosp03_numerator) + "/" + str(
            hosp03_denominator) + ' (' + comment(cqmlist[k][2], "fewer Palliative Care Consults (MA)", "more Palliative Care Consults (MA)") + ') '
        cqmcomment = cqmcomment + "\n Hosp19 Target: " + str(
            cqmlist[k][3] + hosp19_numerator) + "/" + str(hosp19_denominator) + ' (' + comment(
            cqmlist[k][3], "fewer 3 Day ED Returns (MA)", "more 3 Day ED Returns (MA)") + ') '
        cqmcomment = cqmcomment + "\n Hosp20 Target: " + str(
            cqmlist[k][4] + hosp20_numerator) + "/" + str(hosp20_denominator) + ' (' + comment(
            cqmlist[k][4], "fewer 3 Day ED Returns (Comm)", "more 3 Day ED Returns (Comm)") + ') '
        cqmcomment = cqmcomment + "\n Risk-Readjusted Readmissions (MA) Target: " + str(
            cqmlist[k][5] + rrama_numerator) + "/" + str(rrama_denominator) + ' (' + comment(
            cqmlist[k][5], "fewer MA Readmissions", "more MA Readmissions") + ') '
        cqmcomment = cqmcomment + "\n Risk-Readjusted Readmissions (Comm) Target: " + str(
            cqmlist[k][6] + rracomm_numerator) + "/" + str(rracomm_denominator) + ' (' + comment(
            cqmlist[k][6], "fewer Commercial Readmissions", "more Commercial Readmissions") + ') '
        cqmcomment = cqmcomment + "\n Hosp04 (Comm) Target: " + str(
            cqmlist[k][7] + hosp04_numerator) + "/" + str(hosp04_denominator) + ' (' + comment(
            cqmlist[k][7], "fewer Palliative Care Consults (Comm)", "more Palliative Care Consults (Comm)") + ') '
        cqmcomment = cqmcomment + "\n Hosp21 Target: " + str(
            cqmlist[k][8] + hosp21_numerator) + "/" + str(hosp21_denominator) + ' (' + comment(
            cqmlist[k][8], "fewer 7 Day Follow-Up", "more 7 Day Follow-Up") + ') '
        # print(cqmcomment)
        #print("rrama_numerator ", rrama_numerator, "/nrrama_den", rrama_denominator, "/nexpected", rrama_e, "/nmarketexpected", rrama_me)
        worksheet.write_comment(
            currow, curcol, cqmcomment, {
                'y_scale': 4.2, 'x_scale': 3})
        if cqmlist[k][1] == cqm_earned:
            worksheet.write_comment(
                currow, curcol, "Current Score", {
                    'y_scale': 1, 'x_scale': 1})

        curcol = curcol + 1
    currow = currow + 1
    curcol = startcol

    for i in range(len(epi_scores2)):
        curcol = startcol
        for j in range(len(qblist)):
            worksheet.write(currow, curcol, epi_scores2[i][0], header)

            # print(epi_scores)

            epicomment = epi_scores2[i][5]
            #epicomment = '$' + comment(
            #    round(
            #        epi_scores[i][2],
            #        2),
            #    "decrease in Major Joint Replacement Costs",
            #    "increase in Major Joint Replacement Costs") + "\n" + '$' + comment(
            #    round(
            #        epi_scores[i][3],
            #        2),
            #    "decrease in COPD costs",
            #    "increase in COPD Costs") + "\n" + '$' + comment(
            #        round(
            #            epi_scores[i][4],
            #            2),
            #    "decrease in Pneumonia costs",
            #    "increase in Pneumonia Costs")

            if epi_scores2[i][0] == epi_earned:
                epicomment = "Current Score"
            if epi_earned == 0:
                epicommentVPF = "Not Scored"
                #worksheet.write(currow, curcol, 'NS', header) ---> VPF
            ##print("row: ", currow)
            ##print("col: ", curcol)
                ##print("epicomment: ", epicomment)
            worksheet.write_comment(
                currow, curcol, epicomment, {
                    'y_scale': 2.2, 'x_scale': 5})
            curcol = curcol + 1

            if qblist[j][3] >= 0:
                qbcomment = "Increase Quality Bundle Star Rating by " + \
                    str(qblist[j][3])
            else:
                qbcomment = "Decrease Quality Bundle Star Rating by " + \
                    str(qblist[j][3])

            worksheet.write(currow, curcol, qblist[j][1], header)
            if qblist[j][2] == 0:
                qbcomment = "Current Score"
            if qb_points == 0:
                qbcomment = "Not Scored"
                worksheet.write(currow, curcol, "NS", header)

            worksheet.write_comment(currow, curcol, qbcomment)

            for k in range(len(cqmlist)):
                curcol = curcol + 1
                # try:
                matrix_score = (epi_scores2[i][0] + qblist[j][1] + cqmlist[k][0]) / (EPI_available_TOTAL4 + \
                                hosp03_points + hosp19_points + hosp20_points + rrama_points + rracomm_points + hosp04_points + \
                                hosp21_points + qb_points)
                matrix_score = round(matrix_score, 2)
                #print(epi_scores[i][0] + qblist[j][1] + cqmlist[k][0])
                #print(MJRpoints + COPDpoints + PNEUpoints + hosp03_points + hosp04_points + hosp19_points + hosp20_points + hosp21_points + 
                #    rrama_points +
                #    rracomm_points +
                #    qb_points)
                # except:
                #matrix_score = 0

                if epi_scores2[i][0] == epi_earned and qblist[j][1] == qb_earned and cqmlist[k][0] == cqm_earned:
                    worksheet.write(
                        currow, curcol, matrix_score, current_score)
                else:
                    if matrix_score < overallbenchmarks[hospsize]['min']:
                        worksheet.write(
                            currow, curcol, matrix_score, table_body_pct_red)
                    elif matrix_score >= overallbenchmarks[hospsize]['min'] and matrix_score < overallbenchmarks[hospsize]['mid']:
                        worksheet.write(
                            currow, curcol, matrix_score, table_body_pct_orange)
                    elif matrix_score >= overallbenchmarks[hospsize]['mid'] and matrix_score < overallbenchmarks[hospsize]['max']:
                        worksheet.write(
                            currow, curcol, matrix_score, table_body_pct_yellow)
                    elif matrix_score >= overallbenchmarks[hospsize]['max']:
                        worksheet.write(
                            currow, curcol, matrix_score, table_body_pct_green)
            curcol = startcol
            currow = currow + 1

    worksheet.autofilter(startrow + 3, startcol, currow, startcol + 1)


def load_stars_history(dates_df, history_df):

    # print(year)

    history_df = history_df[['year', 'PyDate', 'date', 'month', 'data']].copy(
        deep=True).dropna().sort_values(by=['PyDate'])

    print(history_df)
    dflist = []
    for i in range(0, len(history_df)):
        data = history_df['data'].values[i]
        date = history_df["PyDate"].values[i]
        date2 = history_df["date"].values[i]
        year = history_df['year'].values[i]
        month = history_df['month'].values[i]
        dflist.append(loadsas("/n04/data/stars/PROSPECTIVE/YR_" +
                              str(int(year)) +
                              "/incentive_reports/programs/data/" +
                              data +
                              "/tab1_qb_level.sas7bdat"))
        # SAS date to Excel conversion
        dflist[i]['Date'] = date2
        dflist[i]['Year'] = year
        dflist[i]['Month'] = month
        if year == 2016:
            dflist[i]['Health_System_ID'] = dflist[i]['qb_system_id']
    df = pd.concat(dflist).sort_values(
        by=['Health_System_ID', 'Year', 'Month'])
    df = df[["Health_System_ID", "Date", 'Year',
             'Month', "aggr_star_rating"]].drop_duplicates()

    return df


def stars_history_graph(
        wb,
        ws,
        df,
        proj_star_rating=None,
        startrow=0,
        startcol=0,
        graph_title=None,
        x_scale=2,
        y_scale=0.75):
    try:
        df["aggr_star_rating"] = df["aggr_star_rating"].apply(
            lambda x: math.floor(x * 100) / 100)
    except BaseException:
        print("empty")

    print(df)
    print(df['Year'].max())
    year = int(df['Year'].max())
    min_mo = int(df['Month'].where(df["Year"] == year).dropna().min())
    try:
        min_score = math.floor(
            df['aggr_star_rating'].where(
                df["Year"] == year).dropna().min() * 10) / 10
    except BaseException:
        min_score = 1
    max_mo = int(df['Month'].where(df["Year"] == year).dropna().max())

    curcol = itercol(ws, startrow, startcol, 0, "Month", header)
    ws.write(startrow + 1, curcol - 1, "CY " + str(int(year)), header)
    ws.write(startrow + 2, curcol - 1, "CY " + str(int(year - 1)), header)
    ws.write(startrow + 3, curcol - 1, "Projected", header)
    ws.write(startrow +
             4, curcol -
             1, str(qbbenchmarks["max"]["points"]) +
             " Points", header)
    ws.write(startrow +
             5, curcol -
             1, str(qbbenchmarks["mid"]["points"]) +
             " Points", header)
    ws.write(startrow +
             6, curcol -
             1, str(qbbenchmarks["min"]["points"]) +
             " Points", header)

    # i <= max_mo or
    for i in range(int(min_mo), 16):
        if i <= 13 or i == 15:
            print(i)
            try:
                cy = df['aggr_star_rating'].where(
                    df['Year'] == year).where(
                    df['Month'] == i).dropna().fillna("").values[0]
            except BaseException:
                cy = ""
            try:
                py = df['aggr_star_rating'].where(
                    df['Year'] == year -
                    1).where(
                    df['Month'] == i).dropna().fillna("").values[0]
            except BaseException:
                py = ""
            curcol = itercol(ws,
                             startrow,
                             curcol,
                             0,
                             "=DATE(" + str(year) + ", " + str(i + 1) + ', 1)-1',
                             table_body_date2)
            if cy != "":
                ws.write(startrow + 1, curcol - 1, cy, table_body_num2)
            if py != "":
                ws.write(startrow + 2, curcol - 1, py, table_body_num2)

            if max_mo != 15 and proj_star_rating is not None and i >= max_mo:
                if i == max_mo:
                    months = 15 - max_mo
                    sr_diff = proj_star_rating - cy
                    m = sr_diff / months
                    b = proj_star_rating - m * 15
                y = m * i + b
                ws.write(startrow + 3, curcol - 1, y, table_body_num2)
            if max_mo != 15 and proj_star_rating is not None and i == 15:
                ws.write(
                    startrow + 3,
                    curcol - 1,
                    proj_star_rating,
                    table_body_num2)
            ws.write(
                startrow + 4,
                curcol - 1,
                qbbenchmarks["max"]["score"],
                table_body_num2)
            ws.write(
                startrow + 5,
                curcol - 1,
                qbbenchmarks["mid"]["score"],
                table_body_num2)
            ws.write(
                startrow + 6,
                curcol - 1,
                qbbenchmarks["min"]["score"],
                table_body_num2)

    series1 = '=' + "'" + ws.get_name() + "'" + '!' + colnum_string(startcol + 1) + \
        str(startrow + 1) + ':' + colnum_string(curcol - 1) + str(startrow + 1)
    series2 = '=' + "'" + ws.get_name() + "'" + '!' + colnum_string(startcol + 1) + \
        str(startrow + 2) + ':' + colnum_string(curcol - 1) + str(startrow + 2)
    series3 = '=' + "'" + ws.get_name() + "'" + '!' + colnum_string(startcol + 1) + \
        str(startrow + 3) + ':' + colnum_string(curcol - 1) + str(startrow + 3)
    series4 = '=' + "'" + ws.get_name() + "'" + '!' + colnum_string(startcol + 1) + \
        str(startrow + 4) + ':' + colnum_string(curcol - 1) + str(startrow + 4)
    print(series2, series3, series4)

    chart = wb.add_chart(
        {'type': 'line'}
    )

    chart.add_series({
        'categories': series1,
        'values': series2,
        'name': "CY " + str(int(year)),
        'line': {
            'color': blue2,
            'width': 2
        },
        'data_labels': {
            'value': True,
            'font': {'name': 'Calibri'}
        },
        'marker': {'type': 'square', 'fill': {'color': blue2}}
    })

    chart.add_series({
        'categories': [ws.get_name(), startrow, startcol + 1, startrow, curcol - 1],
        'values': [ws.get_name(), startrow + 2, startcol + 1, startrow + 2, curcol - 1],
        'name': [ws.get_name(), startrow + 2, startcol],
        'line': {
            'color': purple1,
            'width': 2
        },
        'marker': {'type': 'square', 'fill': {'color': purple1}}
    })

    chart.add_series({
        'categories': [ws.get_name(), startrow, startcol + 1, startrow, curcol - 1],
        'values': [ws.get_name(), startrow + 3, startcol + 1, startrow + 3, curcol - 1],
        'name': [ws.get_name(), startrow + 3, startcol],
        'line': {
            'color': orange2,
            'width': 2,
            'dash_type': 'dash'
        },
        'data_labels': {
            'value': False,
            'font': {'name': 'Calibri'}
        },
        'marker': {'type': 'square', 'fill': {'color': orange2}}
    })

    chart.add_series({
        'categories': [ws.get_name(), startrow, startcol + 1, startrow, curcol - 1],
        'values': [ws.get_name(), startrow + 4, startcol + 1, startrow + 4, curcol - 1],
        'name': [ws.get_name(), startrow + 4, startcol],
        'line': {
            'color': green2,
            'width': 1.5
        }
    })

    chart.add_series({
        'categories': [ws.get_name(), startrow, startcol + 1, startrow, curcol - 1],
        'values': [ws.get_name(), startrow + 5, startcol + 1, startrow + 5, curcol - 1],
        'name': [ws.get_name(), startrow + 5, startcol],
        'line': {
            'color': yellow1,
            'width': 1.5
        }
    })

    chart.add_series({
        'categories': [ws.get_name(), startrow, startcol + 1, startrow, curcol - 1],
        'values': [ws.get_name(), startrow + 6, startcol + 1, startrow + 6, curcol - 1],
        'name': [ws.get_name(), startrow + 6, startcol],
        'line': {
            'color': red2,
            'width': 1.5
        }
    })

    chart.set_y_axis({
        'max': 5,
        'min': 1
    })

    chart.set_x_axis({
        'date_axis': True,
        'num_format': 'mm/dd/yyyy',
        'minor_unit': 1,
        'minor_unit_type': 'months',
        'max': date(year + 1, 3, 31),
        'min': date(year, 2, 28)
    })

    if graph_title is None:
        chart.set_title({'none': True})
    else:
        chart.set_title({'name': "Historical Performance"})

    ws.insert_chart(colnum_string(startcol) + str(startrow + 1),
                    chart, {'x_scale': x_scale, 'y_scale': y_scale})


def na(testvalue, fmt, na_fmt):
    if isinstance(testvalue, str) == False:
        return(fmt)
    else:
        if testvalue == "":
            return(na_fmt)
        else:
            return(fmt)


def bundle_formats(qb_earned=0, qb_points=0):
    if qb_earned == qbbenchmarks['bonus5']['points']:
        return table_body_green
    elif qb_earned == qbbenchmarks['bonus2']['points']:
        return table_body_green
    elif qb_earned == qbbenchmarks['max']['points']:
        return table_body_green
    elif qb_earned == qbbenchmarks['mid']['points']:
        return table_body_yellow
    elif qb_earned == qbbenchmarks['min']['points']:
        return table_body_orange
    elif qb_earned == 0 and qb_points != 0:
        return table_body_red
    else:
        return table_body


def stars_summary(
        ws,
        startrow=0,
        startcol=0,
        current_star_rating="",
        points_earned=0,
        points_available=0,
        predicted_star_rating="",
        py_star_rating="",
        prac_count="",
        py_prac_count="",
        attr_count="",
        py_attr_count="",
        min_prac_stars="",
        py_min_stars="",
        max_prac_stars="",
        py_max_stars=""):
    columns = {
        'Attribute': 2,
        'Current YTD': 1,
        'Predicted Year End': 1,
        'Prior Year End': 1
    }

    curcol = itercol(
        ws,
        startrow + 1,
        startcol,
        columns['Attribute'],
        "",
        headerwrap)
    curcol = itercol(
        ws,
        startrow + 1,
        curcol,
        columns['Current YTD'],
        "Current YTD",
        headerwrap_num)
    curcol = itercol(
        ws,
        startrow + 1,
        curcol,
        columns['Predicted Year End'],
        'Predicted Year End',
        headerwrap_num)
    curcol = itercol(
        ws,
        startrow + 1,
        curcol,
        columns['Prior Year End'],
        'Prior Year End',
        headerwrap_num)
    lastcol = curcol - 1
    ws.merge_range(
        startrow,
        startcol,
        startrow,
        lastcol,
        "QUALITY BUNDLE",
        table_title)

    currow = startrow + 2

    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Attribute'],
        "Aggregate Star Rating",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Current YTD'],
        current_star_rating,
        bundle_formats(
            points_earned,
            points_available))
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Predicted Year End'],
        predicted_star_rating,
        table_body_num2)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Prior Year End'],
        py_star_rating,
        table_body_num2)

    currow = currow + 1

    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Attribute'],
        "Number of Practices",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Current YTD'],
        prac_count,
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Predicted Year End'],
        "",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Prior Year End'],
        py_prac_count,
        table_body)

    currow = currow + 1

    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Attribute'],
        "Attributed MA Members",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Current YTD'],
        attr_count,
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Predicted Year End'],
        "",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Prior Year End'],
        py_attr_count,
        table_body)

    currow = currow + 1

    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Attribute'],
        "Maximum Practice Rating",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Current YTD'],
        max_prac_stars,
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Predicted Year End'],
        "",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Prior Year End'],
        py_prac_count,
        table_body)

    currow = currow + 1

    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Attribute'],
        "Minimum Practice Rating",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Current YTD'],
        min_prac_stars,
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Predicted Year End'],
        "",
        table_body)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Prior Year End'],
        py_prac_count,
        table_body)

    currow = currow + 1

    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Attribute'],
        "POINTS EARNED:",
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Current YTD'],
        points_earned,
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Predicted Year End'],
        "",
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Prior Year End'],
        "",
        table_title)

    qbtables(
        worksheet=ws,
        startrow=startrow,
        startcol=lastcol + 2,
        qb_star_rating=current_star_rating,
        qb_points=points_available,
        qb_earned=points_earned,
        score=False)


def star_rating_fmt(star_rating):
    if star_rating != "":
        if star_rating == 5:
            return table_body_green
        if star_rating == 4:
            return table_body_yellow
        if star_rating == 3:
            return table_body_orange
        if star_rating == 2:
            return table_body_red
        if star_rating == 1:
            return table_body_red
        if star_rating == "":
            return table_body


def stars_table(df, ws, startrow, startcol):

    df.sort_values(by=['row'])
    df_temp = df[['Measure Code',
                  'Class',
                  'Measure',
                  'Weight',
                  'Denominator',
                  'Numerator',
                  'Gaps Addressed',
                  'Beyond Remediation',
                  'YTD Compliance',
                  'Trend Compliance',
                  'Star Rating',
                  'Maximum Potential Compliance',
                  '2 Stars',
                  '3 Stars',
                  '4 Stars',
                  '5 Stars']]

    columns = {
        'Class': 1,
        'Measure': 6,
        'Weight': 0,
        'Denominator': 1,
        'Numerator': 1,
        'Gaps Addressed': 1,
        'Beyond Remediation': 1,
        'YTD Compliance': 1,
        'Trend Compliance': 1,
        'Star Rating': 1,
        'Maximum Potential Compliance': 1,
        '2 Stars': 0,
        '3 Stars': 0,
        '4 Stars': 0,
        'needed4': 1,
        '5 Stars': 0,
        'needed5': 1
    }
    print(df_temp)
    print(len(df_temp))
    currow = startrow + 1
    ws.set_row(currow, 30)
    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Class'],
        "Class",
        headerwrap)
    #ws.write(currow, startcol, "Class", header)
    #ws.merge_range(currow, startcol+1, currow, startcol+5, "Measure", header)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Measure'],
        "Measure",
        headerwrap)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Weight'],
        "Weight",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Denominator'],
        "Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Numerator'],
        "Numerator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Gaps Addressed'],
        "Gaps Addressed",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Beyond Remediation'],
        "Beyond Remediation",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['YTD Compliance'],
        "YTD Compliance",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Trend Compliance'],
        "Trend Compliance",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Star Rating'],
        "Star Rating",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Maximum Potential Compliance'],
        "Maximum Potential Rate",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['4 Stars'],
        "4 Stars",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['needed4'],
        "4 Stars (Gaps Needed)",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['5 Stars'],
        "5 Stars",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['needed5'],
        "5 Stars (Gaps Needed)",
        headerwrap_num)
    for i in range(0, len(df_temp)):
        currow = currow + 1
        curcol = 0
        if df["Measure Code"].values[i] in ['pcr', 'hpc']:
            compliance_fmt = table_body_num2
        else:
            compliance_fmt = table_body_pct
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Class'],
            df["Class"].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Measure'],
            df["Measure"].values[i],
            table_body)
        curcol = itercol(
            ws, currow, curcol, columns['Weight'], df["Weight"].values[i], na(
                df["Weight"].values[i], table_body, table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['Denominator'], df["Denominator"].values[i], na(
                df["Denominator"].values[i], table_body, table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['Numerator'], df["Numerator"].values[i], na(
                df["Numerator"].values[i], table_body, table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['Gaps Addressed'], df["Gaps Addressed"].values[i], na(
                df["Gaps Addressed"].values[i], table_body, table_body_na))
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Beyond Remediation'],
                         df["Beyond Remediation"].values[i],
                         na(df["Beyond Remediation"].values[i],
                             table_body,
                             table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['YTD Compliance'], df["YTD Compliance"].values[i], na(
                df["YTD Compliance"].values[i], table_body_pct2, table_body_na))
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Trend Compliance'],
                         df["Trend Compliance"].values[i],
                         na(df["Trend Compliance"].values[i],
                             compliance_fmt,
                             table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['Star Rating'], df["Star Rating"].values[i], na(
                df["Star Rating"].values[i], table_body, table_body_na))
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Maximum Potential Compliance'],
                         df["Maximum Potential Compliance"].values[i],
                         na(df["Maximum Potential Compliance"].values[i],
                             compliance_fmt,
                             table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['4 Stars'], df["4 Stars"].values[i], na(
                df["4 Stars"].values[i], compliance_fmt, table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['needed4'], df["needed4"].values[i], na(
                df["needed4"].values[i], table_body, table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['5 Stars'], df["5 Stars"].values[i], na(
                df["5 Stars"].values[i], compliance_fmt, table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['needed5'], df["needed5"].values[i], na(
                df["needed5"].values[i], table_body, table_body_na))
        if i == len(df_temp) - 1:
            ws.merge_range(
                startrow,
                startcol,
                startrow,
                curcol - 1,
                "STARS PERFORMANCE",
                table_title)
    currow = currow + 1
    for i in range(0, len(footnotes)):
        ws.merge_range(currow, startcol, currow, curcol - 1, footnotes[i])
        currow = currow + 1


def stars_measure_summary(df, ws, startrow, startcol):

    df.sort_values(by=['row'])
    df_temp = df[['Measure Code',
                  'Measure',
                  'Weight',
                  'Denominator',
                  'compliance',
                  'Star Rating',
                  'pot_star_rating',
                  'missing_points',
                  'adj_aggr_difference',
                  'adj_star_rating',
                  'needed',
                  'gap_worth',
                  'trend_wt',
                  'pot_trend_wt']]
    df_temp = df_temp.copy(deep=True).where(df_temp["Weight"] != "").dropna()

    columns = {
        'Measure': 7,
        'Weight': 0,
        'Denominator': 1,
        'YTD Compliance': 1,
        'Star Rating': 1,
        'Maximum Potential Star Rating': 1,
        'Weighted Points Not Earned': 1,
        'Unearned Points Impact on Star Rating': 1,
        'Gaps Needed to Reach Max Potential': 1,
        'Adjusted Bundle Rating (Max Potential)': 1,
        'Gap Impact on Star Rating': 1
    }

    currow = startrow + 1
    ws.set_row(currow, 45)

    #ws.write(currow, startcol, "Class", header)
    #ws.merge_range(currow, startcol+1, currow, startcol+5, "Measure", header)
    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Measure'],
        "Measure",
        headerwrap)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Weight'],
        "Weight",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Denominator'],
        "Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['YTD Compliance'],
        "YTD Compliance",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Star Rating'],
        "Star Rating",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Maximum Potential Star Rating'],
        "Maximum Potential Star Rating",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Weighted Points Not Earned'],
        "Weighted Points Not Earned",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Unearned Points Impact on Star Rating'],
        "Unearned Points Impact on Star Rating",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Adjusted Bundle Rating (Max Potential)'],
        "Adjusted Bundle Rating (Max Potential)",
        headerwrap_num)

    #curcol = itercol(ws, currow, curcol, columns['Gaps Needed to Reach Max Potential'], "Gaps Needed to Reach Max Potential", headerwrap_num)
    #curcol = itercol(ws, currow, curcol, columns['Gap Impact on Star Rating'], "Gap Value to Star Rating", headerwrap_num)
    for i in range(0, len(df_temp)):
        currow = currow + 1
        curcol = startcol
        if df_temp["Measure Code"].values[i] in ['pcr', 'hpc']:
            compliance_fmt = table_body_num2
        else:
            compliance_fmt = table_body_pct
            'compliance', 'Star Rating', 'pot_star_rating', 'missing_points', 'adj_aggr_difference', 'needed', 'gap_worth'
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Measure'],
            df_temp["Measure"].values[i],
            table_body)
        curcol = itercol(
            ws, currow, curcol, columns['Weight'], df_temp["Weight"].values[i], na(
                df_temp["Weight"].values[i], table_body, table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['Denominator'], df_temp["Denominator"].values[i], na(
                df_temp["Denominator"].values[i], table_body, table_body_na))
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['YTD Compliance'],
                         df_temp["compliance"].values[i],
                         na(df_temp["compliance"].values[i],
                             compliance_fmt,
                             table_body_na))
        curcol = itercol(
            ws, currow, curcol, columns['Star Rating'], df_temp["Star Rating"].values[i], na(
                df_temp["Star Rating"].values[i], star_rating_fmt(
                    df_temp["Star Rating"].values[i]), table_body_na))
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Maximum Potential Star Rating'],
                         df_temp["pot_star_rating"].values[i],
                         na(df_temp["pot_star_rating"].values[i],
                             star_rating_fmt(df_temp["pot_star_rating"].values[i]),
                             table_body_na))
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Weighted Points Not Earned'],
                         df_temp["missing_points"].values[i],
                         na(df_temp["missing_points"].values[i],
                             table_body,
                             table_body_na))
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Unearned Points Impact on Star Rating'],
                         df_temp["adj_aggr_difference"].values[i],
                         na(df_temp["adj_aggr_difference"].values[i],
                             table_body_num2,
                             table_body_na))
        try:
            curcol = itercol(ws,
                             currow,
                             curcol,
                             columns['Adjusted Bundle Rating (Max Potential)'],
                             math.floor(df_temp["adj_star_rating"].values[i] * 100) / 100,
                             na(df_temp["adj_star_rating"].values[i],
                                 table_body_num2,
                                 table_body_na))
        except BaseException:
            curcol = itercol(
                ws, currow, curcol, columns['Adjusted Bundle Rating (Max Potential)'], '', na(
                    df_temp["adj_star_rating"].values[i], table_body_num2, table_body_na))
        #curcol = itercol(ws, currow, curcol, columns['Gaps Needed to Reach Max Potential'], df_temp["needed"].values[i], na(df_temp["needed"].values[i], table_body, table_body_na))
        #curcol = itercol(ws, currow, curcol, columns['Gap Impact on Star Rating'], df_temp["gap_worth"].values[i], na(df_temp["gap_worth"].values[i], table_body, table_body_na))

    ws.merge_range(
        startrow,
        startcol,
        startrow,
        curcol - 1,
        "STARS MEASURE SUMMARY",
        table_title)
    currow = currow + 1
    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Measure'],
        "Total:",
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Weight'],
        df_temp["Weight"].where(
            df_temp["Weight"] != "").dropna().sum(),
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Denominator'],
        "",
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['YTD Compliance'],
        "",
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Star Rating'],
        df_temp["trend_wt"].where(
            df_temp["trend_wt"] != "").dropna().sum(),
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Maximum Potential Star Rating'],
        df_temp["pot_trend_wt"].where(
            df_temp["pot_trend_wt"] != "").dropna().sum(),
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Weighted Points Not Earned'],
        df_temp["missing_points"].where(
            df_temp["missing_points"] != "").dropna().sum(),
        table_title)
    #curcol = itercol(ws, currow, curcol, columns['Unearned Points Impact on Star Rating'], df_temp["adj_aggr_difference"].where(df_temp["adj_aggr_difference"]!="").dropna().sum(), table_title)
    # Left blank due to rounding error
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Unearned Points Impact on Star Rating'],
        "",
        table_title)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Adjusted Bundle Rating (Max Potential)'],
        math.floor(
            df_temp["pot_trend_wt"].where(
                df_temp["pot_trend_wt"] != "").dropna().sum() /
            df_temp["Weight"].where(
                df_temp["Weight"] != "").dropna().sum() *
            100) /
        100,
        table_title)


def qb_practice_summary(ws, df, startrow=0, startcol=0):
    print(df)
    df.sort_values(by=["practice_id2", "row"])
    print(df)
    measures = df[["Measure"]].copy(deep=True).drop_duplicates()
    print(measures)

    practices = df[["practice_id2", 'Practice Name', 'mbr_count']].copy(
        deep=True).drop_duplicates()

    curcol = itercol(
        ws,
        startrow,
        startcol,
        1,
        "Practice ID",
        headerwrap,
        endrow=startrow +
        4)
    curcol = itercol(
        ws,
        startrow,
        curcol,
        4,
        "Practice Name",
        headerwrap,
        endrow=startrow +
        4)
    curcol = itercol(
        ws,
        startrow,
        curcol,
        1,
        "Attributed MA Members",
        headerwrap_num,
        endrow=startrow + 4)
    curcol = itercol(
        ws,
        startrow,
        curcol,
        1,
        "Practice Star Rating",
        headerwrap_num,
        endrow=startrow + 4)

    for i in range(0, len(measures)):
        curcol = itercol(
            ws,
            startrow,
            curcol,
            7,
            measures["Measure"].values[i],
            headerwrap_group,
            endrow=startrow + 3)
        ws.merge_range(
            startrow + 4,
            curcol - 8,
            startrow + 4,
            curcol - 7,
            "Numerator",
            header_num)
        ws.merge_range(
            startrow + 4,
            curcol - 6,
            startrow + 4,
            curcol - 5,
            "Denominator",
            header_num)
        ws.merge_range(
            startrow + 4,
            curcol - 4,
            startrow + 4,
            curcol - 3,
            "Compliance",
            header_num)
        ws.set_column(curcol - 8, curcol - 3, None, None,
                      {'level': 1, 'collapsed': True, 'hidden': True})
        ws.merge_range(
            startrow + 4,
            curcol - 2,
            startrow + 4,
            curcol - 1,
            "Star Rating",
            header_num)
        currow = startrow + 5

    for j in range(0, len(practices)):
        practice_measures = df.copy(deep=True).where(
            df["practice_id2"] == practices["practice_id2"].values[j]).dropna()
        curcol = itercol(
            ws,
            currow,
            startcol,
            1,
            practices["practice_id2"].values[j],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            4,
            practices["Practice Name"].values[j],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            1,
            practices["mbr_count"].values[j],
            table_body)
        curcol = itercol(
            ws, currow, curcol, 1, practice_measures["aggr_star_rating"].values[0], na(
                practice_measures["aggr_star_rating"].values[0], table_body_num2, table_body_na))
        for k in range(0, len(practice_measures)):
            if practice_measures["Measure Code"].values[k] in ['pcr', 'hpc']:
                compliance_fmt = table_body_num2
            else:
                compliance_fmt = table_body_pct
            if practice_measures["Measure Code"].values[k] in [
                    'cdc2', 'mah1a', 'mah2a', 'mah3a', 'hrm']:
                numerator = "Gaps Addressed"
            elif practice_measures["Measure Code"].values[k] in ['pcr', 'hpc']:
                numerator = "Beyond Remediation"
            else:
                numerator = "Numerator"
            curcol = itercol(
                ws, currow, curcol, 1, practice_measures[numerator].values[k], na(
                    practice_measures[numerator].values[k], table_body, table_body_na))
            curcol = itercol(
                ws, currow, curcol, 1, practice_measures["Denominator"].values[k], na(
                    practice_measures["Denominator"].values[k], table_body, table_body_na))
            curcol = itercol(
                ws, currow, curcol, 1, practice_measures["compliance"].values[k], na(
                    practice_measures["compliance"].values[k], compliance_fmt, table_body_na))
            curcol = itercol(
                ws, currow, curcol, 1, practice_measures["Star Rating"].values[k], na(
                    practice_measures["Star Rating"].values[k], star_rating_fmt(
                        practice_measures["Star Rating"].values[k]), table_body_na))
        currow = currow + 1

    # Doesnt currently work with merged cells
    #ws.autofilter(startrow+4, startcol, currow-1, curcol-1)


def hosp03_detail(ws, df, startrow=0, startcol=0, hospital=False):
    df = df.fillna("")
    print("HOSP03 DATAFRAME")
    print(df)
    width = 20
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 22
        ws.write(startrow + 2, startcol, "Hospital ID", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol + width,
        cqmbenchmarks['hosp03']['title'],
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name  ", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name  ", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date  ", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID  ", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product  ", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 9,
        "Index Admission  ",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number  ", header)
    ws.set_column(startcol + 5, startcol + 5, 15)
    ws.write(startrow + 2, startcol + 6, "Admission Date  ", header)
    ws.set_column(startcol + 6, startcol + 6, 15)
    ws.write(startrow + 2, startcol + 7, "Discharge Date  ", header)
    ws.set_column(startcol + 7, startcol + 7, 15)
    ws.write(startrow + 2, startcol + 8, "Claim Number  ", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    ws.write(startrow + 2, startcol + 9, "Type of Bill  ", header)
    ws.set_column(startcol + 9, startcol + 9, 10)

    ws.merge_range(
        startrow + 1,
        startcol + 10,
        startrow + 1,
        startcol + 16,
        "Denominator Criteria  ",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 10, "Description  ", header)
    ws.set_column(startcol + 10, startcol + 10, 25)
    ws.write(startrow + 2, startcol + 11, "Code  ", header)
    ws.set_column(startcol + 11, startcol + 11, 6)
    ws.write(startrow + 2, startcol + 12, "Service Date  ", header)
    ws.set_column(startcol + 12, startcol + 12, 12)
    ws.write(startrow + 2, startcol + 13, "Claim Number  ", header)
    ws.set_column(startcol + 13, startcol + 13, 20)
    ws.write(startrow + 2, startcol + 14, "Code  ", header)
    ws.set_column(startcol + 14, startcol + 14, 6)
    ws.write(startrow + 2, startcol + 15, "Service Date  ", header)
    ws.set_column(startcol + 15, startcol + 15, 12)
    ws.write(startrow + 2, startcol + 16, "Claim Number  ", header)
    ws.set_column(startcol + 16, startcol + 16, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 17,
        startrow + 1,
        startcol + 20,
        "Numerator Criteria",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 17, "AIS  ", header)
    ws.set_column(startcol + 17, startcol + 17, 6)
    ws.write(startrow + 2, startcol + 18, "Code  ", header)
    ws.set_column(startcol + 18, startcol + 18, 6)
    ws.write(startrow + 2, startcol + 19, "Service Date  ", header)
    ws.set_column(startcol + 19, startcol + 19, 12)
    ws.write(startrow + 2, startcol + 20, "Claim Number  ", header)
    ws.set_column(startcol + 20, startcol + 20, 20)
    ws.write(startrow + 2, startcol + 21, 'Numerator', header)
    ws.set_column(startcol + 21, startcol + 2, None, None, {'hidden': 1})

    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['EACM_LA_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_FST_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_BIR_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACAG_UNQ_MBR_ID'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_PAT_CL_NO'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EAC_ADMM_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EAC_DCG_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['EAC_SRCSY_ASND_CLM_NO'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACFBT_CD'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['description'].values[i],
            table_body)
        if df['description'].values[i] == "COPD/CF with O2":
            ws.write(
                currow,
                startcol + 11,
                df["dx2_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 12,
                df["dx2_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 13,
                df["dx2_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 14,
                df["dx3_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 15,
                df["dx3_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 16,
                df["dx3_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
        elif df['description'].values[i] == "Substantial Risk of Death":
            ws.write(
                currow,
                startcol + 11,
                df["proc_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 12,
                df["proc_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 13,
                df["proc_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
            ws.write(currow, startcol + 14, "", table_body)
            ws.write(currow, startcol + 15, "", table_body)
            ws.write(currow, startcol + 16, "", table_body)
        else:
            ws.write(
                currow,
                startcol + 11,
                df["dx_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 12,
                df["dx_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 13,
                df["dx_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
            ws.write(currow, startcol + 14, "", table_body)
            ws.write(currow, startcol + 15, "", table_body)
            ws.write(currow, startcol + 16, "", table_body)

        if df['ais'].values[i] == 1:
            ws.write(currow, startcol + 17, "Y", table_body)
        else:
            ws.write(currow, startcol + 17, "", table_body)

        ws.write(
            currow,
            startcol + 18,
            df["pall_care_code"].values[i],
            table_body)
        ws.write(
            currow,
            startcol + 19,
            df["pall_care_svce_dt"].values[i],
            table_body_date2)
        ws.write(
            currow,
            startcol + 20,
            df["pall_care_eac_srcsy_asnd_clm_no"].values[i],
            table_body)
        def num(int):
            if int == 1:
                return 1
            else:
                return ""

        ws.write(
            currow,
            startcol + 21,
            num(df["hosp03_num"].values[i]),
            table_body)
        currow = currow + 1



def hosp04_detail(ws, df, startrow=0, startcol=0, hospital=False):
    df = df.fillna("")
    print("HOSP04 DATAFRAME")
    print(df)
    width = 20
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 22
        ws.write(startrow + 2, startcol, "Hospital ID", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol + width,
        cqmbenchmarks['hosp04']['title'],
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name  ", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name  ", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date  ", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID  ", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product  ", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 9,
        "Index Admission  ",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number  ", header)
    ws.set_column(startcol + 5, startcol + 5, 15)
    ws.write(startrow + 2, startcol + 6, "Admission Date  ", header)
    ws.set_column(startcol + 6, startcol + 6, 15)
    ws.write(startrow + 2, startcol + 7, "Discharge Date  ", header)
    ws.set_column(startcol + 7, startcol + 7, 15)
    ws.write(startrow + 2, startcol + 8, "Claim Number  ", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    ws.write(startrow + 2, startcol + 9, "Type of Bill  ", header)
    ws.set_column(startcol + 9, startcol + 9, 10)

    ws.merge_range(
        startrow + 1,
        startcol + 10,
        startrow + 1,
        startcol + 16,
        "Denominator Criteria  ",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 10, "Description  ", header)
    ws.set_column(startcol + 10, startcol + 10, 25)
    ws.write(startrow + 2, startcol + 11, "Code  ", header)
    ws.set_column(startcol + 11, startcol + 11, 6)
    ws.write(startrow + 2, startcol + 12, "Service Date  ", header)
    ws.set_column(startcol + 12, startcol + 12, 12)
    ws.write(startrow + 2, startcol + 13, "Claim Number  ", header)
    ws.set_column(startcol + 13, startcol + 13, 20)
    ws.write(startrow + 2, startcol + 14, "Code  ", header)
    ws.set_column(startcol + 14, startcol + 14, 6)
    ws.write(startrow + 2, startcol + 15, "Service Date  ", header)
    ws.set_column(startcol + 15, startcol + 15, 12)
    ws.write(startrow + 2, startcol + 16, "Claim Number  ", header)
    ws.set_column(startcol + 16, startcol + 16, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 17,
        startrow + 1,
        startcol + 20,
        "Numerator Criteria",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 17, "AIS  ", header)
    ws.set_column(startcol + 17, startcol + 17, 6)
    ws.write(startrow + 2, startcol + 18, "Code  ", header)
    ws.set_column(startcol + 18, startcol + 18, 6)
    ws.write(startrow + 2, startcol + 19, "Service Date  ", header)
    ws.set_column(startcol + 19, startcol + 19, 12)
    ws.write(startrow + 2, startcol + 20, "Claim Number  ", header)
    ws.set_column(startcol + 20, startcol + 20, 20)
    ws.write(startrow + 2, startcol + 21, 'Numerator', header)
    ws.set_column(startcol + 21, startcol + 2, None, None, {'hidden': 1})

    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['EACM_LA_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_FST_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_BIR_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACAG_UNQ_MBR_ID'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_PAT_CL_NO'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EAC_ADMM_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EAC_DCG_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['EAC_SRCSY_ASND_CLM_NO'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACFBT_CD'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['description'].values[i],
            table_body)
        if df['description'].values[i] == "COPD/CF with O2":
            ws.write(
                currow,
                startcol + 11,
                df["dx2_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 12,
                df["dx2_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 13,
                df["dx2_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 14,
                df["dx3_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 15,
                df["dx3_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 16,
                df["dx3_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
        elif df['description'].values[i] == "Substantial Risk of Death":
            ws.write(
                currow,
                startcol + 11,
                df["proc_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 12,
                df["proc_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 13,
                df["proc_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
            ws.write(currow, startcol + 14, "", table_body)
            ws.write(currow, startcol + 15, "", table_body)
            ws.write(currow, startcol + 16, "", table_body)
        else:
            ws.write(
                currow,
                startcol + 11,
                df["dx_code"].values[i],
                table_body)
            ws.write(
                currow,
                startcol + 12,
                df["dx_svce_dt"].values[i],
                table_body_date2)
            ws.write(
                currow,
                startcol + 13,
                df["dx_eac_srcsy_asnd_clm_no"].values[i],
                table_body)
            ws.write(currow, startcol + 14, "", table_body)
            ws.write(currow, startcol + 15, "", table_body)
            ws.write(currow, startcol + 16, "", table_body)

        if df['ais'].values[i] == 1:
            ws.write(currow, startcol + 17, "Y", table_body)
        else:
            ws.write(currow, startcol + 17, "", table_body)

        ws.write(
            currow,
            startcol + 18,
            df["pall_care_code"].values[i],
            table_body)
        ws.write(
            currow,
            startcol + 19,
            df["pall_care_svce_dt"].values[i],
            table_body_date2)
        ws.write(
            currow,
            startcol + 20,
            df["pall_care_eac_srcsy_asnd_clm_no"].values[i],
            table_body)
        def num(int):
            if int == 1:
                return 1
            else:
                return ""

        ws.write(
            currow,
            startcol + 21,
            num(df["hosp04_num"].values[i]),
            table_body)
        currow = currow + 1
        
        

def hosp19_detail(ws, df, startrow=0, startcol=0, title='', hospital=False):
    df = df.fillna("")
    width = 11
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 13
        ws.write(startrow + 2, startcol, "Hospital ID  ", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name  ", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol +
        width,
        title,
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name  ", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name  ", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date  ", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID  ", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product  ", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 9,
        "Index ED Visit",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number  ", header)
    ws.set_column(startcol + 5, startcol + 5, 20)
    ws.write(startrow + 2, startcol + 6, "Service Date  ", header)
    ws.set_column(startcol + 6, startcol + 6, 10)
    ws.write(startrow + 2, startcol + 7, "Claim Number  ", header)
    ws.set_column(startcol + 7, startcol + 7, 15)
    ws.write(startrow + 2, startcol + 8, "Diagnosis  ", header)
    ws.set_column(startcol + 8, startcol + 8, 10)
    ws.write(startrow + 2, startcol + 9, "Diagnosis Description  ", header)
    ws.set_column(startcol + 9, startcol + 9, 100)

    ws.merge_range(
        startrow + 1,
        startcol + 10,
        startrow + 1,
        startcol + 14,
        "Return ED Visit  ",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 10, "Service Date  ", header)
    ws.set_column(startcol + 10, startcol + 10, 10)
    ws.write(startrow + 2, startcol + 11, "Claim Number  ", header)
    ws.set_column(startcol + 11, startcol + 11, 15)
    ws.write(startrow + 2, startcol + 12, "Diagnosis  ", header)
    ws.set_column(startcol + 12, startcol + 12, 10)
    ws.write(startrow + 2, startcol + 13, "Diagnosis Description  ", header)
    ws.set_column(startcol + 13, startcol + 13, 100)
    ws.write(startrow + 2, startcol + 14, "Place of Capture  ", header)
    ws.set_column(startcol + 14, startcol + 14, 35)

    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['EACM_LA_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_FST_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_BIR_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACAG_UNQ_MBR_ID'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_PAT_CL_NO'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['SVCE_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['EAC_SRCSY_ASND_CLM_NO'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRI_DIAG_CD'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            str(
                df['description_de'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['num_svce_dt'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['num_claim'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['num_diag'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            str(
                df['description_nu'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['num_provname'].values[i],
            table_body)

        currow = currow + 1


def hosp21_detail(ws, df, startrow=0, startcol=0, hospital=False):
    df = df.fillna("")
    #print("hosp21 DATAFRAME")
    # print(df)
    width = 16
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 18
        ws.write(startrow + 2, startcol, "Hospital ID  ", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name  ", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol + width,
        cqmbenchmarks['hosp21']['title'],
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 12,
        "Index Admission",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number", header)
    ws.set_column(startcol + 5, startcol + 5, 15)
    ws.write(startrow + 2, startcol + 6, "Admission Date", header)
    ws.set_column(startcol + 6, startcol + 6, 10)
    ws.write(startrow + 2, startcol + 7, "Discharge Date", header)
    ws.set_column(startcol + 7, startcol + 7, 10)
    ws.write(startrow + 2, startcol + 8, "Claim Number", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    ws.write(startrow + 2, startcol + 9, "Type of Bill", header)
    ws.set_column(startcol + 9, startcol + 9, 10)
    ws.write(startrow + 2, startcol + 10, "Description", header)
    ws.set_column(startcol + 10, startcol + 10, 30)
    ws.write(startrow + 2, startcol + 11, "DRG", header)
    ws.set_column(startcol + 11, startcol + 11, 10)
    ws.write(startrow + 2, startcol + 12, "Discharge Status", headerwrap)
    ws.set_column(startcol + 12, startcol + 12, 10)
    ws.set_row(startrow + 2, 30, headerwrap)
    ws.merge_range(
        startrow + 1,
        startcol + 13,
        startrow + 1,
        startcol + 16,
        "Follow-up Visit",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 13, "Service Date", headerwrap)
    ws.set_column(startcol + 13, startcol + 13, 10)
    ws.write(startrow + 2, startcol + 14, "Claim Number", headerwrap)
    ws.set_column(startcol + 14, startcol + 14, 20)
    ws.write(startrow + 2, startcol + 15, "Description", header)
    ws.set_column(startcol + 15, startcol + 15, 30)
    ws.write(startrow + 2, startcol + 16, "Code", header)
    ws.set_column(startcol + 16, startcol + 16, 15)

    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['EACM_LA_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_FST_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_BIR_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACAG_UNQ_MBR_ID'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_PAT_CL_NO'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EAC_ADMM_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EAC_DCG_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['EAC_SRCSY_ASND_CLM_NO'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACFBT_CD'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['description'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['CMN_EACDRG_CD'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACDS_CD'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['follow_up_svce_dt'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['follow_up_clm_no'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['follow_up_description'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['follow_up_proc_code'].values[i],
            table_body)
        currow = currow + 1


def reads_detail(ws, df, startrow=0, startcol=0, title="", hospital=False):
    df = df.fillna("")
    #print("hosp21 DATAFRAME")
    # print(df)
    width = 12
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 14
        ws.write(startrow + 2, startcol, "Hospital ID", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol +
        width,
        title,
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 11,
        "Index Admission",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number", header)
    ws.set_column(startcol + 5, startcol + 5, 15)
    ws.write(startrow + 2, startcol + 6, "Admission Date", header)
    ws.set_column(startcol + 6, startcol + 6, 10)
    ws.write(startrow + 2, startcol + 7, "IESD  ", header)
    ws.set_column(startcol + 7, startcol + 7, 10)
    ws.write(startrow + 2, startcol + 8, "Diagnosis", header)
    ws.set_column(startcol + 8, startcol + 8, 10)
    ws.write(startrow + 2, startcol + 9, "Description", headerwrap)
    ws.set_column(startcol + 9, startcol + 9, 100)
    ws.write(startrow + 2, startcol + 10, "MDC Description", headerwrap)
    ws.set_column(startcol + 10, startcol + 10, 100)
    ws.write(startrow + 2, startcol + 11, "DRG Description", headerwrap)
    ws.set_column(startcol + 11, startcol + 11, 100)

    ws.set_row(startrow + 2, 30, headerwrap)
    ws.merge_range(
        startrow + 1,
        startcol + 12,
        startrow + 1,
        startcol + 15,
        "Readmission",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 12, "Date", headerwrap)
    ws.set_column(startcol + 12, startcol + 12, 10)
    ws.write(startrow + 2, startcol + 13, "Diagnosis", headerwrap)
    ws.set_column(startcol + 13, startcol + 13, 10)
    ws.write(startrow + 2, startcol + 14, "Description", headerwrap)
    ws.set_column(startcol + 14, startcol + 14, 100)
    ws.write(startrow + 2, startcol + 15, "Place of Capture", headerwrap)
    ws.set_column(startcol + 15, startcol + 15, 35)

    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['MEM_LNAME'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['MEM_FNAME'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['eacm_bir_dt'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['UMI'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PCN'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['ADM_DT2'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['IESD2'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['DIAG_I_1'].values[i],
            table_body)
        curcol = itercol(ws, currow, curcol, 0, str(
            df['dx_description'].values[i]), table_body)
        curcol = itercol(ws, currow, curcol, 0, str(
            df['MDC_Description'].values[i]), table_body)
        curcol = itercol(ws, currow, curcol, 0, str(
            df['DRG_Description'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['READMITDATE30'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['DIAG_I_1_30'].values[i],
            table_body)
        curcol = itercol(
            ws, 
            currow, 
            curcol, 
            0, 
            str(
                df['dx_30_description'].values[i]), table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_NM30readmit'].values[i],
            table_body)
        #curcol = itercol(ws, currow, curcol, 0, df['PROV_NBR30readmit'].values[i], table_body)
        #curcol = itercol(ws, currow, curcol, 0, df['PRV_NM30readmit'].values[i], table_body)
        # df['dx_30_description'].values[i]
        currow = currow + 1

def hosp22_detail(ws, df, startrow=0, startcol=0, hospital=False):
    df = df.fillna("")
    print("HOSP22 DATAFRAME")
    print(df)
    width = 8
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 8
        ws.write(startrow + 2, startcol, "Hospital ID", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol + width,
        cqmbenchmarks['hosp22']['title'],
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name  ", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name  ", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date  ", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID  ", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product  ", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 7,
        "Denominator Criteria",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number  ", header)
    ws.set_column(startcol + 5, startcol + 5, 15)
    ws.write(startrow + 2, startcol + 6, "Service Date  ", header)
    ws.set_column(startcol + 6, startcol + 6, 15)
    ws.write(startrow + 2, startcol + 7, "Claim Number  ", header)
    ws.set_column(startcol + 7, startcol + 7, 20)

    ws.write(startrow + 1, startcol + 8, "Numerator Criteria", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    ws.write(startrow + 2, startcol + 8, "Indicator", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    
    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['EACM_LA_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_FST_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_BIR_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACAG_UNQ_MBR_ID'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_PAT_CL_NO'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['svce_dt'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['EAC_SRCSY_ASND_CLM_NO'].values[i]), table_body)

        def num(int):
            if int == 1:
                return 1
            else:
                return ""

        ws.write(
            currow,
            startcol + 8,
            num(df["hosp22_num"].values[i]),
            table_body)
        currow = currow + 1


def hosp23_detail(ws, df, startrow=0, startcol=0, hospital=False):
    df = df.fillna("")
    print("HOSP23 DATAFRAME")
    print(df)
    width = 8
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 8
        ws.write(startrow + 2, startcol, "Hospital ID", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol + width,
        cqmbenchmarks['hosp23']['title'],
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name  ", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name  ", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date  ", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID  ", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product  ", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 7,
        "Denominator Criteria",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number  ", header)
    ws.set_column(startcol + 5, startcol + 5, 15)
    ws.write(startrow + 2, startcol + 6, "Service Date  ", header)
    ws.set_column(startcol + 6, startcol + 6, 15)
    ws.write(startrow + 2, startcol + 7, "Claim Number  ", header)
    ws.set_column(startcol + 7, startcol + 7, 20)

    ws.write(startrow + 1, startcol + 8, "Numerator Criteria", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    ws.write(startrow + 2, startcol + 8, "Indicator", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    
    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['EACM_LA_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_FST_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_BIR_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACAG_UNQ_MBR_ID'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_PAT_CL_NO'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['svce_dt'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['EAC_SRCSY_ASND_CLM_NO'].values[i]), table_body)

        def num(int):
            if int == 1:
                return 1
            else:
                return ""

        ws.write(
            currow,
            startcol + 8,
            num(df["hosp23_num"].values[i]),
            table_body)
        currow = currow + 1


def hosp24_detail(ws, df, startrow=0, startcol=0, hospital=False):
    df = df.fillna("")
    print("HOSP24 DATAFRAME")
    print(df)
    width = 8
    length = len(df)
    autostartcol = startcol
    if hospital:
        width = 8
        ws.write(startrow + 2, startcol, "Hospital ID", header)
        ws.set_column(startcol, startcol, 10)
        ws.write(startrow + 2, startcol + 1, "Hospital Name", header)
        ws.set_column(startcol + 1, startcol + 1, 30)
        startcol = startcol + 2

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol + width,
        cqmbenchmarks['hosp24']['title'],
        table_title)
    ws.write(startrow + 2, startcol, "Member Last Name  ", header)
    ws.set_column(startcol, startcol, 20)
    ws.write(startrow + 2, startcol + 1, "Member First Name  ", header)
    ws.set_column(startcol + 1, startcol + 1, 20)
    ws.write(startrow + 2, startcol + 2, "Birth Date  ", header)
    ws.set_column(startcol + 2, startcol + 2, 10)
    ws.write(startrow + 2, startcol + 3, "Unique Member ID  ", header)
    ws.set_column(startcol + 3, startcol + 3, 20)
    ws.write(startrow + 2, startcol + 4, "Product  ", header)
    ws.set_column(startcol + 4, startcol + 4, 20)

    ws.merge_range(
        startrow + 1,
        startcol + 5,
        startrow + 1,
        startcol + 7,
        "Denominator Criteria",
        headerwrap_group)
    ws.write(startrow + 2, startcol + 5, "Patient Number  ", header)
    ws.set_column(startcol + 5, startcol + 5, 15)
    ws.write(startrow + 2, startcol + 6, "Service Date  ", header)
    ws.set_column(startcol + 6, startcol + 6, 15)
    ws.write(startrow + 2, startcol + 7, "Claim Number  ", header)
    ws.set_column(startcol + 7, startcol + 7, 20)

    ws.write(startrow + 1, startcol + 8, "Numerator Criteria", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    ws.write(startrow + 2, startcol + 8, "Indicator", header)
    ws.set_column(startcol + 8, startcol + 8, 20)
    
    currow = startrow + 3
    for i in range(0, len(df)):
        if hospital:
            startcol = startcol - 2
            curcol = itercol(
                ws,
                currow,
                startcol,
                0,
                df['quality_blue_id'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                0,
                df['quality_blue_name'].values[i],
                table_body)
            startcol = startcol + 2
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['EACM_LA_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_FST_NM'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACM_BIR_DT'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['EACAG_UNQ_MBR_ID'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['product'].values[i],
            table_body_date)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['PRV_PAT_CL_NO'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['svce_dt'].values[i],
            table_body_date2)
        curcol = itercol(
            ws, currow, curcol, 0, str(
                df['EAC_SRCSY_ASND_CLM_NO'].values[i]), table_body)
        def num(int):
            if int == 1:
                return 1
            else:
                return ""

        ws.write(
            currow,
            startcol + 8,
            num(df["hosp24_num"].values[i]),
            table_body)
        currow = currow + 1



def cqm_mbr_detail(ws, df, startrow=0, startcol=0):
    df = df.fillna("")
    #print("hosp21 DATAFRAME")
    # print(df)
    width = 28
    length = len(df)
    autostartcol = startcol

    ws.autofilter(
        startrow + 2,
        autostartcol,
        startrow + 2 + length,
        autostartcol + width)
    ws.freeze_panes(startrow + 3, 0)

    ws.merge_range(
        startrow,
        autostartcol,
        startrow,
        autostartcol +
        width,
        "MEMBER SUMMARY",
        table_title)
    ws.write(startrow + 2, startcol, "Practice ID", header)
    ws.set_column(startcol, startcol, 12)
    ws.write(startrow + 2, startcol + 1, "Practice Name", header)
    ws.set_column(startcol + 1, startcol + 1, 30)
    ws.write(startrow + 2, startcol + 2, "Physician NPI", header)
    ws.set_column(startcol + 2, startcol + 2, 15)
    ws.write(startrow + 2, startcol + 3, "Physician Name", header)
    ws.set_column(startcol + 3, startcol + 3, 30)
    ws.write(startrow + 2, startcol + 4, "Member Last Name", header)
    ws.set_column(startcol + 4, startcol + 4, 20)
    ws.write(startrow + 2, startcol + 5, "Member First Name", header)
    ws.set_column(startcol + 5, startcol + 5, 20)
    ws.write(startrow + 2, startcol + 6, "Birth Date", header)
    ws.set_column(startcol + 6, startcol + 6, 10)
    ws.write(startrow + 2, startcol + 7, "Unique Member ID", header)
    ws.set_column(startcol + 7, startcol + 7, 20)

    ws.set_column(startcol + 4, startcol + 27, 27)
    ws.merge_range(
        startrow + 1,
        startcol + 8,
        startrow + 1,
        startcol + 9,
        cqmbenchmarks['hosp03']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 8, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 9, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 10,
        startrow + 1,
        startcol + 11,
        cqmbenchmarks['hosp04']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 10, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 11, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 12,
        startrow + 1,
        startcol + 13,
        cqmbenchmarks['hosp19']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 12, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 13, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 14,
        startrow + 1,
        startcol + 15,
        cqmbenchmarks['hosp20']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 14, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 15, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 16,
        startrow + 1,
        startcol + 17,
        cqmbenchmarks['rrama']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 16, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 17, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 18,
        startrow + 1,
        startcol + 19,
        cqmbenchmarks['rracomm']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 18, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 19, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 20,
        startrow + 1,
        startcol + 21,
        cqmbenchmarks['hosp21']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 20, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 21, "Numerator  ", header_center)
    

    ws.merge_range(
        startrow + 1,
        startcol + 22,
        startrow + 1,
        startcol + 23,
        cqmbenchmarks['hosp22']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 22, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 23, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 24,
        startrow + 1,
        startcol + 25,
        cqmbenchmarks['hosp23']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 24, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 25, "Numerator  ", header_center)

    ws.merge_range(
        startrow + 1,
        startcol + 26,
        startrow + 1,
        startcol + 27,
        cqmbenchmarks['hosp24']['title'],
        headerwrap_group)
    ws.write(startrow + 2, startcol + 26, "Denominator  ", header_center)
    ws.write(startrow + 2, startcol + 27, "Numerator  ", header_center)

    ws.write(startrow + 2, startcol + 28, "Last PCP Office Visit  ", header_center)
    ws.set_column(startcol + 28, startcol + 28, 25)

    currow = startrow + 3
    for i in range(0, len(df)):
        curcol = itercol(
            ws,
            currow,
            startcol,
            0,
            df['practice_id'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['practice_name'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['physician_npi'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['physician_name'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['mbr_last_nm'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['mbr_frst_nm'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['mbr_bir_dt'].values[i],
            table_body_date2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['umi'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp03_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp03_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp04_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp04_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp19_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp19_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp20_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp20_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['rrama_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['rrama_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['rracomm_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['rracomm_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp21_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp21_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp22_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp22_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp23_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp23_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp24_den'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['hosp24_num'].values[i],
            table_body_center)
        curcol = itercol(
            ws,
            currow,
            curcol,
            0,
            df['LAST_PCP_VISIT_DATE'].values[i],
            table_body_date2)

        currow = currow + 1





def cqmtable(worksheet, startrow, df):
    worksheet.merge_range(startrow,0,startrow,9,'CLINICAL QUALITY METRICS',table_title)
    worksheet.merge_range(startrow + 1,0,startrow + 1,5,'Clinical Quality Metric',header)
    worksheet.write(startrow + 1, 6, 'Rate', header_num)
    worksheet.write(startrow + 1, 7, 'Earned', header_num)
    worksheet.merge_range(startrow + 1,8,startrow + 1,9,'Available',header_num)
    worksheet.merge_range(startrow,11,startrow,18,'THRESHOLDS',table_title)
    worksheet.write(startrow + 1, 11, 'Max', header_num)
    worksheet.merge_range(startrow + 1,12,startrow + 1,13,'Points',header_num)
    worksheet.write(startrow + 1, 14, 'Mid', header_num)
    worksheet.merge_range(startrow + 1,15,startrow + 1,16,'Points',header_num)
    worksheet.write(startrow + 1, 17, 'Min', header_num)
    worksheet.merge_range(startrow + 1,18,startrow + 1,19,'Points',header_num)

    # Function to create the rows for Clinical Quality Metrics, otherwise this
    # code would be repetitive
    print("***************************len(df)**********************")
    print("***************************len(df)**********************")
    print("***************************len(df)**********************")
    print(len(df))
    print(len(df))
    print(len(df))
    print("***************************startrow**********************")
    print("***************************startrow**********************")
    print("***************************startrow**********************")
    print(startrow)
    print(startrow)
    print(startrow)
    print("***************************PRINT # Function to create the rows for Clinical Quality Metrics, otherwise this**********************")
    print("***************************PRINT # Function to create the rows for Clinical Quality Metrics, otherwise this**********************")
    print("***************************PRINT # Function to create the rows for Clinical Quality Metrics, otherwise this**********************")
    def cqmrow(
            row,
            metric,
            rate,
            points,
            available,
            maxpoints,
            midpoints,
            minpoints,
            measure_cd):

        print("***************************PRINT # cqmrow FOR LOOPING:**********************")
        print(row)
        print(metric)
        print(rate)
        print(points)
        print(available)
        print(maxpoints)
        print(midpoints)
        print(minpoints)
        print(row)
        print("***************************PRINT # cqmrow FOR LOOPING:**********************")
        print(row)
        print(metric)
        print(row)
        print("***************************PRINT # cqmrow FOR LOOPING:**********************")
        maxpoints = round(maxpoints, 4)
        midpoints = round(midpoints, 4)
        minpoints = round(minpoints, 4)
        
        print("***************************PRINT # if maxpoints > midpoints:**********************")
        print("***************************PRINT # if maxpoints > midpoints:**********************")
        print("***************************PRINT # if maxpoints > midpoints:**********************")
        if maxpoints > midpoints:
            gtlt = u'\u2265'
        elif maxpoints < midpoints:
            gtlt = u'\u2264'

        worksheet.merge_range(row, 0, row, 5, metric, table_body)
        print("***************************PRINT # if points == 0 and available != 0:**********************")
        print("***************************PRINT # if points == 0 and available != 0:**********************")
        print("***************************PRINT # if points == 0 and available != 0:**********************")
        if points == 0 and available != 0:
            worksheet.write(row, 6, rate, table_body_pct2_red)
            worksheet.write(row, 7, points, table_body_num2)
            worksheet.merge_range(row, 8, row, 9, available, table_body)
        else:
            print("********************WITHIN ELSE points == 0 and available != 0:**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:**************************************")
            worksheet.write(row, 6, rate, table_body_pct2)
            worksheet.write(row, 7, points, table_body_num2)
            worksheet.merge_range(row, 8, row, 9, available, table_body)
        print("***************************PRINT # if points == 0 and available != 0:22222**********************")
        print("***************************PRINT # if points == 0 and available != 0:22222**********************")
        print("***************************PRINT # if points == 0 and available != 0:22222**********************")
        print(points)
        print(available)
        print(maxpoints)
        if points == available and available != 0:
            worksheet.write(row, 6, rate, table_body_pct2_green)
            worksheet.write(row,11,gtlt + '{:.2%}'.format(maxpoints),table_body_pct2_green)
            worksheet.merge_range(row,12,row,13,str(available) + " Points",table_body_pct2_green)
        else:
            print("********************WITHIN ELSE points == 0 and available != 0:22222**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:22222**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:22222**************************************")
            if available != 0:
                worksheet.write(row,11,gtlt + '{:.2%}'.format(maxpoints),table_body_pct2)
                print("********************WITHIN ELSE points == 0 and available != 0:22222 PART II**************************************")
                print("********************WITHIN ELSE points == 0 and available != 0:22222 PART II**************************************")
                print("********************WITHIN ELSE points == 0 and available != 0:22222 PART II**************************************")
                worksheet.merge_range(row, 12, row, 13, str(available) + " Points", table_body_pct2)
        print("***************************PRINT # if points == 0 and available != 0:33333**********************")
        print("***************************PRINT # if points == 0 and available != 0:33333**********************")
        print("***************************PRINT # if points == 0 and available != 0:33333**********************")
        if points == cqmbenchmarks[measure_cd]['pointsmid'] and available != 0:
            worksheet.write(row, 6, rate, table_body_pct2_yellow)
            worksheet.write(row,14,gtlt + '{:.2%}'.format(midpoints),table_body_pct2_yellow)
            worksheet.merge_range(row, 15, row, 16, str(cqmbenchmarks[measure_cd]['pointsmid']) + " Points", table_body_pct2_yellow)
        else:
            print("********************WITHIN ELSE points == 0 and available != 0:33333**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:33333**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:33333**************************************")
            if available != 0:
                worksheet.write(row,14,gtlt + '{:.2%}'.format(midpoints),table_body_pct2)
                worksheet.merge_range(row, 15, row, 16, str(cqmbenchmarks[measure_cd]['pointsmid']) + " Points", table_body_pct2)

        print("***************************PRINT # if points == 0 and available != 0:44444**********************")
        print("***************************PRINT # if points == 0 and available != 0:44444**********************")
        print("***************************PRINT # if points == 0 and available != 0:44444**********************")
        print(points)
        print(available)
        print(minpoints)
        if points == available * .5 and available != 0:
            worksheet.write(row, 6, rate, table_body_pct2_orange)
            worksheet.write(row,17,gtlt + '{:.2%}'.format(minpoints),table_body_pct2_orange)
            worksheet.merge_range(row, 18, row, 19, str(cqmbenchmarks[measure_cd]['pointsmin']) + " Points", table_body_pct2_orange)
        else:
            print("********************WITHIN ELSE points == 0 and available != 0:44444**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:44444**************************************")
            print("********************WITHIN ELSE points == 0 and available != 0:44444**************************************")
            if available != 0:
                worksheet.write(row,17,gtlt + '{:.2%}'.format(minpoints),table_body_pct2)
                print("********************WITHIN ELSE points == 0 and available != 0:44444 - 222222**************************************")
                print("********************WITHIN ELSE points == 0 and available != 0:44444 - 222222**************************************")
                print("********************WITHIN ELSE points == 0 and available != 0:44444 - 222222**************************************")
                print(row)
                print(table_body_pct2)
                worksheet.merge_range(row, 18, row, 19, str(cqmbenchmarks[measure_cd]['pointsmin']) + " Points", table_body_pct2)
                
    for i in range(0, len(df)):
            cqmrow(startrow + 2 + i,
               cqmbenchmarks[df["measure_cd"].values[i]]["title"],
               df["rate"].values[i],
               df["points_earned"].values[i],
               df["points_available"].values[i],
               df["max"].values[i],
               df["mid"].values[i],
               df["min"].values[i],
               df["measure_cd"].values[i])
               
    print("***************************PRINT # worksheet.write(startrow + len(df) + 2, 0, TOTAL, table_title:**********************")
    print("***************************PRINT # worksheet.write(startrow + len(df) + 2, 0, TOTAL, table_title:**********************")
    print("***************************PRINT # worksheet.write(startrow + len(df) + 2, 0, TOTAL, table_title:**********************")
    worksheet.write(startrow + len(df) + 2, 0, "TOTAL", table_title)
    worksheet.merge_range(startrow + len(df) + 2,1, startrow + len(df) + 2,7, df["total_points_earned"].values[0],table_title)
    worksheet.merge_range(startrow + len(df) + 2,8, startrow + len(df) + 2,9, df["total_points_available"].values[0],table_title)
    worksheet.merge_range(startrow + len(df) + 2,11,startrow + len(df) + 2,19,"",table_title)

    return startrow + len(df) + 4


def cqm_monthly(wb, ws, df, measure_cd, startrow):
    df = df.copy(
        deep=True).where(
        df['measure_cd'] == measure_cd).dropna(
            how='all').sort_values(
                by=['month'])
    print(measure_cd, df)
    if measure_cd in ['rrama', 'rracomm']:
        ws.merge_range(
            startrow,
            0,
            startrow,
            14,
            cqmbenchmarks[measure_cd]['title'],
            table_title)
    else:
        ws.merge_range(
            startrow,
            0,
            startrow,
            8,
            cqmbenchmarks[measure_cd]['title'],
            table_title)
    currow = startrow + 1
    ws.merge_range(currow, 0, currow, 1, "Period", header)
    ws.merge_range(currow, 2, currow, 3, "Denominator", header_num)
    ws.merge_range(currow, 4, currow, 5, "Numerator", header_num)

    if measure_cd in ['rrama', 'rracomm']:
        ws.merge_range(currow, 6, currow, 7, "Observed", header_num)
        ws.merge_range(currow, 8, currow, 9, "Expected", header_num)
        ws.merge_range(currow, 10, currow, 11, "Market Rate", header_num)
        ws.merge_range(
            currow,
            12,
            currow,
            13,
            "Risk-Adjusted Rate",
            header_num)
        ws.write(currow, 14, "Tier", header)
    else:
        ws.merge_range(currow, 6, currow, 7, "Rate", header_num)
        ws.write(currow, 8, "Tier", header)

    for i in range(0, 14):
        currow = currow + 1
        try:
            ws.merge_range(currow, 0, currow, 1, "=date(" +
                           str(df['year'].values[i]) +
                           ", " +
                           str(df['month'].values[i] +
                               1) +
                           ", 1)-1", table_body_date)
            ws.merge_range(
                currow,
                2,
                currow,
                3,
                df['denominator'].values[i],
                table_body)
            ws.merge_range(
                currow,
                4,
                currow,
                5,
                df['numerator'].values[i],
                table_body)
            if measure_cd in ['rrama', 'rracomm']:
                ws.merge_range(
                    currow,
                    6,
                    currow,
                    7,
                    df["observed"].values[i],
                    table_body_pct2)
                ws.merge_range(
                    currow,
                    8,
                    currow,
                    9,
                    df["expected"].values[i],
                    table_body_pct2)
                ws.merge_range(
                    currow,
                    10,
                    currow,
                    11,
                    df["mkt_expected"].values[i],
                    table_body_pct2)
                ws.merge_range(
                    currow,
                    12,
                    currow,
                    13,
                    df["rate"].values[i],
                    table_body_pct2)
                ws.write(currow, 14, df["tier"].values[i], table_body)
            else:
                ws.merge_range(
                    currow,
                    6,
                    currow,
                    7,
                    df['rate'].values[i],
                    table_body_pct2)
                ws.write(currow, 8, df['tier'].values[i], table_body)
        except BaseException:
            ws.merge_range(currow, 0, currow, 1, "", table_body_date)
            ws.merge_range(currow, 2, currow, 3, '', table_body)
            ws.merge_range(currow, 4, currow, 5, '', table_body)
            if measure_cd in ['rrama', 'rracomm']:
                ws.merge_range(currow, 6, currow, 7, '', table_body_pct2)
                ws.merge_range(currow, 8, currow, 9, '', table_body_pct2)
                ws.merge_range(currow, 10, currow, 11, '', table_body_pct2)
                ws.merge_range(currow, 12, currow, 13, '', table_body_pct2)
                ws.write(currow, 14, '', table_body)
            else:
                ws.merge_range(currow, 6, currow, 7, '', table_body_pct2)
                ws.write(currow, 8, '', table_body)

    series1 = '=' + "'" + ws.get_name() + "'" + '!' + colnum_string(0) + str(startrow + 3) + \
        ':' + colnum_string(0) + str(startrow + 2 + len(df['month']))
    series2 = '=' + "'" + ws.get_name() + "'" + '!' + colnum_string(6) + str(startrow + 3) + \
        ':' + colnum_string(6) + str(startrow + 2 + len(df['month']))

    if measure_cd in ['rrama', 'rracomm']:

        chart = wb.add_chart(
            {'type': 'line'}
        )

        chart.add_series({
            'categories': series1,
            'values': series2,
            'name': "Observed",
            'marker': {'type': 'square', 'fill': {'color': blue2}}
        })

        chart.add_series({
            'categories': [ws.get_name(), startrow + 2, 0, startrow + 1 + len(df['month']), 0],
            'values': [ws.get_name(), startrow + 2, 8, startrow + 1 + len(df['month']), 8],
            'line': {'color': red1},
            'name': "Expected"
        })

        chart.add_series({
            'categories': [ws.get_name(), startrow + 2, 0, startrow + 1 + len(df['month']), 0],
            'values': [ws.get_name(), startrow + 2, 10, startrow + 1 + len(df['month']), 10],
            'line': {'color': purple1},
            'name': "Market Expected"
        })

        chart.add_series({
            'categories': [ws.get_name(), startrow + 2, 0, startrow + 1 + len(df['month']), 0],
            'values': [ws.get_name(), startrow + 2, 12, startrow + 1 + len(df['month']), 12],
            'line': {'color': orange1},
            'name': "Risk-Adjusted Rate"
        })

        chart.set_x_axis({
            'date_axis': True,
        })

        chart.set_title({'name': cqmbenchmarks[measure_cd]['title']})

        ws.insert_chart(startrow, 16, chart, {'x_scale': 1.5, 'y_scale': 1.1})
    else:

        series2 = '=' + "'" + ws.get_name() + "'" + '!' + colnum_string(6) + str(startrow + 3) + \
            ':' + colnum_string(6) + str(startrow + 2 + len(df['month']))
        chart = wb.add_chart(
            {'type': 'line'}
        )

        chart.add_series({
            'categories': series1,
            'values': series2,
            'name': "Rate",
            'line': {
                'color': blue2,
                'width': 2
            },
            'marker': {'type': 'square', 'fill': {'color': blue2}}
        })

        chart.set_x_axis({
            'date_axis': True,
        })

        chart.set_title({'name': cqmbenchmarks[measure_cd]['title']})

        ws.insert_chart(startrow, 10, chart, {'x_scale': 1.5, 'y_scale': 1.1})

    return currow + 2


def hosp03_drivers(ws, df, startrow=6, startcol=0):

    columns = {
        'Condition': 5,
        'Denominator': 1,
        'Numerator': 1,
        'Rate': 0,
        'QB Program Rate': 1,
        'Similar Size Rate': 1,
        'Tier': 0,
        '% of Total Denominator': 1,
        '% of Total Numerator': 1
    }

    currow = startrow + 1
    ws.set_row(currow, 45)

    #ws.write(currow, startcol, "Class", header)
    #ws.merge_range(currow, startcol+1, currow, startcol+5, "Measure", header)
    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Condition'],
        "Condition",
        headerwrap)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Denominator'],
        "Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Numerator'],
        "Numerator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Rate'],
        "Rate",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['QB Program Rate'],
        "QB Program Rate",
        headerwrap_num)
    #curcol = itercol(ws, currow, curcol, columns['Similar Size Rate'], "Similar Size Rate", headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Denominator'],
        "% of Total Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Numerator'],
        "% of Total Numerator",
        headerwrap_num)

    ws.merge_range(
        startrow,
        startcol,
        startrow,
        curcol - 1,
        "KEY DRIVERS",
        table_title)

    for i in range(0, len(df)):
        curcol = startcol
        currow = currow + 1
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Condition'],
            df['description'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Denominator'],
            df['hosp03_den'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Numerator'],
            df['hosp03_num'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Rate'],
            df['rate'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['QB Program Rate'],
            df['program_rate'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Denominator'],
            df['pct_total_den'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Numerator'],
            df['pct_total_num'].values[i],
            table_body_pct2)

    return currow + 2

def hosp04_drivers(ws, df, startrow=6, startcol=0):

    columns = {
        'Condition': 5,
        'Denominator': 1,
        'Numerator': 1,
        'Rate': 0,
        'QB Program Rate': 1,
        'Similar Size Rate': 1,
        'Tier': 0,
        '% of Total Denominator': 1,
        '% of Total Numerator': 1
    }

    currow = startrow + 1
    ws.set_row(currow, 45)

    #ws.write(currow, startcol, "Class", header)
    #ws.merge_range(currow, startcol+1, currow, startcol+5, "Measure", header)
    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Condition'],
        "Condition",
        headerwrap)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Denominator'],
        "Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Numerator'],
        "Numerator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Rate'],
        "Rate",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['QB Program Rate'],
        "QB Program Rate",
        headerwrap_num)
    #curcol = itercol(ws, currow, curcol, columns['Similar Size Rate'], "Similar Size Rate", headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Denominator'],
        "% of Total Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Numerator'],
        "% of Total Numerator",
        headerwrap_num)

    ws.merge_range(
        startrow,
        startcol,
        startrow,
        curcol - 1,
        "KEY DRIVERS",
        table_title)

    for i in range(0, len(df)):
        curcol = startcol
        currow = currow + 1
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Condition'],
            df['description'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Denominator'],
            df['hosp04_den'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Numerator'],
            df['hosp04_num'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Rate'],
            df['rate'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['QB Program Rate'],
            df['program_rate'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Denominator'],
            df['pct_total_den'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Numerator'],
            df['pct_total_num'].values[i],
            table_body_pct2)

    return currow + 2

def hosp19_summary(ws, measure_cd, df, startrow=6, startcol=0):
    ws.merge_range(
        startrow,
        startcol,
        startrow,
        startcol + 9,
        cqmbenchmarks[measure_cd]['title'],
        table_title)
    currow = startrow + 1
    ws.merge_range(currow, startcol, currow, startcol + 4, "Scenario", header)
    ws.merge_range(
        currow,
        startcol + 5,
        currow,
        startcol + 6,
        "Denominator",
        header_num)
    ws.merge_range(
        currow,
        startcol + 7,
        currow,
        startcol + 8,
        "Numerator",
        header_num)
    ws.write(currow, startcol + 9, "Rate", header_num)
    currow = currow + 1
    ws.merge_range(
        currow,
        startcol,
        currow,
        startcol + 4,
        "Hospital Rate",
        table_body)
    ws.merge_range(currow, startcol + 5, currow, startcol + 6,
                   df[measure_cd + '_den'].values[0], table_body)
    ws.merge_range(currow, startcol + 7, currow, startcol + 8,
                   df[measure_cd + '_num'].values[0], table_body)
    try:
        ws.write(currow,
                 startcol + 9,
                 df[measure_cd + '_rate'].values[0],
                 table_body_pct2)
    except BaseException:
        ws.write(currow,
                 startcol + 9,
                 df[measure_cd + '_rate'].values[0],
                 table_body_pct2)
    currow = currow + 1
    ws.merge_range(
        currow,
        startcol,
        currow,
        startcol + 4,
        "Same Hospital Rate",
        table_body)
    ws.merge_range(currow, startcol + 5, currow, startcol + 6,
                   df[measure_cd + '_num'].values[0], table_body)
    ws.merge_range(
        currow,
        startcol + 7,
        currow,
        startcol + 8,
        df['same_hosp_num'].values[0],
        table_body)
    try:
        ws.write(currow, startcol +
                 9, df['same_hosp_num'].values[0] /
                 df[measure_cd +
                    '_num'].values[0], table_body_pct2)
    except BaseException:
        ws.write(currow, startcol + 9, '', table_body_pct2)
    currow = currow + 1
    ws.merge_range(
        currow,
        startcol,
        currow,
        startcol + 4,
        "Same Diagnosis Rate",
        table_body)
    ws.merge_range(currow, startcol + 5, currow, startcol + 6,
                   df[measure_cd + '_num'].values[0], table_body)
    ws.merge_range(currow, startcol + 7, currow, startcol + 8,
                   df['same_dx_' + measure_cd + '_num'].values[0], table_body)
    try:
        ws.write(currow, startcol +
                 9, df['same_dx_' +
                       measure_cd +
                       '_num'].values[0] /
                 df[measure_cd +
                    '_num'].values[0], table_body_pct2)
    except BaseException:
        ws.write(currow, startcol + 9, '', table_body_pct2)

    return currow + 2


def hosp19_drivers(ws, df, measure_cd, startrow=6, startcol=0):

    columns = {
        'Primary Diagnosis': 5,
        'Code': 0,
        'Denominator': 1,
        'Numerator': 1,
        'Rate': 0,
        'QB Program Rate': 1,
        'Similar Size Rate': 1,
        'Tier': 0,
        '% of Total Denominator': 1,
        '% of Total Numerator': 1
    }

    currow = startrow + 1
    ws.set_row(currow, 45)

    #ws.write(currow, startcol, "Class", header)
    #ws.merge_range(currow, startcol+1, currow, startcol+5, "Measure", header)
    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Primary Diagnosis'],
        "Primary Diagnosis",
        headerwrap)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Code'],
        "Code",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Denominator'],
        "Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Numerator'],
        "Numerator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Rate'],
        "Rate",
        headerwrap_num)
    #curcol = itercol(ws, currow, curcol, columns['QB Program Rate'], "QB Program Rate", headerwrap_num)
    #curcol = itercol(ws, currow, curcol, columns['Similar Size Rate'], "Similar Size Rate", headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Denominator'],
        "% of Total Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Numerator'],
        "% of Total Numerator",
        headerwrap_num)

    ws.merge_range(startrow, startcol, startrow, curcol -
                   1, "Top 20 Diagnosis Codes by Volume", table_title)

    for i in range(0, len(df)):
        curcol = startcol
        currow = currow + 1
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Primary Diagnosis'],
            df['description'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Code'],
            df['PRI_DIAG_CD'].values[i],
            table_body)
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Denominator'],
                         df['dx_' + measure_cd + '_den'].values[i],
                         table_body)
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Numerator'],
                         df['dx_' + measure_cd + '_num'].values[i],
                         table_body)
        curcol = itercol(ws,
                         currow,
                         curcol,
                         columns['Rate'],
                         df['dx_' + measure_cd + '_rate'].values[i],
                         table_body_pct2)
        #curcol = itercol(ws, currow, curcol, columns['QB Program Rate'], "", table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Denominator'],
            df['pct_total_den'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Numerator'],
            df['pct_total_num'].values[i],
            table_body_pct2)

    return currow + 2


def readm_top5(ws, df, category, startrow=6, startcol=0):
    df = df.copy(deep=True).where(df['pcr_den'] >= 5).dropna().sort_values(
        by=['risk_adjust_rate', 'pcr_den'], ascending=[True, False])
    df2 = df.copy(deep=True).where(df['pcr_den'] >= 5).dropna().sort_values(
        by=['risk_adjust_rate', 'pcr_den'], ascending=[False, True])

    categories = {
        'mdc': {
            'title': "Major Diagnostic Category",
            'codevar': 'CMN_EACMDC_CD'},
        'drg': {
            'title': "Diagnostic-Related Group",
            'codevar': 'CMN_EACDRG_CD'},
        'dx': {
            'title': "Diagnosis",
            'codevar': 'PRI_DIAG_CD'}}

    columns = {
        'Description': 5,
        'Code': 0,
        'Denominator': 1,
        'Numerator': 1,
        'Observed Rate': 1,
        'Expected Rate': 1,
        'Market Observed Rate': 1,
        'Risk-Adjusted Rate': 1,
        '% of Total Denominator': 1,
        '% of Total Numerator': 1
    }

    def header(titletext, startrow, startcol):
        currow = startrow + 1
        curcol = itercol(
            ws,
            currow,
            startcol,
            columns['Description'],
            'Description',
            headerwrap)
        #curcol = itercol(ws, currow, curcol, columns['Code'], 'Code', headerwrap)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Denominator'],
            'Denominator',
            headerwrap_num)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Numerator'],
            'Numerator',
            headerwrap_num)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Observed Rate'],
            'Observed Rate',
            headerwrap_num)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Expected Rate'],
            'Expected Rate',
            headerwrap_num)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Market Observed Rate'],
            'Market Rate',
            headerwrap_num)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Risk-Adjusted Rate'],
            'Risk-Adjusted Rate',
            headerwrap_num)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Denominator'],
            "% of Total Denominator",
            headerwrap_num)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Numerator'],
            "% of Total Numerator",
            headerwrap_num)

        ws.merge_range(
            currow - 1,
            startcol,
            currow - 1,
            curcol - 1,
            titletext,
            table_title)
        ws.set_row(currow, 45)

        return currow + 1

    if len(df) < 10:
        currow = header(categories[category]['title'], startrow, startcol)
        for i in range(0, len(df)):
            curcol = itercol(ws, currow, startcol, columns['Description'], str(
                df[category + '_description'].values[i]), table_body)
            #curcol = itercol(ws, currow, curcol, columns['Code'], df[categories[category]['codevar']].values[i], table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Denominator'],
                df['pcr_den'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Numerator'],
                df['pcr_num'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Observed Rate'],
                df['observed_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Expected Rate'],
                df['expected_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Market Observed Rate'],
                df['reg_observed_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Risk-Adjusted Rate'],
                df['risk_adjust_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['% of Total Denominator'],
                df['pct_tot_den'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['% of Total Numerator'],
                df['pct_tot_num'].values[i],
                table_body_pct2)
            currow = currow + 1
    else:
        currow = header(
            "Top 5 by " +
            categories[category]['title'],
            startrow,
            startcol)
        for i in range(0, 5):
            curcol = itercol(ws, currow, startcol, columns['Description'], str(
                df[category + '_description'].values[i]), table_body)
            #curcol = itercol(ws, currow, curcol, columns['Code'], df[categories[category]['codevar']].values[i], table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Denominator'],
                df['pcr_den'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Numerator'],
                df['pcr_num'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Observed Rate'],
                df['observed_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Expected Rate'],
                df['expected_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Market Observed Rate'],
                df['reg_observed_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Risk-Adjusted Rate'],
                df['risk_adjust_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['% of Total Denominator'],
                df['pct_tot_den'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['% of Total Numerator'],
                df['pct_tot_num'].values[i],
                table_body_pct2)
            currow = currow + 1
        currow = currow + 1
        currow = header(
            "Bottom 5 by " +
            categories[category]['title'],
            currow,
            startcol)
        for i in range(0, 5):
            curcol = itercol(ws, currow, startcol, columns['Description'], str(
                df2[category + '_description'].values[i]), table_body)
            #curcol = itercol(ws, currow, curcol, columns['Code'], df[categories[category]['codevar']].values[i], table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Denominator'],
                df2['pcr_den'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Numerator'],
                df2['pcr_num'].values[i],
                table_body)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Observed Rate'],
                df2['observed_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Expected Rate'],
                df2['expected_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Market Observed Rate'],
                df2['reg_observed_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['Risk-Adjusted Rate'],
                df2['risk_adjust_rate'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['% of Total Denominator'],
                df2['pct_tot_den'].values[i],
                table_body_pct2)
            curcol = itercol(
                ws,
                currow,
                curcol,
                columns['% of Total Numerator'],
                df2['pct_tot_num'].values[i],
                table_body_pct2)
            currow = currow + 1

    return currow + 1


def hosp21_drivers(ws, df, startrow=6, startcol=0):

    columns = {
        'Condition': 5,
        'Denominator': 1,
        'Numerator': 1,
        'Rate': 0,
        'QB Program Rate': 1,
        'Similar Size Rate': 1,
        'Tier': 0,
        '% of Total Denominator': 1,
        '% of Total Numerator': 1
    }

    currow = startrow + 1
    ws.set_row(currow, 45)

    #ws.write(currow, startcol, "Class", header)
    #ws.merge_range(currow, startcol+1, currow, startcol+5, "Measure", header)
    curcol = itercol(
        ws,
        currow,
        startcol,
        columns['Condition'],
        "Episode DRG",
        headerwrap)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Denominator'],
        "Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Numerator'],
        "Numerator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['Rate'],
        "Rate",
        headerwrap_num)
    #curcol = itercol(ws, currow, curcol, columns['QB Program Rate'], "QB Program Rate", headerwrap_num)
    #curcol = itercol(ws, currow, curcol, columns['Similar Size Rate'], "Similar Size Rate", headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Denominator'],
        "% of Total Denominator",
        headerwrap_num)
    curcol = itercol(
        ws,
        currow,
        curcol,
        columns['% of Total Numerator'],
        "% of Total Numerator",
        headerwrap_num)

    ws.merge_range(
        startrow,
        startcol,
        startrow,
        curcol - 1,
        "KEY DRIVERS",
        table_title)

    for i in range(0, len(df)):
        curcol = startcol
        currow = currow + 1
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Condition'],
            df['description'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Denominator'],
            df['hosp21_den'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Numerator'],
            df['hosp21_num'].values[i],
            table_body)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['Rate'],
            df['hosp21_rate'].values[i],
            table_body_pct2)
        #curcol = itercol(ws, currow, curcol, columns['QB Program Rate'], "", table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Denominator'],
            df['pct_tot_den'].values[i],
            table_body_pct2)
        curcol = itercol(
            ws,
            currow,
            curcol,
            columns['% of Total Numerator'],
            df['pct_tot_num'].values[i],
            table_body_pct2)

    return currow + 2


def dennum_table(
        ws,
        msr,
        den=0,
        num=0,
        obs=0,
        exp=0,
        mkt=0,
        rate=0,
        startrow=6,
        startcol=0):

    if msr in ['rrama', 'rracomm']:
        ws.merge_range(
            startrow,
            startcol,
            startrow,
            startcol + 11,
            cqmbenchmarks[msr]['title'],
            table_title)
        ws.set_row(startrow + 1, 30)
        ws.merge_range(
            startrow + 1,
            startcol,
            startrow + 1,
            startcol + 1,
            'Denominator',
            header_num)
        ws.merge_range(
            startrow + 1,
            startcol + 2,
            startrow + 1,
            startcol + 3,
            'Numerator',
            header_num)
        ws.merge_range(
            startrow + 1,
            startcol + 4,
            startrow + 1,
            startcol + 5,
            'Observed Rate',
            header_num)
        ws.merge_range(
            startrow + 1,
            startcol + 6,
            startrow + 1,
            startcol + 7,
            'Expected Rate',
            header_num)
        ws.merge_range(
            startrow + 1,
            startcol + 8,
            startrow + 1,
            startcol + 9,
            'Market Expected Rate',
            headerwrap_num)
        ws.merge_range(
            startrow + 1,
            startcol + 10,
            startrow + 1,
            startcol + 11,
            'Risk-Adjusted Rate',
            header_num)
        ws.merge_range(
            startrow + 2,
            startcol,
            startrow + 2,
            startcol + 1,
            den,
            table_body)
        ws.merge_range(
            startrow + 2,
            startcol + 2,
            startrow + 2,
            startcol + 3,
            num,
            table_body)
        ws.merge_range(
            startrow + 2,
            startcol + 4,
            startrow + 2,
            startcol + 5,
            obs,
            table_body_pct2)
        ws.merge_range(
            startrow + 2,
            startcol + 6,
            startrow + 2,
            startcol + 7,
            exp,
            table_body_pct2)
        ws.merge_range(
            startrow + 2,
            startcol + 8,
            startrow + 2,
            startcol + 9,
            mkt,
            table_body_pct2)
        ws.merge_range(
            startrow + 2,
            startcol + 10,
            startrow + 2,
            startcol + 11,
            rate,
            table_body_pct2)
    else:
        ws.merge_range(
            startrow,
            startcol,
            startrow,
            startcol + 5,
            cqmbenchmarks[msr]['title'],
            table_title)
        ws.merge_range(
            startrow + 1,
            startcol,
            startrow + 1,
            startcol + 1,
            'Denominator',
            header_num)
        ws.merge_range(
            startrow + 1,
            startcol + 2,
            startrow + 1,
            startcol + 3,
            'Numerator',
            header_num)
        ws.merge_range(
            startrow + 1,
            startcol + 4,
            startrow + 1,
            startcol + 5,
            'Rate',
            header_num)
        ws.merge_range(
            startrow + 2,
            startcol,
            startrow + 2,
            startcol + 1,
            den,
            table_body)
        ws.merge_range(
            startrow + 2,
            startcol + 2,
            startrow + 2,
            startcol + 3,
            num,
            table_body)
        ws.merge_range(
            startrow + 2,
            startcol + 4,
            startrow + 2,
            startcol + 5,
            rate,
            table_body_pct2)

    return startrow + 4


def dates(month, year):
    if month < 12:
        claims_incurred = datetime.date(
            year, month + 1, 1) - datetime.timedelta(1)
        claims_paid = claims_incurred
    elif month == 12:
        claims_incurred = datetime.date(year, 12, 31)
        claims_paid = claims_incurred
    elif month > 12:
        claims_incurred = datetime.date(year, 12, 31)
        claims_paid = datetime.date(
            year + 1, month - 12 + 1, 1) - datetime.timedelta(1)

    claims_incurred = claims_incurred.strftime("%m/%d/%y")
    claims_paid = claims_paid.strftime("%m/%d/%y")
    print(claims_incurred, claims_paid)
    return {"Claims Incurred": claims_incurred, "Claims Paid": claims_paid}
