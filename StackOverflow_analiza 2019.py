import pandas as pd
import xlsxwriter
import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw


## Load the DATA and PERFORM BASIC DATA CLEANING:

df = pd.read_csv(r"PATH:\survey_results_public.csv")



# STEP 1. SELECT ONLY THE COLUMNS NEEDED for the analysis

data = df[["MainBranch", "Hobbyist","OpenSourcer","Employment","Country","EdLevel","UndergradMajor","EduOther","DevType",
           "YearsCode","YearsCodePro","FizzBuzz","CurrencyDesc","CompTotal","CompFreq","ConvertedComp","LanguageWorkedWith",
           "LanguageDesireNextYear","Age","Gender","Ethnicity",]]

# STEP 2. RENAME THE COLUMNS for easier use; ASSIGN NEW COL LABELS TO DATAFRAME

new_cols = ['main_branch','hobbyist','open_sourcer','employment','country','ed_level','undergrad_major',
            'edu_other',"dev_type","years_code","years_code_pro","fizzbuzz","currency_desc","comp_total",
            "comp_freq","converted_comp","language_worked_with","lang_desire_next_year","age","gender","ethnicity"]

data.columns = new_cols

# STEP 3. CONVERT COUNTRY NAMES TO LOWERCASE

languages = data["language_worked_with"].astype(str)
data.loc[:, "country"] = data["country"].str.lower()
countries_unique = data["country"].unique()

  
# # # ********** PART I - LANGUAGES BY COUNTRY ********** # # #
#   Goal: Extract all the user language entries to make a list of all languages and number of respondents who use them   #

# STEP 1. LOOP THAT RETURNS A DICTIONARY with LANGUAGE NAME AS KEY and 0 AS VALUE, to be used for merging series into ONE DataFrame
languages_dict = {}
for lang in languages:
    lang = lang.split(";")
    for i in lang:
        languages_dict[i] = languages_dict.get(i, 0) + 1
individual_lang = languages_dict.fromkeys(languages_dict,0)
individual_lang = pd.Series(languages_dict)


# STEP 1a. ADDING SELECTED COUNTRIES LANGUAGE USAGE VALUES (NUMBER OF RESPONDENTS) TO ONE DATAFRAME
lang_by_country = pd.DataFrame(individual_lang)
lang_by_country.rename({ 0 :"serbia"},axis =1,inplace=True)

# STEP 2. CREATING A FUNCTION that will AGGREGATE LANGUAGE_WORKED_WITH COLUMN DATA BY COUNTRY and EXTRACT INDIVIDUAL LANGUAGES into SEPARATE DICTIONARIES
#         for the selected countries and CONVERT them to a PANDAS SERIES class  
def lang(cty):
    cty_dict = {}
    cty = (data["country"] == cty)
    cty_lang = data.loc[cty, "language_worked_with"].dropna()
    for s in cty_lang:
        s = s.split(';')
        for i in s:
            cty_dict[i] = cty_dict.get(i,0) + 1
    cty_lang = pd.Series(cty_dict)
    return cty_lang

# STEP 2a. LOOP THAT RUNS THE FUNCTION IN STEP 2 for the selected countries CREATING NEW COLUMNS in the lang_by_country DATAFRAME

for c in countries_unique:
    if c == "germany":
        germany = lang(c)
        lang_by_country["germany"] = germany
    elif c == "romania":
        romania = lang(c)
        lang_by_country["romania"] = romania
    elif c == "united states":
        usa = lang(c)
        lang_by_country["usa"] = usa
    elif c == "ukraine":
        ukraine = lang(c)
        lang_by_country["ukraine"] = ukraine
    elif c == "austria":
        austria = lang(c)
        lang_by_country["austria"] = austria
    elif c == "croatia":
        croatia = lang(c)
        lang_by_country["croatia"] = croatia
    elif c == "serbia":
        serbia = lang(c)
        lang_by_country["serbia"] = serbia

lang_by_country = lang_by_country.drop("nan", axis=0)



# Number of Respondents for percentage calculations   
countries = data["country"]
respondents_dict = {}
count = 0 

for c in countries:
    respondents_dict[c] = respondents_dict.get(c,0) + 1
respondents = pd.Series(respondents_dict)
sorted_resp = respondents.sort_values(ascending= False)
sorted_resp.rename({'0' : "Number of Respondents"})


# # # ********** PART Ia - TOP 10 LANGUAGES ********** # # #


# STEP 1. Find 10 largest values in individual countries of the lang_by_country dataframe and assign to separate series
top10_lang_serbia = lang_by_country["serbia"].nlargest(n=10) 
top10_lang_ukraine = lang_by_country["ukraine"].nlargest(n=10)
top10_lang_romania = lang_by_country["romania"].nlargest(n=10)
top10_lang_germany = lang_by_country["germany"].nlargest(n=10)
top10_lang_usa = lang_by_country["usa"].nlargest(n=10)
top10_lang_austria = lang_by_country["austria"].nlargest(n=10)
top10_lang_croatia = lang_by_country["croatia"].nlargest(n=10)

# STEP 1a. WORKAROUND - Assign individual Series to individual dataframes so that the top 10 languages can be assigned to a COUNTRY KEY in STEP 2
top10_srb = pd.DataFrame(top10_lang_serbia)
top10_de = pd.DataFrame(top10_lang_germany)
top10_ro = pd.DataFrame(top10_lang_romania) 
top10_ukr = pd.DataFrame(top10_lang_ukraine) 
top10_us = pd.DataFrame(top10_lang_usa) 
top10_at = pd.DataFrame(top10_lang_austria) 
top10_cro = pd.DataFrame(top10_lang_croatia) 

# STEP 2 - to LABEL the countries' TOP 10 a COUNTRY CODE KEY is used
frames = [top10_srb["serbia"],top10_de["germany"], top10_ro["romania"], top10_ukr["ukraine"],top10_us["usa"],top10_at["austria"],top10_cro["croatia"]]
top10_all_concat = pd.concat(frames, keys=['SRB','DE','RO','UKR','US','AT','CRO'])


# VISUALIZATION OF % SHARE AMONG RESONDENTS (STACKED BAR PLOT)

# STEP 1. PERCENTAGE CALCULATION for individual countries
top10_percentage = pd.DataFrame(top10_lang_serbia) / sorted_resp.loc["serbia"].round(decimals=2) * 100
top10_percentage["germany"] = pd.DataFrame(top10_lang_germany) / sorted_resp.loc["germany"].round(decimals=2) * 100
top10_percentage["romania"] = pd.DataFrame(top10_lang_romania) / sorted_resp.loc["romania"].round(decimals=2) * 100
top10_percentage["ukraine"] = pd.DataFrame(top10_lang_ukraine) /sorted_resp.loc["ukraine"].round(decimals=2) * 100
top10_percentage["usa"] = pd.DataFrame(top10_lang_usa) / sorted_resp.loc["united states"].round(decimals=2) * 100
top10_percentage["austria"] = pd.DataFrame(top10_lang_austria) / sorted_resp.loc["austria"].round(decimals=2) * 100
top10_percentage["croatia"] = pd.DataFrame(top10_lang_croatia) / sorted_resp.loc["croatia"].round(decimals=2) * 100
top10_percentage = top10_percentage.round(decimals=0)


# STEP 2. - STACKED BAR PLOT
        # transpose the data so that countrie names are turned into row labels
bar_data_trans = top10_percentage.transpose()
bar_data = bar_data_trans[["Python","JavaScript","SQL","Java","HTML/CSS","PHP","Bash/Shell/PowerShell"]]

fig2 = plt.figure(figsize=(7,6))
N= 7 
ind = np.arange(7)
width = 0.35

bar1 = fig2.add_subplot(1,1,1)

p1 = bar1.bar(ind, bar_data["Python"], width,color='cornflowerblue')
p2 = bar1.bar(ind, bar_data["Java"], width, bottom=bar_data["Python"], color='gold' )
p3 = bar1.bar(ind, bar_data["JavaScript"], width, bottom=bar_data["Java"], color='orange' )
p4 = bar1.bar(ind, bar_data["SQL"], width, bottom=bar_data["JavaScript"],color='tab:olive')

plt.xticks(ind, ('Serbia', 'Germany','Romania','Ukraine','USA','Austria','Croatia'), rotation=60)
bar1.tick_params(pad=15)
plt.legend((p1[0], p2[0], p3[0], p4[0]), ('Python','Java', 'JavaScript','SQL'),loc=1)
bar1.set_title("% Share Among Respondents", pad=12)
# bar1.set_xlabel("SAMPLE_DATA")

plt.show()


# # # ********** PART Ib - TOP 5 LANGUAGES VISUALIZATION in the SELECTED COUNTRIES ********** # # #

# STEP 1. DIVIDE TOP 5 VALUES from TOP10 DATAFRAMES with NUMBER OF RESPONDENTS and multiplies by 1000 to NORLAMIZE DATA.
#         The LOOP EXTRACTS ROW (LANGUAGE) LABELS for the country CREATING a LIST OF TOP 5 in that country

# Horizontal BAR Top 5 - #1

srb_labels = []
top5_srb_norm = top10_srb.head(5) / sorted_resp.loc["serbia"].round(decimals=2) * 1000 # normalized to top5 languages per 1000 respondents
for row in top5_srb_norm.index:
    srb_labels.append(row)

bar_width_srb = top5_srb_norm.iloc[:,0].astype(int)
bar_positions = np.arange(5) + 1.0
ytick_positions = range(1,6)

fig = plt.figure(figsize = (30,20))
plt.style.use('fivethirtyeight')

ax_srb = fig.add_subplot(3,3,1)
ax_srb.barh(bar_positions, bar_width_srb[::-1], 0.5)
ax_srb.set_yticks(ytick_positions)
ax_srb.set_yticklabels(srb_labels[::-1])


plt.yticks(fontsize=8)
plt.xticks(fontsize=7)
plt.xlabel("Language Usage per 1000 Respondents", fontsize=10, fontstyle='italic', labelpad = 10)
plt.ylabel("")
plt.title("Top 5 Languages - SERBIA", fontsize=12, pad=10)
# print(bar_width_srb.iloc[::-1])
# print(srb_labels)

# Horizontal BAR Top 5 - #2

de_labels = []
top5_de_norm = top10_de.head(5) / sorted_resp.loc["germany"].round(decimals=2) * 1000 # normalized to top5 languages per 1000 respondents
for row in top5_de_norm.index:
    de_labels.append(row)
    
bar_width_de = top5_de_norm.iloc[:,0].astype(int)
bar_positions = np.arange(5) + 1.0
ytick_positions = range(1,6)

ax_de = fig.add_subplot(3,3,2)
ax_de.barh(bar_positions, bar_width_de[::-1], 0.5)
ax_de.set_yticks(ytick_positions)
ax_de.set_yticklabels(de_labels[::-1])


plt.yticks(fontsize=8)
plt.xticks(fontsize=7)
plt.xlabel("Language Usage per 1000 Respondents", fontsize=10, fontstyle='italic', labelpad = 10)
plt.ylabel("")
plt.title("Top 5 Languages - GERMANY", fontsize=12, pad=10)
# print(bar_width_de.iloc[::-1])
# print(de_labels)

# Horizontal BAR Top 5 - #3

ro_labels = []
top5_ro_norm = top10_ro.head(5) / sorted_resp.loc["romania"].round(decimals=2) * 1000 # normalized to top5 languages per 1000 respondents
for row in top5_ro_norm.index:
    ro_labels.append(row)
  
bar_width_ro = top5_ro_norm.iloc[:,0].astype(int)
bar_positions = np.arange(5) + 1.0
ytick_positions = range(1,6)

ax_ro = fig.add_subplot(3,3,3)
ax_ro.barh(bar_positions, bar_width_ro[::-1], 0.5)
ax_ro.set_yticks(ytick_positions)
ax_ro.set_yticklabels(ro_labels[::-1])


plt.yticks(fontsize=8)
plt.xticks(fontsize=7)
plt.xlabel("Language Usage per 1000 Respondents", fontsize=10, fontstyle='italic', labelpad = 10)
plt.ylabel("")
plt.title("Top 5 Languages - ROMANIA", fontsize=12, pad=10)
# print(bar_width_ro.iloc[::-1])
# print(de_labels)

# Horizontal BAR Top 5 - #4

ukr_labels = []
top5_ukr_norm = top10_ukr.head(5) / sorted_resp.loc["ukraine"].round(decimals=2) * 1000 # normalized to top5 languages per 1000 respondents
for row in top5_ukr_norm.index:
    ukr_labels.append(row)
   
bar_width_ukr = top5_ukr_norm.iloc[:,0].astype(int)
bar_positions = np.arange(5) + 1.0
ytick_positions = range(1,6)

ax_ukr = fig.add_subplot(3,3,4)
ax_ukr.barh(bar_positions, bar_width_ukr[::-1], 0.5)
ax_ukr.set_yticks(ytick_positions)
ax_ukr.set_yticklabels(ukr_labels[::-1])


plt.yticks(fontsize=8)
plt.xticks(fontsize=7)
plt.xlabel("Language Usage per 1000 Respondents", fontsize=10, fontstyle='italic', labelpad = 10)
plt.ylabel("")
plt.title("Top 5 Languages - UKRAINE", fontsize=12, pad=10)
# print(bar_width_ukr.iloc[::-1])
# print(ukr_labels)

# Horizontal BAR Top 5 - #5

us_labels = []
top5_us_norm = top10_us.head(5) / sorted_resp.loc["united states"].round(decimals=2) * 1000 # normalized to top5 languages per 1000 respondents
for row in top5_us_norm.index:
    us_labels.append(row)
    
bar_width_us = top5_us_norm.iloc[:,0].astype(int)
bar_positions = np.arange(5) + 1.0
ytick_positions = range(1,6)

ax_us = fig.add_subplot(3,3,5)
ax_us.barh(bar_positions, bar_width_us[::-1], 0.5)
ax_us.set_yticks(ytick_positions)
ax_us.set_yticklabels(us_labels[::-1])


plt.yticks(fontsize=8)
plt.xticks(fontsize=7)
plt.xlabel("Language Usage per 1000 Respondents", fontsize=10, fontstyle='italic', labelpad = 10)
plt.ylabel("")
plt.title("Top 5 Languages - UNITED STATES", fontsize=12, pad=10)

# Horizontal BAR Top 5 - #6

at_labels = []
top5_at_norm = top10_at.head(5) / sorted_resp.loc["austria"].round(decimals=2) * 1000 # normalized to top5 languages per 1000 respondents
for row in top5_at_norm.index:
    at_labels.append(row)
    
bar_width_at = top5_at_norm.iloc[:,0].astype(int)
bar_positions = np.arange(5) + 1.0
ytick_positions = range(1,6)

ax_at = fig.add_subplot(3,3,6)
ax_at.barh(bar_positions, bar_width_at[::-1], 0.5)
ax_at.set_yticks(ytick_positions)
ax_at.set_yticklabels(at_labels[::-1])


plt.yticks(fontsize=8)
plt.xticks(fontsize=7)
plt.xlabel("Language Usage per 1000 Respondents", fontsize=10, fontstyle='italic', labelpad = 10)
plt.ylabel("")
plt.title("Top 5 Languages - AUSTRIA", fontsize=12, pad=10)

# Horizontal BAR Top 5 - #7

cro_labels = []
top5_cro_norm = top10_cro.head(5) / sorted_resp.loc["croatia"].round(decimals=2) * 1000 # normalized to top5 languages per 1000 respondents
for row in top5_cro_norm.index:
   cro_labels.append(row)
    

bar_width_cro = top5_cro_norm.iloc[:,0].astype(int)
bar_positions = np.arange(5) + 1.0
ytick_positions = range(1,6)

ax_cro = fig.add_subplot(3,3,7)
ax_cro.barh(bar_positions, bar_width_cro[::-1], 0.5)
ax_cro.set_yticks(ytick_positions)
ax_cro.set_yticklabels(cro_labels[::-1])


plt.yticks(fontsize=8)
plt.xticks(fontsize=7)
plt.xlabel("Language Usage per 1000 Respondents", fontsize=10, fontstyle='italic', labelpad = 10)
plt.ylabel("")
plt.title("Top 5 Languages - CROATIA", fontsize=12, pad=10)
plt.show()


# # ************************************************************************************************** # #



# # # ********** PART II - FUTURE LEARNING TRENDS ********** # # #

        # 84088 non-null lang_desire_next_year
       
des_languages = data["lang_desire_next_year"].astype(str)

# STEP 1. Create a DICTIONARY to serve as an EMPTY TEMPLATE with all LANGUAGES AS ROW LABELS and CONVERT to a PANDAS SERIES
des_languages_dict = {}
for lang in des_languages:
    lang = lang.split(";")
    for i in lang:
        des_languages_dict[i] = des_languages_dict.get(i, 0) + 1

desired_languages = pd.Series(des_languages_dict, name="Desired Language Next Year - Top 20")

# STEP 2. CONVERT SERIES TO DATAFRAME AND CALCULATE THE PERCENTAGES

desired_lang_df = pd.DataFrame(desired_languages.nlargest(n=20))
desired_df_perc = desired_languages.nlargest(n=20) / 84088 * 100

desired_lang_df["Desired Language Next Year - %"] = desired_df_perc.round(decimals=0).astype(int)


# # ************************************************************************************************** # #



# # # ********** PART III - DEVELOPER TYPES ********** # # #

# STEP 1. Create a DICTIONARY to serve as A template with all DEVELOPER TYPE as row LABELS and CONVERT to SERIES then DATAFRAME
devtype_clean = data["dev_type"].dropna(axis=0, inplace=True)
devtype_dict = {}

for line in data["dev_type"]:
    line = line.split(';')
    for devt in line:
        devtype_dict[devt] = devtype_dict.get(devt,0) + 1
    
devtype = pd.Series(devtype_dict, name='Respondents ALL')
# print(devtype.sort_values(ascending=False))

devtype_key = devtype_dict.fromkeys(devtype_dict,0)
devtype_base_ser = pd.Series(devtype_key)
devtype_df = pd.DataFrame(devtype_base_ser)

devtype_df.rename({0 : 'serbia'}, axis=1, inplace=True)



# # # ********** PART IIIa - DEVELOPER TYPES BY COUNTRY ********** # # #

# STEP 1. CREATING A FUNCTION that will AGGREGATE DEV_TYPE COLUMN DATA BY COUNTRY and EXTRACT DEVELOPER TYPE VALUES into SEPARATE DICTIONARIES for
#         the selected countries and CONVERT them TO A PANDAS SERIES 

def dev_by_cty(devcty):
    dev_cty_dict = {}
    devcty = (data["country"] == devcty)
    devcty_type = data.loc[devcty, "dev_type"]
    for d in devcty_type:
        d = d.split(';')
        for t in d:
            dev_cty_dict[t] = dev_cty_dict.get(t,0)+1
    devcty_type = pd.Series(dev_cty_dict)
    return devcty_type
    
# STEP 1a. LOOP THAT RUNS THE FUNCTION IN STEP 2 for the selected countries CREATING NEW COLUMNS in the devtype_df DATAFRAME   
for c in countries_unique:
    if c == "germany":
        germany_dev = dev_by_cty(c)
        devtype_df["germany"] = germany_dev
    elif c == "romania":
        romania_dev = dev_by_cty(c)
        devtype_df["romania"] = romania_dev
    elif c == "united states":
        usa_dev = dev_by_cty(c)
        devtype_df["usa"] = usa_dev
    elif c == "ukraine":
        ukraine_dev = dev_by_cty(c)
        devtype_df["ukraine"] = ukraine_dev
    elif c == "austria":
        austria_dev = dev_by_cty(c)
        devtype_df["austria"] = austria_dev
    elif c == "croatia":
        croatia_dev = dev_by_cty(c)
        devtype_df["croatia"] = croatia_dev
    elif c == "serbia":
        serbia_dev = dev_by_cty(c)
        devtype_df["serbia"] = serbia_dev



# # # ********** PART IIIb - DEVELOPER TYPES BY COUNTRY PERCENTAGES ********** # # #
      
        # # number of rows - 81335
# STEP 1. LOOP THAT AGGREGATES DEV_TYPE DATA BY SELECTED COUNTRIES THEN CONVERTS TO SERIES
countries_lst = ["serbia","germany","romania","ukraine","united states", "austria", "croatia"]
dev_respondents = {}

for c in countries_lst:
    dev_bool_cty = (data["country"] == c)
    dev_cty = data.loc[dev_bool_cty,"dev_type"].value_counts().sum()
    dev_respondents[c] = dev_respondents.get(c,dev_cty)

devtype_resp = pd.Series(dev_respondents, name= "DevType Respondents")
devtype_resp = devtype_resp.rename({"united states" : "usa"})
   

# STEP 2. PERCENTAGE CALCULATION - LOOP THAT RUNS COUNTRY NAMES and CREATES NEW COLUMNS IN AN EMPTY DATAFRAME by DIVIDING devtype_df COUNTRY VALUES by 
#         the NUMBER OF TOTAL RESPONDENTS entries in the given country (devtype_resp)

devtype_by_cty = pd.DataFrame()
for i in devtype_df.columns:
    devtype_by_cty[i + " %"] = (devtype_df[i] / devtype_resp[i]) * 100
   
devtype_by_cty = devtype_by_cty.round(decimals=2)

# STEP 2. TOTAL RESPONDENT PERCENTAGE CALCULATION (used for DEV TYPE ALL Visualization)
devtype_perc = devtype.sort_values() / 81335 * 100
devtype_perc = devtype_perc.rename({"Respondents ALL" : "Respondents % Share"})
#print(devtype_perc, devtype_by_cty)


# HORIZONTAL BAR PLOT - DEV TYPE ALL Visualization

dt_bar, ax = plt.subplots(figsize=(6.62,5.2))
plt.style.use('fivethirtyeight')

tick_positions = range(1,25)
ax.set_yticks(tick_positions)

ytick_labels = ['Developer, full-stack', 'Developer, back-end', 'Developer, front-end',
       'Developer, desktop or enterprise applications', 'Developer, mobile',
       'Student', 'Database administrator', 'Designer', 'System administrator',
       'DevOps specialist', 'Developer, embedded applications or devices',
       'Data scientist or machine learning specialist',
       'Developer, QA or test', 'Data or business analyst',
       'Academic researcher', 'Engineer, data', 'Educator',
       'Developer, game or graphics', 'Engineering manager', 'Product manager',
       'Scientist', 'Engineer, site reliability', 'Senior executive/VP',
       'Marketing or sales professional']


ax.set_yticklabels(ytick_labels[::-1], fontsize=7.5)
xtick_labels = ['0%', '10%', '20%', '30%', '40%', '50%']
ax.set_xticklabels(xtick_labels)

bar_positions = np.arange(24) + 1.0
bar_width = devtype_perc.values.round(decimals=2)
ax.barh(bar_positions, bar_width, 0.5, color='darkblue')

ax.set_title("Dev Type", fontsize=12, pad=12, fontstyle='oblique')
ax.set_xlabel("Developers %", fontsize=11, labelpad=8, fontstyle='italic')

plt.show()


# # ************************************************************************************************** # #


# # # ********** PART IV - EDUCATION LEVEL IN SELECTED COUNTRIES ********** # # #

# STEP 1. CREATE A FUNCTION that AGGREGATES ED_LEVEL DATA BY COUNTRY, CREATES A DICTIONARY with EDUCATION TYPE as ROW LABELS CONVERTING TO A SERIES
def ed_level(ed):
    ed_level = {}
    selected_rows_bool = (data["country"] == ed)
    selected_rows = data.loc[selected_rows_bool,"ed_level"] 
    for s in selected_rows:
        ed_level[s] = ed_level.get(s,0)+1
    ed_level_ser = pd.Series(ed_level)
    return ed_level_ser
 
# STEP 1a. LOOP THAT RUNS THE FUNCTION IN STEP 1 for the selected countries CREATING NEW COLUMNS in the ed_level_df DATAFRAME    
ed_level_df = pd.DataFrame()   
for c in countries_unique:
    if c == "germany":
        germany_ed = ed_level(c)
        ed_level_df["germany"] = germany_ed
    elif c == "romania":
        romania_ed = ed_level(c)
        ed_level_df["romania"] = romania_ed
    elif c == "united states":
        usa_ed = ed_level(c)
        ed_level_df["usa"] = usa_ed
    elif c == "ukraine":
        ukraine_ed = ed_level(c)
        ed_level_df["ukraine"] = ukraine_ed
    elif c == "austria":
        austria_ed = ed_level(c)
        ed_level_df["austria"] = austria_ed
    elif c == "croatia":
        croatia_ed = ed_level(c)
        ed_level_df["croatia"] = croatia_ed
    elif c == "serbia":
        serbia_ed = ed_level(c)
        ed_level_df["serbia"] = serbia_ed


# STEP 2. LOOP THAT CREATES A DICTIONARY with NUMBER OF RESPONDENTS ENTRIES FOR EDUCATION LEVEL in the SELECTED COUNTRIES
edlvl_resp = {}
for c in countries_lst:
    edlvl_bool_cty = (data["country"] == c)
    ed_cty = data.loc[edlvl_bool_cty,"ed_level"].dropna().value_counts().sum() 
    edlvl_resp[c] = edlvl_resp.get(c,ed_cty)

# STEP 2a. CONVERTS edlvl_resp DICTIONARY TO A PANDAS SERIES
edlevel_resp = pd.Series(edlvl_resp, name= "Education Respondents")
edlevel_resp = edlevel_resp.rename({"united states" : "usa"})

# STEP 2b. REMOVE the INDEX LABELS THAT ARE NOT REPRESENTED IN THE PIE CHARTS
ed_level_pie_df = ed_level_df.drop([np.nan,"Associate degree", "Professional degree (JD, MD, etc.)", "I never completed any formal education"])


# STEP 3. PERCENTAGE CALCULATION - LOOP THAT RUNS COUNTRY NAMES and CREATES NEW COLUMNS IN AN EMPTY DATAFRAME by DIVIDING ed_level_df COUNTRY VALUES by 
#         the NUMBER OF TOTAL RESPONDENTS entries in the given country (edlevel_resp)
edlevel_by_cty = pd.DataFrame()

for i in ed_level_df.columns:
    edlevel_by_cty[i + " %"] = (ed_level_df[i] / edlevel_resp[i]) * 100
   
edlevel_by_cty = edlevel_by_cty.round(decimals=1)


# # Pie FIGURE #1 - PIE CHART

textprops = {"fontsize":9}
colors = ['#008fd5', 'tab:red', 'gold', 'tab:green', 'tab:gray', 'tab:purple']

piefig1 = plt.figure(figsize=(6.60,3.69))
plt.style.use('fivethirtyeight')
srb = piefig1.add_subplot(1,1,1)

labels = ["Bachelor's Degree", "Master's Degree", "Some University, no Degree",
          "Other Doctoral Degree", "Secondary School", "Elementary School"]  
 
       
srb_slices = [ed_level_pie_df.loc[:, "serbia"].values ]
explode = [0.1, 0, 0, 0, 0, 0]
srb.pie(srb_slices, labels = labels, shadow=True, explode=explode, colors=colors, startangle=80, 
        wedgeprops={'edgecolor':'black'}, autopct = '%1.1f%%',textprops=textprops)

srb.set_title("Education Level - Serbia", fontsize=15)

# # Pie FIGURE #2 - PIE CHART

plt.tight_layout()

piefig2 = plt.figure(figsize=(6.60,3.69))
plt.style.use('fivethirtyeight')
de = piefig2.add_subplot(1,1,1)


de_slices = [ed_level_pie_df.loc[:, "germany"].values ]
explode = [0, 0.1, 0, 0, 0, 0]
de.pie(de_slices, labels = labels, shadow=True, explode=explode, colors=colors, startangle=90, 
       wedgeprops={'edgecolor':'black'}, autopct = '%1.1f%%',textprops=textprops)

de.set_title("Education Level - Germany", fontsize=15)

plt.tight_layout()
plt.show()

# # Pie FIGURE #3 - PIE CHART

piefig3 = plt.figure(figsize=(6.60,3.69))
plt.style.use('fivethirtyeight')
ro = piefig3.add_subplot(1,1,1)  


ro_slices = [ed_level_pie_df.loc[:, "romania"].values ]
explode = [0.1, 0, 0, 0, 0, 0]
ro.pie(ro_slices, labels = labels, shadow=True, explode=explode, colors=colors, startangle=70,
       wedgeprops={'edgecolor':'black'}, autopct = '%1.1f%%',textprops=textprops)

ro.set_title("Education Level - Romania", fontsize=15)

plt.tight_layout()
plt.show()

# # Pie FIGURE #4 - PIE CHART

piefig4 = plt.figure(figsize=(6.60,3.69))
plt.style.use('fivethirtyeight')
ukr = piefig4.add_subplot(1,1,1)  


ukr_slices = [ed_level_pie_df.loc[:, "ukraine"].values ]
explode = [0, 0.1, 0, 0, 0, 0]
ukr.pie(ukr_slices, labels = labels, shadow=True, explode=explode,colors=colors, startangle=65, 
        wedgeprops={'edgecolor':'black'}, autopct = '%1.1f%%',textprops=textprops)

ukr.set_title("Education Level - Ukraine", fontsize=15)

plt.tight_layout()
plt.show()

# # Pie FIGURE #5- PIE CHART

piefig5 = plt.figure(figsize=(6.60,3.69))
plt.style.use('fivethirtyeight')
usa = piefig5.add_subplot(1,1,1)  


usa_slices = [ed_level_pie_df.loc[:, "usa"].values ]
explode = [0.1, 0, 0, 0, 0, 0]
usa.pie(usa_slices, labels = labels, shadow=True, explode=explode, colors=colors, startangle=40, 
        wedgeprops={'edgecolor':'black'}, autopct = '%1.1f%%',textprops=textprops)

usa.set_title("Education Level - United States", fontsize=15)

plt.tight_layout()
plt.show()

# # Pie FIGURE #6 - PIE CHART

piefig6 = plt.figure(figsize=(6.60,3.69))
plt.style.use('fivethirtyeight')
at = piefig6.add_subplot(1,1,1)  


at_slices = [ed_level_pie_df.loc[:, "austria"].values ]
explode = [0, 0.1, 0, 0, 0, 0]
at.pie(at_slices, labels = labels, shadow=True, explode=explode, colors=colors, startangle=75, 
       wedgeprops={'edgecolor':'black'}, autopct = '%1.1f%%',textprops=textprops)

at.set_title("Education Level - Austria", fontsize=15)

plt.tight_layout()
plt.show()

# # Pie FIGURE #7 - PIE CHART

piefig7 = plt.figure(figsize=(6.60,3.69))
plt.style.use('fivethirtyeight')
cro = piefig7.add_subplot(1,1,1)  


cro_slices = [ed_level_pie_df.loc[:, "croatia"].values ]
explode = [0, 0.1, 0, 0, 0, 0]
cro.pie(cro_slices, labels = labels, shadow=True, explode=explode, colors=colors, startangle=60, 
        wedgeprops={'edgecolor':'black'}, autopct = '%1.1f%%',textprops=textprops)

cro.set_title("Education Level - Croatia", fontsize=15)

plt.tight_layout()
plt.show()




# # *** DATA CLEANING FOR EXCEL ...

cols_empty = []
df_empty = pd.DataFrame(columns = cols_empty)

lang_by_country.columns = lang_by_country.columns.str.capitalize()
lang_by_country.rename({"Usa" : "USA"},axis=1, inplace=True)
top10_all_concat = top10_all_concat.rename("Number of Respondents", inplace=True)

sorted_resp.index = sorted_resp.index.str.upper()
sorted_resp = sorted_resp.rename("Number of Respondents", inplace=True)

devtype_df.columns = devtype_df.columns.str.capitalize()
devtype_df.rename({"Usa" : "USA"},axis=1, inplace=True)

devtype_by_cty.columns = devtype_by_cty.columns.str.capitalize()
devtype_by_cty.rename({"Usa %" : "USA %"},axis=1, inplace=True)



# # *******************  EXPORT OF DATA TO EXCEL ******************* # #


dfs = {"RAW DATA":data, "Languages by Country":lang_by_country, "Top 10 Languages":top10_all_concat, 
       "Top 5 by Country":df_empty, "Future Learning Trends":desired_lang_df,
       "Dev Type ALL" : devtype.sort_values(ascending=False), "Dev Type by Country" : devtype_df,
       "Education by Country":df_empty, "Respondents Total":sorted_resp}

filename = r"(enter PATH:test_data.xlsx)"


writer = pd.ExcelWriter(filename, engine='xlsxwriter')
for sheet_name in dfs.keys():
    dfs[sheet_name].to_excel(writer,sheet_name=sheet_name)

writer.save()

# *** XLWings - ADDING DataFrames and Images TO SPECIFIC CELLS *** #

sht = xw.Book(filename).sheets["Top 5 by Country"]
sht.pictures.add(fig, name = "Top 5 by Country", update=True)

sht2 = xw.Book(filename).sheets["Top 10 Languages"]
sht2.pictures.add(fig2, name= "Top 4 Share", update=True, left=sht.range('F2').left, top=sht.range('F2').top)

sht3 = xw.Book(filename).sheets["Dev Type by Country"]
sht3.range("K1").options(index=False).value = devtype_by_cty

sht4 = xw.Book(filename).sheets["Dev Type ALL"]
sht4.pictures.add(dt_bar, name = "Dev Type % Share", update=True, left=sht4.range('D1').left, top=sht4.range('D1').top)
         
sht5 = xw.Book(filename).sheets["Education by Country"]
sht5.pictures.add(piefig1, name = "Picture 1", update=True, left=sht.range('A1').left, top=sht.range('A1').top)                                                                                                   
sht5.pictures.add(piefig2, left=sht.range('K1').left, top=sht.range('K1').top)
sht5.pictures.add(piefig3, left=sht.range('U1').left, top=sht.range('U1').top)
sht5.pictures.add(piefig4, left=sht.range('A19').left, top=sht.range('A19').top)
sht5.pictures.add(piefig5, left=sht.range('K19').left, top=sht.range('K19').top)
sht5.pictures.add(piefig6, left=sht.range('U19').left, top=sht.range('U19').top)
sht5.pictures.add(piefig7, left=sht.range('A37').left, top=sht.range('A37').top)

# # ************************ # #
