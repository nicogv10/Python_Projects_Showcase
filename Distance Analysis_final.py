import numpy as np
import pandas as pd
import geopy.distance

#here I am reading in the pt list & facility list csv files
#FILE PATHS AND NAMES NEED TO BE COPIED IN BETWEEN THE ' '
pt_file = pd.read_csv(r'\\patients_file_path\file.csv',
                      sep=",", encoding='mac_roman')
facility_file = pd.read_csv(r'\\facility_file_path\file.csv',
                            sep=",", encoding='mac_roman')

#here I am breaking apart the file by Payor
#ENTER PAYOR YOU'RE INTERESTED IN THE TEXT SPACE BELOW AFTER THE = and within " "
payor = "XYZ"
payor_pts = pt_file[pt_file['PAYOR'] == f"{payor}"]

#here we set our assumption for minutes traveled per mile (typically we use 2)
mins_per_mile = 2

#here I create a concat column in the facilities dataframe (longitude + latitude) that will allow us to remove
#duplicates (some facilities show up multiple times with slightly diff names)
facility_file['CONCAT'] = facility_file['latitude'].astype(str) +"-"+ facility_file['longitude'].astype(str)
facility_file.drop_duplicates(subset=['CONCAT'], inplace=True)

#here I am creating a list of DVA facilities
dva_facilities = facility_file[facility_file['ProviderGroup'] == 'DVA']
#here I am creating a list of competitor (non-DVA) facilities
non_dva_facilities = facility_file[facility_file['ProviderGroup'] != 'DVA']

#here I am going to find the closest DVA facility
#Create an empty dataframe to save the closest DVA facility for each pt
DVA = pd.DataFrame(columns=['Pt_MPI', 'Plan Type', 'Pt_coords', 'DVA Fac #', 'Distance'])

#Iterate over each patient
for index, row in payor_pts.iterrows():
    mpi = row['PATIENT_MPI']
    plan = row['PLAN_TYPE']
    lat, lon = row['Latitude'], row['Longitude']
    # Calculate distance to each DVA facility
    distances = []
    for _, row in dva_facilities.iterrows():
        fac_coords = (row['latitude'], row['longitude'])
        miles_to_dva_fac = geopy.distance.geodesic((lat, lon), fac_coords).miles
        dva_fac_nbr = row['Facility Nbr']
        distances.append((dva_fac_nbr, miles_to_dva_fac))

    #Sort distances in ascending order
    distances.sort(key=lambda x: x[1])

    #Save the closest facility and distance for the current patient
    DVA = DVA.append({
        'Pt_MPI': mpi,
        'Plan Type': plan,
        'Pt_coords': (lat, lon),
        'DVA Fac #': distances[0][0],
        'Distance': distances[0][1],
    }, ignore_index=True)


#here I am going to find the 1st, 2nd & 3rd closest competitor (nonDVA) facilities
#Create an empty dataframe to save the 1st, 2nd & 3rd closest nonDVA facs for each patient
nonDVA = pd.DataFrame(columns=['Pt_coords', '1st_nd', '2nd_nd', '3rd_nd', 'Distance_1', 'Distance_2', 'Distance_3',
                               'Provider_1', 'Provider_2', 'Provider_3', 'Stations_1', 'Stations_2', 'Stations_3',
                               '20_HEMO_1', '20_HEMO_2', '20_HEMO_3'])

#Iterate over each patient
for lat, lon in payor_pts[['Latitude', 'Longitude']].values:
    # Calculate distance to each nonDVA facility
    distances = []
    for _, row in non_dva_facilities.iterrows():
        fac_coords = (row['latitude'], row['longitude'])
        miles_to_nd = geopy.distance.geodesic((lat, lon), fac_coords).miles
        nd_fac = row['Facility']
        nd_provider = row['ProviderGroup']
        nd_stations = row['Stations']
        nd_20_hemo = row['20_HEMO']
        distances.append((nd_fac, miles_to_nd, nd_provider, nd_stations, nd_20_hemo))

    #Sort distances in ascending order
    distances.sort(key=lambda x: x[1])

    #Save the closest facilities and distances for the current patient
    nonDVA = nonDVA.append({
        'Pt_coords': (lat, lon),
        '1st_nd': distances[0][0],
        '2nd_nd': distances[1][0],
        '3rd_nd': distances[2][0],
        'Distance_1': distances[0][1],
        'Distance_2': distances[1][1],
        'Distance_3': distances[2][1],
        'Provider_1': distances[0][2],
        'Provider_2': distances[1][2],
        'Provider_3': distances[2][2],
        'Stations_1': distances[0][3],
        'Stations_2': distances[1][3],
        'Stations_3': distances[2][3],
        '20_HEMO_1': distances[0][4],
        '20_HEMO_2': distances[1][4],
        '20_HEMO_3': distances[2][4]
    }, ignore_index=True)

#this new merged df will have each pts MPI, plan type, coordinates, closest DVA Fac #, Distance to DVA fac, closest
#3 nonDVA Facs and distance to those facs
DVA_and_nonDVA = DVA.merge(nonDVA, on="Pt_coords")
DVA_and_nonDVA.drop_duplicates(subset=['Pt_MPI'], inplace=True)
#we will output this dataframe as one of the sheets in the Excel file

#here I will add more columns to our final output Excel file - the first determines whether the DVA facility is the closest
#option for each pt
DVA_and_nonDVA['DVA Closest?'] = np.where(DVA_and_nonDVA['Distance'].values <= DVA_and_nonDVA['Distance_1'].values, 'Y', 'N')
#the next few columns represent estimated drive time to each facility (using our minutes per mile assumption)
DVA_and_nonDVA['Drive time to DVA'] = DVA_and_nonDVA['Distance']*mins_per_mile
DVA_and_nonDVA['Time_nd_1'] = DVA_and_nonDVA['Distance_1']*mins_per_mile
DVA_and_nonDVA['Time_nd_2'] = DVA_and_nonDVA['Distance_2']*mins_per_mile
DVA_and_nonDVA['Time_nd_3'] = DVA_and_nonDVA['Distance_3']*mins_per_mile
#the following columns calculate the time difference between driving to the closest DVA facility vs competitor facilities
DVA_and_nonDVA['Time_diff_1'] = DVA_and_nonDVA['Time_nd_1'] - DVA_and_nonDVA['Drive time to DVA']
DVA_and_nonDVA['Time_diff_2'] = DVA_and_nonDVA['Time_nd_2'] - DVA_and_nonDVA['Drive time to DVA']
DVA_and_nonDVA['Time_diff_3'] = DVA_and_nonDVA['Time_nd_3'] - DVA_and_nonDVA['Drive time to DVA']

#these next few columns will guide our facility capacity calculations...the logic is as follows: if a patient is less
#than 30 minutes incremental/extra drive time to all 3 nonDVA facilities, then we assume they are just as likely/can just
#as conveniently treat at all 3, so we assign 1/3 patient to all 3. If a patient is less than 30 minutes extra drive to
#the 2 closest nonDVA facilities (but >30 minutes to the 3rd), then we assign 1/2 patient to the 2 closest. Otherwise (if
#the 2nd and 3rd closest nonDVA facilities are >30 minutes extra drive time), we assign 1 full patient to the 1st closest
#nonDVA facility
conditions_count_1 = [
    DVA_and_nonDVA['Time_diff_3'] < 30,
    DVA_and_nonDVA['Time_diff_2'] < 30,
    (DVA_and_nonDVA['Time_diff_2'] >= 30) & (DVA_and_nonDVA['Time_diff_3'] >= 30)
]

values_count_1 = [1/3, 1/2, 1]

DVA_and_nonDVA['Pt count_1'] = np.select(conditions_count_1, values_count_1, default='Unknown')

conditions_count_2 = [
    DVA_and_nonDVA['Time_diff_3'] < 30,
    DVA_and_nonDVA['Time_diff_2'] < 30,
    (DVA_and_nonDVA['Time_diff_2'] >= 30) & (DVA_and_nonDVA['Time_diff_3'] >= 30)
]

values_count_2 = [1/3, 1/2, 0]

DVA_and_nonDVA['Pt count_2'] = np.select(conditions_count_2, values_count_2, default='Unknown')

DVA_and_nonDVA['Pt count_3'] = np.where(DVA_and_nonDVA['Time_diff_3'].values < 30, 1/3, 0)

#now I am going to create a new dataframe by breaking the data apart for the 3 non-DVA facs and stacking them, which will
#later allow me to pivot the data to determine how many new pts are mapping to each nonDVA facility
nonDVA_1 = DVA_and_nonDVA[['1st_nd', 'Provider_1', 'Pt count_1', 'Stations_1', '20_HEMO_1']].copy()
dict_1 = {'1st_nd':'Non DVA Fac', 'Provider_1':'Provider', 'Pt count_1':'Pt_count', 'Stations_1':'Stations', '20_HEMO_1':'20_HEMO'}
nonDVA_1.rename(columns=dict_1, inplace=True)

nonDVA_2 = DVA_and_nonDVA[['2nd_nd', 'Provider_2', 'Pt count_2', 'Stations_2', '20_HEMO_2']].copy()
dict_2 = {'2nd_nd':'Non DVA Fac', 'Provider_2':'Provider', 'Pt count_2':'Pt_count', 'Stations_2':'Stations', '20_HEMO_2':'20_HEMO'}
nonDVA_2.rename(columns=dict_2, inplace=True)

nonDVA_3 = DVA_and_nonDVA[['3rd_nd', 'Provider_3', 'Pt count_3', 'Stations_3', '20_HEMO_3']].copy()
dict_3 = {'3rd_nd':'Non DVA Fac', 'Provider_3':'Provider', 'Pt count_3':'Pt_count', 'Stations_3':'Stations', '20_HEMO_3':'20_HEMO'}
nonDVA_3.rename(columns=dict_3, inplace=True)

nonDVA_facs = pd.concat([nonDVA_1, nonDVA_2, nonDVA_3], axis=0)


#currently my columns in nonDVA_facs are saved as object data type (strings or mixed data)
#I am going to convert the columns that I need to be in numeric format
nonDVA_facs['Pt_count'] = pd.to_numeric(nonDVA_facs['Pt_count'], errors='coerce')
nonDVA_facs['Stations'] = pd.to_numeric(nonDVA_facs['Stations'], errors='coerce')
nonDVA_facs['20_HEMO'] = pd.to_numeric(nonDVA_facs['20_HEMO'], errors='coerce')

#here I pivot the nonDVA_facs dataframe (pivot by nonDVA fac, sum pt counts, average stations & 20_HEMO) then
#calc OG operating utilization and new utilization in new columns and output that new dataframe
capacity_pivot = pd.pivot_table(nonDVA_facs,
                                index=['Non DVA Fac', 'Provider'],
                                values=['Pt_count', 'Stations', '20_HEMO'],
                                aggfunc={'Pt_count':'sum', 'Stations':'mean', '20_HEMO':'mean'})

#reset the index to get the facility and provider columns as separate columns
capacity_pivot.reset_index(inplace=True)

#here I add columns for operating capacity, new patient count, old utilization, and new utilization
capacity_pivot['Operating_Capacity'] = capacity_pivot['Stations']*6
capacity_pivot['New_Pt_Count'] = capacity_pivot['20_HEMO'] + capacity_pivot['Pt_count']
capacity_pivot['Old Utilization'] = capacity_pivot['20_HEMO'] / capacity_pivot['Operating_Capacity']
capacity_pivot['New Utilization'] = capacity_pivot['New_Pt_Count'] / capacity_pivot['Operating_Capacity']

#this next bit of code creates a summary tab that will go in our Excel file
#here I calculate patient counts and average distance for each facility
pt_count = DVA_and_nonDVA.groupby(["DVA Fac #"], as_index=False).agg(pt_count=pd.NamedAgg(column="DVA Fac #", aggfunc="count"))
DVA_avg = DVA_and_nonDVA.groupby(["DVA Fac #"]).agg(avg_dist=pd.NamedAgg(column="Distance", aggfunc="mean"))
non_DVA_avg = DVA_and_nonDVA.groupby(["DVA Fac #"]).agg(avg_dist_nonDVA=pd.NamedAgg(column="Distance_1", aggfunc="mean"))
merge_1 = pt_count.merge(DVA_avg, on="DVA Fac #")
merge_2 = merge_1.merge(non_DVA_avg, on="DVA Fac #")

final_df = merge_2.sort_values(by=['pt_count'], ascending=False)

#saving the results in an Excel file
#ENTER THE FOLDER/FILE PATH
with pd.ExcelWriter(rf'\\Output_file_path\{payor}_Distance Analysis.xlsx') as writer:
    DVA_and_nonDVA.to_excel(writer, sheet_name='Data', index=False)
    final_df.to_excel(writer, sheet_name='Summary', index=False)
    capacity_pivot.to_excel(writer, sheet_name='Capacity Data', index=False)


#HERE CREATE THE PLAN GROUPINGS THAT YOU WANT DISTANCE GRIDS CREATED FOR - I'VE CREATED A BUNCH...JUST TEXT (#) OUT THE ONES YOU DON'T NEED
#YOU CAN ALSO MANUALLY EDIT THE ONES I'VE CREATED TO REMOVE OR ADD A TYPE
all_plans = DVA_and_nonDVA.loc[DVA_and_nonDVA['Plan Type'].isin(['HMO_EPO', 'POS', 'PPO', 'INDEM', 'MCASGN', 'MAASGN', 'MEDICAID', 'MEDICARE'])]
hmo_plan = DVA_and_nonDVA[DVA_and_nonDVA['Plan Type'] == 'HMO_EPO']
ppo_plan = DVA_and_nonDVA[DVA_and_nonDVA['Plan Type'] == 'PPO']
mcasgn_plan = DVA_and_nonDVA[DVA_and_nonDVA['Plan Type'] == 'MCASGN']
maasgn_plan = DVA_and_nonDVA[DVA_and_nonDVA['Plan Type'] == 'MAASGN']
medicare_plan = DVA_and_nonDVA[DVA_and_nonDVA['Plan Type'] == 'MEDICARE']
medicaid_plan = DVA_and_nonDVA[DVA_and_nonDVA['Plan Type'] == 'MEDICAID']

#HERE WE SAVE THE NAMES OF EACH GROUP (W/IN THE 'QUOTES' IF YOU WANT TO CHANGE), AND THE ACTUAL GROUP DATAFRAMES IN A DICTIONARY...CAN ADJUST IF NEEDED
d = {'all':all_plans, 'hmo':hmo_plan, 'ppo':ppo_plan, 'mcasgn':mcasgn_plan, 'maasgn':maasgn_plan, 'mcare':medicare_plan, 'mcaid':medicaid_plan}

#this loop will create each individual distance grid and output them onto their own respective tab of an Excel file
for k, group in d.items():
    if group.empty:
        print(f"No patients in {k}")
    else:
        #here I am creating distance groups/buckets for both DVA & nonDVA
        conditions_1 = [
            (group['Distance'] <= 3),
            (group['Distance'] > 3) & (group['Distance'] <= 6),
            (group['Distance'] > 6) & (group['Distance'] < 10),
            (group['Distance'] >= 10)
        ]
        values_1 = ['0-3', '3-6', '6-10', '10+']
        conditions_2 = [
            (group['Distance_1'] <= 3),
            (group['Distance_1'] > 3) & (group['Distance_1'] <= 6),
            (group['Distance_1'] > 6) & (group['Distance_1'] < 10),
            (group['Distance_1'] >= 10)
        ]
        values_2 = ['0-3', '3-6', '6-10', '10+']

        group['DVA_group'] = np.select(conditions_1, values_1)
        group['nonDVA_group'] = np.select(conditions_2, values_2)
        #here I am going to use the pivot table function to create our distance pivot output
        distance_pivot = pd.pivot_table(group, values='Pt_MPI',
                                        index='nonDVA_group',
                                        columns='DVA_group',
                                        aggfunc=np.count_nonzero,
                                        margins=True, margins_name='Total')

        #this sorts the columns and rows of the pivot so that it outputs in an easy to read/final format
        distance_pivot = distance_pivot.reindex(['0-3','3-6','6-10','10+','Total'], axis=0)
        distance_pivot = distance_pivot.reindex(['0-3','3-6','6-10','10+','Total'], axis=1)

        #this is the grouping table that goes in the PowerPoint - there is code at the bottom that will export this to Excel
        #this line appends row/column names to the Pivot, so when we write to an Excel file it'll be easy to read/digest in Excel
        dp_final = pd.DataFrame(columns=distance_pivot.columns, index=[distance_pivot.index.name]).append(distance_pivot)

        #here I'm exporting/writing the grids to their own Excel files
        with pd.ExcelWriter(rf'\\Output_file_path\{payor}_{k}_Grid.xlsx') as writer:
            dp_final.to_excel(writer, sheet_name=f'{k}', index_label=distance_pivot.columns.name)
