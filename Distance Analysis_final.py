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
nonDVA = pd.DataFrame(columns=['Pt_coords', '1st_nd', '2nd_nd', '3rd_nd', 'Distance_1', 'Distance_2', 'Distance_3', 'Provider_1', 'Provider_2', 'Provider_3'])

#Iterate over each patient
for lat, lon in payor_pts[['Latitude', 'Longitude']].values:
    # Calculate distance to each nonDVA facility
    distances = []
    for _, row in non_dva_facilities.iterrows():
        fac_coords = (row['latitude'], row['longitude'])
        miles_to_nd = geopy.distance.geodesic((lat, lon), fac_coords).miles
        nd_fac = row['Facility']
        nd_provider = row['ProviderGroup']
        distances.append((nd_fac, miles_to_nd, nd_provider))

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
        'Provider_3': distances[2][2]
    }, ignore_index=True)

#this new merged df will have each pts MPI, plan type, coordinates, closest DVA Fac #, Distance to DVA fac, closest
#3 nonDVA Facs and distance to those facs
DVA_and_nonDVA = DVA.merge(nonDVA, on="Pt_coords")
DVA_and_nonDVA.drop_duplicates(subset=['Pt_MPI'], inplace=True)
#we will output this dataframe as one of the sheets in the Excel file


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