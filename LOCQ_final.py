import pandas as pd
import glob
import re

def process_industry_data(industry_code):
    dataframes = {}  #dictionary to store the dataframes

    #setting up a loop to loop through all the CBP txt files in my folder
    path = "/Users/nicog-v/Documents/Python Mapping GPDs/CBP_Data_5.21.23/SIC_86_97/*.txt"
    for fname in glob.glob(path):
        #here I am creating a variable that picks out the year from the fname
        year = re.findall("_(\d+).txt", fname)[0]

        #reading in the CBP county employment files
        cbpco = pd.read_csv(fname, delimiter=',', quotechar='"')

        #removing subtotal & total rows (wherever the SIC code has a '-' or a '\')
        cbpco = cbpco[~cbpco['sic'].str.contains('-')]
        cbpco = cbpco[~cbpco['sic'].str.contains(r"\\")]

        #keeping only the columns needed & getting rid of all the extra columns
        cbp = cbpco.loc[:,['fipstate', 'fipscty', 'sic', 'emp', 'est']]

        #turning the state and county columns into 2-digit and 3-digit strings, respectively
        cbp['fipstate'] = cbp['fipstate'].apply(lambda x: str(x).zfill(2))
        cbp['fipscty'] = cbp['fipscty'].apply(lambda x: str(x).zfill(3))

        #creating a new FIPS value that is the concat of state + county
        cbp['FIPS'] = cbp['fipstate'] + cbp['fipscty']

        #calculating county level total employment
        cbp['cty_emp'] = cbp.groupby('FIPS')['emp'].transform('sum')

        #filter the dataframe based on the specified industry code
        industry = cbp[cbp['sic'] == industry_code]

        #calculating total nationwide employment across all industries
        total_emp = cbp['emp'].sum()

        #calculating employment for the specified industry
        industry_emp = industry['emp'].sum()

        #calculating the LOCQ based on employment numbers
        industry[f'locq_{year}'] = (industry['emp'] / industry['cty_emp']) / (industry_emp / total_emp)

        #removing unneeded columns for easy merge/join later
        industry.drop(['fipstate', 'fipscty', 'sic', 'emp', 'est', 'cty_emp'], axis=1, inplace=True)

        #adding the industry dataframe to the dataframes dictionary with the corresponding year as the key
        industry_name = f"industry_{industry_code}_{year}"
        dataframes[industry_name] = industry

    #create an empty dataframe to store the merged data
    merged_industry = pd.DataFrame()

    #loop through the industry dataframes
    for year in range(86, 98):
        #get the dataframe for the current year and specified industry code
        industry_name = f"industry_{industry_code}_{year}"
        industry_year = dataframes[industry_name]

        #merge the current year's dataframe with the merged dataframe
        if merged_industry.empty:
            merged_industry = industry_year
        else:
            merged_industry = pd.merge(merged_industry, industry_year, on='FIPS', how='outer')

    #sort the merged dataframe by the 'FIPS' column
    merged_industry = merged_industry.sort_values('FIPS')

    #reset the index of the merged dataframe
    merged_industry = merged_industry.reset_index(drop=True)

    #replace NaN values with 0s
    merged_industry = merged_industry.fillna(0)

    #save the results in an Excel file
    output_file = f'/Users/nicog-v/Documents/Python Mapping GPDs/CBP_Data_5.21.23/SIC_86_97/LOCQs_{industry_code}.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        merged_industry.to_excel(writer, sheet_name='comp_soft', index=False)

#specify the industry code we're interested in
industry_code = '3573'  #Computer Hardware industry code

#call the function to process the specified industry code
process_industry_data(industry_code)