import plotly.figure_factory as ff
import numpy as np
import pandas as pd

#setting up the color scale for the output maps
colorscale = ["#f7fbff", "#ebf3fb", "#d2e3f3", "#c6dbef", "#9ecae1",
              "#85bcdb", "#57a0ce", "#4292c6", "#2171b5", "#1361a9",
              "#0b4083", "#08306b"
              ]

def create_establishment_map(industry_code):
    #reading in the LOCQ csv file
    locq_data = pd.read_csv('/Users/nicog-v/Documents/Python Mapping GPDs/ALL_LOC_QUOT_1974_1980.csv')

    #turning the state and county columns into 2-digit and 3-digit strings, respectively
    locq_data['fipstate'] = locq_data['fipstate'].apply(lambda x: str(x).zfill(2))
    locq_data['fipscty'] = locq_data['fipscty'].apply(lambda x: str(x).zfill(3))

    #creating a new FIPS value that is the concatenation of state + county
    locq_data['FIPS'] = locq_data['fipstate'] + locq_data['fipscty']

    #breaking the file apart by year
    years = locq_data['year'].unique()

    #create a dictionary to store the maps for each year
    maps = {}

    for year in years:
        #filter the data for the current year
        year_data = locq_data[(locq_data['year'] == year)]

        #designate the establishment column to a variable and calculate the 10th and 90th percentile
        est_column = f'est_{industry_code}'
        est_values = year_data[est_column]
        non_zero_values = est_values[est_values > 0]
        tenth = non_zero_values.quantile(0.1)
        ninety = non_zero_values.quantile(0.9)

        #set up the scaling for the map
        endpts = list(np.linspace(1.0, 11.0, len(colorscale) - 1))

        #create and plot the map
        fig = ff.create_choropleth(
            fips=year_data['FIPS'].tolist(),
            values=est_values.tolist(),
            scope=['usa'],
            binning_endpoints=endpts,
            colorscale=colorscale,
            county_outline={'color': 'black', 'width': 0.05},
            show_state_data=False,
            show_hover=True,
            asp=2.9,
            title_text=f"Figure A. Concentration of {industry_code} in 19{year}: establishments by county",
            legend_title='# of Establishments'
        )

        fig.layout.template = None

        #update layout features of the map
        fig.update_layout(legend=dict(x=1.05))

        #save the map as a PDF file
        file_path = f"/Users/nicog-v/Documents/Python Mapping GPDs/New title estb maps/Est_{industry_code}_19{year}_v2.pdf"
        fig.write_image(file_path)

        #store the map in the dictionary
        maps[year] = fig

    return maps

#specify the industry code we want to create the map for
industry_code = '7372'  #Computer Software industry code

#call the function to create the map for the specified industry code
industry_maps = create_establishment_map(industry_code)