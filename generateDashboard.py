from pyzillow.pyzillow import ZillowWrapper, GetDeepSearchResults
import pandas as pd
import xlsxwriter
import datetime as d
import numpy as np
#This is your zillow api key
#if you don't have one, sign up at http://www.zillow.com/howto/api/APIOverview.htm
zillow_data = ZillowWrapper('zillowAPIKey')
dfResults = pd.DataFrame({  'zillow_id':[],
                            'home_type':[],
                            'home_detail_link':[],
                            'graph_data_link':[],
                            'map_this_home_link':[],
                            'tax_year':[],
                            'tax_value':[],
                            'year_built':[],
                            'property_size':[],
                            'home_size':[],
                            'bathrooms':[],
                            'bedrooms':[],
                            'last_sold_date':[],
                            'last_sold_price_currency':[],
                            'last_sold_price':[],
                            'zestimate_amount':[],
                            'zestimate_last_updated':[],
                            'zestimate_value_change':[],
                            'zestimate_valuation_range_high':[],
                            'zestimate_valuationRange_low':[],
                            'zestimate_percentile':[]})

ppd =  [    
        ['1104 Shallowbrook Trail S',37013],
        ['1108 Shallowbrook Trail S',37013],
        ['1112 Shallowbrook Trail S',37013],
        ['1116 Shallowbrook Trail S',37013],
        ['1120 Shallowbrook Trail S',37013],
        ['1124 Shallowbrook Trail S',37013],
        ['1128 Shallowbrook Trail S',37013],
        ['1132 Shallowbrook Trail S',37013],
        ['1136 Shallowbrook Trail S',37013],
        ['1140 Shallowbrook Trail S',37013],
        ['1144 Shallowbrook Trail S',37013],
        ['1148 Shallowbrook Trail S',37013],
        ['1152 Shallowbrook Trail S',37013],
        ['1105 Shallowbrook Trail S',37013],
        ['1109 Shallowbrook Trail S',37013],
        ['1113 Shallowbrook Trail S',37013],
        ['1117 Shallowbrook Trail S',37013],
        ['1121 Shallowbrook Trail S',37013],
        ['1125 Shallowbrook Trail S',37013],
        ['1129 Shallowbrook Trail S',37013],
        ['1133 Shallowbrook Trail S',37013],
        ['1137 Shallowbrook Trail S',37013],
        ['1141 Shallowbrook Trail S',37013],
        ['1000 Shallowbrook Trail N',37013],
        ['1001 Shallowbrook Trail N',37013], 
        ['1004 Shallowbrook Trail N',37013], 
        ['1005 Shallowbrook Trail N',37013], 
        ['1008 Shallowbrook Trail N',37013], 
        ['1009 Shallowbrook Trail N',37013], 
        ['1012 Shallowbrook Trail N',37013], 
        ['1013 Shallowbrook Trail N',37013],    
        ['5012 Preserve Blvd',37013],
        ['5016 Preserve Blvd',37013],
        ['5020 Preserve Blvd',37013],
        ['5024 Preserve Blvd',37013],
        ['5028 Preserve Blvd',37013],
        ['5032 Preserve Blvd',37013],
        ['5036 Preserve Blvd',37013],
        ['5040 Preserve Blvd',37013],
        ['5044 Preserve Blvd',37013],
        ['5048 Preserve Blvd',37013],
        ['5052 Preserve Blvd',37013],
        ['5056 Preserve Blvd',37013],
        ['5060 Preserve Blvd',37013],
        ['5064 Preserve Blvd',37013],
        ['5068 Preserve Blvd',37013],
        ['5072 Preserve Blvd',37013],
        ['5076 Preserve Blvd',37013],
        ['5080 Preserve Blvd',37013],
        ['5084 Preserve Blvd',37013],
        ['5088 Preserve Blvd',37013],
        ['5092 Preserve Blvd',37013],
        ['5096 Preserve Blvd',37013],
        ['5100 Preserve Blvd',37013],
        ['5104 Preserve Blvd',37013],
        ['5108 Preserve Blvd',37013],
        ['5112 Preserve Blvd',37013],
        ['5116 Preserve Blvd',37013],
        ['5120 Preserve Blvd',37013],
        ['5124 Preserve Blvd',37013],
        ['5128 Preserve Blvd',37013],
        ['5132 Preserve Blvd',37013],
        ['5136 Preserve Blvd',37013],
        ['5140 Preserve Blvd',37013],
        ['5144 Preserve Blvd',37013],
        ['5148 Preserve Blvd',37013],
        ['5152 Preserve Blvd',37013],
        ['5156 Preserve Blvd',37013],
        ['5160 Preserve Blvd',37013],
        ['5164 Preserve Blvd',37013],
        ['5168 Preserve Blvd',37013],
        ['5172 Preserve Blvd',37013],
        ['5176 Preserve Blvd',37013],
        ['5180 Preserve Blvd',37013],
        ['5184 Preserve Blvd',37013],
        ['5188 Preserve Blvd',37013],
        ['5192 Preserve Blvd',37013],
        ['5011 Preserve Blvd',37013],
        ['5015 Preserve Blvd',37013],
        ['5019 Preserve Blvd',37013],
        ['5023 Preserve Blvd',37013],
        ['5027 Preserve Blvd',37013],
        ['5031 Preserve Blvd',37013],
        ['5035 Preserve Blvd',37013],
        ['5039 Preserve Blvd',37013],
        ['5043 Preserve Blvd',37013],
        ['5047 Preserve Blvd',37013],
        ['5051 Preserve Blvd',37013],
        ['5055 Preserve Blvd',37013],
        ['5059 Preserve Blvd',37013],
        ['5063 Preserve Blvd',37013],
        ['5067 Preserve Blvd',37013],
        ['5071 Preserve Blvd',37013],
        ['5075 Preserve Blvd',37013],
        ['5079 Preserve Blvd',37013],
        ['5083 Preserve Blvd',37013],
        ['5087 Preserve Blvd',37013],
        ['5091 Preserve Blvd',37013],
        ['5095 Preserve Blvd',37013],
        ['5099 Preserve Blvd',37013],
        ['5103 Preserve Blvd',37013],
        ['5105 Preserve Blvd',37013],
        ['5109 Preserve Blvd',37013],
        ['5113 Preserve Blvd',37013],
        ['5117 Preserve Blvd',37013],
        ['5121 Preserve Blvd',37013],
        ['5125 Preserve Blvd',37013],
        ['5129 Preserve Blvd',37013],
        ['5133 Preserve Blvd',37013],
        ['5137 Preserve Blvd',37013],
        ['5141 Preserve Blvd',37013],
        ['5145 Preserve Blvd',37013],
        ['5149 Preserve Blvd',37013],
        ['5153 Preserve Blvd',37013],
        ['5157 Preserve Blvd',37013],
        ['5161 Preserve Blvd',37013],
        ['5165 Preserve Blvd',37013],
        ['5169 Preserve Blvd',37013],
        ['5173 Preserve Blvd',37013],
        ['5177 Preserve Blvd',37013],
        ['5181 Preserve Blvd',37013],
        ['5185 Preserve Blvd',37013],
        ['5189 Preserve Blvd',37013],
        ['5193 Preserve Blvd',37013],
        ['5197 Preserve Blvd',37013],
        ['704 Candlecreek Way',37013],
        ['708 Candlecreek Way',37013],
        ['705 Candlecreek Way',37013],
        ['709 Candlecreek Way',37013],
        ['713 Candlecreek Way',37013],
        ['717 Candlecreek Way',37013],
        ['805 Birchmill Point N',37013],
        ['809 Birchmill Point N',37013],
        ['813 Birchmill Point N',37013],
        ['817 Birchmill Point N',37013],
        ['821 Birchmill Point N',37013],
        ['825 Birchmill Point N',37013],
        ['804 Birchmill Point N',37013],
        ['808 Birchmill Point N',37013],
        ['812 Birchmill Point N',37013],
        ['816 Birchmill Point N',37013],
        ['820 Birchmill Point N',37013],
        ['900 Birchmill Point S',37013],
        ['904 Birchmill Point S',37013],
        ['908 Birchmill Point S',37013],
        ['912 Birchmill Point S',37013],
        ['916 Birchmill Point S',37013],
        ['920 Birchmill Point S',37013],
        ['924 Birchmill Point S',37013],
        ['928 Birchmill Point S',37013],
        ['932 Birchmill Point S',37013],
        ['905 Birchmill Point S',37013],
        ['909 Birchmill Point S',37013],
        ['913 Birchmill Point S',37013],
        ['917 Birchmill Point S',37013],
        ['921 Birchmill Point S',37013],
        ['925 Birchmill Point S',37013],
        ['929 Birchmill Point S',37013],
        ['933 Birchmill Point S',37013],
        ['937 Birchmill Point S',37013],
        ['1205 Barkhill Pl',37013],
        ['1209 Barkhill Pl',37013],
        ['1213 Barkhill Pl',37013],
        ['1204 Barkhill Pl',37013],
        ['1208 Barkhill Pl',37013],
        ['1301 Rainglen Cove',37013],
        ['1305 Rainglen Cove',37013],
        ['1309 Rainglen Cove',37013],
        ['1313 Rainglen Cove',37013],
        ['1317 Rainglen Cove',37013],
        ['1321 Rainglen Cove',37013],
        ['1304 Rainglen Cove',37013],
        ['1308 Rainglen Cove',37013],
        ['1312 Rainglen Cove',37013],
        ['1316 Rainglen Cove',37013]]
j=-1
for i in ppd:
    address = i[0]
    zipcode = i[1]
    j += 1
    try:
        deep_search_response = zillow_data.get_deep_search_results(address, zipcode)
        result = GetDeepSearchResults(deep_search_response)
        dfResultsToAppend = pd.DataFrame({  'zillow_id':result.zillow_id,
                                'home_type':result.home_type,
                                'home_detail_link':result.home_detail_link,
                                'graph_data_link':result.graph_data_link,
                                'map_this_home_link':result.map_this_home_link,
                                'tax_year':result.tax_year,
                                'tax_value':result.tax_value,
                                'year_built':result.year_built,
                                'property_size':result.property_size,
                                'home_size':result.home_size,
                                'bathrooms':result.bathrooms,
                                'bedrooms':result.bedrooms,
                                'last_sold_date':result.last_sold_date,
                                'last_sold_price_currency':result.last_sold_price_currency,
                                'last_sold_price':result.last_sold_price,
                                'zestimate_amount':result.zestimate_amount,
                                'zestimate_last_updated':result.zestimate_last_updated,
                                'zestimate_value_change':result.zestimate_value_change,
                                'zestimate_valuation_range_high':result.zestimate_valuation_range_high,
                                'zestimate_valuationRange_low':result.zestimate_valuationRange_low,
                                'zestimate_percentile':result.zestimate_percentile},index=[j])
        dfResults = dfResults.append(dfResultsToAppend)
    except:
        print address
def floatConversion(x):
    if x is None:
        return None
    else:
        return float(x)
dfResults['last_sold_date'] = dfResults['last_sold_date'].apply(lambda x: d.datetime.strptime(x,'%m/%d/%Y').date())
dfResults['property_size'] = dfResults['property_size'].apply(lambda x: floatConversion(x))
dfResults['last_sold_price'] = dfResults['last_sold_price'].apply(lambda x: floatConversion(x))
dfResults['year_built'] = dfResults['year_built'].apply(lambda x: floatConversion(x))
dfResults['home_size'] = dfResults['home_size'].apply(lambda x: floatConversion(x))
dfResults['bathrooms'] = dfResults['bathrooms'].apply(lambda x: floatConversion(x))
dfResults['bedrooms'] = dfResults['bedrooms'].apply(lambda x: floatConversion(x))
dfResults = dfResults.dropna(subset=['property_size', 'last_sold_price', 'year_built','home_size','bathrooms','bedrooms'], how='any')
dfResults = dfResults.sort('last_sold_date', axis=0, ascending=True, na_position='first')
dfResults=dfResults.set_index([range(len(dfResults.index))])
dfResultsBathroomCount = dfResults[['bathrooms','bedrooms']].groupby(['bathrooms']).count()
dfResultsBathroomCount.columns = [['bathroomCount']]
dfResultsBedroomCount = dfResults[['bathrooms','bedrooms']].groupby(['bedrooms']).count()
dfResultsBedroomCount.columns = [['bedroomCount']]
bins=int(round(np.sqrt(len(dfResults['home_size'])),0))
dfResultsHomeSize = np.histogram(np.array(dfResults['home_size']), bins, range=None, normed=False, weights=None, density=None)
a = []
i=0
for i in range(bins):
    a.append(str(int(dfResultsHomeSize[1][i]))+' to '+str(int(dfResultsHomeSize[1][i+1])))
workbook = xlsxwriter.Workbook('housingTrendsAndComparisonDashboard.xlsx')
worksheetDashboard = workbook.add_worksheet('Dashboard')
worksheetData = workbook.add_worksheet('Data')
worksheetDataAggregated = workbook.add_worksheet('DataAggregated')
for row in dfResults.itertuples():
    for j in range(len(dfResults.columns)):
        worksheetData.write(row[0]+1,j,row[j+1])
        if j == len(dfResults.columns) - 1:
            #price per square foot            
            worksheetData.write(row[0]+1,j+1,'=H'+str(row[0]+2)+'/E'+str(row[0]+2))
            if row[0] > 9:
                worksheetData.write(row[0]+1,j+2,'=average(V'+str(row[0]-8)+':V'+str(row[0]+2)+')')
                worksheetData.write(row[0]+1,j+3,'=stdev(V'+str(row[0]-8)+':V'+str(row[0]+2)+')')
                worksheetData.write(row[0]+1,j+4,'=average(H'+str(row[0]-8)+':H'+str(row[0]+2)+')')
for i in range(len(dfResults.columns)):
    worksheetData.write(0,i,dfResults.columns[i])
worksheetData.write(0,i+1,'price_per_square_foot')
worksheetData.write(0,i+2,'roll_mean_price_per_square_foot')
worksheetData.write(0,i+3,'roll_stdev_price_per_square_foot')
worksheetData.write(0,i+4,'roll_mean_home_price')
i=1 
for row1,row2 in dfResultsBathroomCount.itertuples():
    worksheetDataAggregated.write(i,1,row1)
    worksheetDataAggregated.write(i,2,row2)
    i+=1
i=1 
for row1,row2 in dfResultsBedroomCount.itertuples():
    worksheetDataAggregated.write(i,3,row1)
    worksheetDataAggregated.write(i,4,row2)
    i+=1
for i in range(bins):
    worksheetDataAggregated.write(i+1,5,a[i])
    worksheetDataAggregated.write(i+1,6,dfResultsHomeSize[0][i])
worksheetDataAggregated.write(0,2,'bathroom_count')
worksheetDataAggregated.write(0,4,'bedroom_count')
worksheetDataAggregated.write(0,6,'home_square_footage_count')
dateFormat = workbook.add_format({'num_format':'mm/yy'})
dollarFormatSmall = workbook.add_format({'num_format':'$0'})
dollarFormatBig = workbook.add_format({'num_format':'$0,000'})
worksheetData.set_column('G:G', 11, dateFormat)
worksheetData.set_column('W:X', 11, dollarFormatSmall)
worksheetData.set_column('Y:Y', 11, dollarFormatBig)
chart1 = workbook.add_chart({'type': 'line','trendline': {'type': 'linear'}})
chart1.add_series({'values': '=Data!W12:W'+str(len(dfResults.index)+1),'categories': '=Data!G12:G'+str(len(dfResults.index)+1)})
worksheetDashboard.insert_chart('A1', chart1)
chart1.set_title({'name': '$ Per Square Foot Rolling Average','name_font': {'name': 'Calibri','color': '#E5E5E5','bold':True,'size':14}})
chart1.set_legend({'position': 'none'})
chart1.set_size({'width': 578, 'height': 400})
chart1.set_x_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
chart1.set_y_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
chart1.set_chartarea({'fill':{'color': '#404040'},'border': {'none': True}})
chart1.set_plotarea({'fill':{'color': '#404040'}})
chart2 = workbook.add_chart({'type': 'line','trendline': {'type': 'linear'}})
chart2.add_series({'values': '=Data!X12:X'+str(len(dfResults.index)+1),'categories': '=Data!G12:G'+str(len(dfResults.index)+1)})
worksheetDashboard.insert_chart('J1', chart2)
chart2.set_title({'name': '$ Per Square Foot Rolling Standard Deviation','name_font': {'name': 'Calibri','color': '#E5E5E5','bold':True,'size':14}})
chart2.set_legend({'position': 'none'})
chart2.set_size({'width': 578, 'height': 400})
chart2.set_x_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
chart2.set_y_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
chart2.set_chartarea({'fill':{'color': '#404040'},'border': {'none': True}})
chart2.set_plotarea({'fill':{'color': '#404040'}})
chart3 = workbook.add_chart({'type': 'line','trendline': {'type': 'linear'}})
chart3.add_series({'values': '=Data!Y12:Y'+str(len(dfResults.index)+1),'categories': '=Data!G12:G'+str(len(dfResults.index)+1)})
chart3.set_title({'name': 'Home Price Rolling Average','name_font': {'name': 'Calibri','color': '#E5E5E5','bold':True,'size':14}})
chart3.set_legend({'position': 'none'})
chart3.set_size({'width': 578, 'height': 400})
chart3.set_x_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
chart3.set_y_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
chart3.set_chartarea({'fill':{'color': '#404040'},'border': {'none': True}})
chart3.set_plotarea({'fill':{'color': '#404040'}})
worksheetDashboard.insert_chart('S1', chart3)
chart4 = workbook.add_chart({'type': 'pie'})
chart4.add_series({
    'values': '=DataAggregated!C2:C'+str(len(dfResultsBathroomCount.index)+1),
    'categories': 'DataAggregated!B2:B'+str(len(dfResultsBathroomCount.index)+1)})
chart4.set_title({'name': 'Distribution of Bathrooms','name_font': {'name': 'Calibri','color': '#E5E5E5','bold':True,'size':14}})
chart4.set_legend({'position': 'bottom',    'font': {
        'name': 'Calibri',
        'color': '#E5E5E5'}})
chart4.set_size({'width': 578, 'height': 400})
chart4.set_chartarea({'fill':{'color': '#404040'},'border': {'none': True}})
chart4.set_plotarea({'fill':{'color': '#404040'}})
worksheetDashboard.insert_chart('A21', chart4)
chart5 = workbook.add_chart({'type': 'pie'})
chart5.add_series({
    'values': '=DataAggregated!E2:E'+str(len(dfResultsBedroomCount.index)+1),
    'categories': 'DataAggregated!D2:D'+str(len(dfResultsBedroomCount.index)+1)})
chart5.set_title({'name': 'Distribution of Bedrooms','name_font': {'name': 'Calibri','color': '#E5E5E5','bold':True,'size':14}})
chart5.set_legend({'position': 'bottom',    'font': {
        'name': 'Calibri',
        'color': '#E5E5E5'}})
chart5.set_size({'width': 578, 'height': 400})
chart5.set_chartarea({'fill':{'color': '#404040'},'border': {'none': True}})
chart5.set_plotarea({'fill':{'color': '#404040'}})
worksheetDashboard.insert_chart('J21', chart5)
chart6 = workbook.add_chart({'type': 'column'})
chart6.add_series({
    'values': '=DataAggregated!G2:G'+str(bins+1),
    'categories': 'DataAggregated!F2:F'+str(bins+1)})
chart6.set_title({'name': 'Distribution of Property Square Footage','name_font': {'name': 'Calibri','color': '#E5E5E5','bold':True,'size':14}})
chart6.set_legend({'position': 'none'})
chart6.set_size({'width': 578, 'height': 400})
chart6.set_chartarea({'fill':{'color': '#404040'},'border': {'none': True}})
chart6.set_plotarea({'fill':{'color': '#404040'}})
chart6.set_x_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
chart6.set_y_axis({
    'minor_tick_mark':'none',
    'major_tick_mark':'none',
    'num_font': {
        'name': 'Calibri',
        'color': '#E5E5E5'
    },'major_gridlines': {
        'visible': True,
        'line': {'width': 0.25},
        'color':'gray'}})
worksheetDashboard.insert_chart('S21', chart6)
workbook.close()