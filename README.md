# Local Housing Market Trends and Structure Analysis

I wrote this script to help people understand the nature of the market they are looking to buy or sell in. Understanding price trends and how a house compares to others in the neighborhood can help significantly when buying or selling. High level extrapolation can lead to making decisions on incomplete data. This dashboard is designed to easily condense all relevant information when making the decision to buy or sell. For a sample see housingTrendsAndComparisonDashboard.xlsx.

The script pulls housing data using the Zillow API and then outputs an Excel based dashboard. The dashboard contains 6 graphs.

# I. Setting up the analysis

1. Obtain a Zillow API key from http://www.zillow.com/howto/api/APIOverview.htm and input the API key into 'zillow_data = ZillowWrapper('insertZillowAPIKey')' on line 8 of generateDashboard.py.
2. Create an array of home addresses you'd like to include in the analysis using the format [['street address 1',zipcode1],['street address 2',zipcode2]. (P.S. I think an interesting add on would be the ability to pull all addresses within a certain distance from an address using the googlemaps API.)
3. Run the code.

# II. Definitions

1. $ Per Square Foot Rolling Average - The average of the previous 10 values ranked on the date. This is done to smooth the graph so that we can more easily assess the trend over time. This is not dissimilar to a time based moving average, with the difference here being that the time distances between 1 to 2 and 2 to 3 are not neccessarily the same. As the sample size increases though, this affect would start to decrease.
2. $ Per Square Foot Rolling Standard Deviation - This helps us see how differentiated houses are in certain neighborhoods. For example, if there is a high standard deviation, you could start to look into why that is the case. What's the difference between a high $ per square foot and a lower $ per square foot? Is it that nicer, renovated houses are pulling in substantially more money? We can use this to estimate the value of renovations in certain areas where there is a higher than average differential. This is useful for both those wanting to buy a fixer upper or those looking to justify a higher than average price on a nicer home.
3. Home Price Rolling Average - The average of the previous 10 values ranked on the date.
4. Distribution of Bathrooms - The breakout of bedrooms and bathrooms in a given neighborhood.
5. Distribution of Bedrooms - Same as above
6. Distribution of Property Square Footage - A simple histogram of home square footages for comparison relative to pricing average.
