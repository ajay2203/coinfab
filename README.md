# cointab
The given  the intention of this code is to calculate the charges and compare the charges billed by the courier company with the expected charges as per the calculations from the given data. The steps of the code are described below:

Required libraries are imported including pandas and numpy.
The data is loaded from various excel files using pandas read_excel function and is stored in different dataframes.
A function named add_rate is defined that takes two dataframes as inputs, and returns a dataframe. It takes the zone-wise rates from one dataframe and appends it to the main dataframe as columns for forward and RTO rates.
Another function named calculate_payments is defined that takes one dataframe as input and returns a dataframe. It calculates the expected charges based on the weight slab, fixed rate, and additional rate from the previous function.
The data is cleaned, and some necessary columns are merged into the main dataframe.
The weight slabs are calculated based on the weight of the product.
The charges are calculated using the calculate_payments function based on the weight slabs and other rates.
The final dataframe is merged with the courier company invoice data, and some final calculations are performed, including the difference between expected charges and billed charges.
The final dataframe is returned.
