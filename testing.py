import pandas as pd
import bin_locations

part_df = pd.DataFrame(['1016130'], columns=['Part'])

bins = bin_locations.get_bins(part_df)

print(bins.columns)