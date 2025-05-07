import pandas as pd
import matplotlib.pyplot as plt

# Load data from Excel file
file_path = "C:/Users/agawande/Desktop/18f4 Risk Workflows/case_study_data.xlsx"
ism_data = pd.read_excel(file_path, sheet_name='ism')
equity_data = pd.read_excel(file_path, sheet_name='assets')

# Convert date columns to datetime objects
ism_data['Date'] = pd.to_datetime(ism_data['ISM Release Date'])
equity_data['Date'] = pd.to_datetime(equity_data['Date'])

# Forward fill NaN values in key columns to ensure no NaN remain
ism_data['ISM Actual Reported'] = ism_data['ISM Actual Reported'].ffill()
ism_data['ISM Median Forecast'] = ism_data['ISM Median Forecast'].ffill()

# Calculate conditions
ism_data['Improved'] = ism_data['ISM Actual Reported'].diff().fillna(0) > 0
ism_data['Exceeds Forecast'] = ism_data['ISM Actual Reported'] > ism_data['ISM Median Forecast']

# Ensure these are boolean and no NaN exist
ism_data['Improved'] = ism_data['Improved'].fillna(False).astype(bool)
ism_data['Exceeds Forecast'] = ism_data['Exceeds Forecast'].fillna(False).astype(bool)

# Perform merge using 'merge_asof'
combined_data = pd.merge_asof(
    equity_data.sort_values('Date'),
    ism_data[['Date', 'Improved', 'Exceeds Forecast']].sort_values('Date'),
    on='Date',
    direction='backward'
)

# Plot the US Equity Index highlighting periods of conditions met
plt.figure(figsize=(14, 7))
plt.plot(combined_data['Date'], combined_data['US Equity Index'], label='US Equity Index', color='gray')

# Highlight periods where ISM improved and contains valid boolean values
plt.scatter(
    combined_data.loc[combined_data['Improved'].fillna(False), 'Date'],
    combined_data.loc[combined_data['Improved'].fillna(False), 'US Equity Index'],
    color='blue',
    label='ISM Improved'
)

# Highlight periods where ISM exceeds forecast
plt.scatter(
    combined_data.loc[combined_data['Exceeds Forecast'].fillna(False), 'Date'],
    combined_data.loc[combined_data['Exceeds Forecast'].fillna(False), 'US Equity Index'],
    color='green',
    label='ISM Exceeds Forecast'
)

plt.title('US Equity Performance with ISM Conditions')
plt.xlabel('Date')
plt.ylabel('US Equity Index')
plt.legend()
plt.show()
