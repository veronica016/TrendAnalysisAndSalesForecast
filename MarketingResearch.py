import pandas as pd
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA

# Load the Excel file from your specified path
data = pd.read_excel(r"D:\PREP\Marketing Analytics\DATA.xlsx")

# Convert 'DATE' column to datetime
data['DATE'] = pd.to_datetime(data['DATE'])

# Set 'DATE' as the index
data.set_index('DATE', inplace=True)

# OPTIONAL: Fill missing values if any
# data['NET AMOUNT'] = data['NET AMOUNT'].fillna(method='ffill')

# Fit ARIMA model (order=(1,1,1) is basic, you can tune it later)
model = ARIMA(data['NET AMOUNT'], order=(1, 1, 1))
model_fit = model.fit()

# Create future dates for the next 30 days
future_days = pd.date_range(start=data.index[-1] + pd.Timedelta(days=1), periods=30)

# Forecast next 30 data points
predicted_sales = model_fit.forecast(steps=30)

# Plot the actual and forecasted data
plt.figure(figsize=(12, 6))
plt.plot(data['NET AMOUNT'], label='Historical Sales')
plt.plot(future_days, predicted_sales, label='Predicted Sales', color='orange')
plt.legend()
plt.xlabel('Date')
plt.ylabel('Net Amount')
plt.title('📈 Sales Forecast for Next 30 Days (ARIMA Model)')
plt.grid(True)
plt.tight_layout()
plt.show()

# Print forecasted values
print("📅 Predicted Sales for the Next 30 Days:")
for date, sales in zip(future_days, predicted_sales):
    print(f"{date.strftime('%Y-%m-%d')}: ₹{sales:.2f}")
