import pandas as pd
import requests
from pandas.tseries.holiday import AbstractHolidayCalendar, Holiday
from datetime import datetime, timedelta

# Define a function to fetch holidays from an online API
def fetch_holidays(country_code, year):
    url = f"https://date.nager.at/api/v3/PublicHolidays/{year}/{country_code}"
    response = requests.get(url)
    if response.status_code == 200:
        holidays = response.json()
        return holidays
    else:
        return []

# Define the custom holiday calendar class using fetched holidays
class CustomHolidayCalendar(AbstractHolidayCalendar):
    def __init__(self, holidays):
        self.rules = [
            Holiday(holiday['localName'], year=int(holiday['date'][:4]), month=int(holiday['date'][5:7]), day=int(holiday['date'][8:10]))
            for holiday in holidays
        ]

# Function to create a holiday calendar for a given country code and year
def create_holiday_calendar(country_code, years):
    """Create a holiday calendar for a country spanning multiple years."""
    if isinstance(years, int):
        years = [years]
    all_holidays = []
    for yr in years:
        all_holidays.extend(fetch_holidays(country_code, yr))
    return CustomHolidayCalendar(holidays=all_holidays)

# Function to calculate working days in a month minus holidays
def calculate_working_days(start_date, end_date, holiday_calendar):
    all_days = pd.date_range(start=start_date, end=end_date, freq='B')
    holidays = holiday_calendar.holidays(start=start_date, end=end_date)
    working_days = all_days.difference(holidays)
    return len(working_days)

# Define the date range and country codes
start_date = '2024-04-01'
end_date = '2025-03-31'
months = pd.date_range(start=start_date, end=end_date, freq='MS')

# Determine the years spanned by the date range
start_year = pd.to_datetime(start_date).year
end_year = pd.to_datetime(end_date).year
holiday_years = list(range(start_year, end_year + 1))

# Country codes for the regions
country_codes = {
    "METAR (Dubai)": "AE",
    "Northern (UK)": "GB",
    "Western (France)": "FR",
    "Central (Germany)": "DE"
}

# Create holiday calendars for each region using all relevant years
dubai_holidays = create_holiday_calendar(country_codes["METAR (Dubai)"], holiday_years)
uk_holidays = create_holiday_calendar(country_codes["Northern (UK)"], holiday_years)
france_holidays = create_holiday_calendar(country_codes["Western (France)"], holiday_years)
germany_holidays = create_holiday_calendar(country_codes["Central (Germany)"], holiday_years)

# Calculate man days for each region
data = []
for month_start in months:
    month_end = month_start + pd.offsets.MonthEnd(0)
    dubai_days = calculate_working_days(month_start, month_end, dubai_holidays)
    uk_days = calculate_working_days(month_start, month_end, uk_holidays)
    france_days = calculate_working_days(month_start, month_end, france_holidays)
    germany_days = calculate_working_days(month_start, month_end, germany_holidays)
    data.append([month_start, month_end, dubai_days, uk_days, france_days, germany_days])

# Create the dataframe
df = pd.DataFrame(data, columns=["First Day of Month", "Last Day of Month", "METAR (Dubai)", "Northern (UK)", "Western (France)", "Central (Germany)"])

# Export the dataframe to a CSV file
output_path = "/mnt/data/man_days_per_month.csv"
df.to_csv(output_path, index=False)

output_path
