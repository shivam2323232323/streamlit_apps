import streamlit as st
from io import BytesIO
import xlsxwriter
from google.oauth2.service_account import Credentials
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import DateRange, Metric, Dimension, Filter, FilterExpression

# Streamlit app title
st.title("GA4 Month-on-Month Report Generator")

# Sidebar inputs
st.sidebar.header("Configuration")
key_file = st.sidebar.file_uploader("Upload Service Account Key", type=["json"])
property_id = st.sidebar.text_input("Enter GA4 Property ID")
generate_button = st.sidebar.button("Generate Report")

# Months configuration
months = [
    {"month": "Apr", "start_date": "2024-04-01", "end_date": "2024-04-30"},
    {"month": "May", "start_date": "2024-05-01", "end_date": "2024-05-31"},
    {"month": "Jun", "start_date": "2024-06-01", "end_date": "2024-06-30"},
    {"month": "Jul", "start_date": "2024-07-01", "end_date": "2024-07-31"},
    {"month": "Aug", "start_date": "2024-08-01", "end_date": "2024-08-31"},
    {"month": "Sep", "start_date": "2024-09-01", "end_date": "2024-09-30"},
    {"month": "Oct", "start_date": "2024-10-01", "end_date": "2024-10-31"},
    {"month": "Nov", "start_date": "2024-11-01", "end_date": "2024-11-30"},
    {"month": "Dec", "start_date": "2024-12-01", "end_date": "2024-12-31"},
]

# Function to fetch data from GA4 API
def fetch_ga4_data(key_file_path, property_id):
    credentials = Credentials.from_service_account_file(key_file_path)
    client = BetaAnalyticsDataClient(credentials=credentials)

    request = {
        "property": f"properties/{property_id}",
        "metrics": [
            Metric(name="sessions"),
            Metric(name="engagedSessions"),
            Metric(name="totalUsers"),
        ],
        "dimensions": [
            Dimension(name="sessionDefaultChannelGrouping"),
        ],
        "dimension_filter": FilterExpression(
            filter=Filter(
                field_name="sessionDefaultChannelGrouping",
                string_filter=Filter.StringFilter(value="organic search")
            )
        )
    }

    all_data = []
    for month in months:
        try:
            request['date_ranges'] = [DateRange(start_date=month['start_date'], end_date=month['end_date'])]
            response = client.run_report(request)

            total_sessions = 0
            total_engaged_sessions = 0
            total_users = 0

            for row in response.rows:
                total_sessions += int(row.metric_values[0].value)
                total_engaged_sessions += int(row.metric_values[1].value)
                total_users += int(row.metric_values[2].value)

            monthly_data = [month['month'], total_sessions, total_engaged_sessions, total_users]
            all_data.append(monthly_data)
        except Exception as e:
            st.error(f"Error fetching data for {month['month']} 2024: {e}")

    return all_data

# Function to generate Excel report with insights and formatting
def generate_excel(data):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("GA4 Data")

    headers = ['Month', 'Sessions', 'Engaged Sessions', 'Users']
    header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'align': 'center', 'border': 1})
    content_format = workbook.add_format({'align': 'center', 'border': 1})
    insight_header_format = workbook.add_format({'bold': True, 'font_color': 'black', 'align': 'left'})
    improvement_format = workbook.add_format({'font_color': 'darkgreen', 'bold': True})
    drop_format = workbook.add_format({'font_color': 'red', 'bold': True})

    # Write headers
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, header_format)

    # Write data
    for row_num, row_data in enumerate(data, start=1):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, cell_data, content_format)

    # Generate insights if data is available
    if len(data) >= 2:
        last_month = data[-1][0]
        second_last_month = data[-2][0]
        previous_month = data[-2]
        current_month = data[-1]
        insights_row_start = len(data) + 2

        # Write insights header
        worksheet.write(insights_row_start - 1, 0, f"Highlights - {last_month}'24 vs {second_last_month}'24", insight_header_format)

        for i in range(1, len(headers)):  # Skip the 'Month' column
            metric_name = headers[i]
            previous_value = previous_month[i]
            current_value = current_month[i]
            percentage_change = ((current_value - previous_value) / previous_value * 100) if previous_value != 0 else 0

            change_type = "improvement" if percentage_change > 0 else "drop"
            percentage_text = f"{change_type} of {abs(percentage_change):.2f}%"
            insight = (
                f"We have observed a {percentage_text} in {metric_name} in {last_month} compared to {second_last_month}."
            )

            worksheet.write_rich_string(
                insights_row_start + i - 1, 0,
                "We have observed a ",
                improvement_format if change_type == "improvement" else drop_format, percentage_text,
                insight[len("We have observed a " + percentage_text):],
                content_format
            )

    # Apply conditional formatting
    for col in range(1, len(headers)):
        format_range = f"{chr(65+col)}2:{chr(65+col)}{len(data)+1}"
        worksheet.conditional_format(
            format_range,
            {
                'type': '3_color_scale',
                'min_color': "#F8696B",
                'mid_color': "#FFEB84",
                'max_color': "#63BE7B",
            }
        )

    worksheet.set_column(0, len(headers) - 1, 15)
    workbook.close()
    output.seek(0)
    return output

# Main logic
if generate_button:
    if not key_file:
        st.error("Please upload a valid service account key file.")
    elif not property_id:
        st.error("Please enter a valid GA4 Property ID.")
    else:
        with st.spinner("Fetching data and generating report..."):
            try:
                # Save key file temporarily
                key_file_path = "temp_key.json"
                with open(key_file_path, "wb") as f:
                    f.write(key_file.read())

                # Fetch data and generate report
                data = fetch_ga4_data(key_file_path, property_id)
                excel_file = generate_excel(data)

                st.success("Report generated successfully!")
                st.download_button(
                    label="Download Excel Report",
                    data=excel_file,
                    file_name="GA4_Report_Insights.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred: {e}")
