import streamlit as st
import pandas as pd

# Load your dataset
dataset = pd.read_excel("./Dataset/cleaned/final_data.xlsx")

# Ensure the "Date" column is in datetime format
dataset["Date"] = pd.to_datetime(dataset["Date"], errors='coerce')

# Function for the Report Baking page
def report_baking():
    # Set the title of the app
    st.markdown("<h1 style='text-align: center;'>Baking Progress</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center;'>Metrics Overview</h4>", unsafe_allow_html=True)

    # Add a horizontal line for separation
    st.markdown("---")

    # Sidebar for logo and date selection
    st.sidebar.image("./assets/logo.jpg", width=500)

    st.sidebar.header("Filter")

    # Set the default start date to December 12, 2024
    start_date = st.sidebar.date_input("Start Date", value=pd.to_datetime("2024-12-12"), min_value=pd.to_datetime("2024-12-12"))
    end_date = st.sidebar.date_input("End Date", value=pd.to_datetime("today"))

    # Convert start_date and end_date to datetime
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # Filter the dataset based on the selected date range
    filtered_data = dataset[(dataset["Date"] >= start_date) & (dataset["Date"] <= end_date)]

    # Path to the report file
    report_download = "./Dataset/raw/12 Desember 2024.xlsx"

    # Create a download button in the sidebar
    st.sidebar.download_button(
        label="Download Excel Report",
        data=open(report_download, "rb").read(),
        file_name="baking_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Initialize variables for entering and exiting the baking room
    entering_baking_room = 0
    out_baking_room = 0
    in_pallete = 0
    out_pallete = 0

    # Calculate the number of entries in the baking room with status "IN"
    if not filtered_data.empty:
        in_data = filtered_data[filtered_data["Status"] == "IN"]
        entering_baking_room = in_data.shape[0] * 20

    # Calculate the number of entries in the baking room with status "OUT"
    if not filtered_data.empty:
        out_data = filtered_data[filtered_data["Status"] == "OUT"]
        out_baking_room = out_data.shape[0] * 20

    if not filtered_data.empty:
        in_pallete= in_data["ID pallete"].nunique()

    if not filtered_data.empty:
        out_pallete= out_data["ID pallete"].nunique()

    # Initialize variables for total baked
    model_value = "N/A"
    on_progress = 0
    target_baking = 25000  # Set your target baking value

    # Calculate on_progress based on paired ID boxes for the filtered dataset
    if not filtered_data.empty:
        model_value = filtered_data["model"].iloc[0] if "model" in filtered_data.columns else "N/A"
        paired_boxes = filtered_data[filtered_data.duplicated('ID box', keep=False)]
        on_progress = paired_boxes["ID box"].nunique() * 20

    # Calculate total baked from the entire dataset (not affected by date filter)
    total_baked = 0
    if not dataset.empty:
        total_paired_boxes = dataset[dataset.duplicated('ID box', keep=False)]
        total_baked = total_paired_boxes["ID box"].nunique() * 20

    # Calculate percentage for total baked
    total_percentage = (total_baked / target_baking) * 100 if target_baking > 0 else 0.0

    # Format numbers with thousands separator
    def format_number(num):
        return f"{num:,.0f}".replace(",", ".")  # Replace comma with dot for thousands separator

    # Display metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Model", value=model_value)
    with col3:
        st.metric(label="Shipment Code", value="LDCWT")


    st.markdown("---")  # Horizontal line for separation

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Target Baked", value=format_number(target_baking))
    with col3:
        st.metric(label="Total Baked", value=format_number(total_baked), delta=f"{total_percentage:.2f}%")

    st.markdown("---")  # Horizontal line for separation

    st.markdown("<h2 style='text-align: center;'>Daily Progress</h2>", unsafe_allow_html=True)
    st.markdown("---")  # Horizontal line for separation
    col4, col5, col6 = st.columns(3)
    with col4:
        st.markdown("<h3>IN</h3>", unsafe_allow_html=True)
        st.metric(label="Panels", value=format_number(entering_baking_room))
        st.metric(label="Palette", value=format_number(in_pallete))
    with col5:
        st.markdown("<h3>OUT</h3>", unsafe_allow_html=True)
        st.metric(label="Panels", value=format_number(out_baking_room))
        st.metric(label="Palette", value=format_number(out_pallete))
    with col6:
        st.metric(label="Baked", value=format_number(on_progress))

    # Add vertical space
    st.markdown("<br>", unsafe_allow_html=True)

# Function for the Report Sorting page
def report_sorting():
    st.title("Report Sorting")
    st.write("This is where you can implement the sorting functionality.")

# Main function to run the app
def main():
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Select a page:", ("Report Baking", "Report Sorting"))

    if page == "Report Baking":
        report_baking()
    elif page == "Report Sorting":
        report_sorting()

if __name__ == "__main__":
    main()