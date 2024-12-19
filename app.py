from io import BytesIO
import pymssql  # Use pymssql instead of pyodbc
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
import logging

# Load environment variables
load_dotenv()

# Set up logging for debug messages
logging.basicConfig(level=logging.DEBUG)

# Database connection function
def get_db_connection():
    try:
        # Use pymssql to connect to the database
        conn = pymssql.connect(
            server='103.153.58.143',  # SQL Server address
            user='sa',                # SQL Server username
            password='FnSDj*38J6Z#949sdgj',  # SQL Server password
            database='z_scope'        # SQL Server database
        )
        logging.debug("Database connection established.")
        return conn
    except Exception as e:
        logging.error(f"Error establishing database connection: {e}")
        st.error(f"Error connecting to the database: {e}")
        return None

# Streamlit app interface
def main():
    st.title("TOPS DATA BASE-MAPPING")
    
    # Fetch brands from the database
    conn = get_db_connection()
    if conn is None:
        return  # Exit if no connection is established

    cursor = conn.cursor()
    cursor.execute("SELECT bigid, vcBrand FROM Brand_Master")
    brands = [{'id': row[0], 'name': row[1]} for row in cursor.fetchall()]
    cursor.close()
    conn.close()

    # Display the brand selection dropdown
    brand_choices = [brand['name'] for brand in brands]
    brand_id = st.selectbox("Select Brand", brand_choices)

    if brand_id:
        selected_brand = next(brand for brand in brands if brand['name'] == brand_id)
        brand_id = selected_brand['id']
        
        # Fetch dealers based on selected brand
        conn = get_db_connection()
        if conn is None:
            return  # Exit if no connection is established

        cursor = conn.cursor()
        cursor.execute("SELECT bigid, vcName FROM Dealer_Master WHERE BrandID = ?", (brand_id,))
        dealers = [{'id': row[0], 'name': row[1]} for row in cursor.fetchall()]
        cursor.close()
        conn.close()

        # Dealer selection dropdown
        dealer_choices = [dealer['name'] for dealer in dealers]
        dealer_id = st.selectbox("Select Dealer", dealer_choices)

        if dealer_id:
            selected_dealer = next(dealer for dealer in dealers if dealer['name'] == dealer_id)
            dealer_id = selected_dealer['id']

            # Option to select stored procedure
            procedure = st.selectbox("Select Procedure", ['Base', 'Mapping'])

            if st.button("Get Result"):
                fetch_result(brand_id, dealer_id, procedure)

# Function to fetch the result based on the selected procedure
def fetch_result(brand_id, dealer_id, procedure):
    try:
        conn = get_db_connection()
        if conn is None:
            return  # Exit if no connection is established

        cursor = conn.cursor()

        if procedure == 'Base':
            # Execute the stored procedure for "base"
            cursor.execute("EXEC z_scope.dbo.tops_vs_scs_norms_base1 @brandid = %s, @dealerid = %s", (brand_id, dealer_id))
            logging.debug("Executed stored procedure: tops_vs_scs_norms_base1")

            # Fetch data from the database
            cursor.execute("SELECT * FROM tops_vs_norms_base")
            results_base = cursor.fetchall()

            if not results_base:
                st.warning("No data available for the selected procedure.")
                return

            columns_base = [column[0] for column in cursor.description]
            data_base = pd.DataFrame([tuple(row) for row in results_base], columns=columns_base)

            # Export DataFrame to Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                data_base.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="Base Excel",
                data=output,
                file_name="base_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif procedure == 'Mapping':
            # Execute the stored procedure for "mapping"
            cursor.execute("EXEC z_scope.dbo.Tops_vs_SCS_Norms_test1 @brandid = %s, @dealerid = %s", (brand_id, dealer_id))
            logging.debug("Executed stored procedure: Tops_vs_SCS_Norms_test1")

            # Fetch data from the database
            cursor.execute("SELECT * FROM TOPs_VS_SCS_Norms")
            results_mapping = cursor.fetchall()

            if not results_mapping:
                st.warning("No data available for the selected procedure.")
                return

            columns_mapping = [column[0] for column in cursor.description]
            data_mapping = pd.DataFrame([tuple(row) for row in results_mapping], columns=columns_mapping)

            # Export DataFrame to Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                data_mapping.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="Mapping Excel",
                data=output,
                file_name="mapping_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        cursor.close()
        conn.close()

    except Exception as e:
        logging.error(f"Error in fetch_result function: {e}")
        st.error(f"Error processing the request: {e}")

if __name__ == '__main__':
    main()
