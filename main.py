import streamlit as st
import io
from excel_generator import InvoiceGenerator
from datetime import datetime

st.set_page_config(page_title="Invoice Generator", layout="wide")

st.title("Invoice Generator")

with st.form("invoice_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Company Details")
        company_name = st.text_input("Company Name")
        address_line1 = st.text_input("Address Line 1")
        address_line2 = st.text_input("Address Line 2")
        city = st.text_input("City")
        state = st.text_input("State")
        pincode = st.text_input("Pincode")
        pan = st.text_input("PAN Number")
        phone = st.text_input("Phone Number")

    with col2:
        st.subheader("Client Details")
        client_name = st.text_input("Client Name")
        client_address1 = st.text_input("Client Address Line 1")
        client_address2 = st.text_input("Client City")
        client_state = st.text_input("Client State")
        client_country = st.text_input("Client Country & Code")
        supply_place = st.text_input("Place of Supply")

    st.subheader("Invoice Items")
    description = st.text_input("Description")
    amount = st.number_input("Amount", value=0.0)
    discount_percentage = st.number_input("Discount Percentage", value=0.0)

    st.subheader("Bank Details")
    bank_name = st.text_input("Bank Name")
    account_name = st.text_input("Account Name")
    account_number = st.text_input("Account Number")
    ifsc_code = st.text_input("IFSC Code")
    swift_code = st.text_input("Swift Code")
    account_type = st.text_input("Account Type")

    submitted = st.form_submit_button("Generate Invoice")

if submitted:
    # Create Excel file
    invoice_gen = InvoiceGenerator()
    
    # Add company details
    company_details = {
        'name': company_name,
        'address_line1': address_line1,
        'address_line2': address_line2,
        'city': city,
        'state': state,
        'pincode': pincode,
        'pan': pan,
        'phone': phone
    }
    invoice_gen.add_company_details(company_details)
    
    # Add client details
    client_details = {
        'name': client_name,
        'address1': client_address1,
        'address2': client_address2,
        'state': client_state,
        'country': client_country,
        'supply': supply_place
    }
    invoice_gen.add_bill_to_section(client_details)
    
    # Add items
    items = [
        {
            'description': description,
            'amount': amount
        },
        {
            'description': 'Discount',
            'percentage': discount_percentage,
            'amount': -amount * (discount_percentage/100)
        }
    ]
    invoice_gen.add_items_section(items)
    
    # Add bank details
    bank_details = {
        'Bank Name': bank_name,
        'Name': account_name,
        'Account Number': account_number,
        'IFSC Code': ifsc_code,
        'Swift Code': swift_code,
        'Account Type': account_type
    }
    invoice_gen.add_bank_details(bank_details)
    
    # Save to buffer
    buffer = io.BytesIO()
    invoice_gen.save(buffer)
    buffer.seek(0)
    
    # Offer download
    st.download_button(
        label="Download Excel Invoice",
        data=buffer,
        file_name=f"invoice_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
