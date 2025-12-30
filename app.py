import streamlit as st
from O365 import Account
import datetime as dt

# --- PAGE SETUP ---
st.set_page_config(page_title="MSK Scheduler", page_icon="📅")
st.title("📅 MSK's Smart Scheduler")

# --- 1. SECURE CONFIGURATION ---
# We pull keys from the Cloud Secrets (not hardcoded)
try:
    CLIENT_ID = st.secrets["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
    TENANT_ID = st.secrets["TENANT_ID"]
    # This URL will be the link to your live app
    REDIRECT_URI = st.secrets["REDIRECT_URI"] 
except FileNotFoundError:
    st.error("Secrets not found! Please configure them in Streamlit Cloud.")
    st.stop()

# --- 2. AUTHENTICATION SETUP ---
credentials = (CLIENT_ID, CLIENT_SECRET)
scopes = ['Calendars.ReadWrite', 'OnlineMeetings.ReadWrite', 'User.Read']
account = Account(credentials, auth_flow_type='authorization', tenant_id=TENANT_ID)

# --- 3. LOGIN LOGIC ---
if not account.is_authenticated:
    st.warning("⚠️ You are not connected to Outlook.")
    
    if st.button("Connect Outlook Account"):
        url, state = account.con.get_authorization_url(
            requested_scopes=scopes, 
            redirect_uri=REDIRECT_URI
        )
        st.markdown(f"### [👉 Click here to Login]({url})", unsafe_allow_html=True)
        st.info(f"1. Click the link above. \n2. Authorize the app. \n3. Copy the URL from the browser address bar. \n4. Paste it below.")

    url_pasted = st.text_input("Paste the full Return URL here:")
    
    if url_pasted:
        try:
            result = account.con.request_token(url_pasted, state=state, redirect_uri=REDIRECT_URI)
            if result:
                st.success("✅ Login Successful! Refresh the page.")
        except Exception as e:
            st.error(f"Error: {e}")

# --- 4. THE BOOKING INTERFACE ---
else:
    st.success("🟢 System Online: Connected to MSK's Calendar")

    with st.form("booking_form"):
        st.subheader("Book a Meeting Slot")
        
        col1, col2 = st.columns(2)
        date = col1.date_input("Select Date", min_value=dt.date.today())
        time = col2.time_input("Select Time", value=dt.time(10, 00))
        duration = st.slider("Duration (minutes)", 15, 60, 30)
        
        subject = st.text_input("Meeting Subject", "Project Discussion")
        attendee_email = st.text_input("Your Email (for the invite)")
        
        submitted = st.form_submit_button("Confirm Booking")

        if submitted:
            with st.spinner("Syncing with Teams..."):
                schedule = account.schedule()
                calendar = schedule.get_default_calendar()
                
                new_event = calendar.new_event()
                new_event.subject = subject
                start_datetime = dt.datetime.combine(date, time)
                new_event.start = start_datetime
                new_event.end = start_datetime + dt.timedelta(minutes=duration)
                
                # The Magic Switch for Teams
                new_event.is_online_meeting = True
                
                if attendee_email:
                    new_event.attendees.add(attendee_email)
                
                if new_event.save():
                    st.balloons()
                    st.success("✅ Meeting Booked! Teams Link Generated.")
                else:
                    st.error("Booking Failed.")