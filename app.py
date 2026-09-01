"""
CareSync - Patient-Hospital-Doctor Booking App
------------------------------------------------
A basic prototype built with Streamlit (pure Python).
Data is stored locally in Excel files (via pandas + openpyxl).

Run with:
    pip install -r requirements.txt
    streamlit run app.py
"""

import os
import random
import smtplib
import hashlib
from email.mime.text import MIMEText
from datetime import date, datetime

import pandas as pd
import streamlit as st

# ------------------------------------------------------------------
# BASIC CONFIG
# ------------------------------------------------------------------
APP_NAME = "CareSync"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo.png")
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

USERS_FILE = os.path.join(DATA_DIR, "users.xlsx")
HOSPITALS_FILE = os.path.join(DATA_DIR, "hospitals.xlsx")
DOCTORS_FILE = os.path.join(DATA_DIR, "doctors.xlsx")
APPOINTMENTS_FILE = os.path.join(DATA_DIR, "appointments.xlsx")

# Fixed daily slots offered for every doctor (kept simple on purpose)
TIME_SLOTS = ["10:00 AM", "11:00 AM", "12:00 PM", "2:00 PM", "3:00 PM", "4:00 PM", "5:00 PM"]

# Very simple keyword -> specialist mapping ("basic AI" symptom checker)
SYMPTOM_MAP = {
    "chest pain": "Cardiologist", "heart": "Cardiologist", "bp": "Cardiologist",
    "skin": "Dermatologist", "rash": "Dermatologist", "acne": "Dermatologist",
    "tooth": "Dentist", "teeth": "Dentist", "gum": "Dentist",
    "bone": "Orthopedic", "fracture": "Orthopedic", "joint": "Orthopedic", "knee": "Orthopedic",
    "eye": "Ophthalmologist", "vision": "Ophthalmologist",
    "child": "Pediatrician", "kid": "Pediatrician", "baby": "Pediatrician",
    "stomach": "Gastroenterologist", "digestion": "Gastroenterologist", "acidity": "Gastroenterologist",
    "mental": "Psychiatrist", "anxiety": "Psychiatrist", "depression": "Psychiatrist", "stress": "Psychiatrist",
    "diabetes": "Endocrinologist", "thyroid": "Endocrinologist", "sugar": "Endocrinologist",
    "fever": "General Physician", "cold": "General Physician", "cough": "General Physician",
    "headache": "Neurologist", "migraine": "Neurologist", "seizure": "Neurologist",
    "kidney": "Nephrologist", "urine": "Nephrologist",
    "pregnan": "Gynecologist", "period": "Gynecologist", "women": "Gynecologist",
    "ear": "ENT Specialist", "nose": "ENT Specialist", "throat": "ENT Specialist",
}

SPECIALIST_LIST = sorted(set(SYMPTOM_MAP.values()))

# ------------------------------------------------------------------
# EXCEL HELPERS
# ------------------------------------------------------------------
def load_table(path: str, columns: list) -> pd.DataFrame:
    if os.path.exists(path):
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame(columns=columns)
    return pd.DataFrame(columns=columns)


def save_row(path: str, columns: list, row: dict):
    df = load_table(path, columns)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(path, index=False)


def seed_demo_data():
    """Pre-fill a couple of hospitals + doctors so the booking flow can be tested
    end-to-end without registering a hospital first."""
    if not os.path.exists(HOSPITALS_FILE):
        hospitals = [
            {"CenterName": "CareSync City Hospital", "City": "Delhi",
             "Address": "MG Road, Delhi", "ContactInfo": "9999900001",
             "Email": "citycontact@caresync.demo"},
            {"CenterName": "CareSync Wellness Clinic", "City": "Delhi",
             "Address": "CP, Delhi", "ContactInfo": "9999900002",
             "Email": "wellness@caresync.demo"},
            {"CenterName": "CareSync Multispeciality", "City": "Mumbai",
             "Address": "Andheri, Mumbai", "ContactInfo": "9999900003",
             "Email": "mumbai@caresync.demo"},
        ]
        pd.DataFrame(hospitals).to_excel(HOSPITALS_FILE, index=False)

    if not os.path.exists(DOCTORS_FILE):
        doctors = [
            {"Name": "Dr. A. Sharma", "Age": 42, "ContactInfo": "9999911111",
             "Proof": "demo", "Speciality": "Cardiologist",
             "Hospital": "CareSync City Hospital", "City": "Delhi",
             "Fee": 800, "Schedule": "Mon-Sat, 10AM-5PM", "Email": "asharma@caresync.demo"},
            {"Name": "Dr. R. Verma", "Age": 35, "ContactInfo": "9999922222",
             "Proof": "demo", "Speciality": "General Physician",
             "Hospital": "CareSync City Hospital", "City": "Delhi",
             "Fee": 400, "Schedule": "Mon-Sun, 10AM-5PM", "Email": "rverma@caresync.demo"},
            {"Name": "Dr. N. Kapoor", "Age": 39, "ContactInfo": "9999933333",
             "Proof": "demo", "Speciality": "Dermatologist",
             "Hospital": "CareSync Wellness Clinic", "City": "Delhi",
             "Fee": 600, "Schedule": "Tue-Sun, 11AM-4PM", "Email": "nkapoor@caresync.demo"},
            {"Name": "Dr. S. Iyer", "Age": 47, "ContactInfo": "9999944444",
             "Proof": "demo", "Speciality": "Cardiologist",
             "Hospital": "CareSync Multispeciality", "City": "Mumbai",
             "Fee": 900, "Schedule": "Mon-Fri, 10AM-5PM", "Email": "siyer@caresync.demo"},
        ]
        pd.DataFrame(doctors).to_excel(DOCTORS_FILE, index=False)


seed_demo_data()

# ------------------------------------------------------------------
# SMALL UTILITIES
# ------------------------------------------------------------------
def hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()


def suggest_specialist(text: str) -> str:
    text = text.lower()
    for keyword, specialist in SYMPTOM_MAP.items():
        if keyword in text:
            return specialist
    return "General Physician"


def send_email(to_email: str, subject: str, body: str):
    """Tries to send a real email if SMTP credentials are set as environment
    variables (SMTP_EMAIL / SMTP_PASSWORD). Otherwise, just displays the email
    on-screen so the flow can still be tested end-to-end."""
    sender = os.environ.get("SMTP_EMAIL")
    password = os.environ.get("SMTP_PASSWORD")
    if sender and password:
        try:
            msg = MIMEText(body)
            msg["Subject"] = subject
            msg["From"] = sender
            msg["To"] = to_email
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(sender, password)
                server.sendmail(sender, [to_email], msg.as_string())
            return True
        except Exception as e:
            st.warning(f"Email could not be sent (showing preview instead). Error: {e}")
    with st.expander(f"📧 Email preview -> {to_email}"):
        st.write(f"**Subject:** {subject}")
        st.write(body)
    return False


def show_logo(width=140):
    if os.path.exists(LOGO_PATH):
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.image(LOGO_PATH, width=width)


def set_page(page_name: str):
    st.session_state.page = page_name


def inject_style():
    st.markdown(
        """
        <style>
        .stButton>button {
            border-radius: 10px;
            background: linear-gradient(90deg, #2f9e8f, #7fbf5f);
            color: white;
            font-weight: 600;
            border: none;
            padding: 0.5em 1.2em;
        }
        .stButton>button:hover {
            opacity: 0.9;
        }
        h1, h2, h3 {
            color: #1c2b4a;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ------------------------------------------------------------------
# SESSION STATE INIT
# ------------------------------------------------------------------
defaults = {
    "page": "landing",
    "logged_in": False,
    "user_email": None,
    "user_data": None,
    "symptom_text": "",
    "suggested_specialist": None,
    "booking_for": None,       # "self" or "other"
    "patient_details": None,   # dict
    "selected_hospital": None,
    "selected_specialist_row": None,
    "visit_date": None,
    "selected_slot": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

st.set_page_config(page_title=APP_NAME, page_icon="💚", layout="centered")
inject_style()

# ------------------------------------------------------------------
# PAGE: LANDING (choose Login / Register)
# ------------------------------------------------------------------
def page_landing():
    show_logo()
    st.markdown("<h1 style='text-align:center;'>CareSync</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;'>Connecting patients to hospitals & doctors</p>", unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Log In", use_container_width=True):
            set_page("login")
    with c2:
        if st.button("Register", use_container_width=True):
            set_page("register")

    st.write("---")
    st.caption("Or continue with")
    o1, o2, o3 = st.columns(3)
    with o1:
        if st.button("🔵 Google", use_container_width=True):
            st.info("Google Sign-In needs a real Google OAuth Client ID/Secret to be wired in here.")
    with o2:
        if st.button("🟦 Microsoft", use_container_width=True):
            st.info("Microsoft Sign-In needs a real Azure AD app registration to be wired in here.")
    with o3:
        if st.button("⬛ Apple", use_container_width=True):
            st.info("Apple Sign-In needs a real Apple Developer Services ID to be wired in here.")

    st.write("---")
    st.caption("Are you a hospital, clinic, or doctor?")
    r1, r2 = st.columns(2)
    with r1:
        if st.button("Register your Hospital / Clinic", use_container_width=True):
            set_page("register_hospital")
    with r2:
        if st.button("Register as a Doctor", use_container_width=True):
            set_page("register_doctor")


# ------------------------------------------------------------------
# PAGE: LOGIN
# ------------------------------------------------------------------
def page_login():
    show_logo(100)
    st.subheader("Log in to CareSync")
    users = load_table(USERS_FILE, [])
    email = st.text_input("Email")
    pw = st.text_input("Password", type="password")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Log In", use_container_width=True):
            if users.empty or "Email" not in users.columns:
                st.error("No registered users yet. Please register first.")
            else:
                match = users[(users["Email"] == email) & (users["Password"] == hash_password(pw))]
                if not match.empty:
                    st.session_state.logged_in = True
                    st.session_state.user_email = email
                    st.session_state.user_data = match.iloc[0].to_dict()
                    set_page("home_ai")
                    st.rerun()
                else:
                    st.error("Incorrect email or password.")
    with c2:
        if st.button("Back", use_container_width=True):
            set_page("landing")

    st.markdown("[Forgot password?](#)", help="Basic demo: password reset flow not wired to a real mail server yet.")
    if st.session_state.get("show_forgot"):
        pass
    if st.button("Forgot password", key="forgot_btn"):
        st.info("A password-reset link would be emailed to you here in the full version.")


# ------------------------------------------------------------------
# PAGE: REGISTER
# ------------------------------------------------------------------
def page_register():
    show_logo(100)
    st.subheader("Create your CareSync account")

    name = st.text_input("Full Name")
    age = st.number_input("Age", min_value=0, max_value=120, step=1)
    email = st.text_input("Email")
    mobile = st.text_input("Mobile Number")
    city = st.text_input("City")
    diseases = st.text_input("Any pre-diagnosed diseases (e.g. diabetes, thyroid) — optional")
    pw = st.text_input("Password", type="password")
    pw2 = st.text_input("Confirm Password", type="password")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Register", use_container_width=True):
            if not (name and email and mobile and city and pw):
                st.error("Please fill in all required fields.")
            elif pw != pw2:
                st.error("Passwords do not match.")
            else:
                users = load_table(USERS_FILE, [])
                if not users.empty and "Email" in users.columns and email in users["Email"].values:
                    st.error("An account with this email already exists.")
                else:
                    row = {
                        "Name": name, "Age": age, "Email": email, "Mobile": mobile,
                        "City": city, "PreExistingDiseases": diseases,
                        "Password": hash_password(pw),
                        "RegisteredOn": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    }
                    save_row(USERS_FILE, list(row.keys()), row)
                    st.success("Registration successful! Please log in.")
                    set_page("login")
                    st.rerun()
    with c2:
        if st.button("Back", use_container_width=True, key="reg_back"):
            set_page("landing")


# ------------------------------------------------------------------
# PAGE: HOME / AI SYMPTOM CHECKER
# ------------------------------------------------------------------
def page_home_ai():
    show_logo(100)
    st.subheader(f"Hi {st.session_state.user_data.get('Name', '')}, what's bothering you today?")
    text = st.text_area("Describe your issue / symptoms", value=st.session_state.symptom_text, height=120)
    st.session_state.symptom_text = text

    if st.button("Ask CareSync AI"):
        if not text.strip():
            st.error("Please describe your symptoms first.")
        else:
            st.session_state.suggested_specialist = suggest_specialist(text)

    if st.session_state.suggested_specialist:
        st.success(f"Based on what you described, we recommend seeing a **{st.session_state.suggested_specialist}**.")
        override = st.selectbox(
            "Not right? Choose a different specialist:",
            ["(keep recommendation)"] + SPECIALIST_LIST,
        )
        if override != "(keep recommendation)":
            st.session_state.suggested_specialist = override

        if st.button("Proceed ➜"):
            set_page("proceed_choice")
            st.rerun()

    st.write("---")
    if st.button("Log out"):
        for k in defaults:
            st.session_state[k] = defaults[k]
        st.rerun()


# ------------------------------------------------------------------
# PAGE: PROCEED CHOICE (self / other)
# ------------------------------------------------------------------
def page_proceed_choice():
    show_logo(100)
    st.subheader("Who is this booking for?")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Booking for Myself", use_container_width=True):
            st.session_state.booking_for = "self"
            u = st.session_state.user_data
            st.session_state.patient_details = {
                "Name": u.get("Name"), "Age": u.get("Age"), "City": u.get("City"),
                "Mobile": u.get("Mobile"), "Diseases": u.get("PreExistingDiseases"),
                "Email": st.session_state.user_email,
            }
            set_page("hospital_select")
            st.rerun()
    with c2:
        if st.button("Booking for Someone Else", use_container_width=True):
            st.session_state.booking_for = "other"
            set_page("patient_details_other")
            st.rerun()


# ------------------------------------------------------------------
# PAGE: PATIENT DETAILS (someone else)
# ------------------------------------------------------------------
def page_patient_details_other():
    show_logo(100)
    st.subheader("Patient details")
    name = st.text_input("Patient's Full Name")
    age = st.number_input("Age", min_value=0, max_value=120, step=1)
    city = st.text_input("City")
    mobile = st.text_input("Mobile Number")
    diseases = st.text_input("Any pre-diagnosed diseases (optional)")
    st.caption(f"Confirmation will be sent to your login email: {st.session_state.user_email}")

    if st.button("Submit"):
        if not (name and city and mobile):
            st.error("Please fill in all required fields.")
        else:
            st.session_state.patient_details = {
                "Name": name, "Age": age, "City": city, "Mobile": mobile,
                "Diseases": diseases, "Email": st.session_state.user_email,
            }
            set_page("hospital_select")
            st.rerun()


# ------------------------------------------------------------------
# PAGE: HOSPITAL SELECT
# ------------------------------------------------------------------
def page_hospital_select():
    show_logo(100)
    specialist = st.session_state.suggested_specialist
    st.subheader(f"Hospitals offering {specialist} near you")

    doctors = load_table(DOCTORS_FILE, [])
    hospitals = load_table(HOSPITALS_FILE, [])
    city = st.session_state.patient_details.get("City", "")

    if doctors.empty:
        st.warning("No doctors registered yet.")
        return

    matches = doctors[doctors["Speciality"] == specialist]
    if city:
        city_matches = matches[matches["City"].str.lower() == str(city).lower()]
        matches = city_matches if not city_matches.empty else matches

    if matches.empty:
        st.warning(f"No {specialist}s found. Try a different specialist or check back later.")
        return

    hospital_names = sorted(matches["Hospital"].unique())
    for h in hospital_names:
        info = hospitals[hospitals["CenterName"] == h]
        addr = info.iloc[0]["Address"] if not info.empty else ""
        with st.container(border=True):
            st.markdown(f"**{h}**")
            st.caption(addr)
            if st.button(f"Select {h}", key=f"sel_{h}"):
                st.session_state.selected_hospital = h
                set_page("specialist_select")
                st.rerun()


# ------------------------------------------------------------------
# PAGE: SPECIALIST + DATE + SLOT SELECT
# ------------------------------------------------------------------
def page_specialist_select():
    show_logo(100)
    hospital = st.session_state.selected_hospital
    specialist = st.session_state.suggested_specialist
    st.subheader(f"{specialist}s at {hospital}")

    doctors = load_table(DOCTORS_FILE, [])
    matches = doctors[(doctors["Hospital"] == hospital) & (doctors["Speciality"] == specialist)]

    if matches.empty:
        st.warning("No doctors available.")
        return

    for i, row in matches.iterrows():
        with st.container(border=True):
            st.markdown(f"**{row['Name']}** — ₹{row['Fee']} per visit")
            st.caption(f"Schedule: {row['Schedule']}")
            if st.button(f"Choose {row['Name']}", key=f"doc_{i}"):
                st.session_state.selected_specialist_row = row.to_dict()
                st.rerun()

    if st.session_state.selected_specialist_row:
        st.write("---")
        st.markdown(f"Selected doctor: **{st.session_state.selected_specialist_row['Name']}**")
        visit_date = st.date_input("Select date of visit", min_value=date.today())
        st.session_state.visit_date = visit_date
        slot = st.radio("Available time slots", TIME_SLOTS, horizontal=True)
        st.session_state.selected_slot = slot

        if st.button("Book"):
            set_page("payment")
            st.rerun()


# ------------------------------------------------------------------
# PAGE: PAYMENT
# ------------------------------------------------------------------
def page_payment():
    show_logo(100)
    doc = st.session_state.selected_specialist_row
    fee = float(doc["Fee"])
    advance = round(fee * 0.10, 2)

    st.subheader("Confirm & Pay")
    st.info(
        "Only 10% of the visit fee is payable online now to confirm this appointment. "
        "The remaining amount is payable offline at the hospital/clinic during your visit."
    )
    st.write(f"**Total consultation fee:** ₹{fee}")
    st.write(f"**Payable now (10%):** ₹{advance}")
    st.write(f"**Payable at center:** ₹{round(fee - advance, 2)}")

    st.write("---")
    st.caption("Choose payment method (demo only — no real transaction is made)")
    st.radio("Payment method", ["UPI", "Debit/Credit Card", "Net Banking"])

    if st.button(f"Pay ₹{advance} Now"):
        code = str(random.randint(1000, 9999))
        patient = st.session_state.patient_details
        row = {
            "BookingID": f"CS{random.randint(100000, 999999)}",
            "PatientName": patient["Name"], "Age": patient["Age"], "City": patient["City"],
            "Mobile": patient["Mobile"], "Diseases": patient.get("Diseases", ""),
            "Email": patient["Email"], "Specialist": doc["Speciality"],
            "DoctorName": doc["Name"], "Hospital": st.session_state.selected_hospital,
            "VisitDate": str(st.session_state.visit_date), "TimeSlot": st.session_state.selected_slot,
            "TotalFee": fee, "AdvancePaid": advance, "BalanceDue": round(fee - advance, 2),
            "ConfirmationCode": code, "BookedOn": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        save_row(APPOINTMENTS_FILE, list(row.keys()), row)

        body = (
            f"Thanks for booking with CareSync at {row['Hospital']} hospital with "
            f"{row['DoctorName']} ({row['Specialist']}) at {row['TimeSlot']} on {row['VisitDate']}.\n\n"
            f"Your appointment confirmation code is {code}.\n\n"
            f"Advance paid: Rs.{advance} | Balance payable at center: Rs.{round(fee - advance, 2)}"
        )
        send_email(patient["Email"], "CareSync Appointment Confirmation", body)
        doctor_email = doc.get("Email")
        if doctor_email:
            send_email(doctor_email, "New CareSync Appointment", body)

        st.session_state.last_confirmation = row
        set_page("confirmation")
        st.rerun()


# ------------------------------------------------------------------
# PAGE: CONFIRMATION
# ------------------------------------------------------------------
def page_confirmation():
    show_logo(100)
    row = st.session_state.get("last_confirmation")
    st.subheader("🎉 Booking Confirmed!")
    if row:
        st.write(f"Thanks for booking with **CareSync** at **{row['Hospital']}** with "
                  f"**{row['DoctorName']}** at **{row['TimeSlot']}** on **{row['VisitDate']}**.")
        st.write(f"### Confirmation Code: `{row['ConfirmationCode']}`")
        st.caption("A confirmation email has been sent to you and the doctor.")
    if st.button("Back to Home"):
        st.session_state.suggested_specialist = None
        st.session_state.symptom_text = ""
        set_page("home_ai")
        st.rerun()


# ------------------------------------------------------------------
# PAGE: REGISTER HOSPITAL / CLINIC
# ------------------------------------------------------------------
def page_register_hospital():
    show_logo(100)
    st.subheader("Register your Hospital / Clinic")

    center_name = st.text_input("Center Name")
    city = st.text_input("City")
    address = st.text_input("Address")
    contact = st.text_input("Contact Info (phone)")
    email = st.text_input("Email")

    st.write("---")
    st.caption("Add at least one doctor working at this center")
    doc_name = st.text_input("Doctor's Name")
    doc_speciality = st.selectbox("Doctor's Speciality", SPECIALIST_LIST)
    doc_fee = st.number_input("Consultation Fee (₹)", min_value=0, step=50)
    doc_schedule = st.text_input("Doctor's Available Schedule (e.g. Mon-Fri, 10AM-5PM)")
    doc_email = st.text_input("Doctor's Email (for booking notifications)")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Register Center", use_container_width=True):
            if not (center_name and city and address and contact and email):
                st.error("Please fill in all center details.")
            else:
                hrow = {"CenterName": center_name, "City": city, "Address": address,
                        "ContactInfo": contact, "Email": email}
                save_row(HOSPITALS_FILE, list(hrow.keys()), hrow)

                if doc_name and doc_speciality:
                    drow = {
                        "Name": doc_name, "Age": "", "ContactInfo": "", "Proof": "",
                        "Speciality": doc_speciality, "Hospital": center_name, "City": city,
                        "Fee": doc_fee, "Schedule": doc_schedule, "Email": doc_email,
                    }
                    save_row(DOCTORS_FILE, list(drow.keys()), drow)

                st.success(f"{center_name} registered successfully!")
                set_page("landing")
                st.rerun()
    with c2:
        if st.button("Back", use_container_width=True, key="hosp_back"):
            set_page("landing")


# ------------------------------------------------------------------
# PAGE: REGISTER DOCTOR (individual)
# ------------------------------------------------------------------
def page_register_doctor():
    show_logo(100)
    st.subheader("Register as a Doctor")

    name = st.text_input("Full Name")
    age = st.number_input("Age", min_value=20, max_value=100, step=1)
    contact = st.text_input("Contact Info")
    email = st.text_input("Email")
    proof = st.file_uploader("Upload proof of certification (PDF/Image)", type=["pdf", "png", "jpg", "jpeg"])
    speciality = st.selectbox("Speciality", SPECIALIST_LIST)
    hospitals = load_table(HOSPITALS_FILE, [])
    hospital_options = sorted(hospitals["CenterName"].unique()) if not hospitals.empty else []
    hospital = st.selectbox("Hospital / Clinic you work under", hospital_options) if hospital_options else st.text_input("Hospital / Clinic name")
    fee = st.number_input("Consultation Fee (₹)", min_value=0, step=50)
    schedule = st.text_input("Available Schedule (e.g. Mon-Fri, 10AM-5PM)")

    city = ""
    if hospital_options and hospital in hospital_options:
        city = hospitals[hospitals["CenterName"] == hospital].iloc[0]["City"]
    else:
        city = st.text_input("City")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Register", use_container_width=True):
            if not (name and contact and email and speciality and hospital):
                st.error("Please fill in all required fields.")
            else:
                row = {
                    "Name": name, "Age": age, "ContactInfo": contact,
                    "Proof": proof.name if proof else "not uploaded",
                    "Speciality": speciality, "Hospital": hospital, "City": city,
                    "Fee": fee, "Schedule": schedule, "Email": email,
                }
                save_row(DOCTORS_FILE, list(row.keys()), row)
                st.success("Doctor registered successfully! Awaiting verification.")
                set_page("landing")
                st.rerun()
    with c2:
        if st.button("Back", use_container_width=True, key="doc_back"):
            set_page("landing")


# ------------------------------------------------------------------
# ROUTER
# ------------------------------------------------------------------
PAGES = {
    "landing": page_landing,
    "login": page_login,
    "register": page_register,
    "home_ai": page_home_ai,
    "proceed_choice": page_proceed_choice,
    "patient_details_other": page_patient_details_other,
    "hospital_select": page_hospital_select,
    "specialist_select": page_specialist_select,
    "payment": page_payment,
    "confirmation": page_confirmation,
    "register_hospital": page_register_hospital,
    "register_doctor": page_register_doctor,
}

# Guard: pages that require login
LOGIN_REQUIRED = {"home_ai", "proceed_choice", "patient_details_other",
                   "hospital_select", "specialist_select", "payment", "confirmation"}

current_page = st.session_state.page
if current_page in LOGIN_REQUIRED and not st.session_state.logged_in:
    st.session_state.page = "landing"
    current_page = "landing"

PAGES[current_page]()
