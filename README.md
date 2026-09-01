# CareSync — Patient ⇄ Hospital ⇄ Doctor Booking App (Prototype)

A basic Python prototype of the CareSync flow you described, built with
[Streamlit](https://streamlit.io) (pure Python, no HTML/CSS/JS needed) and
Excel files as the data store (via `pandas` + `openpyxl`).

## What's included

- **Landing page** — Log In / Register, social login buttons (Google, Microsoft,
  Apple — placeholders, see "What still needs real setup" below), plus
  "Register your Hospital/Clinic" and "Register as a Doctor".
- **Login** — email + password, with a "Forgot password" link (placeholder).
- **Register** — Name, Age, Email, Mobile, City, pre-diagnosed diseases,
  password. Saved to `data/users.xlsx`.
- **AI symptom checker (basic)** — keyword-based specialist suggestion
  (e.g. "chest pain" → Cardiologist). You can swap this for a real LLM call later.
- **Booking flow** — self / someone else → patient details → hospital list
  (filtered by specialist + city) → doctor + fee → date + time slot → Book.
- **Payment page** — shows total fee, 10% payable online now, balance payable
  offline, with a note at the top exactly as you asked. Mock "Pay Now" button
  (no real payment gateway wired in — see below).
- **Confirmation** — random 4-digit confirmation code, on-screen + emailed to
  patient and doctor.
- **Hospital/Clinic registration** — center name, city, address, contact,
  email, plus at least one doctor + schedule. Saved to `data/hospitals.xlsx`
  and `data/doctors.xlsx`.
- **Doctor self-registration** — name, age, contact, proof upload, speciality,
  which hospital/clinic they work under, schedule, fee.

All data lives in `data/*.xlsx` — open these directly in Excel to see/edit records.
A few demo hospitals/doctors are pre-seeded so you can test the whole booking
flow immediately.

## Running it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## What still needs real setup (can't be faked in a basic prototype)

1. **Google / Microsoft / Apple login** — needs a registered OAuth app with
   each provider (Client ID + Secret, redirect URLs). I left clearly-labeled
   placeholder buttons; happy to wire in real OAuth (e.g. via `streamlit-oauth`
   or `authlib`) once you have those credentials.
2. **Real online payments** — needs a payment gateway account (Razorpay,
   Stripe, PayU, etc.) and their API keys. The "Pay Now" button currently just
   records the booking and moves to confirmation — no money actually moves.
3. **Real emails** — the app tries to send via Gmail SMTP if you set two
   environment variables before running:
   ```bash
   export SMTP_EMAIL="youraddress@gmail.com"
   export SMTP_PASSWORD="your-16-char-app-password"
   ```
   Without these, it just shows the email content on-screen so you can still
   test the flow.
4. **Password reset** — the "Forgot password" link is a placeholder; a real
   version needs the email sending above plus a reset-token flow.
5. **Doctor certification verification** — currently just stores the uploaded
   file name; a real system would need an admin review step.

## Suggested next steps

- Swap the keyword-based symptom checker for a real AI call (e.g. the
  Anthropic API) for smarter specialist suggestions.
- Move from Excel files to a proper database (SQLite to start, then
  Postgres) once you have more than a handful of concurrent users — Excel
  files aren't safe for simultaneous read/write from many users.
- Add basic input validation (e.g. valid email/mobile format) and stronger
  password rules.
