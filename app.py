import streamlit as st

st.set_page_config(
    page_title="FFT Excel Formula Generator",
    layout="wide"
)

st.title("FFT Excel Formula Generator – Shake Table Analysis")

st.markdown("""
This app generates **step-by-step Excel formulas** for FFT-based shake table analysis.

⚠️ Assumptions:
- Time and **corrected** acceleration data already exist
- FFT is performed **externally** using Excel Data Analysis ToolPak
- FFT output is **one complex column**
""")

# =========================
# USER INPUTS
# =========================
st.sidebar.header("Excel Column Mapping")

time_col = st.sidebar.text_input("Time column", "F")
signal_col = st.sidebar.text_input("Corrected acceleration column (X/Y/Z)", "G")
fft_col = st.sidebar.text_input("FFT output column (complex)", "K")

index_col = st.sidebar.text_input("FFT index column", "M")
freq_col = st.sidebar.text_input("Frequency column", "N")
mag_col = st.sidebar.text_input("FFT magnitude column", "L")
filter_col = st.sidebar.text_input("Filtered FFT column", "P")
recon_col = st.sidebar.text_input("Reconstructed signal column", "Q")

helper_col = st.sidebar.text_input(
    "Helper / calculation column (must be empty)", "J"
)

start_row = st.sidebar.number_input("Start row", value=2, step=1)
N = st.sidebar.number_input("Number of samples (N)", value=64, step=1)
motor_rpm = st.sidebar.number_input("Motor speed (RPM)", value=128)

st.divider()

# =========================
# STEP DEFINITIONS
# =========================
steps = []

# STEP 1
steps.append({
    "title": "Arrange Experimental Data (Prerequisite)",
    "formula": None,
    "paste": None,
    "explanation": f"""
• Time → Column **{time_col}**  
• Corrected acceleration → Column **{signal_col}**  
• FFT output (complex) → Column **{fft_col}**  

FFT must be performed externally using Excel Data Analysis ToolPak.
"""
})

# STEP 2
steps.append({
    "title": "Calculate Time Interval (Δt)",
    "formula": f"={time_col}{start_row+1}-{time_col}{start_row}",
    "paste": f"Paste in {helper_col}{start_row}",
    "explanation": "Computes time difference between two consecutive samples."
})

# STEP 3
steps.append({
    "title": "Calculate Sampling Frequency (Fs)",
    "formula": f"=1/{helper_col}{start_row}",
    "paste": f"Paste in {helper_col}{start_row+1}",
    "explanation": "Sampling frequency is the inverse of sampling interval."
})

# STEP 4
steps.append({
    "title": "Fix Number of Samples (N)",
    "formula": f"{N}",
    "paste": f"Enter manually in {helper_col}{start_row+2}",
    "explanation": "FFT requires number of samples to be a power of 2."
})

# STEP 5
steps.append({
    "title": "Generate FFT Index (Bin Number)",
    "formula": f"=ROW({index_col}{start_row})-ROW(${index_col}${start_row})",
    "paste": f"Paste in {index_col}{start_row} and drag down to row {start_row+N-1}",
    "explanation": "Generates FFT bin numbers from 0 to N−1."
})

# STEP 6
steps.append({
    "title": "Calculate Frequency Resolution (Δf)",
    "formula": f"={helper_col}{start_row+1}/{helper_col}{start_row+2}",
    "paste": f"Paste in {helper_col}{start_row+3}",
    "explanation": "Computes frequency spacing between FFT bins."
})

# STEP 7
steps.append({
    "title": "Create Frequency Axis",
    "formula": f"={index_col}{start_row}*{helper_col}{start_row+3}",
    "paste": f"Paste in {freq_col}{start_row} and drag down to row {start_row+N-1}",
    "explanation": "Assigns a physical frequency value to each FFT bin."
})

# STEP 8
steps.append({
    "title": "Calculate Motor Excitation Frequency",
    "formula": f"={motor_rpm}/60",
    "paste": f"Paste in {helper_col}{start_row+4}",
    "explanation": "Converts motor speed from RPM to Hz."
})

# STEP 9
steps.append({
    "title": "Extract FFT Magnitude",
    "formula": f"=IMABS({fft_col}{start_row})",
    "paste": f"Paste in {mag_col}{start_row} and drag down to row {start_row+N-1}",
    "explanation": "Computes magnitude of FFT complex values."
})

# STEP 10
steps.append({
    "title": "Lower Band-Pass Limit (−10%)",
    "formula": f"=0.9*{helper_col}{start_row+4}",
    "paste": f"Paste in {helper_col}{start_row+5}",
    "explanation": "Lower cutoff frequency for band-pass filtering."
})

# STEP 11
steps.append({
    "title": "Upper Band-Pass Limit (+10%)",
    "formula": f"=1.1*{helper_col}{start_row+4}",
    "paste": f"Paste in {helper_col}{start_row+6}",
    "explanation": "Upper cutoff frequency for band-pass filtering."
})

# STEP 12
steps.append({
    "title": "Apply Frequency-Domain Band-Pass Filter",
    "formula": (
        f"=IF(AND({freq_col}{start_row}>={helper_col}{start_row+5},"
        f"{freq_col}{start_row}<={helper_col}{start_row+6}),"
        f"{fft_col}{start_row},0)"
    ),
    "paste": f"Paste in {filter_col}{start_row} and drag down to row {start_row+N-1}",
    "explanation": "Retains FFT values only within ±10% of motor frequency."
})

# STEP 13
steps.append({
    "title": "Verify Single Non-Zero FFT Bin",
    "formula": f"=COUNTIF({filter_col}:{filter_col},\"<>0\")",
    "paste": f"Paste in {helper_col}{start_row+7}",
    "explanation": "Confirms that only one FFT bin remains after filtering."
})

# STEP 14
steps.append({
    "title": "Extract Dominant FFT Magnitude",
    "formula": f"=MAX(IMABS({filter_col}:{filter_col}))",
    "paste": f"Paste in {helper_col}{start_row+8}",
    "explanation": "Extracts magnitude of the isolated frequency component."
})

# STEP 15
steps.append({
    "title": "Calculate Time-Domain Amplitude",
    "formula": f"=2*{helper_col}{start_row+8}/{helper_col}{start_row+2}",
    "paste": f"Paste in {helper_col}{start_row+9}",
    "explanation": "Converts FFT magnitude into physical acceleration amplitude."
})

# STEP 16
steps.append({
    "title": "Reconstruct Time-Domain Acceleration",
    "formula": (
        f"={helper_col}{start_row+9}"
        f"*SIN(2*PI()*{helper_col}{start_row+4}*{time_col}{start_row})"
    ),
    "paste": f"Paste in {recon_col}{start_row} and drag down",
    "explanation": "Reconstructs clean sinusoidal acceleration response."
})

# =========================
# DISPLAY STEPS
# =========================
for i, step in enumerate(steps, start=1):
    st.subheader(f"Step {i}: {step['title']}")
    if step["formula"]:
        st.code(step["formula"], language="excel")
        st.markdown(f"📍 **Paste:** {step['paste']}")
    st.markdown(step["explanation"])
    st.divider()
