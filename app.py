import streamlit as st

st.set_page_config(page_title="FFT Excel Formula Generator", layout="wide")

st.title("FFT Excel Formula Generator – Shake Table Analysis")

st.markdown("""
This tool generates **step-by-step Excel formulas** for FFT-based shake table analysis.
FFT is assumed to be performed **externally in Excel** using the Data Analysis ToolPak.
""")

# -----------------------------
# USER INPUTS
# -----------------------------
st.sidebar.header("Input Data Layout")

time_col = st.sidebar.text_input("Time column", "F")
signal_col = st.sidebar.text_input("Corrected acceleration column (X/Y/Z)", "G")
fft_col = st.sidebar.text_input("FFT output column (complex)", "K")

start_row = st.sidebar.number_input("Start row", value=2, step=1)
N = st.sidebar.number_input("Number of samples (N)", value=64, step=1)

motor_rpm = st.sidebar.number_input("Motor speed (RPM)", value=128)

helper_col = st.sidebar.text_input("Helper column (Δt, Fs, etc.)", "J")
index_col = st.sidebar.text_input("FFT index column", "M")
freq_col = st.sidebar.text_input("Frequency column", "N")
mag_col = st.sidebar.text_input("FFT magnitude column", "L")
filter_col = st.sidebar.text_input("Filtered FFT column", "P")
recon_col = st.sidebar.text_input("Reconstructed signal column", "Q")

st.divider()

# -----------------------------
# STEP DEFINITIONS
# -----------------------------

steps = []

# Step 1
steps.append({
    "title": "Arrange Experimental Data (Prerequisite)",
    "formula": None,
    "paste": None,
    "explanation": f"""
Time data must be in column **{time_col}**.
Corrected acceleration (mean removed) must be in column **{signal_col}**.
FFT is applied externally and output stored in **{fft_col}**.
"""
})

# Step 2
steps.append({
    "title": "Calculate Time Interval (Δt)",
    "formula": f"={time_col}{start_row+1}-{time_col}{start_row}",
    "paste": f"Paste in {helper_col}{start_row}",
    "explanation": "Calculates time difference between consecutive samples."
})

# Step 3
steps.append({
    "title": "Calculate Sampling Frequency (Fs)",
    "formula": f"=1/{helper_col}{start_row}",
    "paste": f"Paste in {helper_col}{start_row+1}",
    "explanation": "Sampling frequency is inverse of sampling interval."
})

# Step 4
steps.append({
    "title": "Fix Number of Samples (N)",
    "formula": f"{N}",
    "paste": f"Enter manually in {helper_col}{start_row+2}",
    "explanation": "FFT requires number of samples to be a power of 2."
})

# Step 5
steps.append({
    "title": "Apply FFT (External Step)",
    "formula": None,
    "paste": None,
    "explanation": f"""
Apply FFT using Excel Data Analysis ToolPak on column **{signal_col}**.
Store FFT output (complex numbers) in column **{fft_col}**.
"""
})

# Step 6
steps.append({
    "title": "Generate FFT Index (Bin Number)",
    "formula": f"=ROW({index_col}{start_row})-ROW(${index_col}${start_row})",
    "paste": f"Paste in {index_col}{start_row} and drag down to row {start_row+N-1}",
    "explanation": "Generates FFT bin numbers from 0 to N−1."
})

# Step 7
steps.append({
    "title": "Calculate Frequency Resolution (Δf)",
    "formula": f"={helper_col}{start_row+1}/{helper_col}{start_row+2}",
    "paste": f"Paste in {helper_col}{start_row+3}",
    "explanation": "Computes frequency spacing between FFT bins."
})

# Step 8
steps.append({
    "title": "Create Frequency Axis",
    "formula": f"={index_col}{start_row}*{helper_col}{start_row+3}",
    "paste": f"Paste in {freq_col}{start_row} and drag down to row {start_row+N-1}",
    "explanation": "Assigns real frequency values to FFT bins."
})

# Step 9
steps.append({
    "title": "Calculate Motor Excitation Frequency",
    "formula": f"={motor_rpm}/60",
    "paste": f"Paste in {helper_col}{start_row+4}",
    "explanation": "Converts motor speed from RPM to Hz."
})

# Step 10
steps.append({
    "title": "Extract FFT Magnitude",
    "formula": f"=IMABS({fft_col}{start_row})",
    "paste": f"Paste in {mag_col}{start_row} and drag down to row {start_row+N-1}",
    "explanation": "Computes magnitude of FFT complex values."
})

# Step 11
steps.append({
    "title": "Identify Dominant FFT Bin Near Excitation",
    "formula": f"=ABS({freq_col}{start_row}-{helper_col}{start_row+4})",
    "paste": f"Paste in column O and drag down",
    "explanation": "Finds FFT bin closest to motor excitation frequency."
})

# Step 12
steps.append({
    "title": "Extract Dominant FFT Magnitude",
    "formula": f"=IMABS(INDEX({fft_col}:{fft_col},MATCH(MIN(O:O),O:O,0)))",
    "paste": f"Paste in {helper_col}{start_row+5}",
    "explanation": "Extracts FFT magnitude at dominant frequency."
})

# Step 13
steps.append({
    "title": "Calculate Time-Domain Amplitude",
    "formula": f"=2*{helper_col}{start_row+5}/{helper_col}{start_row+2}",
    "paste": f"Paste in {helper_col}{start_row+6}",
    "explanation": "Converts FFT magnitude into physical acceleration amplitude."
})

# Step 14
steps.append({
    "title": "Reconstruct Time-Domain Signal",
    "formula": f"={helper_col}{start_row+6}*SIN(2*PI()*{helper_col}{start_row+4}*{time_col}{start_row})",
    "paste": f"Paste in {recon_col}{start_row} and drag down",
    "explanation": "Reconstructs clean sinusoidal response in time domain."
})

# -----------------------------
# DISPLAY STEPS
# -----------------------------
for i, step in enumerate(steps, start=1):
    st.subheader(f"Step {i}: {step['title']}")
    if step["formula"]:
        st.code(step["formula"], language="excel")
        st.markdown(f"📍 **Paste:** {step['paste']}")
    st.markdown(step["explanation"])
    st.divider()
