import os
import shutil
import tempfile
import importlib.util
from concurrent.futures import ProcessPoolExecutor, as_completed
import pandas as pd
import streamlit as st

# Dynamically load 'OCR HC.py'
spec = importlib.util.spec_from_file_location("ocr_hc", "OCR HC.py")
ocr_hc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ocr_hc)

st.set_page_config(page_title="Telecom OCR Processor", layout="wide")
st.title("📡 Parallel In-Memory OCR Processor")
st.markdown("Upload **RAR / ZIP** archives containing **G-NetTrack** & **Speedtest** screenshots/videos to extract KPIs to CSV.")

# Detect Linux binaries for Tesseract and 7-Zip
tesseract_bin = shutil.which("tesseract") or (ocr_hc.TESSERACT_PATH if ocr_hc.configure_tesseract() else None)
seven_zip_bin = shutil.which("7z") or shutil.which("7za") or ocr_hc.find_7zip()

if ocr_hc.rarfile is not None and seven_zip_bin:
    ocr_hc.rarfile.UNRAR_TOOL = seven_zip_bin

with st.sidebar:
    st.header("⚙️ OCR Settings")
    # Default to 2 workers on Hugging Face's free 2-vCPU tier, with option up to 10
    max_workers = st.selectbox("Select OCR_MAX_WORKERS", options=[2, 4, 6, 8, 10], index=0)
    st.caption(f"Tesseract: `{'Ready' if tesseract_bin else 'Missing'}`")
    st.caption(f"7-Zip / RAR: `{'Ready' if seven_zip_bin else 'Missing'}`")

uploaded_files = st.file_uploader(
    "Drag and drop RAR or ZIP files here",
    type=["rar", "zip"],
    accept_multiple_files=True
)

if uploaded_files and st.button("🚀 Start Processing", type="primary"):
    if not tesseract_bin:
        st.error("Tesseract OCR executable not found on server.")
    else:
        with tempfile.TemporaryDirectory() as temp_dir:
            tasks = []
            for uf in uploaded_files:
                file_path = os.path.join(temp_dir, uf.name)
                with open(file_path, "wb") as f:
                    f.write(uf.getbuffer())
                tasks.append((file_path, tesseract_bin, seven_zip_bin))

            tasks.sort(key=lambda x: os.path.basename(x[0]).lower())
            worker_count = min(max_workers, len(tasks))

            st.info(f"Processing **{len(tasks)}** archive(s) with **{worker_count}** worker(s)...")
            progress_bar = st.progress(0)
            status_text = st.empty()

            results = []
            completed = 0

            with ProcessPoolExecutor(max_workers=worker_count) as executor:
                futures = {
                    executor.submit(ocr_hc.process_archive_task, task): task[0]
                    for task in tasks
                }

                for future in as_completed(futures):
                    archive_path = futures[future]
                    try:
                        archive_id, res = future.result()
                        results.append(res)
                        completed += 1
                        progress_bar.progress(completed / len(tasks))
                        status_text.text(f"[{completed}/{len(tasks)}] Completed: {archive_id}")
                    except Exception as e:
                        st.error(f"Error processing {os.path.basename(archive_path)}: {e}")

            if results:
                results.sort(key=lambda r: str(r.get("RAR ID", "")).lower())
                csv_path = os.path.join(temp_dir, ocr_hc.OUTPUT_CSV)
                ocr_hc.write_csv(csv_path, results)

                st.success(f"✅ Processed {len(results)} archive(s)!")
                df = pd.DataFrame(results)
                st.dataframe(df, use_container_width=True)

                with open(csv_path, "rb") as f:
                    st.download_button(
                        label="📥 Download ocr_result.csv",
                        data=f.read(),
                        file_name="ocr_result.csv",
                        mime="text/csv",
                        type="primary"
                    )