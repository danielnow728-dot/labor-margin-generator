import os
from tempfile import NamedTemporaryFile
from fastapi import FastAPI, UploadFile, File, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from processor import build_report

app = FastAPI(title="Labor Margin Report API")

# Allow CORS for local development if frontend is served differently
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Adjust this in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Helper function to save an uploaded file to a temporary location
def save_temp_file(upload_file: UploadFile) -> str:
    # Use delete=False so we can read it later by path
    temp = NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp.write(upload_file.file.read())
    temp.close()
    return temp.name

# Helper function to clean up files after request
def cleanup_files(*file_paths: str):
    for path in file_paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception as e:
                print(f"Error cleaning up file {path}: {e}")

@app.post("/generate-report")
async def generate_report(
    background_tasks: BackgroundTasks,
    gl: UploadFile = File(...),
    inventory: UploadFile = File(...),
    cost: UploadFile = File(...),
    master: UploadFile = File(...)
):
    # Save Uploaded Files Temporarily
    gl_path = save_temp_file(gl)
    inv_path = save_temp_file(inventory)
    cost_path = save_temp_file(cost)
    master_path = save_temp_file(master)
    
    # An output path for the finalized Excel
    output_path = NamedTemporaryFile(delete=False, suffix=".xlsx").name

    # We register a background task to delete all files AFTER the response is sent back to the user
    background_tasks.add_task(cleanup_files, gl_path, inv_path, cost_path, master_path, output_path)

    try:
        # Pass the temp files to your existing Pandas Logic
        final_file_path = build_report(
            gl_path=gl_path,
            inventory_path=inv_path,
            job_cost_path=cost_path,
            job_master_path=master_path,
            output_path=output_path
        )

        return FileResponse(
            path=final_file_path,
            filename="Labor_Margin_Report.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        # If it fails during building, we still need to cleanup the input temp files
        cleanup_files(gl_path, inv_path, cost_path, master_path, output_path)
        from fastapi import HTTPException
        raise HTTPException(status_code=400, detail=str(e))

# Mount the static frontend directory right into FastAPI for easy centralized deployment
# Access at http://localhost:8000/
app.mount("/", StaticFiles(directory="frontend", html=True), name="frontend")
