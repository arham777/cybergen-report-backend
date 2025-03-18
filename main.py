from fastapi import FastAPI, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
import uuid
from datetime import datetime
import zipfile
from typing import Dict, Optional
import asyncio
from pathlib import Path
import psutil  # Add this import for memory monitoring

# Import our document processing module
from document_processor import process_document, JobStatus, get_job_status, cleanup_job

# Memory limit for free tier (450MB to leave some headroom)
MEMORY_LIMIT = 450 * 1024 * 1024  # 450MB in bytes
FILE_SIZE_LIMIT = 10 * 1024 * 1024  # 10MB for free tier

def check_memory_usage():
    """Check if memory usage is approaching free tier limit"""
    try:
        process = psutil.Process(os.getpid())
        memory_use = process.memory_info().rss
        return {
            "ok": memory_use < MEMORY_LIMIT,
            "current_usage": memory_use,
            "limit": MEMORY_LIMIT,
            "percentage": (memory_use / MEMORY_LIMIT) * 100
        }
    except Exception as e:
        print(f"Error checking memory: {str(e)}")
        return {"ok": True}  # Default to True if can't check

app = FastAPI(
    title="Document Processing API",
    description="API for processing DOCX and PDF files using CyberGen template",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

# Store job information
JOBS: Dict[str, Dict] = {}

# Get base directory from environment variable or use current directory
BASE_DIR = Path(os.getenv("RENDER_WORKSPACE", os.getcwd()))
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

# Create necessary directories
UPLOAD_DIR.mkdir(exist_ok=True, parents=True)
OUTPUT_DIR.mkdir(exist_ok=True, parents=True)

# More aggressive cleanup for free tier
def cleanup_old_files():
    """Clean up files older than 1 hour for free tier"""
    try:
        current_time = datetime.utcnow()
        # Clean uploads directory
        if UPLOAD_DIR.exists():
            for job_dir in UPLOAD_DIR.iterdir():
                if job_dir.is_dir():
                    dir_time = datetime.fromtimestamp(job_dir.stat().st_mtime)
                    # Changed to 1 hour instead of 24 hours for free tier
                    if (current_time - dir_time).total_seconds() > 3600:  # 1 hour
                        shutil.rmtree(job_dir)
        
        # Clean outputs directory
        if OUTPUT_DIR.exists():
            for job_dir in OUTPUT_DIR.iterdir():
                if job_dir.is_dir():
                    dir_time = datetime.fromtimestamp(job_dir.stat().st_mtime)
                    if (current_time - dir_time).total_seconds() > 3600:  # 1 hour
                        shutil.rmtree(job_dir)
    except Exception as e:
        print(f"Error during cleanup: {str(e)}")

@app.on_event("startup")
async def startup_event():
    """Run startup tasks"""
    cleanup_old_files()

@app.get("/")
async def root():
    """Root endpoint for health check"""
    memory_status = check_memory_usage()
    return {
        "status": "healthy",
        "message": "Document Processing API is running",
        "memory_status": memory_status
    }

@app.post("/upload-files/")
async def upload_files(file: UploadFile, background_tasks: BackgroundTasks):
    """
    Upload a DOCX or PDF file for processing.
    """
    try:
        # Check memory usage before accepting new file
        memory_status = check_memory_usage()
        if not memory_status["ok"]:
            raise HTTPException(
                status_code=503,
                detail="Server is currently under heavy load. Please try again later."
            )
            
        # Check file size (limit to 10MB for free tier)
        file_size = 0
        chunk_size = 1024 * 1024  # 1MB chunks
        while chunk := await file.read(chunk_size):
            file_size += len(chunk)
            if file_size > FILE_SIZE_LIMIT:
                raise HTTPException(
                    status_code=413,
                    detail="File too large. Free tier limited to 10MB files."
                )
        
        # Reset file position for later reading
        await file.seek(0)
        
        # Validate file type
        if not file.filename.lower().endswith(('.docx', '.pdf')):
            raise HTTPException(
                status_code=400,
                detail="Only DOCX and PDF files are supported"
            )
        
        # Generate unique job ID
        job_id = str(uuid.uuid4())
        
        # Create job directories
        job_upload_dir = UPLOAD_DIR / job_id
        job_output_dir = OUTPUT_DIR / job_id
        job_upload_dir.mkdir(exist_ok=True, parents=True)
        job_output_dir.mkdir(exist_ok=True, parents=True)
        
        # Save uploaded file
        file_path = job_upload_dir / file.filename
        try:
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
        except Exception as e:
            # Clean up created directories if file save fails
            shutil.rmtree(job_upload_dir, ignore_errors=True)
            shutil.rmtree(job_output_dir, ignore_errors=True)
            raise HTTPException(
                status_code=500,
                detail=f"Failed to save uploaded file: {str(e)}"
            )
        
        # Initialize job information
        JOBS[job_id] = {
            "status": JobStatus.PENDING,
            "created_at": datetime.utcnow().isoformat(),
            "input_file": file.filename,
            "output_files": [],
            "error": None
        }
        
        # Start processing in background
        background_tasks.add_task(
            process_document,
            job_id,
            file_path,
            job_output_dir,
            JOBS
        )
        
        return {
            "job_id": job_id, 
            "status": "Processing started",
            "memory_status": memory_status
        }
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )

@app.get("/job-status/{job_id}")
async def job_status(job_id: str):
    """
    Get the status of a processing job.
    """
    try:
        status = get_job_status(job_id, JOBS)
        if status is None:
            raise HTTPException(
                status_code=404,
                detail="Job not found"
            )
        return status
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )

@app.get("/download/{job_id}/{filename}")
async def download_file(job_id: str, filename: str):
    """
    Download a specific processed file.
    """
    try:
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(
                status_code=404,
                detail="Job not found"
            )
        
        if job["status"] != JobStatus.COMPLETED:
            raise HTTPException(
                status_code=400,
                detail="Job processing not completed"
            )
        
        file_path = OUTPUT_DIR / job_id / filename
        if not file_path.exists():
            raise HTTPException(
                status_code=404,
                detail="File not found"
            )
        
        return FileResponse(
            path=str(file_path),
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )

@app.get("/download-all/{job_id}")
async def download_all(job_id: str):
    """
    Download all processed files as a zip file.
    """
    try:
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(
                status_code=404,
                detail="Job not found"
            )
        
        if job["status"] != JobStatus.COMPLETED:
            raise HTTPException(
                status_code=400,
                detail="Job processing not completed"
            )
        
        # Create zip file
        zip_filename = f"processed_files_{job_id}.zip"
        zip_path = OUTPUT_DIR / job_id / zip_filename
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for filename in job["output_files"]:
                file_path = OUTPUT_DIR / job_id / filename
                if file_path.exists():
                    zipf.write(file_path, filename)
        
        return FileResponse(
            path=str(zip_path),
            filename=zip_filename,
            media_type="application/zip"
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )

@app.delete("/job/{job_id}")
async def delete_job(job_id: str):
    """
    Delete a job and its associated files.
    """
    try:
        if job_id not in JOBS:
            raise HTTPException(
                status_code=404,
                detail="Job not found"
            )
        
        # Clean up job files and data
        await cleanup_job(job_id, JOBS, UPLOAD_DIR, OUTPUT_DIR)
        
        return {"status": "Job deleted successfully"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 