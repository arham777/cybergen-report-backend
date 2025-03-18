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
async def upload_files(files: list[UploadFile], background_tasks: BackgroundTasks):
    """
    Upload multiple DOCX or PDF files for processing.
    """
    try:
        # Check memory usage before accepting new files
        memory_status = check_memory_usage()
        if not memory_status["ok"]:
            raise HTTPException(
                status_code=503,
                detail="Server is currently under heavy load. Please try again later."
            )
        
        # Generate unique job ID
        job_id = str(uuid.uuid4())
        
        # Create job directories
        job_upload_dir = UPLOAD_DIR / job_id
        job_output_dir = OUTPUT_DIR / job_id
        job_upload_dir.mkdir(exist_ok=True, parents=True)
        job_output_dir.mkdir(exist_ok=True, parents=True)
        
        processed_files = []
        error_files = []
        
        # Process each file
        for file in files:
            try:
                # Check file size
                file_size = 0
                chunk_size = 1024 * 1024  # 1MB chunks
                while chunk := await file.read(chunk_size):
                    file_size += len(chunk)
                    if file_size > FILE_SIZE_LIMIT:
                        error_files.append({
                            "filename": file.filename,
                            "error": "File too large. Free tier limited to 10MB files."
                        })
                        break
                
                # Reset file position
                await file.seek(0)
                
                # Skip if file is too large
                if file_size > FILE_SIZE_LIMIT:
                    continue
                
                # Validate file type
                if not file.filename.lower().endswith(('.docx', '.pdf')):
                    error_files.append({
                        "filename": file.filename,
                        "error": "Invalid file type. Only DOCX and PDF files are supported."
                    })
                    continue
                
                # Save uploaded file
                file_path = job_upload_dir / file.filename
                try:
                    with open(file_path, "wb") as buffer:
                        shutil.copyfileobj(file.file, buffer)
                    processed_files.append(file.filename)
                except Exception as e:
                    error_files.append({
                        "filename": file.filename,
                        "error": f"Failed to save file: {str(e)}"
                    })
                    continue
                
            except Exception as e:
                error_files.append({
                    "filename": file.filename,
                    "error": str(e)
                })
        
        # If no files were processed successfully
        if not processed_files:
            # Clean up created directories
            shutil.rmtree(job_upload_dir, ignore_errors=True)
            shutil.rmtree(job_output_dir, ignore_errors=True)
            raise HTTPException(
                status_code=400,
                detail={
                    "message": "No files were processed successfully",
                    "errors": error_files
                }
            )
        
        # Initialize job information
        JOBS[job_id] = {
            "status": JobStatus.PENDING,
            "created_at": datetime.utcnow().isoformat(),
            "input_files": processed_files,
            "output_files": [],
            "error_files": error_files,
            "error": None
        }
        
        # Start processing in background
        background_tasks.add_task(
            process_document,
            job_id,
            job_upload_dir,
            job_output_dir,
            JOBS
        )
        
        return {
            "job_id": job_id,
            "status": "Processing started",
            "processed_files": processed_files,
            "error_files": error_files,
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

@app.get("/download/{job_id}")
async def download_file(job_id: str):
    """
    Download the first processed file from a job.
    If multiple files were processed, use /download-all endpoint.
    """
    try:
        # Check if job exists
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(
                status_code=404,
                detail="Job not found"
            )
        
        # Check job status
        if job["status"] != JobStatus.COMPLETED:
            raise HTTPException(
                status_code=400,
                detail=f"Job processing not completed. Current status: {job['status']}"
            )
        
        # Check if there are any output files
        if not job.get("output_files"):
            raise HTTPException(
                status_code=404,
                detail="No processed files found for this job"
            )
        
        # Get the first processed file
        filename = job["output_files"][0]
        file_path = OUTPUT_DIR / job_id / filename
        
        # Check if file exists
        if not file_path.exists():
            raise HTTPException(
                status_code=404,
                detail=f"File not found at path: {file_path}"
            )
        
        # Determine media type based on file extension
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if filename.lower().endswith('.pdf'):
            media_type = "application/pdf"
        
        return FileResponse(
            path=str(file_path),
            filename=filename,
            media_type=media_type,
            headers={
                "Cache-Control": "no-cache, no-store, must-revalidate",
                "Pragma": "no-cache",
                "Expires": "0"
            }
        )
    except Exception as e:
        print(f"Error in download_file: {str(e)}")  # Log the error
        raise HTTPException(
            status_code=500,
            detail=f"Error downloading file: {str(e)}"
        )

@app.get("/download-all/{job_id}")
async def download_all(job_id: str):
    """
    Download all processed files as a zip file.
    """
    try:
        # Check if job exists
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(
                status_code=404,
                detail="Job not found"
            )
        
        # Check job status
        if job["status"] != JobStatus.COMPLETED:
            raise HTTPException(
                status_code=400,
                detail=f"Job processing not completed. Current status: {job['status']}"
            )
        
        # Check if there are any output files
        if not job.get("output_files"):
            raise HTTPException(
                status_code=404,
                detail="No processed files found for this job"
            )
        
        # Create zip file
        zip_filename = f"processed_files_{job_id}.zip"
        job_output_dir = OUTPUT_DIR / job_id
        zip_path = job_output_dir / zip_filename
        
        # Ensure the output directory exists
        job_output_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Add all processed files to zip
                for filename in job["output_files"]:
                    file_path = job_output_dir / filename
                    if file_path.exists():
                        # Add file to zip with just the filename
                        zipf.write(file_path, arcname=filename)
                        print(f"Added {filename} to zip")
                    else:
                        print(f"Warning: File not found: {file_path}")
                
                # Add images if they exist
                images_dir = job_output_dir / "images"
                if images_dir.exists():
                    for image_dir in images_dir.iterdir():
                        if image_dir.is_dir():  # Each input file has its own image directory
                            for image_file in image_dir.glob("*.*"):
                                image_zip_path = f"images/{image_dir.name}/{image_file.name}"
                                zipf.write(image_file, arcname=image_zip_path)
                                print(f"Added {image_zip_path} to zip")
                
                # If there were any error files, add an error report
                if job.get("error_files"):
                    error_report = "Error Report.txt"
                    error_content = "\n".join(
                        f"File: {err['filename']}\nError: {err['error']}\n"
                        for err in job["error_files"]
                    )
                    zipf.writestr(error_report, error_content)
                    print("Added error report to zip")
        except Exception as zip_error:
            # Clean up the zip file if it exists and there was an error
            if zip_path.exists():
                zip_path.unlink()
            raise Exception(f"Failed to create zip file: {str(zip_error)}")
        
        # Verify the zip file was created and has content
        if not zip_path.exists() or zip_path.stat().st_size == 0:
            raise HTTPException(
                status_code=500,
                detail="Failed to create valid zip file"
            )
        
        print(f"Successfully created zip file with {len(job['output_files'])} documents")
        
        # Return the zip file
        return FileResponse(
            path=str(zip_path),
            filename=zip_filename,
            media_type="application/zip",
            headers={
                "Cache-Control": "no-cache, no-store, must-revalidate",
                "Pragma": "no-cache",
                "Expires": "0"
            }
        )
    except Exception as e:
        print(f"Error in download_all: {str(e)}")  # Log the error
        raise HTTPException(
            status_code=500,
            detail=f"Error creating zip file: {str(e)}"
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