# Document Processing API

This FastAPI-based API provides document processing capabilities for DOCX and PDF files. It converts and formats documents according to a specified template.

## Features

- Upload DOCX and PDF files for processing
- Convert PDF files to DOCX format
- Apply consistent formatting to documents
- Track processing status
- Download processed files
- Clean up job files

## API Endpoints

- `POST /upload-files/` - Upload DOCX and PDF files for processing
- `GET /job-status/{job_id}` - Check the status of a processing job
- `GET /download/{job_id}/{filename}` - Download a specific processed file
- `GET /download-all/{job_id}` - Download all processed files as a zip file
- `DELETE /job/{job_id}` - Delete a job and its associated files

## Setup

1. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the server:
```bash
uvicorn main:app --reload
```

The API will be available at `http://localhost:8000`

## API Documentation

Once the server is running, you can access:
- Interactive API documentation: `http://localhost:8000/docs`
- Alternative API documentation: `http://localhost:8000/redoc`

## Usage Examples

### Upload a File

```bash
curl -X POST "http://localhost:8000/upload-files/" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@document.docx"
```

### Check Job Status

```bash
curl -X GET "http://localhost:8000/job-status/{job_id}" \
  -H "accept: application/json"
```

### Download Processed File

```bash
curl -X GET "http://localhost:8000/download/{job_id}/{filename}" \
  -H "accept: application/json" \
  --output processed_file.docx
```

### Download All Files

```bash
curl -X GET "http://localhost:8000/download-all/{job_id}" \
  -H "accept: application/json" \
  --output processed_files.zip
```

### Delete Job

```bash
curl -X DELETE "http://localhost:8000/job/{job_id}" \
  -H "accept: application/json"
```

## Directory Structure

```
.
├── main.py                 # FastAPI application
├── document_processor.py   # Document processing logic
├── requirements.txt       # Project dependencies
├── uploads/              # Directory for uploaded files
└── outputs/             # Directory for processed files
```

## Error Handling

The API includes comprehensive error handling for:
- Invalid file types
- Processing failures
- Missing files/jobs
- Server errors

## Notes

- Supported input formats: DOCX and PDF
- Output format: DOCX
- Files are processed asynchronously
- Job files are automatically cleaned up when deleted
- All processed documents maintain consistent formatting 