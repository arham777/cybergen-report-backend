from enum import Enum
import shutil
from pathlib import Path
import asyncio
from typing import Dict, Optional
import os
import docx
from docx.shared import Pt, Inches, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.enum.section import WD_ORIENTATION
from pdf2docx import Converter
import fitz
from PIL import Image
import io

class JobStatus(str, Enum):
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"

def get_job_status(job_id: str, jobs: Dict) -> Optional[Dict]:
    """
    Get the status of a job.
    """
    if job_id not in jobs:
        return None
    return jobs[job_id]

async def cleanup_job(job_id: str, jobs: Dict, upload_dir: Path, output_dir: Path):
    """
    Clean up job files and data.
    """
    # Remove job directories
    job_upload_dir = upload_dir / job_id
    job_output_dir = output_dir / job_id
    
    if job_upload_dir.exists():
        shutil.rmtree(job_upload_dir)
    if job_output_dir.exists():
        shutil.rmtree(job_output_dir)
    
    # Remove job from jobs dict
    if job_id in jobs:
        del jobs[job_id]

def set_page_size_and_margins(doc):
    """
    Set the page size to A4 and appropriate margins.
    """
    for section in doc.sections:
        section.page_height = Mm(297)
        section.page_width = Mm(210)
        section.orientation = WD_ORIENTATION.PORTRAIT
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.header_distance = Inches(0.5)
        section.footer_distance = Inches(0.5)

def convert_pdf_to_docx(pdf_path: Path, output_path: Path) -> Optional[Path]:
    """
    Convert PDF to DOCX format.
    """
    try:
        cv = Converter(str(pdf_path))
        cv.convert(str(output_path))
        cv.close()
        return output_path
    except Exception as e:
        print(f"Error converting PDF to DOCX: {str(e)}")
        return None

def process_document(job_id: str, input_file: Path, output_dir: Path, jobs: Dict):
    """
    Process the uploaded document.
    """
    try:
        # Update job status
        jobs[job_id]["status"] = JobStatus.PROCESSING
        
        # Create template document
        template_doc = docx.Document()
        set_page_size_and_margins(template_doc)
        
        # Process based on file type
        if input_file.suffix.lower() == '.pdf':
            # Convert PDF to DOCX
            docx_path = output_dir / f"{input_file.stem}_converted.docx"
            if not convert_pdf_to_docx(input_file, docx_path):
                raise Exception("Failed to convert PDF to DOCX")
            
            # Process the converted DOCX
            source_doc = docx.Document(docx_path)
        else:
            # Process DOCX directly
            source_doc = docx.Document(input_file)
        
        # Process document content
        for element in source_doc.element.body:
            if element.tag.endswith('}p'):  # Paragraph
                # Find corresponding paragraph
                para = None
                for p in source_doc.paragraphs:
                    if p._element == element:
                        para = p
                        break
                
                if not para or not para.text.strip():
                    continue
                
                # Add paragraph to template
                new_para = template_doc.add_paragraph()
                for run in para.runs:
                    new_run = new_para.add_run(run.text)
                    new_run.font.size = Pt(12)
                    new_run.font.color.rgb = RGBColor(0, 0, 0)
                
                new_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                new_para.paragraph_format.space_after = Pt(12)
            
            elif element.tag.endswith('}tbl'):  # Table
                # Find corresponding table
                table = None
                for tbl in source_doc.tables:
                    if tbl._element == element:
                        table = tbl
                        break
                
                if table:
                    # Copy table
                    new_table = template_doc.add_table(
                        rows=len(table.rows),
                        cols=len(table.columns)
                    )
                    
                    # Copy table content
                    for i, row in enumerate(table.rows):
                        for j, cell in enumerate(row.cells):
                            new_cell = new_table.cell(i, j)
                            new_cell.text = cell.text
                            
                            # Format cell text
                            paragraph = new_cell.paragraphs[0]
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            for run in paragraph.runs:
                                run.font.size = Pt(11)
                                run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Save processed document
        output_file = output_dir / f"processed_{input_file.name}"
        if output_file.suffix.lower() != '.docx':
            output_file = output_file.with_suffix('.docx')
        
        template_doc.save(output_file)
        
        # Update job information
        jobs[job_id].update({
            "status": JobStatus.COMPLETED,
            "output_files": [output_file.name]
        })
        
    except Exception as e:
        # Update job with error information
        jobs[job_id].update({
            "status": JobStatus.FAILED,
            "error": str(e)
        })
        print(f"Error processing document: {str(e)}") 