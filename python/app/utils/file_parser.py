import io
import fitz
import docx
from fastapi import UploadFile, HTTPException, status

async def extract_text_from_file(file: UploadFile) -> str:

    """
    Extracts raw text from an uploaded file.
    Supports: .txt, .pdf, .docx
    """

    filename = file.filename.lower()
    content = await file.read()

    text = ""

    try:
        if filename.endswith(".txt"):
            text = content.decode("utf-8")

        elif filename.endswith(".pdf"):
            # Open PDF using PyMuPDF from memory stream
            pdf_document = fitz.open(stream = content, filetype = "pdf")
            for page in pdf_document:
                text += page.get_text("text") + "\n"
            pdf_document.close()

        elif filename.endswith(".docx"):
            doc = docx.Document(io.BytesIO(content))
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])

        else:
            raise HTTPException(
                status_code = status.HTTP_415_UNSUPPORTED_MEDIA_TYPE, 
                detail = f"Unsupported file extension for {filename}. Only .pdf, .docx, and .txt are allowed."
            )
        
        if not text.strip():
            raise HTTPException(
                status_code = status.HTTP_400_BAD_REQUEST, 
                detail = f"No text could be extracted from {filename}."
            )
        
        return text

    except Exception as e:
        # Rethrow HTTPExceptions directly
        if isinstance(e, HTTPException):
            raise e
        
        raise HTTPException(
            status_code = status.HTTP_500_INTERNAL_SERVER_ERROR, 
            detail = f"Failed to parse file {filename}: {str(e)}"
        )

