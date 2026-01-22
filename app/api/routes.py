"""API routes for the Programa Schedule Parser.

This module defines the REST API endpoints:
- POST /parse: Parse an Excel schedule into structured JSON
- GET /health: Health check endpoint
"""

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import JSONResponse

from app.core.models import ErrorResponse, ParseResponse

# Create router instance
router = APIRouter()


@router.get(
    "/health",
    summary="Health Check",
    description="Returns the health status of the API service.",
    response_description="Health status object",
    responses={
        200: {
            "description": "Service is healthy",
            "content": {
                "application/json": {
                    "example": {"status": "ok"}
                }
            }
        }
    }
)
async def health_check() -> dict:
    """Health check endpoint.

    Returns:
        dict: Health status with "status": "ok"
    """
    return {"status": "ok"}


@router.post(
    "/parse",
    response_model=ParseResponse,
    summary="Parse Excel Schedule",
    description=(
        "Upload an Excel (.xlsx) schedule file and parse it into structured JSON. "
        "The parser extracts product information including codes, names, brands, "
        "colours, finishes, materials, dimensions, quantities, and prices."
    ),
    response_description="Parsed schedule with products",
    responses={
        200: {
            "description": "Successfully parsed schedule",
            "model": ParseResponse,
        },
        400: {
            "description": "Invalid file format or parsing error",
            "model": ErrorResponse,
        },
    }
)
async def parse_schedule(
    file: UploadFile = File(
        ...,
        description="Excel schedule file (.xlsx format)",
    )
) -> ParseResponse | JSONResponse:
    """Parse an uploaded Excel schedule into structured JSON.

    Args:
        file: Uploaded Excel file (.xlsx format)

    Returns:
        ParseResponse: Parsed schedule containing schedule name and products
    """
    # Validate file extension
    if not file.filename:
        return JSONResponse(
            status_code=400,
            content=ErrorResponse(
                error="Invalid file",
                detail="No filename provided",
            ).model_dump(),
        )

    if not file.filename.lower().endswith(".xlsx"):
        extension = file.filename.split(".")[-1] if "." in file.filename else "no extension"
        return JSONResponse(
            status_code=400,
            content=ErrorResponse(
                error="Invalid file format",
                detail=f"Expected .xlsx file, got '{extension}'",
            ).model_dump(),
        )

    # TODO: In Phase 2+, implement actual parsing logic
    # For now, return dummy response with hardcoded schedule_name and empty products

    # Extract schedule name from filename (remove extension)
    schedule_name = file.filename.rsplit(".", 1)[0] if file.filename else "Unknown Schedule"

    # Return dummy ParseResponse
    return ParseResponse(
        schedule_name=schedule_name,
        products=[]
    )
