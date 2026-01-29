"""API routes for the Programa Schedule Parser.

This module defines the REST API endpoints:
- POST /parse: Parse an Excel schedule into structured JSON
- GET /health: Health check endpoint
"""

import logging

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import JSONResponse

from app.core.config import settings
from app.core.models import ErrorResponse, ParseResponse
from app.parser.service import ScheduleParser, ScheduleParserConfig
from app.parser.workbook import WorkbookLoadError, load_workbook_safe
from app.parser.sheet_detector import get_schedule_sheets

# Create router instance
router = APIRouter()
logger = logging.getLogger(__name__)

# Create parser instance at startup using app settings.
# LLM client will be wired in Phase 2 (currently uses default no-op behavior).
_parser_config = ScheduleParserConfig(
    use_llm=settings.use_llm,
    llm_mode=settings.llm_mode,
    llm_min_missing_fields=settings.llm_min_missing_fields,
)
schedule_parser = ScheduleParser(config=_parser_config)


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
        422: {
            "description": "Request validation error (e.g., missing required form field)",
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
        extension = file.filename[file.filename.rfind(".") :] if "." in file.filename else "no extension"
        return JSONResponse(
            status_code=400,
            content=ErrorResponse(
                error="Invalid file format",
                detail=f"Expected .xlsx file, got '{extension}'",
            ).model_dump(),
        )

    try:
        file_bytes = await file.read()
    except Exception as e:
        return JSONResponse(
            status_code=400,
            content=ErrorResponse(
                error="Failed to read upload",
                detail=f"{type(e).__name__}: {e}",
            ).model_dump(),
        )
    finally:
        try:
            await file.close()
        except Exception:
            pass

    try:
        wb = load_workbook_safe(file_bytes)
        schedule_sheets = get_schedule_sheets(wb)
        parsed = schedule_parser.parse_workbook(wb, filename=file.filename)

        logger.info(
            "Parsed workbook filename=%s sheets_total=%d schedule_sheets=%d products=%d",
            file.filename,
            len(wb.sheetnames),
            len(schedule_sheets),
            len(parsed.products),
        )

        return parsed
    except WorkbookLoadError as e:
        return JSONResponse(
            status_code=400,
            content=ErrorResponse(
                error=e.message,
                detail=e.detail,
            ).model_dump(),
        )
    except Exception as e:
        return JSONResponse(
            status_code=400,
            content=ErrorResponse(
                error="Failed to parse workbook",
                detail=f"{type(e).__name__}: {e}",
            ).model_dump(),
        )
