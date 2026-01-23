"""Programa Schedule Parser - FastAPI Application Entry Point.

This module creates and configures the FastAPI application instance,
sets up CORS middleware, and includes the API routes.
"""

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.routes import router
from app.core.models import ErrorResponse


def create_app() -> FastAPI:
    """Create and configure the FastAPI application.

    Returns:
        Configured FastAPI application instance.
    """
    app = FastAPI(
        title="Programa Schedule Parser",
        description=(
            "REST API that parses interior designer Excel schedules (.xlsx) "
            "into structured JSON for import into Programa's platform."
        ),
        version="1.0.0",
        docs_url="/docs",
        redoc_url="/redoc",
        openapi_url="/openapi.json",
    )

    # Configure CORS middleware (allow all origins for development)
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=False,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.exception_handler(RequestValidationError)
    async def request_validation_exception_handler(  # type: ignore[misc]
        request: Request,
        exc: RequestValidationError,
    ) -> JSONResponse:
        missing_fields: list[str] = []
        for err in exc.errors():
            if err.get("type") == "missing":
                loc = err.get("loc") or ()
                if loc:
                    missing_fields.append(str(loc[-1]))

        detail = "Request body validation failed"
        if "file" in missing_fields:
            detail = "Missing required form field: file"

        return JSONResponse(
            status_code=422,
            content=ErrorResponse(
                error="Validation error",
                detail=detail,
            ).model_dump(),
        )

    # Include API routes
    app.include_router(router)

    return app


# Create the application instance
app = create_app()
