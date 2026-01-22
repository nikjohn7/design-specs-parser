"""Programa Schedule Parser - FastAPI Application Entry Point.

This module creates and configures the FastAPI application instance,
sets up CORS middleware, and includes the API routes.
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.routes import router


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
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Include API routes
    app.include_router(router)

    return app


# Create the application instance
app = create_app()
