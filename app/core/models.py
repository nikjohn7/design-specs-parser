"""Pydantic models for the Programa Schedule Parser API.

This module defines the request/response models for the parsing API:
- Product: Represents a single parsed product from a schedule
- ParseResponse: Successful parse result with schedule name and products
- ErrorResponse: Error response for failed requests
"""

from pydantic import BaseModel, Field


class Product(BaseModel):
    """Represents a single product extracted from a design schedule.

    All fields are optional to support graceful degradation — missing data
    becomes null rather than causing validation errors.
    """

    doc_code: str | None = Field(
        default=None,
        description="Specification code / reference (e.g., 'FCA-01 A', 'L1', 'PTF-*K')",
    )
    product_name: str | None = Field(
        default=None,
        description="Product name extracted from PRODUCT/NAME/RANGE fields",
    )
    brand: str | None = Field(
        default=None,
        description="Manufacturer or brand name",
    )
    colour: str | None = Field(
        default=None,
        description="Product colour (normalized from COLOUR/COLOR)",
    )
    finish: str | None = Field(
        default=None,
        description="Surface finish (e.g., 'MATT', 'GLOSS', 'SATIN')",
    )
    material: str | None = Field(
        default=None,
        description="Material composition or species",
    )
    width: int | None = Field(
        default=None,
        description="Width in millimeters",
        ge=0,
    )
    length: int | None = Field(
        default=None,
        description="Length in millimeters",
        ge=0,
    )
    height: int | None = Field(
        default=None,
        description="Height/depth/thickness in millimeters",
        ge=0,
    )
    qty: int | None = Field(
        default=None,
        description="Quantity",
        ge=0,
    )
    rrp: float | None = Field(
        default=None,
        description="Recommended retail price (numeric value only)",
        ge=0,
    )
    feature_image: str | None = Field(
        default=None,
        description="Base64-encoded product image or URL (optional)",
    )
    product_description: str | None = Field(
        default=None,
        description="Combined section context and item location",
    )
    product_details: str | None = Field(
        default=None,
        description="Additional specification details as pipe-separated key-value pairs",
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "doc_code": "FCA-01 A",
                    "product_name": "ICONIC",
                    "brand": "VICTORIA CARPETS",
                    "colour": "SILVER SHADOW",
                    "finish": None,
                    "material": "80% WOOL 20% SYNTHETIC",
                    "width": 3660,
                    "length": None,
                    "height": None,
                    "qty": None,
                    "rrp": 45.50,
                    "feature_image": None,
                    "product_description": "FLOORING - CARPET | APARTMENTS | GOLD SCHEME",
                    "product_details": "PRODUCT: ICONIC | CODE: 50/2833 | STYLE: TWIST",
                }
            ]
        }
    }


class ParseResponse(BaseModel):
    """Successful response from the /parse endpoint.

    Contains the extracted schedule name and list of products.
    """

    schedule_name: str = Field(
        description="Name of the schedule (from title row or filename)"
    )
    products: list[Product] = Field(
        description="List of products extracted from the schedule"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "schedule_name": "12006: GEM, WATERLINE PLACE, WILLIAMSTOWN",
                    "products": [
                        {
                            "doc_code": "FCA-01 A",
                            "product_name": "ICONIC",
                            "brand": "VICTORIA CARPETS",
                            "colour": "SILVER SHADOW",
                            "finish": None,
                            "material": "80% WOOL 20% SYNTHETIC",
                            "width": 3660,
                            "length": None,
                            "height": None,
                            "qty": None,
                            "rrp": 45.50,
                            "feature_image": None,
                            "product_description": "FLOORING - CARPET | APARTMENTS | GOLD SCHEME",
                            "product_details": "PRODUCT: ICONIC | CODE: 50/2833 | STYLE: TWIST",
                        }
                    ],
                }
            ]
        }
    }


class ErrorResponse(BaseModel):
    """Error response for failed requests.

    Returned with appropriate HTTP status codes (400, 422, 500).
    """

    error: str = Field(
        description="Brief error message describing what went wrong"
    )
    detail: str | None = Field(
        default=None,
        description="Additional error details or stack trace (if available)",
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "error": "Invalid file format",
                    "detail": "Expected .xlsx file, got .csv",
                },
                {
                    "error": "Failed to parse workbook",
                    "detail": "File appears to be corrupted or password-protected",
                },
            ]
        }
    }
