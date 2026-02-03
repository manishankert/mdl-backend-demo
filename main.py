# main.py
"""
MDL DOCX Builder - FastAPI Application Entry Point

This module serves as the entry point for the MDL DOCX Builder API.
It initializes the FastAPI application, configures middleware, and includes all routes.
"""

import logging
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.exceptions import RequestValidationError

from routes import router

# Configure logging
logging.basicConfig(level=logging.INFO)

# ------------------------------------------------------------------------------
# FastAPI app
# ------------------------------------------------------------------------------
app = FastAPI(title="MDL DOCX Builder")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routes
app.include_router(router)


# ------------------------------------------------------------------------------
# Middleware
# ------------------------------------------------------------------------------
@app.middleware("http")
async def log_requests(request: Request, call_next):
    if request.url.path.endswith("/build-mdl-docx-auto"):
        raw = await request.body()
        try:
            logging.info("== /build-mdl-docx-auto RAW BODY ==")
        except Exception:
            pass
        # re-create the request stream for downstream
        request._receive = (lambda b=raw: {"type": "http.request", "body": b, "more_body": False})
    return await call_next(request)


# ------------------------------------------------------------------------------
# Exception handlers
# ------------------------------------------------------------------------------
@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    logging.info("== Pydantic Validation Errors ==")
    logging.info(exc.errors())
    return JSONResponse(status_code=422, content={"ok": False, "errors": exc.errors()})
