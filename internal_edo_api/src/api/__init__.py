"""API package."""
from fastapi import APIRouter
from .v1 import agreemeant

router_v1 = APIRouter()
router_v1.include_router(agreemeant.router, prefix="/path/v1")