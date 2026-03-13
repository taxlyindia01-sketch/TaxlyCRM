from fastapi import APIRouter, HTTPException, Depends, UploadFile, File
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, EmailStr
from typing import Optional
from pathlib import Path
import shutil

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import Business, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/business", tags=["Business"])

ROOT_DIR = Path(__file__).parent.parent
LOGO_UPLOAD_DIR = ROOT_DIR / "uploads" / "logos"
QR_UPLOAD_DIR = ROOT_DIR / "uploads" / "qrcodes"
SIG_UPLOAD_DIR = ROOT_DIR / "uploads" / "signatures"
LOGO_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
SIG_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
QR_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


class BusinessCreate(BaseModel):
    business_name: str
    gstin: str
    email: EmailStr
    phone: str
    address: str
    city: str
    state: str
    pincode: str
    bank_name: Optional[str] = None
    account_number: Optional[str] = None
    ifsc_code: Optional[str] = None
    swift_code: Optional[str] = None
    authorised_signatory: Optional[str] = None
    terms_and_conditions: Optional[str] = None
    upi_id: Optional[str] = None


@router.post("/logo")
async def upload_logo(file: UploadFile = File(...), user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    if not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image")
    file_ext = file.filename.split(".")[-1]
    filename = f"{user['tenant_id']}_logo.{file_ext}"
    with open(LOGO_UPLOAD_DIR / filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    logo_url = f"/api/business/logo/{filename}"
    result = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = result.scalar_one_or_none()
    if business:
        business.logo_url = logo_url
        await db.commit()
    return {"logo_url": logo_url}


@router.get("/logo/{filename}")
async def get_logo(filename: str):
    file_path = LOGO_UPLOAD_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Logo not found")
    return StreamingResponse(open(file_path, "rb"), media_type="image/jpeg")


@router.post("/qrcode")
async def upload_qrcode(file: UploadFile = File(...), user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    if not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image")
    file_ext = file.filename.split(".")[-1]
    filename = f"{user['tenant_id']}_qrcode.{file_ext}"
    with open(QR_UPLOAD_DIR / filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    qr_url = f"/api/business/qrcode/{filename}"
    result = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = result.scalar_one_or_none()
    if business:
        business.qr_code_url = qr_url
        await db.commit()
    return {"qr_code_url": qr_url}


@router.get("/qrcode/{filename}")
async def get_qrcode(filename: str):
    file_path = QR_UPLOAD_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="QR Code not found")
    return StreamingResponse(open(file_path, "rb"), media_type="image/png")


@router.get("")
async def get_business(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = result.scalar_one_or_none()
    if not business:
        raise HTTPException(status_code=404, detail="Business profile not found")
    return serialize_doc(business)


@router.post("")
async def create_business(business_data: BusinessCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    if result.scalar_one_or_none():
        raise HTTPException(status_code=400, detail="Business profile already exists")
    business = Business(tenant_id=user["tenant_id"], **business_data.model_dump())
    db.add(business)
    await db.commit()
    await db.refresh(business)
    return serialize_doc(business)


@router.put("")
async def update_business(business_data: BusinessCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    existing = result.scalar_one_or_none()
    update_data = business_data.model_dump()

    if existing:
        for k, v in update_data.items():
            setattr(existing, k, v)
        await db.commit()
        await db.refresh(existing)
        return serialize_doc(existing)
    else:
        business = Business(tenant_id=user["tenant_id"], **update_data)
        db.add(business)
        await db.commit()
        await db.refresh(business)
        return serialize_doc(business)


@router.post("/signature")
async def upload_signature(file: UploadFile = File(...), user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    if not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image")
    file_ext = file.filename.split(".")[-1]
    filename = f"{user['tenant_id']}_signature.{file_ext}"
    with open(SIG_UPLOAD_DIR / filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    sig_url = f"/api/business/signature/{filename}"
    result = await db.execute(select(Business).where(Business.tenant_id == user["tenant_id"]))
    business = result.scalar_one_or_none()
    if business:
        business.signature_url = sig_url
        await db.commit()
    return {"signature_url": sig_url}


@router.get("/signature/{filename}")
async def serve_signature(filename: str):
    from fastapi.responses import FileResponse
    path = SIG_UPLOAD_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="Signature not found")
    return FileResponse(path)
