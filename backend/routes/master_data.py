"""master_data.py"""
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from typing import Optional
from datetime import datetime, timezone
import uuid

from sqlalchemy import select, delete
from sqlalchemy.ext.asyncio import AsyncSession

from database import MasterData, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/master-data", tags=["Master Data"])


class MasterDataCreate(BaseModel):
    description: str
    hsn_sac_code: Optional[str] = None
    gst_rate: Optional[float] = None


@router.get("")
async def get_master_data(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(MasterData).where(MasterData.tenant_id == user["tenant_id"]))
    return [serialize_doc(m) for m in result.scalars().all()]


@router.post("")
async def create_master_data(master_data_item: MasterDataCreate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    item = MasterData(
        master_id=f"master_{uuid.uuid4().hex[:12]}",
        tenant_id=user["tenant_id"],
        description=master_data_item.description,
        hsn_sac_code=master_data_item.hsn_sac_code,
        gst_rate=master_data_item.gst_rate,
        type="unified",
        value=master_data_item.description,
        created_at=datetime.now(timezone.utc)
    )
    db.add(item)
    await db.commit()
    await db.refresh(item)
    return serialize_doc(item)


@router.delete("/{master_id}")
async def delete_master_data(master_id: str, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(delete(MasterData).where(MasterData.master_id == master_id, MasterData.tenant_id == user["tenant_id"]))
    if result.rowcount == 0:
        raise HTTPException(status_code=404, detail="Master data not found")
    await db.commit()
    return {"message": "Master data deleted successfully"}
