from fastapi import APIRouter, Depends
from pydantic import BaseModel

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from database import SeriesConfig, serialize_doc, get_db
from routes.auth import get_current_user

router = APIRouter(prefix="/series-config", tags=["Series Config"])


class SeriesConfigUpdate(BaseModel):
    invoice_prefix: str = "INV"
    estimate_prefix: str = "EST"
    credit_note_prefix: str = "CN"


@router.get("")
async def get_series_config(user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(SeriesConfig).where(SeriesConfig.tenant_id == user["tenant_id"]))
    config = result.scalar_one_or_none()
    if not config:
        config = SeriesConfig(
            tenant_id=user["tenant_id"],
            invoice_prefix="INV", invoice_counter=0,
            estimate_prefix="EST", estimate_counter=0,
            credit_note_prefix="CN", credit_note_counter=0
        )
        db.add(config)
        await db.commit()
        await db.refresh(config)
    return serialize_doc(config)


@router.put("")
async def update_series_config(config_data: SeriesConfigUpdate, user: dict = Depends(get_current_user), db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(SeriesConfig).where(SeriesConfig.tenant_id == user["tenant_id"]))
    config = result.scalar_one_or_none()
    if config:
        config.invoice_prefix = config_data.invoice_prefix
        config.estimate_prefix = config_data.estimate_prefix
        config.credit_note_prefix = config_data.credit_note_prefix
    else:
        config = SeriesConfig(
            tenant_id=user["tenant_id"],
            invoice_prefix=config_data.invoice_prefix,
            invoice_counter=0,
            estimate_prefix=config_data.estimate_prefix,
            estimate_counter=0,
            credit_note_prefix=config_data.credit_note_prefix,
            credit_note_counter=0
        )
        db.add(config)
    await db.commit()
    await db.refresh(config)
    return serialize_doc(config)
