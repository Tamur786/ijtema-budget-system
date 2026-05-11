from sqlalchemy import Column, Integer, String, Float, DateTime
from datetime import datetime
from database import Base

class Budget(Base):
    __tablename__ = "budget"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, default="Nationales Ijtema")
    total = Column(Float, default=0.0)


class Voucher(Base):
    __tablename__ = "vouchers"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String)
    department = Column(String)
    purpose = Column(String)
    amount = Column(Float)
    supplier = Column(String)
    status = Column(String, default="eingereicht")
    created_at = Column(DateTime, default=datetime.utcnow)
    approved_at = Column(DateTime, nullable=True)
    approved_by = Column(String, nullable=True)


class History(Base):
    __tablename__ = "history"

    id = Column(Integer, primary_key=True, index=True)
    action = Column(String)
    voucher_id = Column(Integer, nullable=True)
    amount = Column(Float, default=0.0)
    old_remaining = Column(Float, nullable=True)
    new_remaining = Column(Float, nullable=True)
    person = Column(String, nullable=True)
    comment = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)