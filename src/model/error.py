from src.model.base import Base
from sqlmodel import Column, Field, Text, ForeignKey

class Error(Base, table=True):
    run_id: int = Field(
        sa_column=Column(ForeignKey("runs.id", ondelete="CASCADE"))
    )
    # ----- #
    error_type: str = Field()
    message: str = Field()
    traceback: str = Field(sa_column=Column(Text))
