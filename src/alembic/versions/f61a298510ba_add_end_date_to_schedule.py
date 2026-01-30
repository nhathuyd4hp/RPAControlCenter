"""add end_date to schedule

Revision ID: f61a298510ba
Revises: 424f429cdd0c
Create Date: 2026-01-30 16:11:57.334156

"""

from typing import Sequence, Union

import sqlalchemy as sa
import sqlmodel
from alembic import op
from sqlalchemy.dialects import mysql

# revision identifiers, used by Alembic.
revision: str = "f61a298510ba"  # pragma: allowlist secret
down_revision: Union[str, Sequence[str], None] = "424f429cdd0c"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    op.add_column("schedule", sa.Column("day", sqlmodel.sql.sqltypes.AutoString(), nullable=True))
    op.drop_column("schedule", "end_date")
    # ### end Alembic commands ###


def downgrade() -> None:
    """Downgrade schema."""
    op.add_column("schedule", sa.Column("end_date", mysql.DATETIME(), nullable=True))
    op.drop_column("schedule", "day")
    op.create_index(op.f("ix_error_run_id"), "error", ["run_id"], unique=False)
    # ### end Alembic commands ###
