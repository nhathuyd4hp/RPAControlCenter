"""remove log table

Revision ID: 4e4711876b58
Revises: 3e57a3c75be6
Create Date: 2025-12-19 10:37:29.677063

"""

from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op
from sqlalchemy.dialects import mysql

# revision identifiers, used by Alembic.
revision: str = "4e4711876b58"
down_revision: Union[str, Sequence[str], None] = "3e57a3c75be6"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    op.drop_table("log")


def downgrade() -> None:
    """Downgrade schema."""
    op.create_table(
        "log",
        sa.Column("id", mysql.VARCHAR(length=255), nullable=False),
        sa.Column("created_at", mysql.DATETIME(), nullable=False),
        sa.Column("updated_at", mysql.DATETIME(), nullable=False),
        sa.Column("run_id", mysql.VARCHAR(length=255), nullable=False),
        sa.Column("timestamp", mysql.DATETIME(), nullable=False),
        sa.Column("level", mysql.VARCHAR(length=255), nullable=False),
        sa.Column("message", mysql.TEXT(), nullable=True),
        sa.ForeignKeyConstraint(["run_id"], ["runs.id"], name=op.f("log_ibfk_1")),
        sa.PrimaryKeyConstraint("id"),
        mysql_collate="utf8mb4_0900_ai_ci",
        mysql_default_charset="utf8mb4",
        mysql_engine="InnoDB",
    )
    op.create_index(op.f("ix_log_run_id"), "log", ["run_id"], unique=False)
    # ### end Alembic commands ###
