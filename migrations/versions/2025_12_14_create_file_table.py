"""create file table

Revision ID: 856dcc834701
Revises: 
Create Date: 2025-12-14 00:57:33.288393

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = '856dcc834701'
down_revision: Union[str, Sequence[str], None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    op.create_table('file',
        sa.Column('uuid', sa.Uuid(), nullable=False),
        sa.Column('status', sa.String(length=50), nullable=False),
        sa.Column('original_s3_link', sa.String(), nullable=False),
        sa.Column('processed_s3_link', sa.String(), nullable=True),
        sa.Column('created_at', sa.DateTime(timezone=True), server_default=sa.text('now()'), nullable=False),
        sa.Column('updated_at', sa.DateTime(timezone=True), nullable=True),
        sa.Column('deleted_at', sa.DateTime(timezone=True), nullable=True),
        sa.PrimaryKeyConstraint('uuid', name=op.f('pk_file')),
        sa.UniqueConstraint('uuid', name=op.f('uq_file_uuid'))
    )


def downgrade() -> None:
    """Downgrade schema."""
    op.drop_table('file')
