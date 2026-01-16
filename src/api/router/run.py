from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlmodel import Session

from src.api.common.response import SuccessResponse
from src.api.dependency import get_session
from src.model.runs import Status
from src.service import RunService

router = APIRouter(prefix="/runs", tags=["Runs"])


@router.get(
    path="",
    name="Lịch sử chạy",
    response_model=SuccessResponse,
)
def get_histories(
    session: Session = Depends(get_session),
    status: Status | None = Query(default=None),
):
    if status:
        histories = RunService(session).findByStatus(status)
    else:
        histories = RunService(session).findMany()
    return SuccessResponse(data=histories)


@router.get(
    path="/{id}",
    name="Lịch sử chạy",
    response_model=SuccessResponse,
)
def get_history(
    id: str,
    session: Session = Depends(get_session),
):
    history = RunService(session).findByID(id)
    if history is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="run not found")
    return SuccessResponse(data=history)
