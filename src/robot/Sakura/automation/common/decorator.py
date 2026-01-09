import functools
import logging
import time
from typing import Any, Callable, Tuple, Type


def retry_if_exception(
    exceptions: Tuple[Type[BaseException], ...] = (),
    max_retries: int = 10,
    delay: float = 2.0,
    failure_return: Any = None,
) -> Callable:
    """
    Decorator thử lại một hàm nếu nó gây ra các ngoại lệ được chỉ định.
    Args:
        exceptions: Tuple các loại ngoại lệ để bắt và thử lại
        max_retries: Số lần thử lại tối đa
        delay: Thời gian chờ ban đầu giữa các lần thử lại (giây)
        failure_return: Giá trị trả về nếu thất bại
        message_on_retry: Message khi retry (default: RETRY: <function_name>)
    """

    def wrapper(func):
        @functools.wraps(func)
        def method(self, *args, **kwargs) -> Any:
            logger: logging.Logger = (
                getattr(self, "logger")
                if hasattr(self, "logger") and isinstance(getattr(self,'logger'),logging.Logger)
                else logging.getLogger()
            )
            retries = 0
            while True:
                try:
                    return func(self, *args, **kwargs)
                except exceptions as e:
                    retries += 1
                    if retries > max_retries:
                        msg = e.msg if hasattr(e, "msg") else str(e)
                        logger.error(
                            f"{func.__name__} failed after {max_retries} retries: {msg}"
                        )
                        return failure_return
                    logger.info(f"RETRY {func.__name__}")
                    time.sleep(delay)
                except Exception as e:
                    msg = e.msg if hasattr(e, "msg") else str(e)
                    logger.error(f"{func.__name__} catch {type(e)}: {msg} ")
                    return failure_return

        return method

    return wrapper