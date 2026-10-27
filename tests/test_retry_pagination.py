from __future__ import annotations

from types import SimpleNamespace

import pytest

from m365_graph_toolkit.errors import GraphError
from m365_graph_toolkit.retry import collect_pages, is_retryable, with_retry
from tests.fakes import page


class FakeHttpError(Exception):
    def __init__(self, status: int, *, retry_after: str | None = None) -> None:
        super().__init__(f"HTTP {status}")
        self.response_status_code = status
        self.response_headers = {"Retry-After": retry_after} if retry_after else {}


def test_is_retryable_statuses() -> None:
    assert is_retryable(FakeHttpError(429))
    assert is_retryable(FakeHttpError(503))
    assert not is_retryable(FakeHttpError(403))
    assert not is_retryable(FakeHttpError(404))
    assert is_retryable(TimeoutError("timed out"))
    assert not is_retryable(ValueError("bad input"))


async def test_with_retry_succeeds_after_transient_failures() -> None:
    attempts = {"count": 0}

    async def flaky() -> str:
        attempts["count"] += 1
        if attempts["count"] < 3:
            raise FakeHttpError(429, retry_after="0")
        return "ok"

    assert await with_retry(flaky, max_retries=3, base_delay=0) == "ok"
    assert attempts["count"] == 3


async def test_with_retry_retries_exactly_max_retries() -> None:
    attempts = {"count": 0}

    async def always_fails() -> None:
        attempts["count"] += 1
        raise FakeHttpError(503, retry_after="0")

    with pytest.raises(FakeHttpError):
        await with_retry(always_fails, max_retries=3, base_delay=0)
    # One initial attempt + exactly 3 retries (v1 ran max_retries + 1 retries).
    assert attempts["count"] == 4


async def test_with_retry_fails_fast_on_non_retryable() -> None:
    attempts = {"count": 0}

    async def forbidden() -> None:
        attempts["count"] += 1
        raise FakeHttpError(403)

    with pytest.raises(FakeHttpError):
        await with_retry(forbidden, max_retries=3, base_delay=0)
    assert attempts["count"] == 1


async def test_collect_pages_follows_next_link() -> None:
    first = page([1, 2], next_link="https://example.com/page2")
    second = page([3])

    async def first_page():
        return first

    async def next_page(url: str):
        assert url == "https://example.com/page2"
        return second

    assert await collect_pages(first_page, next_page) == [1, 2, 3]


async def test_collect_pages_tolerates_null_value() -> None:
    async def first_page():
        return page(None)

    assert await collect_pages(first_page, lambda url: None) == []  # type: ignore[arg-type]


async def test_collect_pages_fails_loud_on_missing_value() -> None:
    async def first_page():
        return SimpleNamespace(odata_next_link=None)

    with pytest.raises(GraphError):
        await collect_pages(first_page, lambda url: None)  # type: ignore[arg-type]
