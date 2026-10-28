"""Planner, calendar, mail, and Teams-messages client tests (fake Graph clients)."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from m365_graph_toolkit.clients.calendar import CalendarClient
from m365_graph_toolkit.clients.mail import MailClient
from m365_graph_toolkit.clients.planner import PlannerClient
from m365_graph_toolkit.clients.teams_messages import TeamsMessagesClient
from m365_graph_toolkit.errors import GraphError, ValidationError
from tests.fakes import page


def _identity(user_id: str, name: str) -> SimpleNamespace:
    return SimpleNamespace(user=SimpleNamespace(id=user_id, display_name=name))


def _planner_task(**overrides) -> SimpleNamespace:
    data = {
        "id": "t1",
        "title": "Do it",
        "bucket_id": "b1",
        "percent_complete": 50,
        "priority": 5,
        "created_date_time": None,
        "start_date_time": None,
        "due_date_time": None,
        "completed_date_time": None,
        "created_by": _identity("u1", "Ada"),
        "completed_by": None,
        "assignments": {"u2": SimpleNamespace()},
        "applied_categories": SimpleNamespace(
            additional_data={"category1": True, "category2": False}
        ),
        "details": SimpleNamespace(
            description="Details here",
            checklist={
                "c1": SimpleNamespace(title="First", is_checked=True),
                "c2": SimpleNamespace(title="Second", is_checked=False),
            },
        ),
    }
    data.update(overrides)
    return SimpleNamespace(**data)


# --- Planner --------------------------------------------------------------


async def test_planner_user_tasks_page_and_normalise(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(PlannerClient, "_tasks_configuration", lambda self, **kwargs: "CONFIG")
    client = MagicMock()
    tasks_builder = client.users.by_user_id.return_value.planner.tasks
    tasks_builder.get = AsyncMock(return_value=page([_planner_task()], next_link="p2"))
    tasks_builder.with_url.return_value.get = AsyncMock(return_value=page([_planner_task(id="t2")]))
    tasks = await PlannerClient(client).get_user_tasks("u1")
    assert [t["id"] for t in tasks] == ["t1", "t2"]
    task = tasks[0]
    assert task["created_by_user_id"] == "u1"
    assert task["assigned_to"] == ["u2"]
    assert task["labels"] == ["category1"]
    assert task["description"] == "Details here"
    assert task["checklist_total"] == 2
    assert task["checklist_completed"] == 1
    assert task["checklist_items"][0] == {"title": "First", "is_checked": True}
    assert tasks_builder.get.call_args.kwargs["request_configuration"] == "CONFIG"
    client.users.by_user_id.assert_called_with("u1")


async def test_planner_user_tasks_validates_top() -> None:
    with pytest.raises(ValidationError, match="top"):
        await PlannerClient(MagicMock()).get_user_tasks("u1", top=0)


async def test_planner_tasks_for_users_skips_failing_user() -> None:
    planner = PlannerClient(MagicMock())

    async def fake_tasks(user_id: str, **kwargs) -> list:  # noqa: ANN001, ANN202
        if user_id == "bad":
            raise GraphError("forbidden")
        return [{"id": "t1"}]

    planner.get_user_tasks = fake_tasks  # type: ignore[method-assign]
    users = [
        {"id": "bad", "display_name": "Bad"},
        {"id": "good", "display_name": "Good"},
    ]
    rows = await planner.get_tasks_for_users(users)
    assert rows == [{"id": "t1", "user_id": "good", "user_display_name": "Good"}]


async def test_planner_plan_tasks_resolve_bucket_names(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        PlannerClient,
        "_plan_tasks_configuration",
        lambda self, **kwargs: "CONFIG",
    )
    client = MagicMock()
    plan_builder = client.planner.plans.by_planner_plan_id.return_value
    bucket = SimpleNamespace(id="b1", name="To do")
    plan_builder.buckets.get = AsyncMock(return_value=page([bucket]))
    plan_builder.tasks.get = AsyncMock(return_value=page([_planner_task()]))
    tasks = await PlannerClient(client).get_plan_tasks("p1")
    assert tasks[0]["bucket_name"] == "To do"
    assert tasks[0]["plan_id"] == "p1"


async def test_planner_group_plans_page() -> None:
    client = MagicMock()
    plans_builder = client.groups.by_group_id.return_value.planner.plans
    plan = SimpleNamespace(
        id="p1",
        title="Plan",
        created_date_time=None,
        created_by=_identity("u1", "Ada"),
        container=SimpleNamespace(container_id="g1", type="group", url="https://example"),
    )
    plans_builder.get = AsyncMock(return_value=page([plan]))
    plans = await PlannerClient(client).get_group_plans("g1")
    assert plans[0]["created_by_user_id"] == "u1"
    assert plans[0]["container_id"] == "g1"
    assert plans[0]["group_name"] is None


# --- Calendar -------------------------------------------------------------


def _event() -> SimpleNamespace:
    return SimpleNamespace(
        id="e1",
        subject="Standup",
        body_preview="Daily.",
        importance="normal",
        is_all_day=False,
        is_cancelled=False,
        is_online_meeting=False,
        online_meeting_url=None,
        created_date_time=None,
        last_modified_date_time=None,
        start=SimpleNamespace(date_time="2026-10-28T14:00:00", time_zone="UTC"),
        end=SimpleNamespace(date_time="2026-10-28T14:30:00", time_zone="UTC"),
        location=SimpleNamespace(display_name="Virtual"),
        organizer=SimpleNamespace(
            email_address=SimpleNamespace(address="ada@example.com", name="Ada")
        ),
        attendees=[SimpleNamespace(email_address=SimpleNamespace(address="ben@example.com"))],
    )


async def test_calendar_events_normalise(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(CalendarClient, "_events_configuration", lambda self, **kwargs: "CONFIG")
    client = MagicMock()
    events_builder = client.users.by_user_id.return_value.calendar.events
    events_builder.get = AsyncMock(return_value=page([_event()], next_link="p2"))
    events_builder.with_url.return_value.get = AsyncMock(return_value=page([]))
    events = await CalendarClient(client).get_user_events("u1")
    assert len(events) == 1
    event = events[0]
    assert event["start_date_time"] == "2026-10-28T14:00:00"
    assert event["start_time_zone"] == "UTC"
    assert event["location"] == "Virtual"
    assert event["organizer_email"] == "ada@example.com"
    assert event["attendee_emails"] == ["ben@example.com"]


async def test_calendar_events_for_users_skips_failing_user() -> None:
    calendar = CalendarClient(MagicMock())

    async def fake_events(user_id: str, **kwargs) -> list:  # noqa: ANN001, ANN202
        if user_id == "bad":
            raise GraphError("forbidden")
        return [{"id": "e1"}]

    calendar.get_user_events = fake_events  # type: ignore[method-assign]
    rows = await calendar.get_events_for_users(
        [{"id": "bad"}, {"id": "good", "display_name": "Good"}]
    )
    assert rows == [{"id": "e1", "user_id": "good", "user_display_name": "Good"}]


# --- Mail -----------------------------------------------------------------


def _message() -> SimpleNamespace:
    return SimpleNamespace(
        id="m1",
        subject="Hi",
        body_preview="preview",
        importance="normal",
        is_read=True,
        has_attachments=False,
        received_date_time=None,
        sent_date_time=None,
        created_date_time=None,
        from_=SimpleNamespace(
            email_address=SimpleNamespace(address="s@example.com", name="Sender")
        ),
        to_recipients=[SimpleNamespace(email_address=SimpleNamespace(address="t@example.com"))],
        cc_recipients=[],
        bcc_recipients=[],
        body=SimpleNamespace(content="<p>hi</p>", content_type="html"),
    )


async def test_mail_messages_normalise(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(MailClient, "_messages_configuration", lambda self, **kwargs: "CONFIG")
    client = MagicMock()
    messages_builder = client.users.by_user_id.return_value.messages
    messages_builder.get = AsyncMock(return_value=page([_message()]))
    messages = await MailClient(client).get_user_messages("u1")
    assert messages[0]["from_email"] == "s@example.com"
    assert messages[0]["from_name"] == "Sender"
    assert messages[0]["to_emails"] == ["t@example.com"]
    assert messages[0]["body"] == "<p>hi</p>"
    assert messages[0]["body_content_type"] == "html"


async def test_mail_messages_for_users_skips_unprovisioned_mailbox() -> None:
    mail = MailClient(MagicMock())

    async def fake_messages(user_id: str, **kwargs) -> list:  # noqa: ANN001, ANN202
        if user_id == "bad":
            raise GraphError("MailboxNotEnabledForRESTAPI")
        return [{"id": "m1"}]

    mail.get_user_messages = fake_messages  # type: ignore[method-assign]
    rows = await mail.get_messages_for_users([{"id": "bad"}, {"id": "good"}])
    assert rows == [{"id": "m1", "user_id": "good", "user_display_name": None}]


# --- Teams messages -------------------------------------------------------


async def test_teams_messages_normalise(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        TeamsMessagesClient, "_teams_configuration", lambda self, **kwargs: "CONFIG"
    )
    monkeypatch.setattr(
        TeamsMessagesClient, "_messages_configuration", lambda self, **kwargs: "CONFIG"
    )
    client = MagicMock()
    team = SimpleNamespace(
        id="team1",
        display_name="Ops",
        description="",
        created_date_time=None,
        is_archived=False,
    )
    client.teams.get = AsyncMock(return_value=page([team]))
    teams = await TeamsMessagesClient(client).get_teams()
    assert teams[0]["display_name"] == "Ops"

    channel = SimpleNamespace(
        id="ch1",
        display_name="General",
        description="",
        membership_type="standard",
        created_date_time=None,
    )
    channels_builder = client.teams.by_team_id.return_value.channels
    channels_builder.get = AsyncMock(return_value=page([channel]))
    channels = await TeamsMessagesClient(client).get_channels("team1")
    assert channels[0]["display_name"] == "General"

    message = SimpleNamespace(
        id="msg1",
        subject=None,
        importance="normal",
        message_type="message",
        created_date_time=None,
        last_modified_date_time=None,
        deleted_date_time=None,
        body=SimpleNamespace(content="Hello", content_type="text"),
        from_=SimpleNamespace(user=SimpleNamespace(id="u1", display_name="Ada")),
        replies=[],
        reactions=[SimpleNamespace(reaction_type="like")],
    )
    messages_builder = channels_builder.by_channel_id.return_value.messages
    messages_builder.get = AsyncMock(return_value=page([message]))
    rows = await TeamsMessagesClient(client).get_channel_messages("team1", "ch1")
    assert rows[0]["sender_id"] == "u1"
    assert rows[0]["sender_display_name"] == "Ada"
    assert rows[0]["sender_type"] == "user"
    assert rows[0]["body"] == "Hello"
    assert rows[0]["reply_count"] == 0
    assert rows[0]["reaction_count"] == 1
    assert rows[0]["reaction_types"] == ["like"]
    assert rows[0]["team_id"] == "team1"
    assert rows[0]["channel_id"] == "ch1"


async def test_teams_messages_validates_ids() -> None:
    client = TeamsMessagesClient(MagicMock())
    with pytest.raises(ValidationError):
        await client.get_channels("")
    with pytest.raises(ValidationError):
        await client.get_channel_messages("team1", "  ")
