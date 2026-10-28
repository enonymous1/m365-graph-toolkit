"""SharePoint STIG import tests (fake Graph clients, no SDK needed except model build)."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from m365_graph_toolkit.clients.sharepoint import SharePointClient
from m365_graph_toolkit.errors import GraphError, ValidationError
from m365_graph_toolkit.stig import load_stig_csv, plan_list_from_csv
from tests.fakes import page

SAMPLE_CSV = Path(__file__).resolve().parents[1] / "examples" / "Fictional Example STIG.csv"

_PROVISIONED_COLUMNS = [
    SimpleNamespace(id="c0", name="Title", display_name="Title", text=None),
    *(
        SimpleNamespace(
            id=f"c{i + 1}",
            name=header,
            display_name=header,
            text=SimpleNamespace(allow_multiple_lines=True),
        )
        for i, header in enumerate(
            [
                "STIG",
                "Benchmark ID",
                "Release Info",
                "Version",
                "Rule ID",
                "STIG ID",
                "Classification",
                "Severity",
                "Rule Title",
                "Fix Text",
                "Discussion",
            ]
        )
    ),
]


def _list_obj(list_id: str = "l1") -> SimpleNamespace:
    return SimpleNamespace(
        id=list_id,
        display_name="Fictional Example STIG",
        name="fictional-example-stig",
        description="",
        created_date_time=None,
        last_modified_date_time=None,
        web_url="https://contoso.example/sites/x/lists/l1",
    )


async def _import_client(
    monkeypatch: pytest.MonkeyPatch,
    *,
    columns: list,
    item_side_effect=None,
) -> tuple[MagicMock, SharePointClient]:
    client = MagicMock()
    sp = SharePointClient(client)
    monkeypatch.setattr(
        SharePointClient,
        "_build_list_model",
        lambda self, display_name, cols: f"MODEL:{display_name}",
    )
    client.sites.by_site_id.return_value.lists.post = AsyncMock(return_value=_list_obj())
    columns_builder = client.sites.by_site_id.return_value.lists.by_list_id.return_value.columns
    columns_builder.get = AsyncMock(return_value=page(list(columns)))
    items_builder = client.sites.by_site_id.return_value.lists.by_list_id.return_value.items
    if item_side_effect is None:
        items_builder.post = AsyncMock(return_value=SimpleNamespace(id="i1"))
    else:
        items_builder.post = AsyncMock(side_effect=item_side_effect)
    return client, sp


async def test_create_list_rejects_title_column() -> None:
    sp = SharePointClient(MagicMock())
    with pytest.raises(ValidationError, match="Title"):
        await sp.create_list("s1", "List", [{"name": "Title"}])


async def test_create_list_posts_model(monkeypatch: pytest.MonkeyPatch) -> None:
    client, sp = await _import_client(monkeypatch, columns=_PROVISIONED_COLUMNS)
    created = await sp.create_list("s1", "Fictional Example STIG", [{"name": "Severity"}])
    assert created["id"] == "l1"
    post = client.sites.by_site_id.return_value.lists.post
    assert post.call_args.args[0] == "MODEL:Fictional Example STIG"


async def test_build_list_model_rejects_duplicates() -> None:
    pytest.importorskip("msgraph")
    sp = SharePointClient(MagicMock())
    with pytest.raises(ValidationError, match="Duplicate"):
        sp._build_list_model("List", [{"name": "Severity"}, {"name": "severity"}])


async def test_get_list_columns_normalises(monkeypatch: pytest.MonkeyPatch) -> None:
    client, sp = await _import_client(monkeypatch, columns=_PROVISIONED_COLUMNS)
    columns = await sp.get_list_columns("s1", "l1")
    assert columns[0] == {
        "id": "c0",
        "name": "Title",
        "display_name": "Title",
        "is_text": False,
        "allow_multiple_lines": None,
    }
    assert columns[1]["is_text"] is True
    assert columns[1]["allow_multiple_lines"] is True


async def test_import_list_from_csv_success(monkeypatch: pytest.MonkeyPatch) -> None:
    _, sp = await _import_client(monkeypatch, columns=_PROVISIONED_COLUMNS)
    plan, rows = load_stig_csv(SAMPLE_CSV)
    summary = await sp.import_list_from_csv("s1", plan, rows, poll_attempts=2, poll_delay_seconds=0)
    assert summary["items_created"] == 2
    assert summary["items_failed"] == 0
    assert summary["list_id"] == "l1"
    assert summary["failures"] == []
    items_builder = sp.client.sites.by_site_id.return_value.lists.by_list_id.return_value.items
    first_call_fields = items_builder.post.call_args_list[0].args[0]
    # The posted model carries a FieldValueSet with the mapped fields.
    assert first_call_fields.fields.additional_data["Title"] == "V-999001"


async def test_import_list_counts_failed_rows(monkeypatch: pytest.MonkeyPatch) -> None:
    _, sp = await _import_client(
        monkeypatch,
        columns=_PROVISIONED_COLUMNS,
        item_side_effect=[SimpleNamespace(id="i1"), GraphError("boom")],
    )
    plan, rows = load_stig_csv(SAMPLE_CSV)
    summary = await sp.import_list_from_csv("s1", plan, rows, poll_attempts=2, poll_delay_seconds=0)
    assert summary["items_created"] == 1
    assert summary["items_failed"] == 1
    assert summary["failures"][0]["row"] == 2
    assert "boom" in summary["failures"][0]["error"]


async def test_import_list_raises_when_columns_never_provision(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _, sp = await _import_client(monkeypatch, columns=[])
    plan, rows = load_stig_csv(SAMPLE_CSV)
    with pytest.raises(GraphError, match="not provisioned"):
        await sp.import_list_from_csv("s1", plan, rows, poll_attempts=2, poll_delay_seconds=0)
    items_builder = sp.client.sites.by_site_id.return_value.lists.by_list_id.return_value.items
    items_builder.post.assert_not_called()


async def test_import_list_rejects_empty_rows(monkeypatch: pytest.MonkeyPatch) -> None:
    _, sp = await _import_client(monkeypatch, columns=_PROVISIONED_COLUMNS)
    plan = plan_list_from_csv(SAMPLE_CSV)
    with pytest.raises(ValidationError, match="no rows"):
        await sp.import_list_from_csv("s1", plan, [], poll_attempts=1, poll_delay_seconds=0)


async def test_import_lists_from_directory_reports_per_file_errors(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    good = tmp_path / "Good.csv"
    good.write_text("Vuln ID,Severity\nV-1,high\n", encoding="utf-8")
    bad = tmp_path / "Bad.csv"
    bad.write_text("Severity\nlow\n", encoding="utf-8")
    sp = SharePointClient(MagicMock())

    async def fake_import(site_id, plan, rows, **kwargs):  # noqa: ANN001, ANN202
        return {
            "list_name": plan["list_name"],
            "list_id": "l1",
            "source_file": plan["source_file"],
            "rows_total": len(rows),
            "items_created": len(rows),
            "items_failed": 0,
            "failures": [],
        }

    sp.import_list_from_csv = fake_import  # type: ignore[method-assign]
    summaries = await sp.import_lists_from_directory("s1", tmp_path)
    assert [s["list_name"] for s in summaries] == ["Bad", "Good"]
    assert "error" in summaries[0]
    assert summaries[1]["items_created"] == 1
