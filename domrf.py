from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from typing import Self

import httpx

ORIGIN = "https://наш.дом.рф"
API_PATH = "/portal-analytics/api/ready-construction/"


@dataclass(frozen=True, order=True)
class ReportPeriod:
    year: int
    month: int

    def __post_init__(self) -> None:
        if not 1 <= self.month <= 12:
            raise ValueError("Отчётный месяц должен быть от 1 до 12")

    def params(self) -> dict[str, int]:
        return {"repYear": self.year, "repMonth": self.month}


class DomRfClient:
    """Клиент отчёта «Распроданность и стройготовность»."""

    def __init__(self) -> None:
        self._client = httpx.Client(
            base_url=ORIGIN,
            trust_env=False,
            follow_redirects=True,
            max_redirects=5,
            timeout=20,
            headers={"User-Agent": "Mozilla/5.0 (Macintosh)"},
        )

    def __enter__(self) -> Self:
        return self

    def __exit__(self, *exc_info: object) -> None:
        self.close()

    def close(self) -> None:
        self._client.close()

    def _get(self, endpoint: str, params: Mapping[str, int | str]) -> dict:
        response = self._client.get(API_PATH + endpoint, params=params)
        response.raise_for_status()
        if "application/json" not in response.headers.get("content-type", ""):
            raise RuntimeError(f"{endpoint}: сервер вернул ответ без JSON")
        return response.json()

    def report_periods(self) -> list[ReportPeriod]:
        """Отчётные месяцы из истории показателя, от новых к старым."""
        data = self._get("dynamics", {"dynamicChartType": "soldPercChart"})
        periods = {
            ReportPeriod(year["repYear"], month["repMonth"])
            for chart in data["dynamicCharts"]
            for year in chart["repYears"]
            for month in year["repMonths"]
        }
        return sorted(periods, reverse=True)

    def regions(self, period: ReportPeriod) -> dict[int, str]:
        """Коды и названия регионов с данными за выбранный месяц."""
        data = self._get("charts", period.params())
        return {
            row["regionCd"]: row["key"]
            for row in self._chart_rows(data, "regionChart")
            if row.get("regionCd") is not None
        }

    def completion_years(self, period: ReportPeriod, region: int) -> list[int]:
        """Годы ввода из диаграммы, включая последнюю группу."""
        params = {**period.params(), "regionCd": region}
        data = self._get("ready-year-charts", params)
        return sorted(row["readyYear"] for row in data["readyYears"])

    def developers(
        self,
        period: ReportPeriod,
        region: int,
        *,
        completion_years: Iterable[int] = (),
    ) -> list[dict]:
        """Строки виджета «Девелоперы» с исходной точностью API."""
        params: dict[str, int | str] = {**period.params(), "regionCd": region}
        values = sorted(set(completion_years))
        if values:
            params["objReady100PercYear"] = ",".join(map(str, values))
        return self._chart_rows(self._get("charts", params), "devChart")

    @staticmethod
    def _chart_rows(data: dict, chart_type: str) -> list[dict]:
        return next(
            (
                chart["data"]
                for chart in data["charts"]
                if chart["chartType"] == chart_type
            ),
            [],
        )

    def fetch_report(self, period: ReportPeriod, region: int) -> dict:
        params = {**period.params(), "regionCd": region}
        datasets = {
            name: self._get(name, params)
            for name in ("ready-year-charts", "filters", "index", "charts")
        }
        return {**params, "datasets": datasets}
