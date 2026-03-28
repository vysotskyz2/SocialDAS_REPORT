from datetime import datetime
from httpx import AsyncClient
from loguru import logger


class AnalyticsClient:
    def __init__(self, http_client: AsyncClient, base_url: str) -> None:
        self._client = http_client
        self._base_url = base_url

    async def _get(self, path: str, params: dict | None = None) -> dict | list | None:
        url = f"{self._base_url}{path}"
        try:
            response = await self._client.get(url, params=params, timeout=30.0)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            logger.error(f"Ошибка обработки API Analytic: {url} - {e}")
            return None

    @staticmethod
    def _date_params(
        date_from: datetime | None, date_to: datetime | None
    ) -> dict:
        params = {}
        if date_from:
            params["date_from"] = date_from.isoformat()
        if date_to:
            params["date_to"] = date_to.isoformat()
        return params

    async def get_overview(
        self, platform: str, account_id: str
    ) -> dict | None:
        return await self._get(f"/{platform}/{account_id}/overview")

    async def get_followers(
        self,
        platform: str,
        account_id: str,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
    ) -> dict | None:
        endpoint = "subscribers" if platform == "youtube" else "followers"
        return await self._get(
            f"/{platform}/{account_id}/{endpoint}",
            params=self._date_params(date_from, date_to),
        )

    async def get_posts(
        self,
        platform: str,
        account_id: str,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
        limit: int = 100,
    ) -> dict | None:
        endpoint = "posts" if platform == "instagram" else "videos"
        params = {**self._date_params(date_from, date_to), "limit": limit}
        return await self._get(
            f"/{platform}/{account_id}/{endpoint}", params=params
        )

    async def get_engagement(
        self,
        platform: str,
        account_id: str,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
    ) -> dict | None:
        return await self._get(
            f"/{platform}/{account_id}/engagement",
            params=self._date_params(date_from, date_to),
        )

    async def get_growth(
        self,
        platform: str,
        account_id: str,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
        projection_days: int = 14,
    ) -> dict | None:
        params = {
            **self._date_params(date_from, date_to),
            "projection_days": projection_days,
        }
        return await self._get(
            f"/{platform}/{account_id}/growth", params=params
        )

    async def get_content_performance(
        self,
        platform: str,
        account_id: str,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
        limit: int = 50,
    ) -> dict | None:
        params = {**self._date_params(date_from, date_to), "limit": limit}
        return await self._get(
            f"/{platform}/{account_id}/content-performance", params=params
        )

    async def get_posting_patterns(
        self,
        platform: str,
        account_id: str,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
    ) -> dict | None:
        return await self._get(
            f"/{platform}/{account_id}/posting-patterns",
            params=self._date_params(date_from, date_to),
        )

    async def get_trends(
        self,
        platform: str,
        account_id: str,
        date_from: datetime | None = None,
        date_to: datetime | None = None,
    ) -> dict | None:
        return await self._get(
            f"/{platform}/{account_id}/trends",
            params=self._date_params(date_from, date_to),
        )
