"""
Data provider abstraction layer.

Supports two backends:
  1. Bloomberg Terminal API (blpapi) - requires Bloomberg license
  2. Yahoo Finance (yfinance) - free fallback

Both backends return data in a standardized pandas DataFrame format.
"""

import logging
from abc import ABC, abstractmethod
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)


class DataProvider(ABC):
    """Abstract base class for market data providers."""

    @abstractmethod
    def get_daily_snapshot(self, tickers: List[str]) -> pd.DataFrame:
        """
        Fetch today's snapshot for the given tickers.

        Returns a DataFrame with columns:
            ticker, name, open, high, low, close, prev_close,
            change, change_pct, volume, avg_volume_10d,
            volume_ratio, market_cap, pe_ratio, sector
        """
        pass

    @abstractmethod
    def get_intraday_prices(self, ticker: str) -> pd.DataFrame:
        """Fetch intraday price data for a single ticker."""
        pass

    @abstractmethod
    def is_connected(self) -> bool:
        """Check if the data provider connection is active."""
        pass


class BloombergProvider(DataProvider):
    """
    Bloomberg Terminal API data provider.

    Requires:
      - Bloomberg Terminal running on the machine (or SAPI access)
      - blpapi Python SDK installed (`pip install blpapi`)
      - Valid Bloomberg license
    """

    def __init__(self, host: str = "localhost", port: int = 8194):
        self.host = host
        self.port = port
        self._session = None
        self._connected = False
        self._connect()

    def _connect(self):
        """Establish connection to Bloomberg Terminal."""
        try:
            import blpapi

            session_options = blpapi.SessionOptions()
            session_options.setServerHost(self.host)
            session_options.setServerPort(self.port)

            self._session = blpapi.Session(session_options)

            if not self._session.start():
                logger.error("Failed to start Bloomberg session")
                return

            if not self._session.openService("//blp/refdata"):
                logger.error("Failed to open //blp/refdata service")
                return

            self._connected = True
            logger.info("Connected to Bloomberg Terminal at %s:%d", self.host, self.port)

        except ImportError:
            logger.error(
                "blpapi package not installed. Install with: pip install blpapi"
            )
            raise
        except Exception as e:
            logger.error("Failed to connect to Bloomberg: %s", e)
            raise

    def is_connected(self) -> bool:
        return self._connected

    def get_daily_snapshot(self, tickers: List[str]) -> pd.DataFrame:
        """Fetch daily snapshot from Bloomberg using //blp/refdata service."""
        import blpapi

        if not self._connected:
            raise ConnectionError("Not connected to Bloomberg Terminal")

        ref_data_service = self._session.getService("//blp/refdata")
        request = ref_data_service.createRequest("ReferenceDataRequest")

        # Convert tickers to Bloomberg format (e.g., "MSFT" -> "MSFT US Equity")
        bbg_tickers = [f"{t} US Equity" for t in tickers]
        for t in bbg_tickers:
            request.getElement("securities").appendValue(t)

        fields = [
            "PX_OPEN",
            "PX_HIGH",
            "PX_LOW",
            "PX_LAST",
            "PREV_CLOSING_PX",
            "CHG_PCT_1D",
            "VOLUME",
            "VOLUME_AVG_10D",
            "CUR_MKT_CAP",
            "PE_RATIO",
            "GICS_SECTOR_NAME",
            "LONG_COMP_NAME",
        ]
        for f in fields:
            request.getElement("fields").appendValue(f)

        self._session.sendRequest(request)

        rows = []
        while True:
            event = self._session.nextEvent(5000)
            for msg in event:
                if msg.messageType() == blpapi.Name("ReferenceDataResponse"):
                    security_data = msg.getElement("securityData")
                    for i in range(security_data.numValues()):
                        sec = security_data.getValueAsElement(i)
                        sec_name = sec.getElementAsString("security")
                        field_data = sec.getElement("fieldData")

                        ticker = sec_name.replace(" US Equity", "")
                        row = self._extract_bloomberg_fields(ticker, field_data)
                        rows.append(row)

            if event.eventType() == blpapi.Event.RESPONSE:
                break

        df = pd.DataFrame(rows)
        df = self._calculate_derived_fields(df)
        return df

    def _extract_bloomberg_fields(self, ticker: str, field_data) -> Dict:
        """Extract field values from Bloomberg response element."""

        def safe_get(element, field, default=None):
            try:
                if element.hasElement(field):
                    return element.getElementAsFloat(field)
            except Exception:
                pass
            return default

        def safe_get_str(element, field, default=""):
            try:
                if element.hasElement(field):
                    return element.getElementAsString(field)
            except Exception:
                pass
            return default

        return {
            "ticker": ticker,
            "name": safe_get_str(field_data, "LONG_COMP_NAME"),
            "open": safe_get(field_data, "PX_OPEN"),
            "high": safe_get(field_data, "PX_HIGH"),
            "low": safe_get(field_data, "PX_LOW"),
            "close": safe_get(field_data, "PX_LAST"),
            "prev_close": safe_get(field_data, "PREV_CLOSING_PX"),
            "change_pct": safe_get(field_data, "CHG_PCT_1D", 0.0),
            "volume": safe_get(field_data, "VOLUME", 0),
            "avg_volume_10d": safe_get(field_data, "VOLUME_AVG_10D", 0),
            "market_cap": safe_get(field_data, "CUR_MKT_CAP"),
            "pe_ratio": safe_get(field_data, "PE_RATIO"),
            "sector": safe_get_str(field_data, "GICS_SECTOR_NAME"),
        }

    def _calculate_derived_fields(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate derived fields from raw Bloomberg data."""
        if "close" in df.columns and "prev_close" in df.columns:
            df["change"] = df["close"] - df["prev_close"]
        if "volume" in df.columns and "avg_volume_10d" in df.columns:
            df["volume_ratio"] = df["volume"] / df["avg_volume_10d"].replace(0, 1)
        return df

    def get_intraday_prices(self, ticker: str) -> pd.DataFrame:
        """Fetch intraday bars from Bloomberg using //blp/refdata IntradayBarRequest."""
        import blpapi

        if not self._connected:
            raise ConnectionError("Not connected to Bloomberg Terminal")

        ref_data_service = self._session.getService("//blp/refdata")
        request = ref_data_service.createRequest("IntradayBarRequest")
        request.set("security", f"{ticker} US Equity")
        request.set("eventType", "TRADE")
        request.set("interval", 5)  # 5-minute bars

        today = datetime.now()
        request.set("startDateTime", today.replace(hour=9, minute=30, second=0))
        request.set("endDateTime", today.replace(hour=16, minute=0, second=0))

        self._session.sendRequest(request)

        bars = []
        while True:
            event = self._session.nextEvent(5000)
            for msg in event:
                bar_data = msg.getElement("barData").getElement("barTickData")
                for i in range(bar_data.numValues()):
                    bar = bar_data.getValueAsElement(i)
                    bars.append(
                        {
                            "time": bar.getElementAsDatetime("time"),
                            "open": bar.getElementAsFloat("open"),
                            "high": bar.getElementAsFloat("high"),
                            "low": bar.getElementAsFloat("low"),
                            "close": bar.getElementAsFloat("close"),
                            "volume": bar.getElementAsInteger("volume"),
                        }
                    )
            if event.eventType() == blpapi.Event.RESPONSE:
                break

        return pd.DataFrame(bars)


class YahooFinanceProvider(DataProvider):
    """
    Yahoo Finance data provider using yfinance.

    Free alternative to Bloomberg - no license required.
    Suitable for EOD and delayed intraday data.
    """

    def __init__(self):
        try:
            import yfinance  # noqa: F401
            self._connected = True
            logger.info("Yahoo Finance provider initialized")
        except ImportError:
            logger.error("yfinance not installed. Install with: pip install yfinance")
            raise

    def is_connected(self) -> bool:
        return self._connected

    def get_daily_snapshot(self, tickers: List[str]) -> pd.DataFrame:
        """
        Fetch daily snapshot data for the given tickers via yfinance.

        Uses batch download for efficiency and supplements with
        individual ticker info for fundamental data.
        """
        import yfinance as yf

        logger.info("Fetching daily data for %d tickers via Yahoo Finance...", len(tickers))

        # Batch download recent price history (2 trading days for prev_close)
        end_date = datetime.now()
        start_date = end_date - timedelta(days=10)  # extra buffer for weekends/holidays

        try:
            data = yf.download(
                tickers,
                start=start_date.strftime("%Y-%m-%d"),
                end=end_date.strftime("%Y-%m-%d"),
                group_by="ticker",
                auto_adjust=False,
                progress=False,
                threads=True,
            )
        except Exception as e:
            logger.error("Failed to download batch data: %s", e)
            data = pd.DataFrame()

        rows = []
        for ticker in tickers:
            try:
                row = self._build_ticker_row(ticker, data, tickers)
                if row is not None:
                    rows.append(row)
            except Exception as e:
                logger.warning("Failed to process %s: %s", ticker, e)

        if not rows:
            logger.error("No data retrieved for any tickers")
            return pd.DataFrame()

        df = pd.DataFrame(rows)
        return df

    def _build_ticker_row(
        self, ticker: str, batch_data: pd.DataFrame, all_tickers: List[str]
    ) -> Optional[Dict]:
        """Build a single row of data for one ticker."""
        import yfinance as yf

        try:
            # Extract price data from batch download
            if len(all_tickers) > 1 and not batch_data.empty:
                try:
                    ticker_data = batch_data[ticker].dropna(how="all")
                except (KeyError, TypeError):
                    ticker_data = pd.DataFrame()
            elif len(all_tickers) == 1 and not batch_data.empty:
                ticker_data = batch_data.dropna(how="all")
            else:
                ticker_data = pd.DataFrame()

            # Get ticker info for fundamentals
            yf_ticker = yf.Ticker(ticker)
            info = {}
            try:
                info = yf_ticker.info or {}
            except Exception:
                pass

            if ticker_data.empty or len(ticker_data) < 1:
                logger.warning("No price data for %s", ticker)
                return None

            # Latest trading day
            latest = ticker_data.iloc[-1]

            # Previous close
            if len(ticker_data) >= 2:
                prev_close = float(ticker_data["Close"].iloc[-2])
            else:
                prev_close = info.get("previousClose") or info.get("regularMarketPreviousClose")
                if prev_close is None:
                    prev_close = float(latest.get("Open", latest["Close"]))

            close_price = float(latest["Close"])
            open_price = float(latest["Open"])
            high_price = float(latest["High"])
            low_price = float(latest["Low"])
            volume = int(latest.get("Volume", 0))

            change = close_price - prev_close
            change_pct = (change / prev_close * 100) if prev_close else 0.0

            # Average volume (10-day)
            avg_vol = info.get("averageDailyVolume10Day") or info.get("averageVolume10days", 0)
            if avg_vol == 0 and len(ticker_data) >= 3:
                avg_vol = int(ticker_data["Volume"].tail(10).mean())

            volume_ratio = (volume / avg_vol) if avg_vol and avg_vol > 0 else 1.0

            return {
                "ticker": ticker,
                "name": info.get("shortName") or info.get("longName", ticker),
                "open": open_price,
                "high": high_price,
                "low": low_price,
                "close": close_price,
                "prev_close": prev_close,
                "change": round(change, 2),
                "change_pct": round(change_pct, 2),
                "volume": volume,
                "avg_volume_10d": int(avg_vol) if avg_vol else 0,
                "volume_ratio": round(volume_ratio, 2),
                "market_cap": info.get("marketCap"),
                "pe_ratio": info.get("trailingPE") or info.get("forwardPE"),
                "sector": info.get("sector", "Technology"),
                "fifty_two_week_high": info.get("fiftyTwoWeekHigh"),
                "fifty_two_week_low": info.get("fiftyTwoWeekLow"),
                "beta": info.get("beta"),
            }

        except Exception as e:
            logger.warning("Error building row for %s: %s", ticker, e)
            return None

    def get_intraday_prices(self, ticker: str) -> pd.DataFrame:
        """Fetch intraday price data using yfinance."""
        import yfinance as yf

        try:
            yf_ticker = yf.Ticker(ticker)
            data = yf_ticker.history(period="1d", interval="5m")
            if data.empty:
                return pd.DataFrame()

            data = data.reset_index()
            data.columns = [c.lower() for c in data.columns]
            return data[["datetime", "open", "high", "low", "close", "volume"]]
        except Exception as e:
            logger.warning("Failed to get intraday data for %s: %s", ticker, e)
            return pd.DataFrame()


def create_provider(config: dict) -> DataProvider:
    """
    Factory function to create the appropriate data provider
    based on configuration.
    """
    provider_type = config.get("data_source", {}).get("provider", "yahoo").lower()

    if provider_type == "bloomberg":
        bbg_config = config.get("data_source", {}).get("bloomberg", {})
        host = bbg_config.get("host", "localhost")
        port = bbg_config.get("port", 8194)
        try:
            return BloombergProvider(host=host, port=port)
        except Exception as e:
            logger.warning(
                "Bloomberg connection failed (%s). Falling back to Yahoo Finance.", e
            )
            return YahooFinanceProvider()
    else:
        return YahooFinanceProvider()
