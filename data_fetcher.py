"""
data_fetcher.py
Responsável por extrair e estruturar dados financeiros via yfinance.
"""

import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import ta
import json
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def _safe_float(val, default=None):
    """Converte valores para float de forma segura."""
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return default
        return float(val)
    except (TypeError, ValueError):
        return default


def _safe_int(val, default=None):
    """Converte valores para int de forma segura."""
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return default
        return int(val)
    except (TypeError, ValueError):
        return default


def _format_number(val, scale=1e9):
    """Formata números em bilhões/milhões para leitura humana."""
    f = _safe_float(val)
    if f is None:
        return None
    return round(f / scale, 3)


def _series_to_dict(series, scale=1e9, years=5):
    """Converte pandas Series em dict com as últimas N entradas."""
    if series is None or series.empty:
        return {}
    result = {}
    for date, val in series.items():
        key = str(date)[:10]
        result[key] = _format_number(val, scale)
    # Retorna apenas os últimos N anos
    sorted_keys = sorted(result.keys(), reverse=True)[:years]
    return {k: result[k] for k in sorted_keys}


def _get_income_statement(ticker_obj):
    """Extrai demonstração de resultados (5 anos)."""
    try:
        inc = ticker_obj.financials
        if inc is None or inc.empty:
            return {}

        rows_of_interest = [
            "Total Revenue", "Gross Profit", "Operating Income",
            "Net Income", "EBITDA", "Basic EPS", "Diluted EPS"
        ]
        result = {}
        for row in rows_of_interest:
            if row in inc.index:
                result[row] = _series_to_dict(inc.loc[row])
        return result
    except Exception as e:
        logger.warning(f"Erro ao extrair Income Statement: {e}")
        return {}


def _get_balance_sheet(ticker_obj):
    """Extrai balanço patrimonial."""
    try:
        bs = ticker_obj.balance_sheet
        if bs is None or bs.empty:
            return {}

        rows_of_interest = [
            "Total Assets", "Total Liabilities Net Minority Interest",
            "Total Equity Gross Minority Interest", "Total Debt",
            "Cash And Cash Equivalents", "Long Term Debt",
            "Current Assets", "Current Liabilities"
        ]
        result = {}
        for row in rows_of_interest:
            if row in bs.index:
                result[row] = _series_to_dict(bs.loc[row])
        return result
    except Exception as e:
        logger.warning(f"Erro ao extrair Balance Sheet: {e}")
        return {}


def _get_cash_flow(ticker_obj):
    """Extrai demonstração de fluxo de caixa."""
    try:
        cf = ticker_obj.cashflow
        if cf is None or cf.empty:
            return {}

        rows_of_interest = [
            "Free Cash Flow", "Operating Cash Flow",
            "Capital Expenditure", "Depreciation And Amortization",
            "Stock Based Compensation", "Changes In Working Capital"
        ]
        result = {}
        for row in rows_of_interest:
            if row in cf.index:
                result[row] = _series_to_dict(cf.loc[row])
        return result
    except Exception as e:
        logger.warning(f"Erro ao extrair Cash Flow: {e}")
        return {}


def _get_multiples(ticker_obj, info):
    """Extrai múltiplos de avaliação TTM e Forward."""
    return {
        "pe_ttm": _safe_float(info.get("trailingPE")),
        "pe_forward": _safe_float(info.get("forwardPE")),
        "ps_ttm": _safe_float(info.get("priceToSalesTrailing12Months")),
        "pb": _safe_float(info.get("priceToBook")),
        "ev_ebitda": _safe_float(info.get("enterpriseToEbitda")),
        "ev_revenue": _safe_float(info.get("enterpriseToRevenue")),
        "peg_ratio": _safe_float(info.get("pegRatio")),
        "market_cap_bn": _format_number(info.get("marketCap")),
        "enterprise_value_bn": _format_number(info.get("enterpriseValue")),
        "revenue_ttm_bn": _format_number(info.get("totalRevenue")),
        "gross_margin": _safe_float(info.get("grossMargins")),
        "operating_margin": _safe_float(info.get("operatingMargins")),
        "net_margin": _safe_float(info.get("profitMargins")),
        "revenue_growth_yoy": _safe_float(info.get("revenueGrowth")),
        "earnings_growth_yoy": _safe_float(info.get("earningsGrowth")),
        "return_on_equity": _safe_float(info.get("returnOnEquity")),
        "return_on_assets": _safe_float(info.get("returnOnAssets")),
        "debt_to_equity": _safe_float(info.get("debtToEquity")),
        "current_ratio": _safe_float(info.get("currentRatio")),
        "quick_ratio": _safe_float(info.get("quickRatio")),
        "shares_outstanding_bn": _format_number(info.get("sharesOutstanding"), scale=1e9),
        "float_shares_bn": _format_number(info.get("floatShares"), scale=1e9),
        "beta": _safe_float(info.get("beta")),
        "52w_high": _safe_float(info.get("fiftyTwoWeekHigh")),
        "52w_low": _safe_float(info.get("fiftyTwoWeekLow")),
        "dividend_yield": _safe_float(info.get("dividendYield")),
        "payout_ratio": _safe_float(info.get("payoutRatio")),
    }


def _get_technical_indicators(ticker_obj):
    """Calcula MA50, MA200, RSI, MACD com histórico de 2 anos."""
    try:
        hist = ticker_obj.history(period="2y")
        if hist.empty:
            return {}

        close = hist["Close"]
        volume = hist["Volume"]

        # Moving Averages
        ma50 = close.rolling(window=50).mean().iloc[-1]
        ma200 = close.rolling(window=200).mean().iloc[-1]
        current_price = close.iloc[-1]

        # RSI
        rsi_series = ta.momentum.RSIIndicator(close=close, window=14).rsi()
        rsi = rsi_series.iloc[-1]

        # MACD
        macd_ind = ta.trend.MACD(close=close)
        macd_line = macd_ind.macd().iloc[-1]
        macd_signal = macd_ind.macd_signal().iloc[-1]
        macd_hist = macd_ind.macd_diff().iloc[-1]

        # Volume avg 20d
        vol_avg_20 = volume.rolling(window=20).mean().iloc[-1]

        # Price history (últimos 12 meses, mensal)
        monthly = hist["Close"].resample("ME").last().tail(12)
        price_history = {str(d)[:10]: round(float(v), 2) for d, v in monthly.items()}

        return {
            "current_price": round(float(current_price), 2),
            "ma50": round(float(ma50), 2) if not np.isnan(ma50) else None,
            "ma200": round(float(ma200), 2) if not np.isnan(ma200) else None,
            "price_vs_ma50_pct": round((float(current_price) / float(ma50) - 1) * 100, 2) if not np.isnan(ma50) else None,
            "price_vs_ma200_pct": round((float(current_price) / float(ma200) - 1) * 100, 2) if not np.isnan(ma200) else None,
            "rsi_14": round(float(rsi), 2) if not np.isnan(rsi) else None,
            "macd_line": round(float(macd_line), 4) if not np.isnan(macd_line) else None,
            "macd_signal": round(float(macd_signal), 4) if not np.isnan(macd_signal) else None,
            "macd_histogram": round(float(macd_hist), 4) if not np.isnan(macd_hist) else None,
            "volume_avg_20d": _safe_int(vol_avg_20),
            "price_history_monthly": price_history,
        }
    except Exception as e:
        logger.warning(f"Erro ao calcular indicadores técnicos: {e}")
        return {}


def _get_insider_data(ticker_obj):
    """Extrai dados de insider transactions."""
    try:
        insider_txn = ticker_obj.insider_transactions
        if insider_txn is None or insider_txn.empty:
            return {"transactions": [], "summary": "No insider data available"}

        # Últimas 10 transações
        recent = insider_txn.head(10)
        transactions = []
        for _, row in recent.iterrows():
            transactions.append({
                "insider": str(row.get("Insider Trading", "")),
                "relation": str(row.get("Relationship", "")),
                "date": str(row.get("Start Date", ""))[:10],
                "transaction": str(row.get("Transaction", "")),
                "shares": _safe_int(row.get("Shares")),
                "value": _format_number(row.get("Value"), scale=1e6),
            })

        # Sumariza compras vs vendas
        buys = sum(1 for t in transactions if "purchase" in t["transaction"].lower())
        sells = sum(1 for t in transactions if "sale" in t["transaction"].lower())
        return {
            "transactions": transactions,
            "recent_buys": buys,
            "recent_sells": sells,
        }
    except Exception as e:
        logger.warning(f"Erro ao extrair dados de insiders: {e}")
        return {"transactions": [], "summary": "Error fetching insider data"}


def _get_company_info(info):
    """Extrai informações gerais da empresa."""
    return {
        "name": info.get("longName", ""),
        "ticker": info.get("symbol", ""),
        "sector": info.get("sector", ""),
        "industry": info.get("industry", ""),
        "country": info.get("country", ""),
        "exchange": info.get("exchange", ""),
        "currency": info.get("currency", "USD"),
        "website": info.get("website", ""),
        "employees": _safe_int(info.get("fullTimeEmployees")),
        "description": info.get("longBusinessSummary", "")[:800],
    }


def fetch_all_data(ticker: str) -> dict:
    """
    Função principal. Recolhe todos os dados para o ticker fornecido.
    Retorna um dicionário JSON estruturado ou lança ValueError se o ticker for inválido.
    """
    ticker = ticker.strip().upper()
    logger.info(f"A recolher dados para: {ticker}")

    try:
        t = yf.Ticker(ticker)
        info = t.info

        # Validação: ticker inválido retorna info sem campos essenciais
        if not info or info.get("quoteType") is None or info.get("regularMarketPrice") is None:
            # Tentativa alternativa
            if not info.get("longName") and not info.get("shortName"):
                raise ValueError(f"Ticker '{ticker}' não encontrado ou sem dados disponíveis.")

        company_info = _get_company_info(info)
        multiples = _get_multiples(t, info)
        income_stmt = _get_income_statement(t)
        balance_sheet = _get_balance_sheet(t)
        cash_flow = _get_cash_flow(t)
        technical = _get_technical_indicators(t)
        insiders = _get_insider_data(t)

        payload = {
            "meta": {
                "ticker": ticker,
                "generated_at": datetime.utcnow().isoformat() + "Z",
                "data_currency": company_info.get("currency", "USD"),
                "unit": "Billions USD (where applicable)",
            },
            "company": company_info,
            "multiples": multiples,
            "income_statement": income_stmt,
            "balance_sheet": balance_sheet,
            "cash_flow": cash_flow,
            "technical": technical,
            "insiders": insiders,
        }

        logger.info(f"Dados recolhidos com sucesso para {ticker}")
        return payload

    except ValueError:
        raise
    except Exception as e:
        logger.error(f"Erro fatal ao recolher dados para {ticker}: {e}")
        raise ValueError(f"Falha ao recolher dados para '{ticker}': {str(e)}")
