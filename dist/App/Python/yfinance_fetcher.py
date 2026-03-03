#!/usr/bin/env python3
"""
yfinance Stock Price Fetcher
為 StockManager 提供股價數據

需要安裝: pip install yfinance
"""

import yfinance as yf
import json
import sys
from datetime import datetime

def get_stock_price(ticker):
    """
    獲取股票的實時價格和前收盤價
    
    Args:
        ticker: 股票代碼 (例如: AAPL, 2330.TW)
    
    Returns:
        JSON 格式的股票數據
    """
    try:
        # 創建 Ticker 對象
        stock = yf.Ticker(ticker)
        
        # 獲取實時數據
        info = stock.info
        
        # 獲取歷史數據（用於獲取前收盤價）
        hist = stock.history(period="2d")
        
        # 提取數據
        current_price = info.get('currentPrice') or info.get('regularMarketPrice')
        previous_close = info.get('previousClose') or info.get('regularMarketPreviousClose')
        
        # 如果 info 中沒有，從歷史數據中獲取
        if previous_close is None and len(hist) >= 2:
            previous_close = float(hist['Close'].iloc[-2])
        
        if current_price is None and len(hist) >= 1:
            current_price = float(hist['Close'].iloc[-1])
        
        # 計算漲跌幅
        change_percent = None
        if current_price and previous_close and previous_close != 0:
            change = current_price - previous_close
            change_percent = (change / previous_close) * 100
        
        # 獲取其他信息
        open_price = info.get('open') or info.get('regularMarketOpen')
        high_price = info.get('dayHigh') or info.get('regularMarketDayHigh')
        low_price = info.get('dayLow') or info.get('regularMarketDayLow')
        volume = info.get('volume') or info.get('regularMarketVolume')
        
        # 市場狀態
        market_state = info.get('marketState', 'REGULAR')
        
        # 構建結果
        result = {
            'success': True,
            'ticker': ticker,
            'data': {
                'current_price': current_price,
                'previous_close': previous_close,
                'change_percent': change_percent,
                'open': open_price,
                'high': high_price,
                'low': low_price,
                'volume': volume,
                'market_state': market_state,
                'timestamp': datetime.now().isoformat()
            }
        }
        
        return result
        
    except Exception as e:
        return {
            'success': False,
            'ticker': ticker,
            'error': str(e),
            'error_type': type(e).__name__
        }

def print_stock_history_lines(ticker, period):
    """
    輸出歷史 K 線資料（給 C# 解析）
    格式:
      HISTORY_OK
      yyyy-MM-dd|open|high|low|close|volume
    """
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period=period, interval="1d")

        if hist is None or hist.empty:
            print("HISTORY_ERROR|No historical data")
            return

        print("HISTORY_OK")
        for idx, row in hist.iterrows():
            date_str = idx.strftime("%Y-%m-%d")
            open_price = float(row.get("Open", 0.0))
            high_price = float(row.get("High", 0.0))
            low_price = float(row.get("Low", 0.0))
            close_price = float(row.get("Close", 0.0))
            volume = int(row.get("Volume", 0))

            print(f"{date_str}|{open_price}|{high_price}|{low_price}|{close_price}|{volume}")
    except Exception as e:
        print(f"HISTORY_ERROR|{str(e)}")

def print_fundamentals_lines(ticker):
    """
    輸出財報/基本面資料（給 C# 解析）
    格式:
      FUNDAMENTALS_OK
      key|value
    """
    try:
        stock = yf.Ticker(ticker)
        info = stock.info or {}

        fields = {
            "trailingPE": info.get("trailingPE"),
            "forwardPE": info.get("forwardPE"),
            "earningsGrowth": info.get("earningsGrowth"),
            "revenueGrowth": info.get("revenueGrowth"),
            "profitMargins": info.get("profitMargins"),
            "returnOnEquity": info.get("returnOnEquity"),
            "debtToEquity": info.get("debtToEquity"),
            "currentRatio": info.get("currentRatio"),
            "marketCap": info.get("marketCap"),
            "regularMarketPrice": info.get("regularMarketPrice"),
        }

        print("FUNDAMENTALS_OK")
        for k, v in fields.items():
            if v is None:
                continue
            print(f"{k}|{v}")
    except Exception as e:
        print(f"FUNDAMENTALS_ERROR|{str(e)}")

def main():
    """主函數：從命令行參數獲取股票代碼"""
    if len(sys.argv) < 2:
        print(json.dumps({
            'success': False,
            'error': 'Usage: python yfinance_fetcher.py <ticker>',
            'error_type': 'InvalidArgument'
        }))
        sys.exit(1)
    
    ticker = sys.argv[1]

    # 歷史模式: python yfinance_fetcher.py <ticker> history <period>
    if len(sys.argv) >= 3 and sys.argv[2].lower() == 'history':
        period = sys.argv[3] if len(sys.argv) >= 4 else '3mo'
        print_stock_history_lines(ticker, period)
        return

    # 財報/基本面模式: python yfinance_fetcher.py <ticker> fundamentals
    if len(sys.argv) >= 3 and sys.argv[2].lower() == 'fundamentals':
        print_fundamentals_lines(ticker)
        return

    result = get_stock_price(ticker)
    
    # 輸出 JSON（供 C# 讀取）
    print(json.dumps(result, indent=2))

if __name__ == '__main__':
    main()
