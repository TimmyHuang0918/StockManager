using System;
using System.Collections.Generic;
using System.Threading;

namespace StockManager.Services
{
        public class MonitorService
        {
                private readonly StockManagerService _stockManager;
                private readonly PriceFetcherService _priceFetcher;
                private Thread _priceThread;
                private bool _running;
                private int _interval = 10;

                public MonitorService(StockManagerService stockManager, PriceFetcherService priceFetcher)
                {
                        _stockManager = stockManager;
                        _priceFetcher = priceFetcher;
                }

                public void StartThreads(int interval = 10)
                {
                        _interval = interval;
                        _running = true;
                        _priceThread = new Thread(PriceUpdateLoop);
                        _priceThread.IsBackground = true;
                        _priceThread.Start();
                }

                public void StopThreads()
                {
                        _running = false;
                        if (_priceThread != null && _priceThread.IsAlive)
                        {
                                _priceThread.Join(2000);
                        }
                }

                private void PriceUpdateLoop()
                {
                        while (_running)
                        {
                                try
                                {
                                        var tickers = _stockManager.GetTickers();
                                        foreach (var ticker in tickers)
                                        {
                                                if (!_running) break;

                                                // 使用新方法：獲取價格和前收盤價
                                                _priceFetcher.UpdatePriceWithPreviousClose(ticker);

                                                Thread.Sleep(600);
                                        }
                                        Thread.Sleep(_interval * 1000);
                                }
                                catch (Exception ex)
                                {
                                        Console.WriteLine($"價格更新錯誤: {ex.Message}");
                                        Thread.Sleep(5000);
                                }
                        }
                }
        }
}
