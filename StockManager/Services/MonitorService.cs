using System;
using System.Threading;

namespace StockManager.Services
{
        public class MonitorService
        {
                private readonly StockManagerService _stockManager;
                private readonly IMarketDataGateway _priceFetcher;
                private Thread _priceThread;
                private bool _running;
                private volatile bool _paused;
                private int _interval = 10;

                public MonitorService(StockManagerService stockManager, IMarketDataGateway priceFetcher)
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

                public void Pause()
                {
                        _paused = true;
                }

                public void Resume()
                {
                        _paused = false;
                }

                private void PriceUpdateLoop()
                {
                        while (_running)
                        {
                                try
                                {
                                        if (_paused)
                                        {
                                                Thread.Sleep(200);
                                                continue;
                                        }

                                        var tickers = _stockManager.GetTickers();
                                        foreach (var ticker in tickers)
                                        {
                                                if (!_running || _paused) break;

                                                // 使用新方法：獲取價格和前收盤價
                                                _priceFetcher.UpdatePriceWithPreviousClose(ticker);

                                                Thread.Sleep(300);
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
