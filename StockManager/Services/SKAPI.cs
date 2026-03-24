using SKCOMLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockManager.Services
{
    public class SKAPI
    {
	/// <summary>
	/// 全域單例。
	/// </summary>
	public static SKAPI Instance { get; } = new SKAPI();

	/// <summary>登入、環境與回傳碼相關元件。</summary>
	public SKCenterLib Center { get; } = new SKCenterLib();
	/// <summary>回報相關元件。</summary>
	public SKReplyLib Reply { get; } = new SKReplyLib();
	/// <summary>下單相關元件。</summary>
	public SKOrderLib Order { get; } = new SKOrderLib();
	/// <summary>國內報價元件。</summary>
	public SKQuoteLib Quote { get; } = new SKQuoteLib();
	/// <summary>海外期貨報價元件。</summary>
	public SKOSQuoteLib OSQuote { get; } = new SKOSQuoteLib();
	/// <summary>海外選擇權報價元件。</summary>
	public SKOOQuoteLib OOQuote { get; } = new SKOOQuoteLib();

	#region Center 事件
	public event _ISKCenterLibEvents_OnShowAgreementEventHandler OnShowAgreement
	{
	    add { Center.OnShowAgreement += value; }
	    remove { Center.OnShowAgreement -= value; }
	}

	public event _ISKCenterLibEvents_OnTimerEventHandler OnTimer
	{
	    add { Center.OnTimer += value; }
	    remove { Center.OnTimer -= value; }
	}

	public event _ISKCenterLibEvents_OnNotifySGXAPIOrderStatusEventHandler OnNotifySGXAPIOrderStatus
	{
	    add { Center.OnNotifySGXAPIOrderStatus += value; }
	    remove { Center.OnNotifySGXAPIOrderStatus -= value; }
	}
	#endregion

	#region Reply 事件
	public event _ISKReplyLibEvents_OnReplyMessageEventHandler OnReplyMessage
	{
	    add { Reply.OnReplyMessage += value; }
	    remove { Reply.OnReplyMessage -= value; }
	}

	public event _ISKReplyLibEvents_OnNewDataEventHandler OnNewData
	{
	    add { Reply.OnNewData += value; }
	    remove { Reply.OnNewData -= value; }
	}

	public event _ISKReplyLibEvents_OnCompleteEventHandler OnComplete
	{
	    add { Reply.OnComplete += value; }
	    remove { Reply.OnComplete -= value; }
	}

	public event _ISKReplyLibEvents_OnReplyClearEventHandler OnReplyClear
	{
	    add { Reply.OnReplyClear += value; }
	    remove { Reply.OnReplyClear -= value; }
	}

	public event _ISKReplyLibEvents_OnReplyClearMessageEventHandler OnReplyClearMessage
	{
	    add { Reply.OnReplyClearMessage += value; }
	    remove { Reply.OnReplyClearMessage -= value; }
	}

	public event _ISKReplyLibEvents_OnSolaceReplyDisconnectEventHandler OnSolaceReplyDisconnect
	{
	    add { Reply.OnSolaceReplyDisconnect += value; }
	    remove { Reply.OnSolaceReplyDisconnect -= value; }
	}

	public event _ISKReplyLibEvents_OnSolaceReplyConnectionEventHandler OnSolaceReplyConnection
	{
	    add { Reply.OnSolaceReplyConnection += value; }
	    remove { Reply.OnSolaceReplyConnection -= value; }
	}

	public event _ISKReplyLibEvents_OnStrategyDataEventHandler OnStrategyData
	{
	    add { Reply.OnStrategyData += value; }
	    remove { Reply.OnStrategyData -= value; }
	}
	#endregion

	#region Order 事件
	public event _ISKOrderLibEvents_OnAccountEventHandler OnAccount
	{
	    add { Order.OnAccount += value; }
	    remove { Order.OnAccount -= value; }
	}

	public event _ISKOrderLibEvents_OnProxyStatusEventHandler OnProxyStatus
	{
	    add { Order.OnProxyStatus += value; }
	    remove { Order.OnProxyStatus -= value; }
	}

	public event _ISKOrderLibEvents_OnTelnetTestEventHandler OnTelnetTest
	{
	    add { Order.OnTelnetTest += value; }
	    remove { Order.OnTelnetTest -= value; }
	}

	public event _ISKOrderLibEvents_OnProxyOrderEventHandler OnProxyOrder
	{
	    add { Order.OnProxyOrder += value; }
	    remove { Order.OnProxyOrder -= value; }
	}

	public event _ISKOrderLibEvents_OnAsyncOrderEventHandler OnAsyncOrder
	{
	    add { Order.OnAsyncOrder += value; }
	    remove { Order.OnAsyncOrder -= value; }
	}

	public event _ISKOrderLibEvents_OnOFSmartStrategyReportEventHandler OnOFSmartStrategyReport
	{
	    add { Order.OnOFSmartStrategyReport += value; }
	    remove { Order.OnOFSmartStrategyReport -= value; }
	}

	public event _ISKOrderLibEvents_OnStopLossReportEventHandler OnStopLossReport
	{
	    add { Order.OnStopLossReport += value; }
	    remove { Order.OnStopLossReport -= value; }
	}

	public event _ISKOrderLibEvents_OnTSSmartStrategyReportEventHandler OnTSSmartStrategyReport
	{
	    add { Order.OnTSSmartStrategyReport += value; }
	    remove { Order.OnTSSmartStrategyReport -= value; }
	}

	public event _ISKOrderLibEvents_OnBalanceQueryEventHandler OnBalanceQuery
	{
	    add { Order.OnBalanceQuery += value; }
	    remove { Order.OnBalanceQuery -= value; }
	}

	public event _ISKOrderLibEvents_OnFutureRightsEventHandler OnFutureRights
	{
	    add { Order.OnFutureRights += value; }
	    remove { Order.OnFutureRights -= value; }
	}

	public event _ISKOrderLibEvents_OnMarginPurchaseAmountLimitEventHandler OnMarginPurchaseAmountLimit
	{
	    add { Order.OnMarginPurchaseAmountLimit += value; }
	    remove { Order.OnMarginPurchaseAmountLimit -= value; }
	}

	public event _ISKOrderLibEvents_OnOFOpenInterestGWReportEventHandler OnOFOpenInterestGWReport
	{
	    add { Order.OnOFOpenInterestGWReport += value; }
	    remove { Order.OnOFOpenInterestGWReport -= value; }
	}

	public event _ISKOrderLibEvents_OnOpenInterestEventHandler OnOpenInterest
	{
	    add { Order.OnOpenInterest += value; }
	    remove { Order.OnOpenInterest -= value; }
	}

	public event _ISKOrderLibEvents_OnOverseaFutureEventHandler OnOverseaFuture
	{
	    add { Order.OnOverseaFuture += value; }
	    remove { Order.OnOverseaFuture -= value; }
	}

	public event _ISKOrderLibEvents_OnOverSeaFutureRightEventHandler OnOverSeaFutureRight
	{
	    add { Order.OnOverSeaFutureRight += value; }
	    remove { Order.OnOverSeaFutureRight -= value; }
	}

	public event _ISKOrderLibEvents_OnOverseaOptionEventHandler OnOverseaOption
	{
	    add { Order.OnOverseaOption += value; }
	    remove { Order.OnOverseaOption -= value; }
	}

	public event _ISKOrderLibEvents_OnProfitLossGWReportEventHandler OnProfitLossGWReport
	{
	    add { Order.OnProfitLossGWReport += value; }
	    remove { Order.OnProfitLossGWReport -= value; }
	}

	public event _ISKOrderLibEvents_OnRealBalanceReportEventHandler OnRealBalanceReport
	{
	    add { Order.OnRealBalanceReport += value; }
	    remove { Order.OnRealBalanceReport -= value; }
	}
	#endregion

	#region Quote 事件
	public event _ISKQuoteLibEvents_OnConnectionEventHandler OnConnection
	{
	    add { Quote.OnConnection += value; }
	    remove { Quote.OnConnection -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyServerTimeEventHandler OnNotifyServerTime
	{
	    add { Quote.OnNotifyServerTime += value; }
	    remove { Quote.OnNotifyServerTime -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyQuoteLONGEventHandler OnNotifyQuoteLONG
	{
	    add { Quote.OnNotifyQuoteLONG += value; }
	    remove { Quote.OnNotifyQuoteLONG -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyTicksLONGEventHandler OnNotifyTicksLONG
	{
	    add { Quote.OnNotifyTicksLONG += value; }
	    remove { Quote.OnNotifyTicksLONG -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyBest5LONGEventHandler OnNotifyBest5LONG
	{
	    add { Quote.OnNotifyBest5LONG += value; }
	    remove { Quote.OnNotifyBest5LONG -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyHistoryTicksLONGEventHandler OnNotifyHistoryTicksLONG
	{
	    add { Quote.OnNotifyHistoryTicksLONG += value; }
	    remove { Quote.OnNotifyHistoryTicksLONG -= value; }
	}

	public event _ISKQuoteLibEvents_OnKLineCompleteEventHandler OnKLineComplete
	{
	    add { Quote.OnKLineComplete += value; }
	    remove { Quote.OnKLineComplete -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyMarketTotEventHandler OnNotifyMarketTot
	{
	    add { Quote.OnNotifyMarketTot += value; }
	    remove { Quote.OnNotifyMarketTot -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyMarketBuySellEventHandler OnNotifyMarketBuySell
	{
	    add { Quote.OnNotifyMarketBuySell += value; }
	    remove { Quote.OnNotifyMarketBuySell -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyMarketHighLowNoWarrantEventHandler OnNotifyMarketHighLowNoWarrant
	{
	    add { Quote.OnNotifyMarketHighLowNoWarrant += value; }
	    remove { Quote.OnNotifyMarketHighLowNoWarrant -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyMACDLONGEventHandler OnNotifyMACDLONG
	{
	    add { Quote.OnNotifyMACDLONG += value; }
	    remove { Quote.OnNotifyMACDLONG -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyBoolTunelLONGEventHandler OnNotifyBoolTunelLONG
	{
	    add { Quote.OnNotifyBoolTunelLONG += value; }
	    remove { Quote.OnNotifyBoolTunelLONG -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyKLineDataEventHandler OnNotifyKLineData
	{
	    add { Quote.OnNotifyKLineData += value; }
	    remove { Quote.OnNotifyKLineData -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyCommodityListWithTypeNoEventHandler OnNotifyCommodityListWithTypeNo
	{
	    add { Quote.OnNotifyCommodityListWithTypeNo += value; }
	    remove { Quote.OnNotifyCommodityListWithTypeNo -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyFutureTradeInfoLONGEventHandler OnNotifyFutureTradeInfoLONG
	{
	    add { Quote.OnNotifyFutureTradeInfoLONG += value; }
	    remove { Quote.OnNotifyFutureTradeInfoLONG -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyStrikePricesEventHandler OnNotifyStrikePrices
	{
	    add { Quote.OnNotifyStrikePrices += value; }
	    remove { Quote.OnNotifyStrikePrices -= value; }
	}

	public event _ISKQuoteLibEvents_OnNotifyOddLotSpreadDealEventHandler OnNotifyOddLotSpreadDeal
	{
	    add { Quote.OnNotifyOddLotSpreadDeal += value; }
	    remove { Quote.OnNotifyOddLotSpreadDeal -= value; }
	}

	#region OOQuote 事件
	public event _ISKOOQuoteLibEvents_OnConnectEventHandler OnOOConnect
	{
	    add { OOQuote.OnConnect += value; }
	    remove { OOQuote.OnConnect -= value; }
	}

	public event _ISKOOQuoteLibEvents_OnProductsEventHandler OnOOProducts
	{
	    add { OOQuote.OnProducts += value; }
	    remove { OOQuote.OnProducts -= value; }
	}

	public event _ISKOOQuoteLibEvents_OnNotifyQuoteLONGEventHandler OnOONotifyQuoteLONG
	{
	    add { OOQuote.OnNotifyQuoteLONG += value; }
	    remove { OOQuote.OnNotifyQuoteLONG -= value; }
	}

	public event _ISKOOQuoteLibEvents_OnNotifyTicksLONGEventHandler OnOONotifyTicksLONG
	{
	    add { OOQuote.OnNotifyTicksLONG += value; }
	    remove { OOQuote.OnNotifyTicksLONG -= value; }
	}

	public event _ISKOOQuoteLibEvents_OnNotifyBest5LONGEventHandler OnOONotifyBest5LONG
	{
	    add { OOQuote.OnNotifyBest5LONG += value; }
	    remove { OOQuote.OnNotifyBest5LONG -= value; }
	}

	public event _ISKOOQuoteLibEvents_OnNotifyBest10LONGEventHandler OnOONotifyBest10LONG
	{
	    add { OOQuote.OnNotifyBest10LONG += value; }
	    remove { OOQuote.OnNotifyBest10LONG -= value; }
	}
	#endregion
	#endregion

	#region Center 方法
	public string SKCenterLib_GetSKAPIVersionAndBit(string id) => Center.SKCenterLib_GetSKAPIVersionAndBit(id);
	public int SKCenterLib_SetAuthority(int authorityFlag) => Center.SKCenterLib_SetAuthority(authorityFlag);
	public string SKCenterLib_GetReturnCodeMessage(int code) => Center.SKCenterLib_GetReturnCodeMessage(code);
	public int SKCenterLib_RequestAgreement(string userId) => Center.SKCenterLib_RequestAgreement(userId);
	public int SKCenterLib_SetLogPath(string path) => Center.SKCenterLib_SetLogPath(path);
	public string SKCenterLib_GetLastLogInfo() => Center.SKCenterLib_GetLastLogInfo();
	public int SKCenterLib_GenerateKeyCert(string logInId, string custCertId) => Center.SKCenterLib_GenerateKeyCert(logInId, custCertId);
	public int SKCenterLib_Login(string userId, string password) => Center.SKCenterLib_Login(userId, password);
	#endregion

	#region Order / Reply / Quote 方法
	public int SKOrderLib_LoadOSCommodity() => Order.SKOrderLib_LoadOSCommodity();
	public int SKOrderLib_LoadOOCommodity() => Order.SKOrderLib_LoadOOCommodity();
	public int UnlockOrder(int marketType) => Order.UnlockOrder(marketType);
	public int SetMaxQty(int marketType, int maxQty) => Order.SetMaxQty(marketType, maxQty);
	public int SetMaxCount(int marketType, int maxCount) => Order.SetMaxCount(marketType, maxCount);
	public int SKOrderLib_GetLoginType(string userId) => Order.SKOrderLib_GetLoginType(userId);
	public int SKOrderLib_GetSpeedyType(string userId) => Order.SKOrderLib_GetSpeedyType(userId);
	public int SKOrderLib_TelnetTest() => Order.SKOrderLib_TelnetTest();
	public int SKReplyLib_ConnectByID(string userId) => Reply.SKReplyLib_ConnectByID(userId);
	public int SKReplyLib_IsConnectedByID(string userId) => Reply.SKReplyLib_IsConnectedByID(userId);
	public int SKReplyLib_SolaceCloseByID(string userId) => Reply.SKReplyLib_SolaceCloseByID(userId);
	public int SKQuoteLib_EnterMonitorLONG() => Quote.SKQuoteLib_EnterMonitorLONG();
	public int SKQuoteLib_LeaveMonitor() => Quote.SKQuoteLib_LeaveMonitor();
	public int SKQuoteLib_IsConnected() => Quote.SKQuoteLib_IsConnected();
	public int SKQuoteLib_GetQuoteStatus(ref int connectionCount, ref bool isOutLimit) => Quote.SKQuoteLib_GetQuoteStatus(ref connectionCount, ref isOutLimit);
	public int SKQuoteLib_RequestStocks(short pageNo, string stockNos) => Quote.SKQuoteLib_RequestStocks(pageNo, stockNos);
	public int SKQuoteLib_RequestTicks(short pageNo, string stockNos) => Quote.SKQuoteLib_RequestTicks(pageNo, stockNos);
	public int SKQuoteLib_Gamma(double s, double k, double r, double t, double sigma, out double gamma) => Quote.SKQuoteLib_Gamma(s, k, r, t, sigma, out gamma);
	public int SKQuoteLib_Vega(double s, double k, double r, double t, double sigma, out double vega) => Quote.SKQuoteLib_Vega(s, k, r, t, sigma, out vega);
	public int SKQuoteLib_Delta(short callPut, double s, double k, double r, double t, double sigma, out double delta) => Quote.SKQuoteLib_Delta(callPut, s, k, r, t, sigma, out delta);
	public int SKQuoteLib_Theta(short callPut, double s, double k, double r, double t, double sigma, out double theta) => Quote.SKQuoteLib_Theta(callPut, s, k, r, t, sigma, out theta);
	public int SKQuoteLib_Rho(short callPut, double s, double k, double r, double t, double sigma, out double rho) => Quote.SKQuoteLib_Rho(callPut, s, k, r, t, sigma, out rho);
	public int SKQuoteLib_RequestKLineAMByDate(string stockNo, short kLineType, short outType, short tradeSession, string startDate, string endDate, short minuteNumber) => Quote.SKQuoteLib_RequestKLineAMByDate(stockNo, kLineType, outType, tradeSession, startDate, endDate, minuteNumber);
	public int SKQuoteLib_RequestStockList(short marketNo) => Quote.SKQuoteLib_RequestStockList(marketNo);
	public int SKQuoteLib_CancelRequestTicks(string stockNos) => Quote.SKQuoteLib_CancelRequestTicks(stockNos);
	public int SKQuoteLib_RequestTicksWithMarketNo(short pageNo, short marketNo, string stockNos) => Quote.SKQuoteLib_RequestTicksWithMarketNo(pageNo, marketNo, stockNos);
	public int SKQuoteLib_CancelRequestStocks(string stockNos) => Quote.SKQuoteLib_CancelRequestStocks(stockNos);
	public int SKQuoteLib_RequestStocksWithMarketNo(short pageNo, short marketNo, string stockNos) => Quote.SKQuoteLib_RequestStocksWithMarketNo(pageNo, marketNo, stockNos);
	public int SKQuoteLib_RequestFutureTradeInfo(short pageNo, string stockNo) => Quote.SKQuoteLib_RequestFutureTradeInfo(pageNo, stockNo);
	public int SKQuoteLib_GetStrikePrices() => Quote.SKQuoteLib_GetStrikePrices();
	public int SKQuoteLib_GetStockByNoLONG(string stockNo, ref SKSTOCKLONG stock) => Quote.SKQuoteLib_GetStockByNoLONG(stockNo, ref stock);
	public int SKQuoteLib_RequestServerTime() => Quote.SKQuoteLib_RequestServerTime();
	public int SKQuoteLib_GetMarketBuySellUpDown() => Quote.SKQuoteLib_GetMarketBuySellUpDown();
	public int SKQuoteLib_RequestMACD(short pageNo, string stockNo) => Quote.SKQuoteLib_RequestMACD(pageNo, stockNo);
	public int SKQuoteLib_RequestBoolTunel(short pageNo, string stockNo) => Quote.SKQuoteLib_RequestBoolTunel(pageNo, stockNo);
	public int SKQuoteLib_GetMACDLONG(short marketNo, int stockIndex, ref SKMACD macd) => Quote.SKQuoteLib_GetMACDLONG(marketNo, stockIndex, ref macd);
	public int SKQuoteLib_GetBoolTunelLONG(short marketNo, int stockIndex, ref SKBoolTunel boolTunel) => Quote.SKQuoteLib_GetBoolTunelLONG(marketNo, stockIndex, ref boolTunel);
	public int SKQuoteLib_GetStockByIndexLONG(short marketNo, int stockIndex, ref SKSTOCKLONG stock) => Quote.SKQuoteLib_GetStockByIndexLONG(marketNo, stockIndex, ref stock);
	public int SKQuoteLib_GetMarketPriceTS() => Quote.SKQuoteLib_GetMarketPriceTS();
	public int SKOOQuoteLib_LeaveMonitor() => OOQuote.SKOOQuoteLib_LeaveMonitor();
	public int SKOOQuoteLib_IsConnected() => OOQuote.SKOOQuoteLib_IsConnected();
	public int SKOOQuoteLib_RequestProducts() => OOQuote.SKOOQuoteLib_RequestProducts();
	public int SKOOQuoteLib_GetStockByIndexLONG(int index, ref SKFOREIGNLONG stock) => OOQuote.SKOOQuoteLib_GetStockByIndexLONG(index, ref stock);
	public int SKOOQuoteLib_GetTickLONG(int index, int ptr, ref SKFOREIGNTICK tick) => OOQuote.SKOOQuoteLib_GetTickLONG(index, ptr, ref tick);
	public int SKOOQuoteLib_GetBest5LONG(int stockIndex, ref SKBEST5 best5) => OOQuote.SKOOQuoteLib_GetBest5LONG(stockIndex, ref best5);
	public int SKOOQuoteLib_GetStockByNoLONG(string stockNo, ref SKFOREIGNLONG stock) => OOQuote.SKOOQuoteLib_GetStockByNoLONG(stockNo, ref stock);
	public int SKOOQuoteLib_RequestStocks(short pageNo, string stockNos) => OOQuote.SKOOQuoteLib_RequestStocks(pageNo, stockNos);
	public int SKOOQuoteLib_RequestTicks(short pageNo, string stockNos) => OOQuote.SKOOQuoteLib_RequestTicks(pageNo, stockNos);
	public int SKOOQuoteLib_RequestMarketDepth(short pageNo, string stockNos) => OOQuote.SKOOQuoteLib_RequestMarketDepth(pageNo, stockNos);
	public int SKOOQuoteLib_EnterMonitorLONG() => OOQuote.SKOOQuoteLib_EnterMonitorLONG();
	public int WithDraw(string logInId, string fullAccountOut, int typeOut, string fullAccountIn, int typeIn, int currency, string dollars, string password, out string message) => Order.WithDraw(logInId, fullAccountOut, typeOut, fullAccountIn, typeIn, currency, dollars, password, out message);
	public int AssembleOptions(string logInId, bool asyncOrder, ref FUTUREORDER order, out string message) => Order.AssembleOptions(logInId, asyncOrder, ref order, out message);
	public int DisassembleOptions(string logInId, bool asyncOrder, ref FUTUREORDER order, out string message) => Order.DisassembleOptions(logInId, asyncOrder, ref order, out message);
	public int CoverAllProduct(string logInId, bool asyncOrder, ref FUTUREORDER order, out string message) => Order.CoverAllProduct(logInId, asyncOrder, ref order, out message);
	public int SendTFOffset(string logInId, bool asyncOrder, string account, int commodity, string yearMonth, int buySell, int qty, out string message) => Order.SendTFOffset(logInId, asyncOrder, account, commodity, yearMonth, buySell, qty, out message);
	public string GetOrderReport(string userId, string account, int format) => Order.GetOrderReport(userId, account, format);
	public string GetFulfillReport(string userId, string account, int format) => Order.GetFulfillReport(userId, account, format);
	public int SKOrderLib_InitialProxyByID(string userId) => Order.SKOrderLib_InitialProxyByID(userId);
	public int ProxyDisconnectByID(string userId) => Order.ProxyDisconnectByID(userId);
	public int ProxyReconnectByID(string userId) => Order.ProxyReconnectByID(userId);
	public int SKOrderLib_LogUpload() => Order.SKOrderLib_LogUpload();
	public int SKOrderLib_Initialize() => Order.SKOrderLib_Initialize();
	public int GetUserAccount() => Order.GetUserAccount();
	public int ReadCertByID(string userId) => Order.ReadCertByID(userId);
	public int SKOrderLib_PingandTracertTest() => Order.SKOrderLib_PingandTracertTest();
	public int AddSGXAPIOrderSocket(string userId, string account) => Order.AddSGXAPIOrderSocket(userId, account);
	public int SendTFOffsetNew(string logInId, bool asyncOrder, string account, int commodity, string yearMonth, int buySell, int qty, int qty2, int qty3, out string message) => Order.SendTFOffsetNew(logInId, asyncOrder, account, commodity, yearMonth, buySell, qty, qty2, qty3, out message);
	#endregion
    }
}
