using Airiti.Common;
using Dapper;
using EpaperWork.DataObject;
using EpaperWork.Service;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;

namespace EpaperWork.DataAccess
{
  /// <summary>
  /// 資料庫相關方法
  /// </summary>
  class DA_Epaper
  {
    /// <summary>
    /// 撈取寄送列表資料庫連接字串
    /// </summary>
    private readonly string ConnectionString = ConfigurationManager.ConnectionStrings["iCheers"].ConnectionString;

    /// <summary>
    /// 寫入Log資料庫名稱
    /// </summary>
    private readonly string Database = ConfigurationManager.AppSettings["Schema_Epaper"] ?? "HistoryLog".ToString().Trim();

    /// <summary>
    /// 電子報發送程式主機數量
    /// </summary>
    private readonly string EpsAPCount = ConfigurationManager.AppSettings["EpsAPCount"] ?? "1";

    /// <summary>
    /// 電子報發送程式主機序號(從0開始)
    /// </summary>
    private readonly string EpsAPNum = ConfigurationManager.AppSettings["EpsAPNum"] ?? "0";

    /// <summary>
    /// 偵錯模式
    /// </summary>
    private readonly bool DebugMode = (ConfigurationManager.AppSettings["Debug"] ?? string.Empty).ToLower() == "true";

    /// <summary>
    /// 取得電子報發信資料
    /// </summary>
    /// <returns></returns>
    public (List<DO_EpaperSend>, List<DO_EpaperSendLog>) GetEpaperSendInfoList()
    {
      #region 預設宣告
      List<DO_EpaperSend> Return1 = new List<DO_EpaperSend>();
      List<DO_EpaperSendLog> Return2 = new List<DO_EpaperSendLog>();
      #endregion
      try
      {
        #region QueryString
        string QueryString = $@"SELECT ES.EpsID, 
                                       ES.EpsSubject, 
                                       ES.EpsBody, 
                                       ES.EpsDispName, 
                                       ES.EpsPMMail, 
                                       ES.EpsActCode,
                                       ES.EpsStatus
                                FROM dbo.EpaperSend AS ES
                                WHERE ES.EpsStatus IN('1', '2') --0::暫存; 1::啟用;2::發送中;3::已發送;9::取消
                                AND ES.EpsEstTime <= GETDATE();
                                SELECT ES.EpsID, 
                                       ESL.EpsMail
                                FROM {Database}.dbo.EpaperSendLog AS ESL
                                     INNER JOIN dbo.EpaperSend AS ES ON ES.EpsID = ESL.EpsID
                                                                        AND ES.EpsStatus IN('1', '2') --0::暫存; 1::啟用;2::發送中;3::已發送;9::取消
                                                                        AND ES.EpsEstTime <= GETDATE()
                                WHERE ESL.EpsSendStatus IN('0', '2', '3') --0::未發送;1::發送成功;2::發送失敗;3::格式錯誤;9::取消發送;
                                      AND EpsOrder % {EpsAPCount} IN({EpsAPNum}) --篩選發送主機處理清單
                                ORDER BY ESL.EpsID, 
                                         ESL.EpsOrder;";
        #endregion
        using (SqlConnection Connection = new SqlConnection(ConnectionString))
        {
          using (var Multi = Connection.QueryMultiple(QueryString))
          {
            Return1 = Multi.Read<DO_EpaperSend>().ToList();
            Return2 = Multi.Read<DO_EpaperSendLog>().ToList();
          }
        }
        Debug($@"GetEpaperSendInfoList執行成功！
                 {QueryString}");
        return (Return1, Return2);
      }
      catch
      {
        throw;
      }
    }

    /// <summary>
    /// 更新電子報郵件發送狀態
    /// </summary>
    /// <param name="MailDetail">電子報內容</param>
    /// <param name="SendMail">電子報郵件發送記錄</param>
    public void UpdateSendMailLog(DO_EpaperSend MailDetail, DO_SendMail SendMail)
    {
      try
      {
        using (TransactionScope Scope = new TransactionScope())
        {
          #region QueryString
          string QueryString = @"DECLARE @SendTime DATETIME= CASE
                                                                 WHEN ISDATE(@EpsSendTime) = 1
                                                                 THEN CAST(@EpsSendTime AS DATETIME)
                                                                 ELSE GETDATE()
                                                             END;
                                 --更新HistoryLog.dbo.EpaperSendLog
                                 UPDATE HistoryLog.dbo.EpaperSendLog
                                   SET 
                                       EpsSendTime = @SendTime, 
                                       EpsSendCount = EpsSendCount + @EpsSendCount, 
                                       EpsSendStatus = @EpsSendStatus
                                 WHERE EpsID = @EpsID
                                       AND EpsMail = @EpsMail;
                                 IF @@ROWCOUNT = 0
                                     BEGIN
                                         INSERT INTO HistoryLog.dbo.EpaperSendLog
                                         (EpsID, 
                                          EpsMail, 
                                          EpsOrder, 
                                          EpsSendTime, 
                                          EpsSendCount, 
                                          EpsSendStatus
                                         )
                                         VALUES
                                         (@EpsID, 
                                          @EpsMail, 
                                          99999, 
                                          @SendTime, 
                                          @EpsSendCount, 
                                          @EpsSendStatus
                                         );
                                 END;                               
                                 --更新HistoryLog.dbo.EpaperMailLog
                                 IF @EpsSendStatus = '1' --1::發送成功
                                     BEGIN
                                         UPDATE HistoryLog.dbo.EpaperMailLog
                                           SET 
                                               EmSendCnt = EmSendCnt + 1, 
                                               EmSendTime = @SendTime, 
                                               EmSendSuccessCnt = EmSendSuccessCnt + 1, 
                                               EmSendSuccessTime = @SendTime
                                         WHERE EmAddr = @EpsMail;
                                         IF @@ROWCOUNT = 0
                                             BEGIN
                                                 INSERT INTO HistoryLog.dbo.EpaperMailLog
                                                 (EmAddr, 
                                                  EmSendCnt, 
                                                  EmSendTime, 
                                                  EmSendSuccessCnt, 
                                                  EmSendSuccessTime
                                                 )
                                                 VALUES
                                                 (@EpsMail, 
                                                  1, 
                                                  @SendTime, 
                                                  1, 
                                                  @SendTime
                                                 );
                                         END;
                                 END;
                                     ELSE
                                     IF @EpsSendStatus IN('2', '3') --2::發送失敗;3::格式錯誤;
                                         BEGIN
                                             UPDATE HistoryLog.dbo.EpaperMailLog
                                               SET 
                                                   EmSendCnt = EmSendCnt + 1, 
                                                   EmSendTime = @SendTime, 
                                                   EmSendFailureCnt = EmSendFailureCnt + 1, 
                                                   EmSendFailureTime = @SendTime
                                             WHERE EmAddr = @EpsMail;
                                             IF @@ROWCOUNT = 0
                                                 BEGIN
                                                     INSERT INTO HistoryLog.dbo.EpaperMailLog
                                                     (EmAddr, 
                                                      EmSendCnt, 
                                                      EmSendTime, 
                                                      EmSendFailureCnt, 
                                                      EmSendFailureTime
                                                     )
                                                     VALUES
                                                     (@EpsMail, 
                                                      1, 
                                                      @SendTime, 
                                                      1, 
                                                      @SendTime
                                                     );
                                             END;
                                     END;";
          #endregion
          using (SqlConnection Connection = new SqlConnection(ConnectionString))
          {
            DynamicParameters Parameters = new DynamicParameters();
            Parameters.Add("@EpsID", SendMail.EpsID, DbType.Int64, ParameterDirection.Input);
            Parameters.Add("@EpsMail", SendMail.EpsMail, DbType.String, ParameterDirection.Input);
            Parameters.Add("@EpsSendTime", SendMail.SendTime, DbType.DateTime, ParameterDirection.Input);
            Parameters.Add("@EpsSendCount", SendMail.SendCount, DbType.Int64, ParameterDirection.Input);
            Parameters.Add("@EpsSendStatus", SendMail.SendStatus, DbType.String, ParameterDirection.Input);
            Connection.Execute(QueryString, Parameters);
          }
          UpdateEpaperStatus(MailDetail);
          Scope.Complete();
          Debug($@"UpdateSendMailLog執行成功！
                   {QueryString}
                   @EpsID = {SendMail.EpsID}
                   @EpsMail = {SendMail.EpsMail}
                   @EpsSendTime = {SendMail.SendTime}
                   @EpsSendCount = {SendMail.SendCount}
                   @EpsSendStatus = {SendMail.SendStatus}");
        }
      }
      catch
      {
        throw;
      }
    }

    /// <summary>
    /// 更新電子報目前狀態
    /// </summary>
    public void UpdateEpaperStatus(DO_EpaperSend MailDetail)
    {
      #region QueryString
      string QueryString = @"SELECT RTRIM(EpsStatus) AS EpsStatus
                             FROM   EpaperSend
                             WHERE  EpsID = @EpsID;";
      #endregion
      using (SqlConnection Connection = new SqlConnection(ConnectionString))
      {
        DynamicParameters Parameters = new DynamicParameters();
        Parameters.Add("@EpsID", MailDetail.EpsID, DbType.Int64, ParameterDirection.Input);
        MailDetail.EpsStatus = Connection.QueryFirst<char>(QueryString, Parameters);
      }
    }

    /// <summary>
    /// 寄送重啟通知信
    /// </summary>
    /// <param name="errorMessage">失敗訊息</param>
    public void SendRebootLetter(int ErrorCount)
    {
      try
      {
        var ReturnCount = 0;
        var MailTitle = "電子報發送程式(iCheersEpaperSendService)程式重啟";
        var ResultBody = $@"<div>電子報發送程式(iCheersEpaperSendService)程式於【{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}】重新啟用。</div>
                            <br/>                            
                            <div>錯誤訊息：內部錯誤已達{ErrorCount}次！</div>";
        using (SqlConnection Connection = new SqlConnection(ConnectionString))
        {
          //準備參數
          DynamicParameters parameters = new DynamicParameters();
          parameters.Add("@MailTitle", MailTitle, DbType.String, ParameterDirection.Input);
          parameters.Add("@ResultBody", ResultBody, DbType.String, ParameterDirection.Input);
          parameters.Add("@ReturnCount", dbType: DbType.Int32, direction: ParameterDirection.ReturnValue);
          //公版共用寄信SP
          Connection.Execute("SP_LETTER_INVOICE_EXPORT_ERROR", parameters, commandType: CommandType.StoredProcedure);
          ReturnCount = parameters.Get<int>("@ReturnCount");
        }
        Debug($@"SendRebootLetter執行成功！
                 @MailTitle = {MailTitle}
                 @ResultBody = {ResultBody}");
      }
      catch
      {
        throw;
      }
    }

    /// <summary>
    /// 寫入Log
    /// </summary>
    /// <param name="text">寫入訊息</param>
    private void Debug(string text)
    {
      if (DebugMode)
      {
        System.Diagnostics.Debug.WriteLine(text);
        Action WriteLog = () =>
        {
          LogFile objLogFile = new LogFile("Epaper", "DA_Epaper");
          objLogFile.Log(text).LogFlush();
        };
        QueueService.QueueList.Enqueue(WriteLog);
      }
    }
  }
}
