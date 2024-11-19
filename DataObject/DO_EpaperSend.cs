using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EpaperWork.DataObject
{
  /// <summary>
  /// 電子報發送記錄
  /// </summary>
  public class DO_EpaperSend
  {
    /// <summary>
    /// 電子報發送ID
    /// </summary>
    public int EpsID;

    /// <summary>
    /// 電子報主旨
    /// </summary>
    public string EpsSubject = string.Empty;

    /// <summary>
    /// 電子報內容
    /// </summary>
    public string EpsBody = string.Empty;

    /// <summary>
    /// 電子報類型名稱
    /// </summary>
    public string EpsDispName = string.Empty;

    /// <summary>
    /// 發送清單上傳檔名
    /// </summary>
    public string EpsListFile = string.Empty;

    /// <summary>
    /// 發送清單上傳路徑
    /// </summary>
    public string EpsListPath = string.Empty;

    /// <summary>
    /// 發送數量
    /// </summary>
    public int EpsSendCount = 0;

    /// <summary>
    /// 預計發送時間
    /// </summary>
    public DateTime? EpsEstTime = null;

    /// <summary>
    /// 實際發送時間
    /// </summary>
    public DateTime? EpsSendTime = null;

    /// <summary>
    /// EpsCompTime
    /// </summary>
    public DateTime? EpsCompTime = null;

    /// <summary>
    /// 電子報說明
    /// </summary>
    public string EpsDesc = string.Empty;

    /// <summary>
    /// 負責人員
    /// </summary>
    public string EpsPM = string.Empty;

    /// <summary>
    /// 發送完成通知信相
    /// </summary>
    public string EpsPMMail = string.Empty;

    /// <summary>
    /// 發送狀態
    /// </summary>
    public char EpsStatus = '0';

    /// <summary>
    /// 關聯活動ID
    /// </summary>
    public int? AdvID = null;

    /// <summary>
    /// 電子報活動代碼
    /// </summary>
    public string EpsActCode = string.Empty;

    /// <summary>
    /// 發送成功數量
    /// </summary>
    public int EpsSuccessCount = 0;

    /// <summary>
    /// 發送失敗數量
    /// </summary>
    public int EpsFailureCount = 0;

    /// <summary>
    /// 最後修改人
    /// </summary>
    public string Userstamp = string.Empty;

    /// <summary>
    /// 最後修改日期
    /// </summary>
    public string Datestamp = string.Empty;
  }
}
