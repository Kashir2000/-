using EpaperSend;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static EpaperWork.DataObject.EnumData;

namespace EpaperWork.DataObject
{
  /// <summary>
  /// 電子報發送清單記錄[HistoryLog]
  /// </summary>
  public class DO_EpaperSendLog
  {
    /// <summary>
    /// 電子報發送ID
    /// </summary>
    public int EpsID;

    /// <summary>
    /// 郵件帳號
    /// </summary>
    public string EpsMail;

    /// <summary>
    /// 發送次序
    /// </summary>
    public int EpsOrder = 1;

    /// <summary>
    /// 最後發信時間
    /// </summary>
    public DateTime? EpsSendTime = null;

    /// <summary>
    /// 發信次數
    /// </summary>
    public int EpsSendCount = 0;

    /// <summary>
    /// 發信成功否
    /// </summary>
    public char EpsSendStatus = '0';

    public  MailSendStatus MailSendStatus
    {
      get
      {
        switch (EpsSendStatus)
        {
          case '0':
            return MailSendStatus.未發送;
          case '1':
            return MailSendStatus.發送成功;
          case '2':
            return MailSendStatus.發送失敗;
          case '3':
            return MailSendStatus.格式錯誤;
          case '9':
            return MailSendStatus.取消發送;
          default:
            return MailSendStatus.未歸類;
        }
      }
    }
  }
}
