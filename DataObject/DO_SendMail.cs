using System;
using static EpaperWork.DataObject.EnumData;

namespace EpaperWork.DataObject
{
  /// <summary>
  /// 電子報信件物件
  /// </summary>
  public class DO_SendMail
  {
    public int EpsID;
    public string EpsMail;
    public int SendCount = 0;
    public DateTime SendTime;
    public MailSendStatus SendStatus = MailSendStatus.未發送;

    /// <summary>
    /// 電子報物件
    /// </summary>
    /// <param name="EpsID">電子報ID</param>
    /// <param name="EpsMail">EMAIL帳號</param>
    public DO_SendMail(int EpsID, string EpsMail)
    {
      this.EpsID = EpsID;
      this.EpsMail = EpsMail;
    }
  }
}
