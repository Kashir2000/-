namespace EpaperWork.DataObject
{
  public class EnumData
  {
    /// <summary>
    /// 發送狀態列舉型別
    /// </summary>
    public enum MailSendStatus
    {
      未歸類 = -1,
      未發送 = 0,
      發送成功 = 1,
      發送失敗 = 2,
      格式錯誤 = 3,
      取消發送 = 9
    }
  }
}
