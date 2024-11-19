using Airiti.Common;
using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace EpaperWork.Service
{
  public class QueueService
  {
    /// <summary>
    /// DB寫入線程
    /// </summary>
    public static ConcurrentQueue<Action> QueueList = new ConcurrentQueue<Action>();

    /// <summary>
    /// 執行狀態
    /// </summary>
    private volatile bool Processing = false;

    /// <summary>
    /// Log檔案
    /// </summary>
    private LogFile objLogFile = new LogFile("Epaper", "QueueService");

    /// <summary>
    /// 開始執行佇列工作
    /// </summary>
    public void StartQueueWork()
    {
      Processing = true;
      Task.Run(() => ProcessQueue());
    }

    /// <summary>
    /// 隊列工作處理
    /// </summary>
    private async Task ProcessQueue()
    {
      while (Processing)
      {
        if (QueueList.TryDequeue(out Action FuncWork))
        {
          try
          {
            FuncWork();
          }
          catch (Exception ex)
          {
            ErrorLog(ex);
          }
        }
        else
        {
          // 如果隊列為空，短暫休眠以減少 CPU 使用率
          await Task.Delay(1000);
        }
      }
    }

    /// <summary>
    /// 結束執行佇列工作
    /// </summary>
    public void StopQueueWorker()
    {
      Processing = false;
    }

    /// <summary>
    /// 寫入錯誤Log
    /// </summary>
    /// <param name="ex">錯誤物件</param>
    private void ErrorLog(Exception ex)
    {
      Action WriteLog = () =>
      {
        objLogFile.Log(ex).LogFlush();
      };
      QueueService.QueueList.Enqueue(WriteLog);
    }
  }
}
