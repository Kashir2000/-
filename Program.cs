using EpaperSend;
using System;
using System.Configuration;
using System.ServiceProcess;

namespace EpaperWork
{
  public class Program
  {
    static void Main(string[] args)
    {
      Airiti.Configuration.Debug = (ConfigurationManager.AppSettings["Debug"] ?? string.Empty).ToLower() == "true";
      Airiti.Configuration.DebugPath = ConfigurationManager.AppSettings["DebugPath"] ?? string.Empty;
      Airiti.Configuration.LogReNew = false; //檔名後面不要有小時區分

      if (Environment.UserInteractive)
      {
        // 開始電子報寄信作業
        var SendWork = new Epaper();
        SendWork.StartWork();
        Console.WriteLine("電子報寄信作業執行中，結束請按任意鍵...");
        Console.ReadKey();
      }
      else
      {
        // Windows服務運行
        ServiceBase[] ServicesToRun;
        ServicesToRun = new ServiceBase[]
        {
            new EpaperServer()
        };
        ServiceBase.Run(ServicesToRun);
      }
    }
  }
}
