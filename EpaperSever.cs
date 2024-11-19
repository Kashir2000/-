using EpaperSend;
using System.Configuration;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace EpaperWork
{
  partial class EpaperServer : ServiceBase
  {
    /// <summary>
    /// 服務使用的電子報程式宣告
    /// </summary>
    private Epaper SendWork;

    /// <summary>
    /// 執行此程式的的服務名稱
    /// </summary>
    private readonly string MyServiceName = ConfigurationManager.AppSettings["ServiceName"] ?? "iCheersEpaperSendService";


    public EpaperServer()
    {
      this.ServiceName = MyServiceName;
      InitializeComponent();
    }

    protected override void OnStart(string[] args)
    {
      Task.Run(() =>
      {
        SendWork = new Epaper();
        SendWork.StartWork();
      });
    }

    protected override void OnStop()
    {
      SendWork?.StopWork();
    }
  }
}
