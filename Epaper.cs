using Airiti.Common;
using Airiti.Extensions;
using EpaperWork.DataAccess;
using EpaperWork.DataObject;
using EpaperWork.Service;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using static EpaperWork.DataObject.EnumData;

namespace EpaperSend
{
  public class Epaper
  {
    #region 系統參數
    private readonly string ServiceName = ConfigurationManager.AppSettings["ServiceName"] ?? "iCheersEpaperSendService";
    private readonly string RestartPath = ConfigurationManager.AppSettings["RestartPath"] ?? @"D:\RestartService\RestartService.exe";
    private readonly int CheckTime = int.TryParse(ConfigurationManager.AppSettings["CheckTime"], out int Value) && Value > 0 ? Value * 1000 : 300000;
    private readonly int CheckErrorTime = int.TryParse(ConfigurationManager.AppSettings["CheckErrorTime"], out int Value) && Value > 0 ? Value * 1000 : 60000;
    private readonly int TaskCount = int.TryParse(ConfigurationManager.AppSettings["TaskCount"], out int Value) ? Math.Max(1, Math.Min(Value, Environment.ProcessorCount)) : Environment.ProcessorCount;
    private readonly int MaxErrorCount = int.TryParse(ConfigurationManager.AppSettings["MaxErrorCount"], out int Value) && Value > 0 ? Value : 0;
    private readonly int DelayTime = int.TryParse(ConfigurationManager.AppSettings["TaskCount"], out int Value) && Value > 0 ? Value * 1000 : 0;
    private readonly int RetryCount = int.TryParse(ConfigurationManager.AppSettings["RetryCount"], out int Value) && Value > 0 ? Value : 3;
    private readonly string Mailhost = ConfigurationManager.AppSettings["Mailhost"] ?? string.Empty;
    private readonly int MailPort = int.TryParse(ConfigurationManager.AppSettings["MailPort"], out int Value) && Value > 0 ? Value : 587;
    private readonly bool MailSSL = (ConfigurationManager.AppSettings["MailSSL"] ?? string.Empty).ToLower() == "true";
    private readonly int Timeout = int.TryParse(ConfigurationManager.AppSettings["Timeout"], out int Value) && Value > 0 ? Value * 1000 : 100000;
    private readonly string MailAccount = ConfigurationManager.AppSettings["MailAccount"] ?? string.Empty;
    private readonly string MailPassword = ConfigurationManager.AppSettings["MailPassword"] ?? string.Empty;
    private readonly string MailFrom = ConfigurationManager.AppSettings["MailFrom"] ?? string.Empty;
    private readonly string MailReply = ConfigurationManager.AppSettings["MailReply"] ?? string.Empty;
    private readonly bool DebugMode = (ConfigurationManager.AppSettings["Debug"] ?? string.Empty).ToLower() == "true";
    #endregion

    /// <summary>
    /// 計時器(自我循環)
    /// </summary>
    private System.Timers.Timer MTimer;

    /// <summary>
    /// 計時器(錯誤檢查)
    /// </summary>
    private System.Timers.Timer MTimer2;

    /// <summary>
    /// 確保CheckEaper()唯一旗標
    /// </summary>
    private bool UniqueEaperWork = false;

    /// <summary>
    /// 線程服務
    /// </summary>
    private QueueService QueueService;

    /// <summary>
    /// 錯誤事件列表
    /// </summary>
    private ConcurrentBag<Exception> Exceptions;

    /// <summary>
    /// Log檔案
    /// </summary>
    private LogFile objLogFile = new LogFile("Epaper", "SendWork");


    public Epaper()
    {
      #region 初始化工作
      QueueService = new QueueService();
      Exceptions = new ConcurrentBag<Exception>();
      #region 電子報寄信作業計時
      MTimer = new System.Timers.Timer(CheckTime);
      MTimer.AutoReset = true;
      MTimer.Elapsed += new ElapsedEventHandler(async (object sender, ElapsedEventArgs e) => { await CheckEaper(); });
      #endregion
      #region 錯誤列表檢查作業計時
      if (MaxErrorCount > 0)
      {
        MTimer2 = new System.Timers.Timer(CheckErrorTime);
        MTimer2.AutoReset = true;
        MTimer2.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs e) =>
        {
          if (Exceptions.Count() >= MaxErrorCount)
          {
            MTimer.Stop();
            MTimer2.Stop();
            Debug($"攔截錯誤已達{MaxErrorCount}次，稍後程式將自動重啟！");
            if (Environment.UserInteractive)
            {
              ProgramReStart();
            }
            else
            {
              ServiceReStart();
            }
          }
        });
      }
      #endregion
      #endregion
    }

    /// <summary>
    /// 檢查發送電子報
    /// </summary>
    private async Task CheckEaper()
    {
      // 防止重入
      if (UniqueEaperWork) return;
      UniqueEaperWork = true;
      try
      {
        MTimer.Stop();
        Debug($"開始電子報發信程序...");
        Stopwatch Watch = new Stopwatch();
        Watch.Start();
        List<Task> WorkList = new List<Task>();
        var Result = new DA_Epaper().GetEpaperSendInfoList();
        if (Result.Item1.Count > 0)
        {
          Debug($"共{Result.Item1.Count}筆，待發送電子報！");
          foreach (var MailDetail in Result.Item1)
          {
            List<DO_SendMail> MailList = Result.Item2.Where(DataRow => DataRow.EpsID == MailDetail.EpsID).Select(Row => new DO_SendMail(Row.EpsID, Row.EpsMail)).ToList();
            if (MailList.Count > 0)
            {
              WorkList.Add(Task.Run(async () => await AsyncSend(MailDetail, MailList)));
            }
            else
            {
              Debug($"電子報編號【{MailDetail.EpsID}】收信者清單數量為0，不需寄送。");
            }
          }
        }
        else
        {
          Debug($"無待發電子報。");
        }
        await Task.WhenAll(WorkList);
        Watch.Stop();
        Debug($"結束電子報發信程序，總共花了{Watch.ElapsedMilliseconds / 1000.0}秒。");
        MTimer.Start();
      }
      catch (Exception ex)
      {
        Exceptions.Add(ex);
        ErrorLog(ex);
        MTimer.Start();
      }
      finally
      {
        // 重置重入旗標
        UniqueEaperWork = false;
      }
    }

    /// <summary>
    /// 非同步寄信程序
    /// </summary>
    /// <returns></returns>
    private async Task AsyncSend(DO_EpaperSend MailDetail, List<DO_SendMail> MailList)
    {
      try
      {
        List<Task> Tasks = new List<Task>();
        List<DO_SendMail> SendList = new List<DO_SendMail>();

        Debug($@"電子報編號【{MailDetail.EpsID}】開始寄送...
                           信件主旨：{MailDetail.EpsSubject}
                           發送數：{MailList.Count}
                           線程數：{TaskCount}
                           重試數：{RetryCount}");

        #region E-mail格式檢查
        foreach (var Email in MailList)
        {
          if (IsValidEmail(Email.EpsMail))
          {
            SendList.Add(Email);
          }
          else
          {
            Email.SendTime = DateTime.Now;
            Email.SendStatus = EnumData.MailSendStatus.格式錯誤;
            Debug($"電子報編號({MailDetail.EpsID})之【{Email.EpsMail}】 非Email格式");
            Action FormatError = () =>
            {
              objLogFile.Log($"電子報編號({MailDetail.EpsID})之【{Email.EpsMail}】 非Email格式").LogFlush();
              new DA_Epaper().UpdateSendMailLog(MailDetail, Email);
            };
            QueueService.QueueList.Enqueue(FormatError);
          }
        }
        #endregion

        #region 寫法一：非同步線程集合        
        // 創建一個 SemaphoreSlim 來限制同時運行的任務數量
        var Semaphore = new SemaphoreSlim(TaskCount);
        #endregion

        #region 寫法二：最大並行線程控制【不適用網路 I/O 】：已廢棄
        //var ParallelSet = new ParallelOptions { MaxDegreeOfParallelism = TaskCount };
        #endregion

        #region 寫法一：非同步線程任務
        foreach (var Mail in SendList)
        {
          // 等待直到 Semaphore 允許運行新的任務
          await Semaphore.WaitAsync();
          if (new List<char> { '1', '2' }.Contains(MailDetail.EpsStatus))
          {
            Tasks.Add(Task.Run(async () =>
            {
              var TaskID = Task.CurrentId;
              try
              {
                Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}中...");
                // 執行發送郵件的操作
                await SendMailKitAsync(MailDetail, Mail);
                Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}完成！");
              }
              catch
              {
                Mail.SendCount++;
                Mail.SendTime = DateTime.Now;
                Mail.SendStatus = MailSendStatus.發送失敗;
                Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}失敗(第{Mail.SendCount}次)！");
              }
              finally
              {
                // 完成後釋放 Semaphore
                Semaphore.Release();
                Debug($"{TaskID}釋放後可進入之執行序數量：{Semaphore.CurrentCount}");
              }
            }));
          }
          else
          {
            Debug($"EpaperID【{MailDetail.EpsID}】狀態不為「1::啟用」「2::發送中」，已中斷發送。");
            break;
          }
        }
        await Task.WhenAll(Tasks);
        #endregion

        #region 寫法二：非同步線程任務【不適用網路 I/O 】：已廢棄
        //Parallel.For(0, SendList.Count, ParallelSet, (i) =>
        //{
        //  var Mail = SendList[i];
        //  Tasks.Add(Task.Run(async () =>
        //  {
        //    var TaskID = Task.CurrentId;
        //    try
        //    {
        //      Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}中...");
        //      // 執行發送郵件的操作
        //      await SendMailKitAsync(MailDetail, Mail);
        //      Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}完成！");
        //    }
        //    catch
        //    {
        //      Mail.SendCount++;
        //      Mail.SendTime = DateTime.Now;
        //      Mail.SendStatus = MailSendStatus.發送失敗;
        //      Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}失敗(第{Mail.SendCount}次)！");
        //    }
        //  }));
        //});
        //await Task.WhenAll(Tasks);
        #endregion

        #region 寫法一：重發未成功清單
        while (new List<char> { '1', '2' }.Contains(MailDetail.EpsStatus) && SendList.Exists(Mail => Mail.SendStatus == MailSendStatus.發送失敗 && Mail.SendCount < RetryCount))
        {
          Tasks.Clear();
          List<DO_SendMail> ReSendList = SendList.Where(Mail => Mail.SendStatus == MailSendStatus.發送失敗 && Mail.SendCount < RetryCount).ToList();
          Debug($"寄送失敗低於{RetryCount}次的含有{ReSendList.Count}封，重試寄信中..");
          foreach (var Mail in ReSendList)
          {
            // 等待直到 Semaphore 允許運行新的任務
            await Semaphore.WaitAsync();
            Tasks.Add(Task.Run(async () =>
            {
              var TaskID = Task.CurrentId;
              try
              {
                Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}中...");
                // 執行發送郵件的操作
                await SendMailKitAsync(MailDetail, Mail);
                Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}完成！");
              }
              catch
              {
                Mail.SendCount++;
                Mail.SendTime = DateTime.Now;
                Mail.SendStatus = MailSendStatus.發送失敗;
                Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}失敗(第{Mail.SendCount}次)！");
              }
              finally
              {
                // 完成後釋放 Semaphore
                Semaphore.Release();
                Debug($"{TaskID}釋放後可進入之執行序數量：{Semaphore.CurrentCount}");
              }
            }));
          }
          await Task.WhenAll(Tasks);
        }
        #endregion

        #region 寫法二：重發未成功清單【不適用網路 I/O 】：已廢棄
        //while (SendList.Exists(Mail => Mail.SendStatus == MailSendStatus.發送失敗 && Mail.SendCount < RetryCount))
        //{
        //  List<DO_SendMail> ReSendList = SendList.Where(Mail => Mail.SendStatus == MailSendStatus.發送失敗 && Mail.SendCount < RetryCount).ToList();
        //  Parallel.For(0, ReSendList.Count, ParallelSet, (i) =>
        //  {
        //    var Mail = ReSendList[i];
        //    Tasks.Add(Task.Run(async () =>
        //    {
        //      var TaskID = Task.CurrentId;
        //      try
        //      {
        //        Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}中...");
        //        // 執行發送郵件的操作
        //        await SendMailKitAsync(MailDetail, Mail);
        //        Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}完成！");
        //      }
        //      catch
        //      {
        //        Mail.SendCount++;
        //        Mail.SendTime = DateTime.Now;
        //        Mail.SendStatus = MailSendStatus.發送失敗;
        //        Debug($"EpaperID【{MailDetail.EpsID}】{TaskID}寄送至 {Mail.EpsMail}失敗(第{Mail.SendCount}次)！");
        //      }
        //    }));
        //  });
        //}
        //await Task.WhenAll(Tasks);
        #endregion

        #region 寫入失敗信件紀錄
        List<DO_SendMail> FailuresList = SendList.Where(Mail => Mail.SendStatus == MailSendStatus.發送失敗 && Mail.SendCount == RetryCount).ToList();
        foreach (var Mail in FailuresList)
        {
          Action SendError = () =>
          {
            objLogFile.Log($"電子報編號({MailDetail.EpsID})之【{Mail.EpsMail}】 發送失敗已達{RetryCount}次！").LogFlush();
            new DA_Epaper().UpdateSendMailLog(MailDetail, Mail);
          };
          QueueService.QueueList.Enqueue(SendError);
        }

        #endregion

        Debug($@"電子報編號【{MailDetail.EpsID}】{(new List<char> { '1', '2' }.Contains(MailDetail.EpsStatus) ? "發送完成" : "中斷發送")}！
                            信件主旨：{MailDetail.EpsSubject}
                            發送清單數：{MailList.Count}
                            未發送數：{MailList.FindAll(r => r.SendStatus == MailSendStatus.未發送).Count}
                            帳號不符數：{MailList.FindAll(r => r.SendStatus == MailSendStatus.格式錯誤).Count}
                            發送失敗數：{MailList.FindAll(r => r.SendStatus == MailSendStatus.發送失敗).Count}
                            發送成功數：{MailList.FindAll(r => r.SendStatus == MailSendStatus.發送成功).Count}");
      }
      catch
      {
        throw;
      }
    }

    /// <summary>
    /// 非同步寄信(SMTP原生方法)
    /// </summary>
    /// <param name="MailDetail">郵件內容</param>
    /// <param name="Mail">收件人內容</param>
    /// <returns></returns>
    private async Task SendEmailAsync(DO_EpaperSend MailDetail, DO_SendMail Mail)
    {
      #region 處理內文標籤參數
      string EpsBody = MailDetail.EpsBody;
      //物件是傳址的，若直接用MailDetail.EpsBody取代參數會變成全部信件都一樣參數
      //[E-mail Address]_[Short Date]_[Hours]:[Minutes]:[Seconds]
      //[電郵地址]_[短式日期顯示]_[小時]:[分鐘]:[秒]
      EpsBody = EpsBody.Replace("[E-mail Address]", Mail.EpsMail).Replace("[電郵地址]", Mail.EpsMail);
      EpsBody = EpsBody.Replace("[Short Date]", DateTime.Now.ToString("yyyy/MM/dd")).Replace("[短式日期顯示]", DateTime.Now.ToString("yyyy/MM/dd"));
      EpsBody = EpsBody.Replace("[Hours]", DateTime.Now.ToString("HH")).Replace("[小時]", DateTime.Now.ToString("HH"));
      EpsBody = EpsBody.Replace("[Minutes]", DateTime.Now.ToString("mm")).Replace("[分鐘]", DateTime.Now.ToString("mm"));
      EpsBody = EpsBody.Replace("[Seconds]", DateTime.Now.ToString("ss")).Replace("[秒]", DateTime.Now.ToString("ss"));
      #endregion      
      try
      {
        using (var SMTP = new System.Net.Mail.SmtpClient(Mailhost, MailPort))
        {
          SMTP.Credentials = new NetworkCredential(MailAccount, MailPassword);
          SMTP.EnableSsl = MailSSL;
          SMTP.Timeout = Timeout;
          ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
          ServicePointManager.ServerCertificateValidationCallback = (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) => { return true; };

          using (MailMessage MailItem = new MailMessage())
          {
            MailItem.From = new MailAddress(MailFrom, MailDetail.EpsDispName);
            MailItem.To.Add(Mail.EpsMail);
            MailItem.ReplyToList.Add(MailReply);
            MailItem.Subject = MailDetail.EpsSubject;
            MailItem.Body = EpsBody;
            MailItem.IsBodyHtml = true;

            await SMTP.SendMailAsync(MailItem);
            #region 寄成功執行DB寫入
            Mail.SendCount++;
            Mail.SendTime = DateTime.Now;
            Mail.SendStatus = MailSendStatus.發送成功;
            Action SendSuccess = () =>
            {
              new DA_Epaper().UpdateSendMailLog(MailDetail, Mail);
            };
            QueueService.QueueList.Enqueue(SendSuccess);
            #endregion
          }
        }
      }
      catch (System.TimeoutException ex)
      {
        MTimer.Start();
        Debug($@"EpaperID【{MailDetail.EpsID}】寄送至 {Mail.EpsMail}回應逾時！
                   錯誤內容：{ex.Message}");
        Exceptions.Add(ex);
        ErrorLog(ex);
        throw;
      }
      catch (Exception ex)
      {
        MTimer.Start();
        Debug($@"EpaperID【{MailDetail.EpsID}】寄送至 {Mail.EpsMail}發生錯誤！
                 錯誤內容：{ex.Message}");
        Exceptions.Add(ex);
        ErrorLog(ex);
        throw;
      }
      finally
      {
        if (DelayTime > 0)
        {
          await Task.Delay(DelayTime);
        }
      }
    }

    /// <summary>
    /// 非同步寄信(MailKit)
    /// </summary>
    /// <param name="MailDetail">郵件內容</param>
    /// <param name="Mail">收件人內容</param>
    /// <returns></returns>
    private async Task SendMailKitAsync(DO_EpaperSend MailDetail, DO_SendMail Mail)
    {
      #region 處理內文標籤參數
      string EpsBody = MailDetail.EpsBody;
      //物件是傳址的，若直接用MailDetail.EpsBody取代參數會變成全部信件都一樣參數
      //[E-mail Address]_[Short Date]_[Hours]:[Minutes]:[Seconds]
      //[電郵地址]_[短式日期顯示]_[小時]:[分鐘]:[秒]
      EpsBody = EpsBody.Replace("[E-mail Address]", Mail.EpsMail).Replace("[電郵地址]", Mail.EpsMail);
      EpsBody = EpsBody.Replace("[Short Date]", DateTime.Now.ToString("yyyy/MM/dd")).Replace("[短式日期顯示]", DateTime.Now.ToString("yyyy/MM/dd"));
      EpsBody = EpsBody.Replace("[Hours]", DateTime.Now.ToString("HH")).Replace("[小時]", DateTime.Now.ToString("HH"));
      EpsBody = EpsBody.Replace("[Minutes]", DateTime.Now.ToString("mm")).Replace("[分鐘]", DateTime.Now.ToString("mm"));
      EpsBody = EpsBody.Replace("[Seconds]", DateTime.Now.ToString("ss")).Replace("[秒]", DateTime.Now.ToString("ss"));
      #endregion            
      using (var Client = new MailKit.Net.Smtp.SmtpClient())
      {
        try
        {
          using (var emailMessage = new MimeMessage())
          {
            emailMessage.From.Add(new MailboxAddress(MailDetail.EpsDispName, MailFrom));
            emailMessage.To.Add(new MailboxAddress("", Mail.EpsMail));
            emailMessage.Subject = MailDetail.EpsSubject;
            emailMessage.Body = new TextPart("html") { Text = EpsBody };
            Client.Timeout = Timeout;
            Client.SslProtocols |= System.Security.Authentication.SslProtocols.Tls12;
            Client.ServerCertificateValidationCallback = (s, c, h, e) => true;
            await Client.ConnectAsync(Mailhost, MailPort, SecureSocketOptions.StartTls);
            await Client.AuthenticateAsync(MailAccount, MailPassword);
            await Client.SendAsync(emailMessage);
            #region 寄成功執行DB寫入
            Mail.SendCount++;
            Mail.SendTime = DateTime.Now;
            Mail.SendStatus = MailSendStatus.發送成功;
            Action SendSuccess = () =>
            {
              new DA_Epaper().UpdateSendMailLog(MailDetail, Mail);
            };
            QueueService.QueueList.Enqueue(SendSuccess);
            #endregion
          }
        }
        catch (System.TimeoutException ex)
        {
          MTimer.Start();
          Debug($@"EpaperID【{MailDetail.EpsID}】寄送至 {Mail.EpsMail}回應逾時！
                   錯誤內容：{ex.Message}");
          Exceptions.Add(ex);
          ErrorLog(ex);
          throw;
        }
        catch (Exception ex)
        {
          MTimer.Start();
          Debug($@"EpaperID【{MailDetail.EpsID}】寄送至 {Mail.EpsMail}發生錯誤！
                   錯誤內容：{ex.Message}");
          Exceptions.Add(ex);
          ErrorLog(ex);
          throw;
        }
        finally
        {
          await Client.DisconnectAsync(true);
          if (DelayTime > 0)
          {
            await Task.Delay(DelayTime);
          }
        }
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
          objLogFile.Log(text).LogFlush();
        };
        QueueService.QueueList.Enqueue(WriteLog);
      }
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

    /// <summary>
    /// E-mail檢查格式
    /// </summary>
    /// <param name="Email"></param>
    /// <returns></returns>
    private bool IsValidEmail(string Email)
    {
      if (string.IsNullOrWhiteSpace(Email))
        return false;

      try
      {
        // 正則表達式來檢查郵件格式的有效性
        return Regex.IsMatch(Email,
                 @"^[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+(\.[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+)*@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z]{2,}$",
                 RegexOptions.IgnoreCase);
      }
      catch (Exception ex)
      {
        Debug($@"E-mail檢查格式發生錯誤！
                 錯誤內容：{ex.Message}");
        Exceptions.Add(ex);
        ErrorLog(ex);
        return false;
      }
    }

    /// <summary>
    /// 手動開始
    /// </summary>
    public async void StartWork()
    {
      if (DebugMode)
      {
        objLogFile.StartLog();
      }
      MTimer2.Start();
      QueueService.StartQueueWork();
      await CheckEaper();
    }

    /// <summary>
    /// 手動停止
    /// </summary>
    public void StopWork()
    {
      if (DebugMode)
      {
        objLogFile.EndLog();
      }
      MTimer.Stop();
      MTimer2.Stop();
      QueueService.StopQueueWorker();
    }

    /// <summary>
    /// 程式重啟
    /// </summary>
    private void ProgramReStart()
    {
      try
      {
        new DA_Epaper().SendRebootLetter(MaxErrorCount);
        string FilePath = Process.GetCurrentProcess().MainModule.FileName;
        ProcessStartInfo StartInfo = new ProcessStartInfo(FilePath);
        Process.Start(StartInfo);
        Environment.Exit(0);
      }
      catch (Exception ex)
      {
        Debug($@"重啟程式發生錯誤！
                 錯誤內容：{ex.Message}");
        ErrorLog(ex);
      }
    }

    /// <summary>
    /// 重啟服務
    /// </summary>
    private void ServiceReStart()
    {
      try
      {
        new DA_Epaper().SendRebootLetter(MaxErrorCount);
        var ServiceControl = new ServiceController(ServiceName);
        if (ServiceControl.Status != ServiceControllerStatus.Stopped)
        {
          var StartInfo = new ProcessStartInfo
          {
            FileName = RestartPath,
            Arguments = ServiceName,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
          };
          Process.Start(StartInfo);
          ServiceControl.Stop();
          ServiceControl.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromMinutes(1));
        }
      }
      catch (Exception ex)
      {
        Debug($@"重啟程式發生錯誤！
                 錯誤內容：{ex.Message}");
        ErrorLog(ex);
      }
    }
  }
}
