using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskScheduler;
using System.IO;
using System.Net;
using System.Net.Mail;
using Newtonsoft.Json;
using System.Configuration;

namespace ScheduledTaskChecker
{
    class tschedule
    {
        //
        // Inspiration came from this post. https://stackoverflow.com/questions/41366002/how-to-list-scheduled-tasks-in-c-sharp

        private StringBuilder htmlBody = new StringBuilder();
        private WatchList watchList;
        private string configStr;
        public tschedule()
        {
            using (StreamReader sr = new StreamReader(ConfigurationManager.AppSettings["watchlist"].ToString()))
            {
                // Read the stream to a string, and write the string to the console.
                configStr = sr.ReadToEnd();
                Console.WriteLine(configStr);
                watchList = JsonConvert.DeserializeObject<WatchList>(configStr);
            }
        }

        public void ProcessTaskFolder(ITaskFolder taskFolder)
        {
            int idx;
            string name, path;
            _TASK_STATE state;
            DateTime lastRun;

            IRegisteredTaskCollection taskCol = taskFolder.GetTasks((int)_TASK_ENUM_FLAGS.TASK_ENUM_HIDDEN);  // include hidden tasks, otherwise 0
            for (idx = 1; idx <= taskCol.Count; idx++)  // browse all tasks in folder
            {
                IRegisteredTask runTask = taskCol[idx];  // 1 based index
                

                name = runTask.Name;
                path = runTask.Path;
                state = runTask.State;
                lastRun = runTask.LastRunTime;

               

                Console.WriteLine(path);
                foreach (var wt in watchList.tasks)
                {
                    if (wt == name)
                    {
                        string stateHtml = "";
                        stateHtml = (state.ToString() != "TASK_STATE_DISABLED") ? "<span style='color:green'>ENABLED</span>" : "<span style='color:red'>DISABLED</span>";
                       WriteContent("DateTime: " + DateTime.Now + " name: " + name + " status: " + state +" last run time " + lastRun.ToString() +" \n");

                        htmlBody.Append("<tr>");
                        htmlBody.AppendFormat("<td><b>{0}</b></td> <td>{1}</td> <td><span style='font-size:9pt'> {2} </span></td>", name, stateHtml, lastRun);
                        htmlBody.Append("<tr>");
                    }
                }
            }

            ITaskFolderCollection taskFolderCol = taskFolder.GetFolders(0);  // 0 = reserved for future use
            for (idx = 1; idx <= taskFolderCol.Count; idx++)  // recursively browse subfolders
                ProcessTaskFolder(taskFolderCol[idx]);  // 1 based index
        }

        public void ParseScheduleTasks()
        {
            ITaskService taskService = new TaskScheduler.TaskScheduler();
            taskService.Connect();

            ProcessTaskFolder(taskService.GetFolder("\\"));
        }

        public void WriteContent(string content)
        {
            string path = ConfigurationManager.AppSettings["txtFile"];
            string file = "tasks.txt";

            // This text is added only once to the file.
            if (!File.Exists(path + file))
            {
                File.WriteAllText(path + file, content);
            }
            else
            {
                File.AppendAllText(path + file, content);
            }
        }

        public  void SendScheduledTaskReport(string subject)
        {
           
            
                MailAddress to = new MailAddress(ConfigurationManager.AppSettings["sendTo"]);
                MailAddress from = new MailAddress(ConfigurationManager.AppSettings["mailFrom"]);
               

                MailMessage message = new MailMessage(from, to);

            string cc = ConfigurationManager.AppSettings["cc"];
            string[] ccList = null;
            if(cc != "")
            {
                ccList = cc.Split(',');
            }
            if(ccList.Count() > 0)
            {
                foreach(var c in ccList)
                {
                    MailAddress ccMember = new MailAddress(c);
                    message.CC.Add(ccMember);
                }
            }


            StringBuilder html = new StringBuilder();
            message.Subject = subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;

            html.AppendFormat("<html><body style='font-family:verdana'><h2>{0}</h2>", subject);
            html.Append("<table><tr><td>TASK</td><td>STATE</td><td>LAST RUN</td></tr>");
            html.Append(htmlBody.ToString());
            html.Append("</table>");

            html.Append("</body></html>");

            message.Body = html.ToString();
            message.IsBodyHtml = true;
            SmtpClient client = new SmtpClient();
            client.Host = ConfigurationManager.AppSettings["mailHost"];

            client.Send(message);
            message.Dispose();

        }
    }

    public class WatchList
    {
        public IList<string> tasks { get; set; }
    }
}
