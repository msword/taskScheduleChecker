using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskScheduler;
using System.Configuration;
namespace ScheduledTaskChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            tschedule t = new tschedule();
            t.ParseScheduleTasks();

            t.SendScheduledTaskReport(ConfigurationManager.AppSettings["serverName"]);
        }


    }


}
