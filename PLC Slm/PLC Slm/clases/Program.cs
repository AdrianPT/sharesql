using PLC_Slm.clases;
using Quartz;
using Quartz.Impl;
using SimpleLogger;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace PLC_Slm
{
    class Program
    {


        public static void Main(string[] args)
        {
            SimpleLog.SetLogFile(logDir: ".\\Log", prefix: "Log_", writeText: false);

            try
            {
                // Common.Logging.LogManager.Adapter = new Common.Logging.Simple.ConsoleOutLoggerFactoryAdapter {Level = Common.Logging.LogLevel.Info};

                // Grab the Scheduler instance from the Factory 
                IScheduler scheduler = StdSchedulerFactory.GetDefaultScheduler();

                // and start it off
                scheduler.Start();
                int numSeg = Int32.Parse(readConfig("Geral", "Intervalo de registo (seg) ", ""));
                // define el trabajo y lo ata a la clase Hacer
                IJobDetail job = JobBuilder.Create<Make>()
                    .WithIdentity("trabajo1", "grupo1")
                    .Build();

                // Inicia el trigger para empezar el trabajo ahora, y repite cada 10 segundos
                // var trigger = new CronTrigger("trigger1", "group1", "job1", "group1", "0 0 1 ? * MON-FRI");
                ITrigger trigger = TriggerBuilder.Create()
                    .WithIdentity("gatillo1", "grupo1")
                    .StartNow()
                    .WithSimpleSchedule(x => x
                        .WithIntervalInSeconds( numSeg )               
                        
                        .WithRepeatCount(0)
                    .RepeatForever()
                        )
                    .Build();

                // Tell quartz to schedule the job using our trigger
                scheduler.ScheduleJob(job, trigger);

                // A los 10 segundos duerme
                Thread.Sleep(TimeSpan.FromSeconds(400));

                // apaga el scheduler cuando estoy listo para cerrar el programa
                scheduler.Shutdown();

            }
            catch (SchedulerException se)
            {

                SimpleLog.Log(se);

            }


            Environment.Exit(0);

        }



        #region ReadConfig
        public static string readConfig(string MainSection, string key, string defaultValue)
        {
            string urlConfig = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            urlConfig = urlConfig + "\\config.ini";

            IniFile inif = new IniFile(urlConfig);
            string value = "";

            value = (inif.IniReadValue(MainSection, key, defaultValue));
            return value;
        }
        #endregion

    }
}
