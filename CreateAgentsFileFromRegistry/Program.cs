using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateAgentsFileFromRegistry
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Console.WriteLine("Name the xml document: ");

            string name_od_the_document = "skills";

            var time_of_event = DateTime.Now;

            var date_of_event = time_of_event.Date;
            var year_of_event = date_of_event.Year;
            var month_of_event = date_of_event.Month;
            var day_of_event = date_of_event.Day;
            string date_string = string.Format("{0}.{1}.{2}", day_of_event, month_of_event, year_of_event);

            var sat = time_of_event.Hour;
            var min = time_of_event.Minute;

           // name_od_the_document = name_od_the_document + "Date" + date_string + "Hour" + sat+"Min"+min;

            Ucitavanje.uzmiPodatkeIzRegEdita(name_od_the_document);

          //  Console.WriteLine("Document " + name_od_the_document + ".xml is in \\pvw1hdcicbecz\\IC_Reports\\Custom\\AgentSkillWorkgroup folder!");

            //Console.WriteLine("Press Enter to exit application...");

           // Console.Read();

        }
    }
}
