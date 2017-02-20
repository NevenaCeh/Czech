using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CitanjeIzRegistra
{
    public class Ucitavanje
    {
        public static void uzmiPodatkeIzRegEdita(string name_od_the_document)
        {
            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);

            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            XmlElement elementroot = doc.CreateElement(string.Empty, "Production", string.Empty);

            try
            {
                XmlElement element1 = doc.CreateElement(string.Empty, "Attendant", string.Empty);
                elementroot.AppendChild(element1);

                string koren = "SOFTWARE\\Wow6432Node\\Interactive Intelligence\\EIC\\Directory Services\\Root\\CIC40\\Production\\Pvw1hdcicbecz\\AttendantData\\Attendant";

                RegistryKey key = Registry.LocalMachine.OpenSubKey(koren);

                if (key != null)
                {

                    string[] profili = key.GetSubKeyNames();

                    foreach (string profil in profili)
                    {

                        XmlElement element2 = doc.CreateElement(string.Empty, "Profile", string.Empty);

                        string[] vrednostiprofila = (string[])key.OpenSubKey(profil).GetValue("FullNodePath");

                        foreach (string vp in vrednostiprofila)
                        {
                            string substringzaprofile = vp.Substring(34);

                            string daudjeuprofil = koren + "\\" + substringzaprofile;

                            XmlElement brojprofila = doc.CreateElement(string.Empty, "Broj", string.Empty);
                            XmlText tekstbroja = doc.CreateTextNode(substringzaprofile);
                            brojprofila.AppendChild(tekstbroja);
                            element2.AppendChild(brojprofila);

                            RegistryKey kljuc = Registry.LocalMachine.OpenSubKey(daudjeuprofil);

                            if (kljuc != null)
                            {
                                string[] kadudjemouprofile = kljuc.GetValueNames();

                                int profilireporting = 0;
                                int profilidnis = 0;
                                int profiliani = 0;
                                int dnisprofili = 0;
                                int audiofajl = 0;
                                int imeprofila = 0;

                                bool imadnis = false;
                                bool imaaudio = false;

                                if (kadudjemouprofile.Contains("DNISString"))
                                {
                                    imadnis = true;
                                }
                                if (kadudjemouprofile.Contains("AudioFile"))
                                {
                                    imaaudio = true;
                                }

                                for (int i = 0; i < kadudjemouprofile.Length; i++)
                                {
                                    if (kadudjemouprofile[i].Equals("Name"))
                                    {
                                        imeprofila = i;
                                    }
                                    if (kadudjemouprofile[i].Equals("Reporting"))
                                    {
                                        profilireporting = i;
                                    }
                                    if (kadudjemouprofile[i].Equals("UseANI"))
                                    {
                                        profiliani = i;
                                    }
                                    if (kadudjemouprofile[i].Equals("UseDNIS"))
                                    {
                                        profilidnis = i;
                                    }
                                    if (imadnis)
                                    {
                                        if (kadudjemouprofile[i].Equals("DNISString"))
                                        {
                                            dnisprofili = i;
                                        }
                                    }
                                    if (imaaudio)
                                    {
                                        if (kadudjemouprofile[i].Equals("AudioFile"))
                                        {
                                            audiofajl = i;
                                        }
                                    }
                                }

                                //name
                                RegistryValueKind rvkimeprofila = kljuc.GetValueKind(kadudjemouprofile[imeprofila]);
                                string[] valuesime = (string[])kljuc.GetValue(kadudjemouprofile[imeprofila]);

                                XmlElement elementimeprofila = doc.CreateElement(string.Empty, "Name", string.Empty);

                                for (int i = 0; i < valuesime.Length; i++)
                                {
                                    XmlText text2 = doc.CreateTextNode(valuesime[i]);
                                    elementimeprofila.AppendChild(text2);
                                    element2.AppendChild(elementimeprofila);
                                }
                                //name

                                ////reporting
                                RegistryValueKind rvkreporting = kljuc.GetValueKind(kadudjemouprofile[profilireporting]);
                                string[] valuesreporting = (string[])kljuc.GetValue(kadudjemouprofile[profilireporting]);

                                XmlElement element4 = doc.CreateElement(string.Empty, "ReportingGroup", string.Empty);

                                for (int i = 0; i < valuesreporting.Length; i++)
                                {
                                    XmlText text2 = doc.CreateTextNode(valuesreporting[i]);
                                    element4.AppendChild(text2);
                                    element2.AppendChild(element4);
                                }
                                
                                ////ani
                                RegistryValueKind rvkani = kljuc.GetValueKind(kadudjemouprofile[profiliani]);
                                string[] valuesasni = (string[])kljuc.GetValue(kadudjemouprofile[profiliani]);

                                XmlElement element5 = doc.CreateElement(string.Empty, "ANI", string.Empty);


                                for (int i = 0; i < valuesasni.Length; i++)
                                {
                                    XmlText text3 = doc.CreateTextNode(valuesasni[i]);
                                    element5.AppendChild(text3);
                                    element2.AppendChild(element5);
                                }
                                
                                XmlElement element6 = doc.CreateElement(string.Empty, "DNIS", string.Empty);
                                ////dnis
                                    RegistryValueKind rvkdnis = kljuc.GetValueKind(kadudjemouprofile[profilidnis]);
                                    string[] valuesdnis = (string[])kljuc.GetValue(kadudjemouprofile[profilidnis]);
                                                                        
                                    XmlElement element7 = doc.CreateElement(string.Empty, "UseDNIS", string.Empty);
                                    for (int i = 0; i < valuesdnis.Length; i++)
                                    {
                                        XmlText text3 = doc.CreateTextNode(valuesdnis[i]);
                                        element7.AppendChild(text3);
                                        element6.AppendChild(element7);
                                        element2.AppendChild(element6);
                                    }


                                //////====if contains dnis========= 
                                if (imadnis)
                                {
                                    XmlElement element8 = doc.CreateElement(string.Empty, "DNISString", string.Empty);

                                    RegistryValueKind rvkdnisstring = kljuc.GetValueKind(kadudjemouprofile[dnisprofili]);

                                    string[] valuesdnisstring = (string[])kljuc.GetValue(kadudjemouprofile[dnisprofili]);

                                    for (int i = 0; i < valuesdnisstring.Length; i++)
                                    {
                                        XmlText textdnisstring = doc.CreateTextNode(valuesdnisstring[i]);
                                        element8.AppendChild(textdnisstring);
                                        element6.AppendChild(element8);
                                    }
                                    element2.AppendChild(element6);

                                }

                                ////audio file
                                if (imaaudio)
                                {
                                    RegistryValueKind rvkaudio = kljuc.GetValueKind(kadudjemouprofile[audiofajl]);
                                    string[] valuesaudio = (string[])kljuc.GetValue(kadudjemouprofile[audiofajl]);

                                    XmlElement elementaudio = doc.CreateElement(string.Empty, "AudioFile", string.Empty);


                                    for (int i = 0; i < valuesaudio.Length; i++)
                                    {
                                        XmlText text3 = doc.CreateTextNode(valuesaudio[i]);
                                        elementaudio.AppendChild(text3);
                                        element2.AppendChild(elementaudio);
                                    }
                                }

                                
                                element1.AppendChild(element2);

                                //schedules
                                foreach (String subkeyName in kljuc.GetSubKeyNames())
                                {

                                    string[] valuessubkeys = (string[])kljuc.OpenSubKey(subkeyName).GetValue("FullNodePath");

                                    for (int i = 0; i < valuessubkeys.Length; i++)
                                    {
                                        string substringsubkeys = valuessubkeys[i].Substring(42);
                                        string putanjadaudjeuschedules = daudjeuprofil + "\\" + substringsubkeys;

                                        RegistryKey kljucevizaschedules = Registry.LocalMachine.OpenSubKey(putanjadaudjeuschedules);

                                        if (kljucevizaschedules != null)
                                        {
                                            
                                            string[] kadudjemouschedule = kljucevizaschedules.GetValueNames();

                                            int schedulename = 0;
                                            int audiosch = 0;
                                            bool schimaaudio = false;

                                            if (kadudjemouschedule.Contains("AudioFile"))
                                            {
                                                schimaaudio = true;
                                            }

                                            for (int j = 0; j < kadudjemouschedule.Length; j++)
                                            {
                                                if (kadudjemouschedule[j].Equals("Name"))
                                                {
                                                    schedulename = j;
                                                }
                                                if (schimaaudio)
                                                {
                                                    if (kadudjemouschedule[j].Equals("AudioFile"))
                                                    {
                                                        audiosch = j;
                                                    }
                                                }

                                            }

                                            XmlElement elementschedule = doc.CreateElement(string.Empty, "Schedule", string.Empty);

                                            XmlElement elementschedulename = doc.CreateElement(string.Empty, "Name", string.Empty);

                                            XmlElement elementscheduleaudio = doc.CreateElement(string.Empty, "AudioFile", string.Empty);

                                            RegistryValueKind rvkschname = kljucevizaschedules.GetValueKind(kadudjemouschedule[schedulename]);

                                            string[] valueschname = (string[])kljucevizaschedules.GetValue(kadudjemouschedule[schedulename]);

                                            for (int k = 0; k < valueschname.Length; k++)
                                            {
                                                XmlText textscheduleime = doc.CreateTextNode(valueschname[k]);
                                                elementschedulename.AppendChild(textscheduleime);
                                                elementschedule.AppendChild(elementschedulename);
                                                
                                            }
                                            if (schimaaudio)
                                            {
                                                RegistryValueKind rvkschaudio = kljucevizaschedules.GetValueKind(kadudjemouschedule[audiosch]);

                                                string[] valueschaudio = (string[])kljucevizaschedules.GetValue(kadudjemouschedule[audiosch]);

                                                for (int k = 0; k < valueschaudio.Length; k++)
                                                {
                                                    XmlText textscheduleime = doc.CreateTextNode(valueschaudio[k]);
                                                    elementscheduleaudio.AppendChild(textscheduleime);
                                                    elementschedule.AppendChild(elementscheduleaudio);
                                                }
                                            }

                                            element2.AppendChild(elementschedule);

                                            //menu actions

                                            foreach (String subkeymenuactions in kljucevizaschedules.GetSubKeyNames())
                                            {
                                                string[] akcije = (string[])kljucevizaschedules.OpenSubKey(subkeymenuactions).GetValue("FullNodePath");

                                                for (int h = 0; h < akcije.Length; h++)
                                                {
                                                    string putanjadomenuactions = putanjadaudjeuschedules + "\\";

                                                    string substringmenuactions = akcije[h].Substring(50);

                                                    string putanjadaudjeumenu = putanjadomenuactions + substringmenuactions;

                                                    RegistryKey kljucevizamenu = Registry.LocalMachine.OpenSubKey(putanjadaudjeumenu);

                                                    if (kljucevizamenu != null)
                                                    {                                                        
                                                        string[] kadudjemoumenu = kljucevizamenu.GetValueNames();

                                                        int menuname = 0;
                                                        int menutype = 0;
                                                        int menudigits = 0;
                                                        int menuaudio = 0;
                                                        bool meniimaaudio = false;
                                                        bool imaworkgroupumeni = false;
                                                        int menuworkgrupa = 0;
                                                        bool meniimanumber = false;
                                                        int menunumber = 0;

                                                        if (kadudjemoumenu.Contains("AudioFile"))
                                                        {
                                                            meniimaaudio = true;
                                                        }
                                                        if (kadudjemoumenu.Contains("Number"))
                                                        {
                                                            meniimanumber = true;
                                                        }
                                                        if (kadudjemoumenu.Contains("Workgroup"))
                                                        {
                                                            imaworkgroupumeni = true;
                                                        }
                                                        for (int n = 0; n < kadudjemoumenu.Length; n++)
                                                        {
                                                            if (kadudjemoumenu[n].Equals("Name"))
                                                            {
                                                                menuname = n;
                                                            }
                                                            if (kadudjemoumenu[n].Equals("Type"))
                                                            {
                                                                menutype = n;
                                                            }
                                                            if (kadudjemoumenu[n].Equals("Digit"))
                                                            {
                                                                menudigits = n;
                                                            }
                                                            if (meniimaaudio)
                                                            {
                                                                if (kadudjemoumenu[n].Equals("AudioFile"))
                                                                {
                                                                    menuaudio = n;
                                                                }
                                                            }
                                                            if (imaworkgroupumeni) {
                                                                if (kadudjemoumenu[n].Equals("Workgroup")) {
                                                                    menuworkgrupa = n;
                                                                }
                                                            }
                                                            if (meniimanumber)
                                                            {
                                                                if (kadudjemoumenu[n].Equals("Number"))
                                                                {
                                                                    menunumber = n;
                                                                }
                                                            }
                                                        }

                                                        XmlElement elementmenu = doc.CreateElement(string.Empty, "Operation", string.Empty);

                                                        XmlElement elementmenuname = doc.CreateElement(string.Empty, "Name", string.Empty);

                                                        XmlElement elementmenutype = doc.CreateElement(string.Empty, "Type", string.Empty);

                                                        XmlElement elementmenudigit = doc.CreateElement(string.Empty, "Digit", string.Empty);

                                                        XmlElement elementmenuaudio = doc.CreateElement(string.Empty, "AudioFile", string.Empty);

                                                        XmlElement elementmenunumber = doc.CreateElement(string.Empty, "Number", string.Empty);

                                                        RegistryValueKind rvkmenuname = kljucevizamenu.GetValueKind(kadudjemoumenu[menuname]);

                                                        string[] valuemenuname = (string[])kljucevizamenu.GetValue(kadudjemoumenu[menuname]);

                                                        for (int a = 0; a < valuemenuname.Length; a++)
                                                        {
                                                            XmlText textmenuname = doc.CreateTextNode(valuemenuname[a]);

                                                            elementmenuname.AppendChild(textmenuname);
                                                            elementmenu.AppendChild(elementmenuname);
                                                        }

                                                        RegistryValueKind rvkmenutype = kljucevizamenu.GetValueKind(kadudjemoumenu[menutype]);

                                                        string[] valuemenutype = (string[])kljucevizamenu.GetValue(kadudjemoumenu[menutype]);

                                                        for (int a = 0; a < valuemenutype.Length; a++)
                                                        {
                                                            XmlText textmenutype = doc.CreateTextNode(valuemenutype[a]);

                                                            elementmenutype.AppendChild(textmenutype);
                                                            elementmenu.AppendChild(elementmenutype);

                                                        }

                                                        RegistryValueKind rvkmenudigit = kljucevizamenu.GetValueKind(kadudjemoumenu[menudigits]);

                                                        string[] valuemenudigit = (string[])kljucevizamenu.GetValue(kadudjemoumenu[menudigits]);

                                                        for (int a = 0; a < valuemenudigit.Length; a++)
                                                        {
                                                            XmlText textmenutype = doc.CreateTextNode(valuemenudigit[a]);

                                                            elementmenudigit.AppendChild(textmenutype);
                                                            elementmenu.AppendChild(elementmenudigit);

                                                        }
                                                        if (meniimaaudio)
                                                        {
                                                            RegistryValueKind rvkmenuaudio = kljucevizamenu.GetValueKind(kadudjemoumenu[menuaudio]);

                                                            string[] valuemenuaudio = (string[])kljucevizamenu.GetValue(kadudjemoumenu[menuaudio]);

                                                            for (int a = 0; a < valuemenuaudio.Length; a++)
                                                            {
                                                                XmlText textmenuname = doc.CreateTextNode(valuemenuaudio[a]);

                                                                elementmenuaudio.AppendChild(textmenuname);
                                                                elementmenu.AppendChild(elementmenuaudio);
                                                            }
                                                        }

                                                        //workgroup of menu
                                                        if (imaworkgroupumeni)
                                                        {
                                                            XmlElement elementmenuworkgroup = doc.CreateElement(string.Empty, "Workgroup", string.Empty);

                                                            RegistryValueKind rvkmenuwrkgrup = kljucevizamenu.GetValueKind(kadudjemoumenu[menuworkgrupa]);

                                                            string[] valuemenuwrkgrup = (string[])kljucevizamenu.GetValue(kadudjemoumenu[menuworkgrupa]);

                                                            for (int a = 0; a < valuemenuwrkgrup.Length; a++)
                                                            {
                                                                XmlText textmenuworkgroup = doc.CreateTextNode(valuemenuwrkgrup[a]);

                                                                elementmenuworkgroup.AppendChild(textmenuworkgroup);
                                                                elementmenu.AppendChild(elementmenuworkgroup);
                                                            }
                                                        }

                                                        //menu number
                                                        if (meniimanumber)
                                                        {
                                                            RegistryValueKind rvkmenuwrkgrup = kljucevizamenu.GetValueKind(kadudjemoumenu[menunumber]);

                                                            string[] valuemenuwrkgrup = (string[])kljucevizamenu.GetValue(kadudjemoumenu[menunumber]);

                                                            for (int a = 0; a < valuemenuwrkgrup.Length; a++)
                                                            {
                                                                XmlText textmenunumber = doc.CreateTextNode(valuemenuwrkgrup[a]);

                                                                elementmenunumber.AppendChild(textmenunumber);
                                                                elementmenu.AppendChild(elementmenunumber);
                                                            }
                                                        }

                                                        elementschedule.AppendChild(elementmenu);

                                                        //menu submenu

                                                        foreach (String menudete in kljucevizamenu.GetSubKeyNames())
                                                        {

                                                            string[] valuesmenudeca = (string[])kljucevizamenu.OpenSubKey(menudete).GetValue("FullNodePath");

                                                            for (int o = 0; o < valuesmenudeca.Length; o++)
                                                            {

                                                                string substringmenudeca = valuesmenudeca[o].Substring(58);

                                                                string putanjadokonacnedece = putanjadaudjeumenu + "\\" + substringmenudeca;

                                                                RegistryKey kljucevizamenudecu = Registry.LocalMachine.OpenSubKey(putanjadokonacnedece);

                                                                if (kljucevizamenudecu != null)
                                                                {

                                                                    string[] kadudjemoumenudecu = kljucevizamenudecu.GetValueNames();

                                                                    int name = 0;
                                                                    int type = 0;
                                                                    int audiomenudete = 0;
                                                                    bool imaaudiodete = false;
                                                                    int value = 0;
                                                                    int subrutina = 0;
                                                                    bool imavalue = false;
                                                                    bool imasubrutinu = false;
                                                                    int workgrupdeteta = 0;
                                                                    bool imadetewrkgrupu = false;
                                                                    int menudetenumber = 0;
                                                                    bool menideteimanumber = false;

                                                                    if (kadudjemoumenudecu.Contains("Workgroup"))
                                                                    {
                                                                        imadetewrkgrupu = true;
                                                                    }
                                                                    if (kadudjemoumenudecu.Contains("Number"))
                                                                    {
                                                                        menideteimanumber = true;
                                                                    }
                                                                    if (kadudjemoumenudecu.Contains("StatValue"))
                                                                    {
                                                                        imavalue = true;
                                                                    }

                                                                    if (kadudjemoumenudecu.Contains("Subroutine"))
                                                                    {
                                                                        imasubrutinu = true;
                                                                    }

                                                                    if (kadudjemoumenudecu.Contains("AudioFile"))
                                                                    {
                                                                        imaaudiodete = true;
                                                                    }

                                                                    for (int n = 0; n < kadudjemoumenudecu.Length; n++)
                                                                    {
                                                                        if (kadudjemoumenudecu[n].Equals("Name"))
                                                                        {
                                                                            name = n;
                                                                        }
                                                                        if (kadudjemoumenudecu[n].Equals("Type"))
                                                                        {
                                                                            type = n;
                                                                        }
                                                                        if (imaaudiodete)
                                                                        {
                                                                            if (kadudjemoumenudecu[n].Equals("AudioFile"))
                                                                            {
                                                                                audiomenudete = n;
                                                                            }
                                                                        }
                                                                        if (imadetewrkgrupu)
                                                                        {
                                                                            if (kadudjemoumenudecu[n].Equals("Workgroup"))
                                                                            {
                                                                                workgrupdeteta = n;
                                                                            }
                                                                        }
                                                                        if (imavalue) {
                                                                            if (kadudjemoumenudecu[n].Equals("StatValue"))
                                                                            {
                                                                                value = n;
                                                                            }
                                                                        }
                                                                        if (imasubrutinu)
                                                                        {
                                                                            if (kadudjemoumenudecu[n].Equals("Subroutine"))
                                                                            {
                                                                                subrutina = n;
                                                                            }
                                                                        }
                                                                        if (menideteimanumber)
                                                                        {
                                                                            if (kadudjemoumenudecu[n].Equals("Number"))
                                                                            {
                                                                                menudetenumber = n;
                                                                            }
                                                                        }
                                                                    }

                                                                    XmlElement elementmenudete = doc.CreateElement(string.Empty, "Operation", string.Empty);

                                                                    XmlElement elementmenudetename = doc.CreateElement(string.Empty, "Name", string.Empty);

                                                                    XmlElement elementmenudetetype = doc.CreateElement(string.Empty, "Type", string.Empty);

                                                                    XmlElement elementmenudeteaudio = doc.CreateElement(string.Empty, "AudioFile", string.Empty);

                                                                    //name submenu
                                                                    RegistryValueKind rvkmenudetename = kljucevizamenudecu.GetValueKind(kadudjemoumenudecu[name]);

                                                                    string[] valuemenudetename = (string[])kljucevizamenudecu.GetValue(kadudjemoumenudecu[name]);

                                                                    for (int a = 0; a < valuemenudetename.Length; a++)
                                                                    {
                                                                        XmlText textmenudetename = doc.CreateTextNode(valuemenudetename[a]);

                                                                        elementmenudetename.AppendChild(textmenudetename);
                                                                        elementmenudete.AppendChild(elementmenudetename);
                                                                        
                                                                    }

                                                                    //type submenu
                                                                    RegistryValueKind rvkmenudetetype = kljucevizamenudecu.GetValueKind(kadudjemoumenudecu[type]);

                                                                    string[] valuemenudetetype = (string[])kljucevizamenudecu.GetValue(kadudjemoumenudecu[type]);

                                                                    for (int a = 0; a < valuemenudetetype.Length; a++)
                                                                    {
                                                                        XmlText textmenudetetype = doc.CreateTextNode(valuemenudetename[a]);

                                                                        elementmenudetetype.AppendChild(textmenudetetype);
                                                                        elementmenudete.AppendChild(elementmenudetetype);
                                                                    }
                                                                    //audio submenu
                                                                    if (imaaudiodete)
                                                                    {
                                                                        RegistryValueKind rvkmenudeteaudio = kljucevizamenudecu.GetValueKind(kadudjemoumenudecu[audiomenudete]);

                                                                        string[] valuemenudeteaudio = (string[])kljucevizamenudecu.GetValue(kadudjemoumenudecu[audiomenudete]);

                                                                        for (int a = 0; a < valuemenudeteaudio.Length; a++)
                                                                        {
                                                                            XmlText textmenudetename = doc.CreateTextNode(valuemenudeteaudio[a]);

                                                                            elementmenudeteaudio.AppendChild(textmenudetename);
                                                                            elementmenudete.AppendChild(elementmenudeteaudio);
                                                                        }
                                                                    }
                                                                    //value submenu
                                                                    if (imavalue)
                                                                    {
                                                                        XmlElement elementmenudetevalue = doc.CreateElement(string.Empty, "Value", string.Empty);

                                                                        RegistryValueKind rvkmenudetevalue = kljucevizamenudecu.GetValueKind(kadudjemoumenudecu[value]);

                                                                        string[] valuemenudetevalue = (string[])kljucevizamenudecu.GetValue(kadudjemoumenudecu[value]);

                                                                        for (int a = 0; a < valuemenudetevalue.Length; a++)
                                                                        {
                                                                            XmlText textmenudetename = doc.CreateTextNode(valuemenudetevalue[a]);

                                                                            elementmenudetevalue.AppendChild(textmenudetename);
                                                                            elementmenudete.AppendChild(elementmenudetevalue);
                                                                            
                                                                        }
                                                                    }
                                                                    //subroutine menu
                                                                    if (imasubrutinu)
                                                                    {
                                                                        XmlElement elementmenudetesubrutina = doc.CreateElement(string.Empty, "Subroutine", string.Empty);

                                                                        RegistryValueKind rvkmenudeteaudio = kljucevizamenudecu.GetValueKind(kadudjemoumenudecu[subrutina]);

                                                                        string[] valuemenudeteaudio = (string[])kljucevizamenudecu.GetValue(kadudjemoumenudecu[subrutina]);

                                                                        for (int a = 0; a < valuemenudeteaudio.Length; a++)
                                                                        {
                                                                            XmlText textmenudetename = doc.CreateTextNode(valuemenudeteaudio[a]);

                                                                            elementmenudetesubrutina.AppendChild(textmenudetename);
                                                                            elementmenudete.AppendChild(elementmenudetesubrutina);
                                                                            
                                                                        }
                                                                    }

                                                                    //workgroup submenu
                                                                    if (imadetewrkgrupu)
                                                                    {
                                                                        XmlElement elementmenudeteworkgroup = doc.CreateElement(string.Empty, "Workgroup", string.Empty);

                                                                        RegistryValueKind rvkmenudeteaudio = kljucevizamenudecu.GetValueKind(kadudjemoumenudecu[workgrupdeteta]);

                                                                        string[] valuemenudeteaudio = (string[])kljucevizamenudecu.GetValue(kadudjemoumenudecu[workgrupdeteta]);

                                                                        for (int a = 0; a < valuemenudeteaudio.Length; a++)
                                                                        {
                                                                            XmlText textmenudetename = doc.CreateTextNode(valuemenudeteaudio[a]);

                                                                            elementmenudeteworkgroup.AppendChild(textmenudetename);
                                                                            elementmenudete.AppendChild(elementmenudeteworkgroup);
                                                                            // elementschedule.AppendChild(elementmenu);
                                                                        }
                                                                    }

                                                                    //number submenu
                                                                    if (menideteimanumber)
                                                                    {
                                                                        XmlElement elementmenudetenumber = doc.CreateElement(string.Empty, "Number", string.Empty);

                                                                        RegistryValueKind rvkmenudetenumber = kljucevizamenudecu.GetValueKind(kadudjemoumenudecu[menudetenumber]);

                                                                        string[] valuemenudetenumber = (string[])kljucevizamenudecu.GetValue(kadudjemoumenudecu[menudetenumber]);

                                                                        for (int a = 0; a < valuemenudetenumber.Length; a++)
                                                                        {
                                                                            XmlText textmenudetenumber = doc.CreateTextNode(valuemenudetenumber[a]);

                                                                            elementmenudetenumber.AppendChild(textmenudetenumber);
                                                                            elementmenudete.AppendChild(elementmenudetenumber);
                                                                        }
                                                                    }
                                                                    elementmenu.AppendChild(elementmenudete);

                                                                    //subsubmenu
                                                                    foreach (String subkeymenuunuciactions in kljucevizamenudecu.GetSubKeyNames())
                                                                    {
                                                                        string[] unucici = (string[])kljucevizamenudecu.OpenSubKey(subkeymenuunuciactions).GetValue("FullNodePath");

                                                                        for (int y = 0; y < unucici.Length; y++) {

                                                                            string putanjadounucica = putanjadokonacnedece + "\\";
                                                                            string substringdounucica = unucici[y].Substring(66);
                                                                           
                                                                            putanjadounucica += substringdounucica;

                                                                            RegistryKey kljucevizaunuke = Registry.LocalMachine.OpenSubKey(putanjadounucica);

                                                                            if (kljucevizaunuke != null) {
                                                                                string[] uunucicima = kljucevizaunuke.GetValueNames();

                                                                                int unuciname = 0;
                                                                                int unucitype = 0;
                                                                                bool unukimanumber = false;
                                                                                int unuknumber = 0;

                                                                                bool imaaudiounuk = false;
                                                                                int audiounuk = 0;
                                                                                if (uunucicima.Contains("AudioFile"))
                                                                                {
                                                                                    imaaudiounuk = true;
                                                                                }

                                                                                bool imaunukvalue = false;
                                                                                int valueunuk = 0;
                                                                                if (uunucicima.Contains("Value"))
                                                                                {
                                                                                    imaunukvalue = true;
                                                                                }

                                                                                bool unukimasubrutinu = false;
                                                                                int subrutinaunuk = 0;
                                                                                if (uunucicima.Contains("Subroutine"))
                                                                                {
                                                                                    unukimasubrutinu = true;
                                                                                }

                                                                                bool unukimawrkgrupu = false;
                                                                                int workgrupunuka = 0;
                                                                                if (uunucicima.Contains("Workgroup"))
                                                                                {
                                                                                    unukimawrkgrupu = true;
                                                                                }

                                                                                if (uunucicima.Contains("Number")) {
                                                                                    unukimanumber = true;
                                                                                }

                                                                                for (int t = 0; t < uunucicima.Length; t++) {
                                                                                    if (uunucicima[t].Equals("Name")) {
                                                                                        unuciname = t;
                                                                                    }
                                                                                    if (uunucicima[t].Equals("Type"))
                                                                                    {
                                                                                        unucitype = t;
                                                                                    }
                                                                                    if (unukimanumber) {
                                                                                        if (uunucicima[t].Equals("Number"))
                                                                                        {
                                                                                            unuknumber = t;
                                                                                        }
                                                                                    }
                                                                                    if (imaaudiounuk)
                                                                                    {
                                                                                        if (uunucicima[t].Equals("AudioFile"))
                                                                                        {
                                                                                            audiounuk = t;
                                                                                        }
                                                                                    }
                                                                                    if (imaunukvalue)
                                                                                    {
                                                                                        if (uunucicima[t].Equals("StatValue"))
                                                                                        {
                                                                                            valueunuk = t;
                                                                                        }
                                                                                    }
                                                                                    if (unukimasubrutinu)
                                                                                    {
                                                                                        if (uunucicima[t].Equals("Subroutine"))
                                                                                        {
                                                                                            subrutinaunuk = t;
                                                                                        }
                                                                                    }
                                                                                    if (unukimawrkgrupu)
                                                                                    {
                                                                                        if (uunucicima[t].Equals("Workgroup"))
                                                                                        {
                                                                                            workgrupunuka = t;
                                                                                        }
                                                                                    }
                                                                                }
                                                                                XmlElement elementunuk = doc.CreateElement(string.Empty, "Operation", string.Empty);

                                                                                XmlElement elementunukname = doc.CreateElement(string.Empty, "Name", string.Empty);

                                                                                XmlElement elementunuktype = doc.CreateElement(string.Empty, "Type", string.Empty);

                                                                                //name

                                                                                RegistryValueKind rvkunukname = kljucevizamenu.GetValueKind(uunucicima[unuciname]);

                                                                                string[] valueunukname = (string[])kljucevizamenu.GetValue(uunucicima[unuciname]);

                                                                                for (int a = 0; a < valueunukname.Length; a++)
                                                                                {
                                                                                    XmlText textmenuname = doc.CreateTextNode(valueunukname[a]);

                                                                                    elementunukname.AppendChild(textmenuname);
                                                                                    elementunuk.AppendChild(elementunukname);
                                                                                }

                                                                                //type
                                                                                RegistryValueKind rvkunuktype = kljucevizamenu.GetValueKind(uunucicima[unucitype]);

                                                                                string[] valueunuktype = (string[])kljucevizamenu.GetValue(uunucicima[unucitype]);

                                                                                for (int a = 0; a < valueunuktype.Length; a++)
                                                                                {
                                                                                    XmlText textmenuname = doc.CreateTextNode(valueunuktype[a]);

                                                                                    elementunuktype.AppendChild(textmenuname);
                                                                                    elementunuk.AppendChild(elementunuktype);
                                                                                }

                                                                                //number unuka
                                                                                if (unukimanumber)
                                                                                {
                                                                                    XmlElement elementunuknumber = doc.CreateElement(string.Empty, "Number", string.Empty);

                                                                                    RegistryValueKind rvkunuknumber = kljucevizaunuke.GetValueKind(uunucicima[unuknumber]);

                                                                                    string[] valueunuknumber = (string[])kljucevizaunuke.GetValue(uunucicima[unuknumber]);

                                                                                    for (int a = 0; a < valueunuknumber.Length; a++)
                                                                                    {
                                                                                        XmlText textmenudetenumber = doc.CreateTextNode(valueunuknumber[a]);

                                                                                        elementunuknumber.AppendChild(textmenudetenumber);
                                                                                        elementunuk.AppendChild(elementunuknumber);
                                                                                    }
                                                                                }

                                                                                //audio unuk
                                                                                if (imaaudiounuk)
                                                                                {
                                                                                    XmlElement elementunukaudio = doc.CreateElement(string.Empty, "AudioFile", string.Empty); ;

                                                                                    RegistryValueKind rvkunukaudio = kljucevizaunuke.GetValueKind(uunucicima[audiounuk]);

                                                                                    string[] valueunukaudio = (string[])kljucevizaunuke.GetValue(uunucicima[audiounuk]);

                                                                                    for (int a = 0; a < valueunukaudio.Length; a++)
                                                                                    {
                                                                                        XmlText textunukname = doc.CreateTextNode(valueunukaudio[a]);

                                                                                        elementunukaudio.AppendChild(textunukname);
                                                                                        elementunuk.AppendChild(elementunukaudio);
                                                                                        // elementschedule.AppendChild(elementmenu);
                                                                                    }
                                                                                }
                                                                                //value deteta
                                                                                if (imaunukvalue)
                                                                                {
                                                                                    XmlElement elementunukvalue = doc.CreateElement(string.Empty, "Value", string.Empty);

                                                                                    RegistryValueKind rvkunukvalue = kljucevizaunuke.GetValueKind(uunucicima[valueunuk]);

                                                                                    string[] valueunukvalue = (string[])kljucevizaunuke.GetValue(uunucicima[valueunuk]);

                                                                                    for (int a = 0; a < valueunukvalue.Length; a++)
                                                                                    {
                                                                                        XmlText textmenudetename = doc.CreateTextNode(valueunukvalue[a]);

                                                                                        elementunukvalue.AppendChild(textmenudetename);
                                                                                        elementunuk.AppendChild(elementunukvalue);
                                                                                    }
                                                                                }
                                                                                //subrutina unuka
                                                                                if (unukimasubrutinu)
                                                                                {
                                                                                    XmlElement elementunuksubrutina = doc.CreateElement(string.Empty, "Subroutine", string.Empty);

                                                                                    RegistryValueKind rvkunuksubrutina = kljucevizaunuke.GetValueKind(uunucicima[subrutinaunuk]);

                                                                                    string[] valueunuksubrutina = (string[])kljucevizaunuke.GetValue(uunucicima[subrutinaunuk]);

                                                                                    for (int a = 0; a < valueunuksubrutina.Length; a++)
                                                                                    {
                                                                                        XmlText textmenudetename = doc.CreateTextNode(valueunuksubrutina[a]);

                                                                                        elementunuksubrutina.AppendChild(textmenudetename);
                                                                                        elementunuk.AppendChild(elementunuksubrutina);
                                                                                        // elementschedule.AppendChild(elementmenu);
                                                                                    }
                                                                                }

                                                                                //workgrupa unuka
                                                                                if (unukimawrkgrupu)
                                                                                {
                                                                                    XmlElement elementunukworkgroup = doc.CreateElement(string.Empty, "Workgroup", string.Empty);

                                                                                    RegistryValueKind rvkunukaudio = kljucevizaunuke.GetValueKind(uunucicima[workgrupunuka]);

                                                                                    string[] valueunukwrkgrupa = (string[])kljucevizaunuke.GetValue(uunucicima[workgrupunuka]);

                                                                                    for (int a = 0; a < valueunukwrkgrupa.Length; a++)
                                                                                    {
                                                                                        XmlText textmenudetename = doc.CreateTextNode(valueunukwrkgrupa[a]);

                                                                                        elementunukworkgroup.AppendChild(textmenudetename);
                                                                                        elementunuk.AppendChild(elementunukworkgroup);
                                                                                        // elementschedule.AppendChild(elementmenu);
                                                                                    }
                                                                                }

                                                                                elementmenudete.AppendChild(elementunuk);
                                                                            }
                                                                            kljucevizaunuke.Close();

                                                                        }
                                                                    }
                                                                        kljucevizamenudecu.Close();
                                                                }
                                                            }
                                                        }

                                                        kljucevizamenu.Close();
                                                    }

                                                }
                                            }

                                            kljucevizaschedules.Close();

                                        }

                                    }



                                }
                                kljuc.Close();
                            }

                        }



                    }

                    key.Close();
                }

                string korenwrkgrp = "SOFTWARE\\Wow6432Node\\Interactive Intelligence\\EIC\\Directory Services\\Root\\CIC40\\Production\\Workgroups";

                RegistryKey keywrkgrp = Registry.LocalMachine.OpenSubKey(korenwrkgrp);

                XmlElement elementwrkgrp = doc.CreateElement(string.Empty, "Workgroups", string.Empty);

                if (keywrkgrp != null)
                {

                    string[] workgrupe = keywrkgrp.GetSubKeyNames();

                    foreach (string grupa in workgrupe)
                    {
                        XmlElement elementwrkgrupa = doc.CreateElement(string.Empty, "Workgroup", string.Empty);
                        XmlElement elementimewrkgrp = doc.CreateElement(string.Empty, "Name", string.Empty);
                        XmlText textimewrkgrp = doc.CreateTextNode(grupa);
                        elementimewrkgrp.AppendChild(textimewrkgrp);
                        elementwrkgrupa.AppendChild(elementimewrkgrp);
                        elementwrkgrp.AppendChild(elementwrkgrupa);

                        string[] queue = (string[])keywrkgrp.OpenSubKey(grupa).GetValue("Queue Management Style");

                        foreach (string mu in queue)
                        {
                            XmlElement elementqueue = doc.CreateElement(string.Empty, "QueueManagementStyle", string.Empty);
                            XmlText textqueue = doc.CreateTextNode(mu);
                            elementqueue.AppendChild(textqueue);
                            elementwrkgrupa.AppendChild(elementqueue);
                        }

                        string daudjeuwrkgrupu = korenwrkgrp + "\\" + grupa;

                        RegistryKey kljuczaworkgroupe = Registry.LocalMachine.OpenSubKey(daudjeuwrkgrupu);

                        if (kljuczaworkgroupe != null)
                        {

                            string[] valuesinworkgroup = kljuczaworkgroupe.GetValueNames();

                            int utilization = 0;
                            bool imautilization = false;
                            int useri = 0;
                            bool imausere = false;
                            int mailbox = 0;
                            bool imamailbox = false;
                            int mailboxdisplayname = 0;
                            int mailoption = 0;

                            if (valuesinworkgroup.Contains("Agent Media % Utilization"))
                            {
                                imautilization = true;
                            }
                            if (valuesinworkgroup.Contains("Users"))
                            {
                                imausere = true;
                            }
                            if (valuesinworkgroup.Contains("Mailbox"))
                            {
                                imamailbox = true;
                            }

                            for (int i = 0; i < valuesinworkgroup.Length; i++)
                            {
                                if (imautilization)
                                {
                                    if (valuesinworkgroup[i].Equals("Agent Media % Utilization"))
                                    {
                                        utilization = i;
                                    }
                                }
                                if (imausere)
                                {
                                    if (valuesinworkgroup[i].Equals("Users"))
                                    {
                                        useri = i;
                                    }
                                }
                                if (imamailbox)
                                {
                                    if (valuesinworkgroup[i].Equals("Mailbox"))
                                    {
                                        mailbox = i;
                                    }
                                    if (valuesinworkgroup.Contains("Mailbox Display Name"))
                                    {
                                        mailboxdisplayname = i;
                                    }
                                    if (valuesinworkgroup[i].Equals("Mailbox Option"))
                                    {
                                        mailoption = i;
                                    }
                                }
                            }

                            ////utilization
                            if (imautilization)
                            {
                                RegistryValueKind rvkutilization = kljuczaworkgroupe.GetValueKind(valuesinworkgroup[utilization]);
                                string[] valuesreporting = (string[])kljuczaworkgroupe.GetValue(valuesinworkgroup[utilization]);

                                XmlElement elementutilization = doc.CreateElement(string.Empty, "Utilization", string.Empty);

                                for (int i = 0; i < valuesreporting.Length; i++)
                                {
                                    XmlText text2 = doc.CreateTextNode(valuesreporting[i]);
                                    elementutilization.AppendChild(text2);
                                    elementwrkgrupa.AppendChild(elementutilization);
                                }
                            }
                            ////utilization

                            ////useri
                            if (imausere)
                            {
                                RegistryValueKind rvkutilization = kljuczaworkgroupe.GetValueKind(valuesinworkgroup[useri]);
                                string[] valuesreporting = (string[])kljuczaworkgroupe.GetValue(valuesinworkgroup[useri]);

                                XmlElement elementuser = doc.CreateElement(string.Empty, "Users", string.Empty);

                                for (int i = 0; i < valuesreporting.Length; i++)
                                {
                                    XmlText textimeusera = doc.CreateTextNode(valuesreporting[i]);
                                    XmlElement elementime = doc.CreateElement(string.Empty, "User", string.Empty);
                                    elementime.AppendChild(textimeusera);
                                    elementuser.AppendChild(elementime);

                                }
                                elementwrkgrupa.AppendChild(elementuser);
                            }

                            //mailbox
                            if (imamailbox)
                            {
                                RegistryValueKind rvkutilization = kljuczaworkgroupe.GetValueKind(valuesinworkgroup[mailbox]);
                                string[] valuesreporting = (string[])kljuczaworkgroupe.GetValue(valuesinworkgroup[mailbox]);

                                XmlElement elementmailbox = doc.CreateElement(string.Empty, "Mailbox", string.Empty);

                                for (int i = 0; i < valuesreporting.Length; i++)
                                {
                                    XmlText textimeusera = doc.CreateTextNode(valuesreporting[i]);
                                    elementmailbox.AppendChild(textimeusera);

                                }

                                RegistryValueKind rvkoption = kljuczaworkgroupe.GetValueKind(valuesinworkgroup[mailoption]);
                                string[] valuesoption = (string[])kljuczaworkgroupe.GetValue(valuesinworkgroup[mailoption]);

                                XmlElement elementoption = doc.CreateElement(string.Empty, "MailboxOption", string.Empty);

                                for (int i = 0; i < valuesoption.Length; i++)
                                {
                                    XmlText textimeusera = doc.CreateTextNode(valuesoption[i]);
                                    elementoption.AppendChild(textimeusera);
                                }

                                elementwrkgrupa.AppendChild(elementmailbox);
                                
                                elementwrkgrupa.AppendChild(elementoption);
                            }
                            }              
                        
                        kljuczaworkgroupe.Close();
                    }

                }
                keywrkgrp.Close();

                elementroot.AppendChild(elementwrkgrp);
            }
            catch (Exception ex) {
                Console.WriteLine(ex.StackTrace);
            }

            doc.AppendChild(elementroot);

            string putanjadofajla = @"D:\I3\IC\Server\Reports\Custom\ProfileReport\" + name_od_the_document + ".xml";

            doc.Save(putanjadofajla);
        }
    }
}
