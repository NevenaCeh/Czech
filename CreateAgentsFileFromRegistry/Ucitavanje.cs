using Microsoft.Win32;
using System;
using System.IO;
using System.Text;
using System.Xml;

namespace CreateAgentsFileFromRegistry
{
    internal class Ucitavanje
    {
        public static void uzmiPodatkeIzRegEdita(string ime)
        {

            XmlDocument doc = new XmlDocument();

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);

            XmlElement root = doc.DocumentElement;

            doc.InsertBefore(xmlDeclaration, root);

            XmlElement elementkoreni = doc.CreateElement(string.Empty, "Users", string.Empty);

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Agent,Skill,Proficiency,Desire To Use");
            try
            {
                string koren = "SOFTWARE\\Wow6432Node\\Interactive Intelligence\\EIC\\Directory Services\\Root\\CIC40\\Production\\Users";

                RegistryKey key = Registry.LocalMachine.OpenSubKey(koren);

                if (key != null)
                {

                    string[] useri = key.GetSubKeyNames();

                    foreach (string user in useri)
                    {
                        XmlElement element1 = doc.CreateElement(string.Empty, "User", string.Empty);
                        elementkoreni.AppendChild(element1);

                        string udjiuusera = koren + "\\" + user;

                        RegistryKey kljucuser = Registry.LocalMachine.OpenSubKey(udjiuusera);

                        XmlElement userelement = doc.CreateElement(string.Empty, "Name", string.Empty);
                        XmlText tekstuserelementa = doc.CreateTextNode(user);
                        userelement.AppendChild(tekstuserelementa);
                        element1.AppendChild(userelement);

                        if (kljucuser != null)
                        {
                            string[] useratributi = kljucuser.GetValueNames();

                            bool imadisplayime = false;
                            int displayime = 0;
                            bool imaskills = false;
                            int skills = 0;

                            for (int i = 0; i < useratributi.Length; i++)
                            {
                                if (useratributi[i].Equals("displayName"))
                                {
                                    imadisplayime = true;
                                    if (imadisplayime)
                                    {
                                        displayime = i;
                                    }
                                }

                                if (useratributi[i].Equals("Skills"))
                                {
                                    imaskills = true;
                                    if (imaskills)
                                    {
                                        skills = i;
                                    }
                                }
                            }

                            if (imaskills)
                            {
                                XmlElement skilovi = doc.CreateElement(string.Empty, "Skills", string.Empty);
                                RegistryValueKind rvk = kljucuser.GetValueKind(useratributi[skills]);

                                string[] vrednostskills = (string[])kljucuser.GetValue(useratributi[skills]);

                                for (int i = 0; i < vrednostskills.Length; i++)
                                {
                                    XmlElement skil = doc.CreateElement(string.Empty, "Skill", string.Empty);
                                    XmlElement imeskila = doc.CreateElement(string.Empty, "Name", string.Empty);
                                    XmlElement proficiencyskila = doc.CreateElement(string.Empty, "Proficiency", string.Empty);
                                    XmlElement desiretouseskila = doc.CreateElement(string.Empty, "DesireToUse", string.Empty);
                                    int index = vrednostskills[i].IndexOf('|');
                                    string imeskilastr = vrednostskills[i].Substring(0, index);
                                    string ostalo = vrednostskills[i].Substring(index + 1);

                                    int indprof = ostalo.IndexOf('|');
                                    string proficiency = ostalo.Substring(0, indprof);
                                    string izakrajdesiretouse = ostalo.Substring(indprof + 1);
                                    izakrajdesiretouse = izakrajdesiretouse.Remove(izakrajdesiretouse.Length - 1);
                                    
                                    sb.AppendLine(user + "," + imeskilastr + "," + proficiency + "," + izakrajdesiretouse);

                                    XmlText text2 = doc.CreateTextNode(imeskilastr);
                                    imeskila.AppendChild(text2);

                                    XmlText textprof = doc.CreateTextNode(proficiency);
                                    proficiencyskila.AppendChild(textprof);

                                    XmlText textdtu = doc.CreateTextNode(izakrajdesiretouse);
                                    desiretouseskila.AppendChild(textdtu);

                                    skil.AppendChild(imeskila);
                                    skil.AppendChild(proficiencyskila);
                                    skil.AppendChild(desiretouseskila);
                                    skilovi.AppendChild(skil);
                                    element1.AppendChild(skilovi);
                                }
                            }
                        }
                        kljucuser.Close();
                    }
                }
                key.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }


            doc.AppendChild(elementkoreni);

           // D:\I3\IC\Server\Reports\Custom\AgentSkillWorkgroup
            string putanjadofajla = "D:\\I3\\IC\\Server\\Reports\\Custom\\AgentSkillWorkgroup\\" + ime + ".csv";
            string putanjadofajlaxml = "D:\\I3\\IC\\Server\\Reports\\Custom\\AgentSkillWorkgroup\\" + ime + ".xml";
            doc.Save(putanjadofajlaxml);
            File.AppendAllText(putanjadofajla, sb.ToString());
        }
    }
}