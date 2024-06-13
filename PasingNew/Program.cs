using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using System.Xml;
using Microsoft.Office.Interop.MSProject;
using Application = Microsoft.Office.Interop.MSProject.Application;
using Task = Microsoft.Office.Interop.MSProject.Task;

namespace PasingNew
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();
            try
            {
                string xmlFilePath = "C:\\Users\\ADMIN\\Desktop\\Список.xml";
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);
                XmlNodeList stages = xmlDoc.SelectNodes("/project/stage");
                Project project = app.Projects.Add();
                foreach (XmlNode stage in stages)
                {
                    string name = stage.SelectSingleNode("Name").InnerText;
                    string dateStart = stage.SelectSingleNode("DateStart").InnerText;
                    string duration = stage.SelectSingleNode("Duration").InnerText;
                    Task task = project.Tasks.Add(name);
                    task.Start = DateTime.Parse(dateStart);
                    task.Duration = duration;
                    XmlNodeList users = stage.SelectNodes("Users/User");
                    List<string> userNames = new List<string>();
                    foreach (XmlNode user in users)
                    {
                        userNames.Add(user.InnerText);
                    }
                    task.ResourceNames = string.Join(";", userNames);

                    XmlNode connectStage = stage.SelectSingleNode("connectStage");
                    if (connectStage != null)
                    {
                        string connect = stage.SelectSingleNode("connectStage").InnerText;
                        task.Predecessors = connect;
                    }
                }
                string filePath = "C:\\Users\\ADMIN\\Desktop\\Новый проект.mpp";
                project.SaveAs(filePath);
                app.Quit();
                Console.WriteLine("Данные успешно добавлены в файл MPP.");
                Console.ReadKey();
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Не корректный xml файл");
                app.Quit();
                Console.ReadKey();
            }
        }
    }
}
