using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using  Spire.Doc;
using System.Threading;
namespace 合并word
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("请将需要合并的word放入一个文件夹（确保文件夹排名为需要合并的文件的排名），并将此文件夹拖入cmd后回车");
            string directory = Console.ReadLine();
            if (!Directory.Exists(directory)) return;
            DirectoryInfo root = new DirectoryInfo(directory);
            Document docTemp = new Document();
            Console.WriteLine(root);
            foreach (var item in root.GetFiles())
            {
                docTemp.InsertTextFromFile(directory + "\\" + item.ToString(), FileFormat.Docx);
                Console.WriteLine(item);
            }

            Console.WriteLine("请输入要保存的docx文件的文件名（默认为---结果.docx）");
            string filename = Console.ReadLine();
            Console.WriteLine("即将保存文件");
            if (filename!=null)
            {
                docTemp.SaveToFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+"\\"+filename+".docx", FileFormat.Docx);
            }
            else
            {
                docTemp.SaveToFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+"\\结果.docx", FileFormat.Docx);
            }
            Console.WriteLine("文件保存完毕");
            Console.ReadLine();
        }




    }
}
