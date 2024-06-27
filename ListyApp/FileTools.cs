using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp
{
    public class FileTools
    {
        public class FileToolsException : Exception
        {
            public FileToolsException(string message, Exception innerException)
                : base(message, innerException)
            { }

            public FileToolsException(string message)
                : base(message)
            { }
        }

        public static void DeleteAllPdfFiles(string folder)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(folder);
                foreach (FileInfo file in dir.GetFiles("*.pdf"))
                    file.Delete();
            }
            catch (Exception ex)
            {
                throw new FileToolsException("Deleting all *.pdf files failed.", ex);
            }
        }

        public static string GetIncFileName(string pathfile)
        {
            string extension = Path.GetExtension(pathfile);
            string folder = Path.GetDirectoryName(pathfile);
            string filename = Path.GetFileNameWithoutExtension(pathfile);

            int number = 1;
            while (File.Exists(pathfile))
            {
                filename = $"{filename}_{number}";
                pathfile = Path.Combine(folder, filename + extension);
            }

            return pathfile;
        }
    }
}
