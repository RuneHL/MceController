using System;
using System.IO;
using System.Linq;

namespace VmcController.AddIn
{
    class GetFileInfo
    {
        public static string GetNewestImage(string image)
        {
            string file = image.Split(Path.DirectorySeparatorChar).Last();
            string dir = image.Substring(0, image.LastIndexOf(Path.DirectorySeparatorChar));

            return dir + Path.DirectorySeparatorChar + GetNewestImageLinq(new DirectoryInfo(@dir), file);
        }

        private static string GetNewestImageLinq(DirectoryInfo directory, string FileMask = "*")
        {
            try
            {
                var myFile = directory.GetFiles(FileMask)
                            .OrderByDescending(f => f.LastWriteTime)
                            .First();

                return myFile.Name;
            }
            catch { return null; }
        }

    }
}
