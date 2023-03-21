using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Xml;

namespace WidenApi
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("");
            Console.WriteLine("------------------------------------------");
            Console.WriteLine("");
            SettingsModel settings = GetSettings();
            foreach (DelVehs d in settings.DelVehList)
            {
                if (!Directory.Exists(d.Path))
                {
                    Console.WriteLine("The path for the Del Veh " + d.Id + " does not exist.");
                    Console.WriteLine(d.Path);
                    Console.WriteLine("Press any key to continue.");
                    Console.ReadLine();
                }
            }

            string incomingFolder = settings.IncomingFolder;
            string errorFolder = settings.ErrorFolder;
            string outputFolder = settings.OutputFolder;
            string doneFolder = settings.DoneFolder;

            Directory.CreateDirectory(errorFolder);
            Directory.CreateDirectory(incomingFolder);
            Directory.CreateDirectory(outputFolder);
            Directory.CreateDirectory(doneFolder);
            string baseUrl = settings.BaseUrl;
            //string baseUrl = "https://private-anon-ad6b562596-widenv1.apiary-mock.com/api/rest/asset/search/";

            Console.WriteLine("For Manual Search type \"m\", otherwise hit any other key");
            string type = Console.ReadLine().ToLower().Trim();
            if (type.Equals("m"))
            {
                ManualInput(outputFolder, baseUrl);
            }
            else
            {
                Console.WriteLine("Starting Excel Sheet Automatic Processing...");
                ExcelSheetInput(incomingFolder, outputFolder, doneFolder, baseUrl);
            }


        }

        private static void ManualInput(string outputFolder, string baseUrl)
        {
            string[] arg = null;
            bool go = true;

            while (go)
            {
                Console.WriteLine("Search For?");
                string search = Console.ReadLine();
                
                CallAPI(baseUrl + search, outputFolder);
            }
            Main(arg);
        }

        private static void ExcelSheetInput(string incomingFolder, string outputFolder, string doneFolder, string baseUrl)
        {
            DirectoryInfo dinfo = new DirectoryInfo(incomingFolder);
            bool go = true;
            string[] arg = null;

            while (go)
            {
                if (dinfo.Exists)
                {
                    FileInfo[] files = dinfo.GetFiles("*xlsx");
                    int count = 1;
                    foreach (FileInfo f in files)
                    {
                        if (!f.Name.StartsWith("~$") || !f.Name.StartsWith("."))
                        {
                            Console.WriteLine(f.Name);
                            CallAPI(baseUrl, outputFolder, GetDownloads(f.FullName), f.FullName);
                        }
                        if (f.Exists)
                        {
                            File.Move(f.FullName, Path.Combine(doneFolder, f.Name), true);
                        }
                        if (count.Equals(files.Length))
                        {
                            Console.WriteLine("-----------------------------------------");
                            Console.WriteLine("Last File Finished...");
                            Console.WriteLine("-----------------------------------------");
                        }
                        count++;
                    }
                    
                }

                Thread.Sleep(10000);
            }
            Main(arg);
        }

        private static List<ExcelSheetModel> GetDownloads(string fileName)
        {
            
            List<ExcelSheetModel> downloadList = new List<ExcelSheetModel>();
            IWorkbook wb = new XSSFWorkbook(fileName);
            ISheet sheet = wb.GetSheetAt(0);
            int rowcount = GetRowCount(sheet);
            for (int i = 1; i < GetRowCount(sheet); i++)
            {
                ExcelSheetModel excelSheet = new ExcelSheetModel();
                
                if (fileName.ToLower().Contains("ft-") || fileName.ToLower().Contains("aa-"))
                {
                    if (fileName.ToLower().Contains("ft-"))
                    {
                        excelSheet.JobName = "FT-" + GetCell(sheet, i, 10);
                    }
                    else
                    {
                        excelSheet.JobName = "AA-" + GetCell(sheet, i, 10);
                    }
                    excelSheet.StyleNumber = GetCell(sheet, i, 0);
                    excelSheet.DelVeh = GetCell(sheet, i, 2);
                    excelSheet.BoardSku = GetCell(sheet, i, 5).TrimStart('0');
                    excelSheet.BoardSku = excelSheet.BoardSku.Trim().TrimStart('0');
                    excelSheet.StyleName = GetCell(sheet, i, 1);
                    excelSheet.CustomerNumber = GetCell(sheet, i, 3);
                    excelSheet.SequenceNumber = GetCell(sheet, i, 4);
                    excelSheet.StampingInstructions = GetCell(sheet, i, 6);
                    excelSheet.FirstColor = GetCell(sheet, i, 7);
                    excelSheet.NumOrders = GetCell(sheet, i, 8);
                    excelSheet.TotalQty = GetCell(sheet, i, 9);
                }
                else
                {
                    excelSheet.JobName = GetCell(sheet, i, 3);
                    excelSheet.StyleNumber = GetCell(sheet, i, 0);
                    excelSheet.DelVeh = GetCell(sheet, i, 1);
                    excelSheet.BoardSku = GetCell(sheet, i, 2).TrimStart('0');
                    excelSheet.BoardSku = excelSheet.BoardSku.Trim().TrimStart('0');
                    excelSheet.CustomerNumber = "";
                    excelSheet.SequenceNumber = "";
                }
                downloadList.Add(excelSheet);
            }
            wb.Close();

            return downloadList;
        }

        public static void CallAPI(string url, string downloadLoc)
        {
            SettingsModel settings = GetSettings();
            string token = settings.APIToken;
            List<UrlModel> urlList = new List<UrlModel>();

            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(url);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    HttpResponseMessage response = client.GetAsync("?options=downloadUrl").Result;

                    if (response.StatusCode.Equals(HttpStatusCode.OK))
                    {
                        var jsonStream = response.Content.ReadAsStream();
                        using (StreamReader reader = new(jsonStream))
                        {
                            string data = reader.ReadToEnd();
                            JObject jsonObject = JObject.Parse(data);
                            int results = (int)jsonObject["numResults"];
                            Console.WriteLine("Number of Results: " + results);
                            for (int i = 0; i < results; i++)
                            {
                                string downloadUrl = (string)jsonObject["assets"][i]["downloadUrl"];
                                string fileName = (string)jsonObject["assets"][i]["name"];
                                string size = (string)jsonObject["assets"][i]["size"];
                                //Console.WriteLine("Download Url: " + downloadUrl);
                                if (results.Equals(1))
                                {
                                    FileDownload(downloadLoc, downloadUrl, fileName, size);
                                }
                                else if (results > 1)
                                {
                                    UrlModel urlModel = new UrlModel();
                                    urlModel.FileName = fileName;
                                    urlModel.URL = downloadUrl;
                                    urlModel.Size = size;
                                    urlList.Add(urlModel);
                                }
                            }
                            if (!urlList.Count.Equals(0))
                            {
                                int urlCount = 0;
                                Console.WriteLine("There are " + urlList.Count + " to choose from. Type the number of the one you want to download.");
                                foreach (UrlModel um in urlList)
                                {
                                    Console.WriteLine("[" + urlCount + "] - " + um.FileName);
                                    urlCount++;
                                }
                                int urlChoice = 0;
                                bool parseOk = false;
                                while (!parseOk)
                                {
                                    parseOk = int.TryParse(Console.ReadLine(), out urlChoice);
                                    if (urlChoice > urlList.Count - 1 || !parseOk)
                                    {
                                        parseOk = false;
                                        Console.WriteLine("Pick a number within range.");
                                    }
                                }
                                FileDownload(downloadLoc, urlList[urlChoice].URL, urlList[urlChoice].FileName, urlList[urlChoice].Size);
                                /*using (var webClient = new WebClient())
                                {
                                    Console.WriteLine("Download started for " + urlList[urlChoice].FileName);
                                    Console.WriteLine("File size: " + urlList[urlChoice].Size);
                                    webClient.DownloadFile(urlList[urlChoice].URL, Path.Combine(downloadLoc, urlList[urlChoice].FileName));
                                    Console.WriteLine("File Download Finished.");
                                    Console.WriteLine("Starting Zip Extration...");
                                    Run(Path.Combine(downloadLoc, urlList[urlChoice].FileName), downloadLoc, true);
                                    Console.WriteLine("Zip Extration Finished...");
                                    Console.WriteLine("-----------------------------------------");
                                    
                                }*/
                            }


                        }
                        string stop = "";
                    }
                    else
                    {
                        Console.WriteLine(string.Format("{0}:{1}", response.StatusCode, response.ReasonPhrase));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void CallAPI(string baseUrl, string outputFolder, List<ExcelSheetModel> excelSheets, string filePath)
        {
            SettingsModel settings = GetSettings();
            string token = settings.APIToken;
            List<UrlModel> urlList = new List<UrlModel>();
            foreach (ExcelSheetModel es in excelSheets)
            {
                if (es.StyleNumber.ToLower().Trim().Equals("null"))
                {
                    es.StyleNumber = "";
                }
                string search = es.StyleNumber + " " + es.BoardSku + " " + es.CustomerNumber + " " + es.SequenceNumber;
                string url = baseUrl + search.Trim();
                try
                {
                    using (var client = new HttpClient())
                    {
                        client.BaseAddress = new Uri(url);
                        client.DefaultRequestHeaders.Accept.Clear();
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                        HttpResponseMessage response = client.GetAsync("?options=downloadUrl&metadata=customerNumber,sequenceNumber").Result;

                        if (response.StatusCode.Equals(HttpStatusCode.OK))
                        {
                            var jsonStream = response.Content.ReadAsStream();
                            using (StreamReader reader = new(jsonStream))
                            {
                                string data = reader.ReadToEnd();
                                JObject jsonObject = JObject.Parse(data);
                                int results = (int)jsonObject["numResults"];
                                List<int> location = new();
                                for (int i = 0; i < results; i++)
                                {
                                    string customerNumber = "";
                                    string sequenceNumber = "";
                                    if (!es.StyleNumber.Equals(""))
                                    {
                                        customerNumber = (string)jsonObject["assets"][i]["metadata"][0]["value"];
                                        sequenceNumber = (string)jsonObject["assets"][i]["metadata"][1]["value"];
                                    }
                                    //customerNumber = String.Format("{0}", customerNumber);
                                    //sequenceNumber = String.Format("{0}", sequenceNumber);                                    

                                    if (customerNumber == null)
                                    {
                                        customerNumber = "";
                                    }
                                    else
                                    {
                                        customerNumber = customerNumber.Replace(".0", "").PadLeft(7, '0');
                                    }
                                    if (sequenceNumber == null)
                                    {
                                        sequenceNumber = "";
                                    }
                                    else
                                    {
                                        sequenceNumber = sequenceNumber.Replace(".0", "").PadLeft(2, '0');
                                    }
                                    if (customerNumber.Equals(es.CustomerNumber.PadLeft(7, '0')) && sequenceNumber.Equals(es.SequenceNumber.PadLeft(2, '0')))
                                    {
                                        location.Add(i);
                                    }
                                    else if (customerNumber.Equals(es.CustomerNumber) && sequenceNumber.Equals(es.SequenceNumber))
                                    {
                                        location.Add(i);
                                    }
                                    else if (((string)jsonObject["assets"][i]["name"]).Contains(es.CustomerNumber) &&
                                    ((string)jsonObject["assets"][i]["name"]).Contains(es.SequenceNumber))
                                    {
                                        location.Add(i);
                                    }
                                    else if (!es.JobName.Contains("ft-"))
                                    {
                                        location.Add(i);
                                    }
                                }
                                Console.WriteLine("Number of Results: " + location.Count);
                                foreach (int i in location)
                                {
                                    string downloadUrl = (string)jsonObject["assets"][i]["downloadUrl"];
                                    string fileName = (string)jsonObject["assets"][i]["name"];
                                    string size = (string)jsonObject["assets"][i]["size"];
                                    //Console.WriteLine("Download Url: " + downloadUrl);
                                    if (location.Count.Equals(1))
                                    {
                                        Directory.CreateDirectory(Path.Combine(outputFolder, es.DelVeh.ToUpper().Trim()));
                                        FileDownload(Path.Combine(outputFolder, es.DelVeh.ToUpper().Trim()), downloadUrl, fileName, size, es.JobName);
                                    }
                                    else
                                    {
                                        if (location.Count > 1)
                                        {
                                            /*Console.WriteLine("ERROR: There are multiple results (" + results + ") for the search (" + search + ")...");
                                            using (StreamWriter batchFile = new StreamWriter(Path.Combine(errorFolder, es.JobName + ".txt"), true))
                                            {
                                                batchFile.WriteLine("ERROR: There are multiple results (" + results + ") for the search (" + search.Trim() + ")...");
                                            }*/

                                            //File.Move(filePath, Path.Combine(errorFolder, Path.GetFileName(filePath)));
                                            UrlModel urlModel = new UrlModel();
                                            urlModel.FileName = fileName;
                                            urlModel.URL = downloadUrl;
                                            urlModel.Size = size;
                                            urlList.Add(urlModel);
                                        }
                                    }
                                }
                                if (location.Count < 1)
                                {
                                    Console.WriteLine("ERROR: There are no results for the search (" + search.Trim() + ")...");
                                    using (StreamWriter batchFile = new StreamWriter(Path.Combine(settings.ErrorFolder, es.JobName + ".txt"), true))
                                    {
                                        batchFile.WriteLine("ERROR: There are no results for the search (" + search.Trim() + ")...");
                                    }
                                }
                                if (urlList.Count > 1)
                                {
                                    Console.WriteLine("ERROR: There are multiple results (" + urlList.Count + ") for the search (" + search.Trim() + ")...");
                                    using (StreamWriter batchFile = new StreamWriter(Path.Combine(settings.ErrorFolder, es.JobName + ".txt"), true))
                                    {
                                        batchFile.WriteLine("ERROR: There are multiple results (" + urlList.Count + ") for the search (" + search.Trim() + ")...");
                                    }
                                    foreach (UrlModel um in urlList)
                                    {
                                        using (StreamWriter batchFile = new StreamWriter(Path.Combine(settings.ErrorFolder, es.JobName + ".txt"), true))
                                        {
                                            batchFile.WriteLine("--" + um.FileName + "...");
                                        }
                                    }
                                    urlList = new List<UrlModel>();
                                }
                                /*if (!urlList.Count.Equals(0))
                                {
                                    int urlCount = 0;
                                    Console.WriteLine("There are " + urlList.Count + " to choose from. Type the number of the one you want to download.");
                                    foreach (UrlModel um in urlList)
                                    {
                                        Console.WriteLine("[" + urlCount + "] - " + um.FileName);
                                        urlCount++;
                                    }
                                    int urlChoice = 0;
                                    bool parseOk = false;
                                    while (!parseOk)
                                    {
                                        parseOk = int.TryParse(Console.ReadLine(), out urlChoice);
                                        if (urlChoice > urlList.Count - 1 || !parseOk)
                                        {
                                            parseOk = false;
                                            Console.WriteLine("Pick a number within range.");
                                        }
                                    }
                                    FileDownload(downloadLoc, urlList[urlChoice].URL, urlList[urlChoice].FileName, urlList[urlChoice].Size);
                                    /*using (var webClient = new WebClient())
                                    {
                                        Console.WriteLine("Download started for " + urlList[urlChoice].FileName);
                                        Console.WriteLine("File size: " + urlList[urlChoice].Size);
                                        webClient.DownloadFile(urlList[urlChoice].URL, Path.Combine(downloadLoc, urlList[urlChoice].FileName));
                                        Console.WriteLine("File Download Finished.");
                                        Console.WriteLine("Starting Zip Extration...");
                                        Run(Path.Combine(downloadLoc, urlList[urlChoice].FileName), downloadLoc, true);
                                        Console.WriteLine("Zip Extration Finished...");
                                        Console.WriteLine("-----------------------------------------");

                                    }*/
                                //}
                            }
                            client.Dispose();
                        }
                        else
                        {
                            Console.WriteLine(string.Format("{0}:{1}", response.StatusCode, response.ReasonPhrase));
                            using (StreamWriter batchFile = new StreamWriter(Path.Combine(settings.ErrorFolder, es.JobName + ".txt"), true))
                            {
                                batchFile.WriteLine(string.Format("{0}:{1}...", response.StatusCode, response.ReasonPhrase));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            
        }

        private static void FileDownload(string outputFolder, string downloadUrl, string fileName, string size, string jobName = "")
        {
            using (var webClient = new WebClient())
            {
                Console.WriteLine("-----------------------------------------");
                Console.WriteLine("Download started for " + fileName);
                Console.WriteLine("File size: " + size);
                webClient.DownloadFile(downloadUrl, Path.Combine(outputFolder, fileName));
                Console.WriteLine("File Download Finished.");
                Console.WriteLine("Starting Zip Extration...");
                Run(Path.Combine(outputFolder, fileName), outputFolder, true, jobName);
                Console.WriteLine("Zip Extration Finished...");
                Console.WriteLine("-----------------------------------------");
            }
        }

        /// <summary>
        /// Runs the main process
        /// </summary>
        /// <param name="fileName">The zip file to process</param>
        /// <param name="to">Where to export the pdfs</param>
        /// <param name="deleteZips">Shoudl the zips be deleted?</param>
        private static void Run(string fileName, string to, bool deleteZips, string jobName)
        {
            SettingsModel settings = GetSettings();
            try
            {
                //Console.WriteLine("Extracting " + fileName + "...");
                UnzipNew(fileName, Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)));
                //Console.WriteLine("-------------------------------------------------------------");
                List<string> files = new List<string>(Directory.GetFiles(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)), "*.pdf", SearchOption.AllDirectories));
                Directory.CreateDirectory(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)));
                int count = 1;
                foreach (string f in files)
                {
                    string delVehs = to;
                    foreach (DelVehs d in settings.DelVehList)
                    {
                        if (d.Id.ToLower().Trim().Equals(Path.GetFileName(to).ToLower().Trim()))
                        {
                            delVehs = d.Path;
                        }
                    }
                    if (jobName.Trim().Equals(""))
                    {
                        File.Copy(f, Path.Combine(delVehs, Path.GetFileName(f)), true);
                    }
                    else if (files.Count < 2)
                    {
                        File.Copy(f, Path.Combine(delVehs, jobName + Path.GetExtension(f)), true);
                    }
                    else
                    {
                        File.Copy(f, Path.Combine(to, jobName + " - " + count + Path.GetExtension(f)), true);
                        Console.WriteLine("Multiple pdf files for " + jobName);
                        count++;
                    }
                    //File.Copy(f, Path.Combine(to, Path.GetFileName(f)), true);
                    //File.Decrypt(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName), Path.GetFileName(f)));
                }
                CleanUp(fileName, to, deleteZips);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// T   Unzips the contents to get the pdfs.
        /// </summary>
        /// <param name="sourceFile">The Zip file</param>
        /// <param name="destination">Where the pdfs need to go.</param>
        private static void UnzipNew(string sourceFile, string destination)
        {
            List<int> linkRemoveList = new List<int>();
            using (ZipArchive archive = ZipFile.OpenRead(sourceFile))
            {
                int zipCount = 0;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    //if (archive.Entries[i].FullName.ToLower().Contains("links"))
                    if (archive.Entries[i].FullName.Any(p => Path.GetDirectoryName(archive.Entries[i].FullName) != Path.GetFileNameWithoutExtension(sourceFile)))
                    {
                        linkRemoveList.Add(i);
                    }
                }
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        zipCount++;
                    }
                }
                int count = 0;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        if ((!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).Contains("._")) && (!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).StartsWith("_")))
                        {
                            count++;
                        }
                    }
                }
                int fileCount = 1;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        string filename = Path.GetFileName(sourceFile);
                        _ = filename.Replace("-", " ");
                        string destinationPath;

                        if ((!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).Contains("._")) && (!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).StartsWith("_")))
                        {
                            if (count < 2)
                            {
                                destinationPath = Path.GetFullPath(Path.Combine(destination, Path.GetFileNameWithoutExtension(filename) + ".pdf"));
                            }
                            else
                            {
                                destinationPath = Path.GetFullPath(Path.Combine(destination, Path.GetFileNameWithoutExtension(filename) + " - " + fileCount.ToString().PadLeft(3, '0') + ".pdf"));
                                fileCount++;
                            }
                            Directory.CreateDirectory(destination);
                            archive.Entries[i].ExtractToFile(destinationPath, true);
                        }

                    }
                }
            }
        }

        /// <summary>
        ///     Delete any left over files
        /// </summary>
        /// <param name="fileName">The zip file name</param>
        /// <param name="to">Where the pdfs were extracted to</param>
        /// <param name="deleteZips">Should the Zips be Deleted?</param>
        private static void CleanUp(string fileName, string to, bool deleteZips)
        {
            try
            {
                bool go = false;
                while (!go)
                {
                    FileInfo dn = new FileInfo(fileName);
                    if (!IsFileLocked(dn))
                    {
                        Directory.Delete(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)), true);
                        if (deleteZips)
                        {
                            File.Delete(fileName);
                        }
                        go = true;
                    }
                    if (IsDirectoryEmpty(to))
                    {
                        Directory.Delete(to);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }

        /// <summary>
        ///     Checks to see if a file is locked
        /// </summary>
        /// <param name="file">The file name.</param>
        /// <returns>Returned true or false depending on if the file is locked.</returns>
        private static bool IsFileLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }

        /// <summary>
        ///     Gets the info from a excel cell.
        /// </summary>
        /// <param name="sheet">The sheet to check in the excel file.</param>
        /// <param name="r">The row.</param>
        /// <param name="c">The Column.</param>
        /// <returns></returns> 
        private static string GetCell(ISheet sheet, int r, int c)
        {
            string value = "";
            try
            {
                IRow row = sheet.GetRow(r);
                ICell cell = CellUtil.GetCell(row, c);
                cell.SetCellType(CellType.String);
                value = cell.StringCellValue;
            }
            catch (Exception e) { }

            return value.Trim();
        }

        /// <summary>
        ///     Checks to see if a cell is blank.
        /// </summary>
        /// <param name="sheet">The sheet to check in the excel files.</param>
        /// <param name="r">The row.</param>
        /// <param name="c">The cell.</param>
        /// <returns>Returns a bool of whether the cell is empty or not.</returns>
        private static bool IsCellBlank(ISheet sheet, int r, int c)
        {
            bool isEmpty = false;
            string value = GetCell(sheet, r, c);

            if (value.Trim().Equals(""))
            {
                isEmpty = true;
            }
            if (value.ToLower().Trim().Contains("when"))
            {
                isEmpty = true;
            }

            return isEmpty;
        }

        /// <summary>
        /// Gets the number of rows in the spreadsheet.
        /// </summary>
        /// <param name="sheet">The sheet to check</param>
        /// <returns>Returns the number of rows.</returns>
        public static int GetRowCount(ISheet sheet)
        {
            int count = 0;
            while (!IsCellBlank(sheet, count, 0))
            {
                count++;
            }
            return count;
        }

        public static string osXPathConversion(string path)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                if (!path.StartsWith("/"))
                {
                    path = "/Volumes" + path.Replace(path.Substring(0, path.IndexOfNth("\\", 2)), "").Replace(":", "/").Replace("\\", "/");
                }
            }

            return path;
        }

        public static SettingsModel GetSettings()
        {
            SettingsModel settings = new();

            XmlDocument doc = new XmlDocument();
            string xmlPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Config.xml");
            doc.Load(xmlPath);

            XmlNode WidenAPINode = doc.DocumentElement.SelectSingleNode("WidenAPISettings");
            settings.BaseUrl = WidenAPINode.SelectSingleNode("BaseUrl").InnerText;
            settings.APIToken = WidenAPINode.SelectSingleNode("APIToken").InnerText;

            XmlNode mainSettingNode = doc.DocumentElement.SelectSingleNode("MainSettings");
            settings.ErrorFolder = osXPathConversion(mainSettingNode.SelectSingleNode("ErrorFolder").InnerText);
            settings.IncomingFolder = osXPathConversion(mainSettingNode.SelectSingleNode("IncomingFolder").InnerText);
            settings.OutputFolder = osXPathConversion(mainSettingNode.SelectSingleNode("OutputFolder").InnerText);
            settings.DoneFolder = osXPathConversion(mainSettingNode.SelectSingleNode("DoneFolder").InnerText);

            XmlNode delVehNode = doc.DocumentElement.SelectSingleNode("DelVehs");
            foreach (XmlNode delNode in delVehNode.ChildNodes)
            {
                DelVehs delVehs = new();
                delVehs.Id = delNode.Attributes["id"].Value;
                delVehs.Path = osXPathConversion(delNode.ChildNodes[0].InnerText);
                settings.DelVehList.Add(delVehs);
                delVehs = new();
            }

            return settings;
        }
    }

    class SettingsModel
    {
        public string BaseUrl { get; set; }
        public string APIToken { get; set; }
        public string ErrorFolder { get; set; }
        public string IncomingFolder { get; set; }
        public string OutputFolder { get; set; }
        public string DoneFolder { get; set; }
        public List<DelVehs> DelVehList { get; set; } = new();
    }
    class DelVehs
    {
        public string Id { get; set; }
        public string Path { get; set; }
    }

    class UrlModel
    {
        public string FileName { get; set; }
        public string URL { get; set; }
        public string Size { get; set; }
    }

    class ExcelSheetModel
    {
        public string StyleNumber { get; set; }
        public string StyleName { get; set; }
        public string DelVeh { get; set; }
        public string CustomerNumber { get; set; }
        public string SequenceNumber { get; set; }
        public string BoardSku { get; set; }
        public string StampingInstructions { get; set; }
        public string FirstColor { get; set; }
        public string NumOrders { get; set; }
        public string TotalQty { get; set; }
        public string JobName { get; set; }
    }
    public static class Extensions
    {
        public static int IndexOfNth(this string str, string value, int nth)
        {
            if (nth < 0)
                throw new ArgumentException("Can not find a negative index of substring in string. Must start with 0");

            int offset = str.IndexOf(value);
            for (int i = 0; i < nth; i++)
            {
                if (offset == -1) return -1;
                offset = str.IndexOf(value, offset + 1);
            }

            return offset;
        }
    }
}
