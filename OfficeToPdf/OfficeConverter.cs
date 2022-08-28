using System;
using System.Diagnostics;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NLog;

namespace OfficeToPdf
{
    public class OfficeConverter : IOfficeConverter
    {
        IFileOperationManager fileOperation;
        ILogger logger;

        public OfficeConverter()
        {
            fileOperation = new GridFSOperationManager();
            logger = LogManager.GetCurrentClassLogger();
        }

        private bool SaveToFile(byte[] sourceBuffer, string path)
        {
            try
            {
                logger.Info($"stream.length:{sourceBuffer.Length}");

                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(sourceBuffer, 0, sourceBuffer.Length);
                    fs.Close();
                    logger.Info("读取源文件到本地成功");
                }
            }
            catch (Exception ex)
            {
                logger.Error("读取源文件到本地失败：" + ex.Message);
                return false;
            }
            
            return true;
        }

        private FileInfo UploadFile(string path,string destName)
        {
            if (!File.Exists(path))
                return null;

            byte[] bytes = null;
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                bytes = new byte[fileStream.Length];
                fileStream.Read(bytes, 0, bytes.Length);
            }

            Stream stream = new MemoryStream(bytes);
            return fileOperation.AddFile(stream, destName, true);
        }

        private bool AddAttachmentAssociate(string fileId, string destFileId)
        {
            try
            {
                string host = EnvironmentHelper.GetEnvValue("ApiHost");
                string api = EnvironmentHelper.GetEnvValue("AssociationApi");
                if (string.IsNullOrEmpty(api))
                {
                    logger.Warn("请检查 AssociationApi 环境变量的配置");
                    return false;
                }
                if (string.IsNullOrEmpty(host))
                {
                    logger.Warn("请检查 ApiHost 环境变量的配置");
                    return false;
                }
                string result = APIHelper.RunApiGet(host, $"{api}/{fileId}/{destFileId}");
            }
            catch (Exception ex)
            {
                logger.Error("关联异常：" + ex.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// excel设置 打印缩放比例
        /// </summary>
        /// <param name="sourceStream"></param>
        /// <param name="outPath"></param>
        /// <returns></returns>
        public bool SetExcelScale(Stream sourceStream, String outPath)
        {
            //读取excel文件
            IWorkbook workbook = null;
            try
            {
                string extension = Path.GetExtension(outPath);
                if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    NPOIFSFileSystem fs = new NPOIFSFileSystem(sourceStream);
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook = new XSSFWorkbook(sourceStream);
                }
            }
            catch (FileNotFoundException e)
            {
                logger.Error("setExcelScale fail: 源文件不存在", e);
                return false;
            }
            catch (IOException e)
            {
                logger.Error("setExcelScale fail: 读取源文件IO异常", e);
                return false;
            }
            try
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {   //获取每个Sheet表
                    ISheet sheet = workbook.GetSheetAt(i);

                    //打印设置
                    sheet.PrintSetup.NoOrientation = false;

                    // 打印方向，true：横向，false：纵向(默认)
                    sheet.PrintSetup.Landscape = true;

                    //设置高度为自动分页
                    sheet.PrintSetup.FitHeight = 0;

                    //设置宽度为一页
                    sheet.PrintSetup.FitWidth = 1;

                    //纸张类型
                    //print.PaperSize = ePaperSize.A4;
                    sheet.PrintSetup.PaperSize = 9;

                    //print.setScale((short)55);//自定义缩放①，此处100为无缩放
                    //启用“适合页面”打印选项的标志
                    sheet.PrintSetup.UsePage = true;
                }
                using (var fs = new MemoryStream())
                {
                    workbook.Write(fs);
                    // Excel文件生成后存储的位置。
                    return SaveToFile(fs.ToArray(), outPath);
                }
            }
            catch (Exception e)
            {
                logger.Error("setExcelScale fail: 创建输出文件IO异常");
                return false;
            }
        }

        public bool OnWork(Messages message)
        {
            LibreOfficeConvertMessage officeMessage = (LibreOfficeConvertMessage)message;
            string sourcePath = string.Empty;
            string destPath = string.Empty;
            try
            {
                if (officeMessage == null)
                    return false;

                Stream sourceStream = fileOperation.GetFile(officeMessage.FileInfo.FileId);
                if (sourceStream == null)
                {
                    logger.Log(LogLevel.Error, $"文件ID：{officeMessage.FileInfo.FileId}，不存在");
                }

                string filename = officeMessage.FileInfo.FileId;
                string extension = System.IO.Path.GetExtension(officeMessage.FileInfo.FileName);

                sourcePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), filename + extension);
                destPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), string.Format("{0}.pdf", filename));

                logger.Log(LogLevel.Info, $"文件原路径：{sourcePath}");
                logger.Log(LogLevel.Info, $"文件目标路径：{destPath}");
                if (extension != null && (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                          extension.Equals(".xls", StringComparison.OrdinalIgnoreCase)))
                {
                    if (!SetExcelScale(sourceStream, sourcePath))
                        return false;
                }
                else
                {
                    byte[] sourceBuffer = new Byte[sourceStream.Length];
                    sourceStream.Read(sourceBuffer, 0, sourceBuffer.Length);
                    sourceStream.Seek(0, SeekOrigin.Begin);
                    if (!SaveToFile(sourceBuffer, sourcePath))
                        return false;
                }

                var psi = new ProcessStartInfo(
                        "libreoffice7.3",
                        string.Format("--invisible --convert-to pdf  {0}", filename + extension))
                    {RedirectStandardOutput = true};

                // 启动
                var proc = Process.Start(psi);
                if (proc == null)
                {
                    logger.Error("请检查 LibreOffice 是否成功安装.");
                    return false;
                }

                logger.Log(LogLevel.Info, "文件转换开始......");
                using (var sr = proc.StandardOutput)
                {
                    while (!sr.EndOfStream)
                    {
                        Console.WriteLine(sr.ReadLine());
                    }
                    if (!proc.HasExited)
                    {
                        proc.Kill();
                    }
                }

                logger.Log(LogLevel.Info, "文件转成完成");
            }
            catch (Exception ex)
            {
                logger.Error($"Error Message:{ex.Message},StackTrace:{ex.StackTrace}");
                return false;
            }
            finally
            {
                if (File.Exists(destPath))
                {
                    var destFileInfo = UploadFile(destPath,
                        string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(officeMessage.FileInfo.FileName)));

                    var empty = destFileInfo == null || string.IsNullOrEmpty(destFileInfo.FileId);
                    string result = empty
                        ? $"文件{officeMessage.FileInfo.FileName}[{officeMessage.FileInfo.FileId}],转目标pdf文件失败.."
                        : $"文件{officeMessage.FileInfo.FileName}[{officeMessage.FileInfo.FileId}],转目标pdf文件[{destFileInfo.FileId}]成功..";

                    logger.Log(LogLevel.Info, result);

                    if (!empty)
                    {
                        if (AddAttachmentAssociate(officeMessage.FileInfo.FileId, destFileInfo.FileId))
                        {
                            logger.Log(LogLevel.Info, "文件关联成功！");
                        }
                    }
                }
                else
                {
                    logger.Log(LogLevel.Info, "---------------没有找到目标路径------------------");
                }

                if (File.Exists(sourcePath))
                    System.IO.File.Delete(sourcePath);

                if (File.Exists(destPath))
                    System.IO.File.Delete(destPath);
            }

            return true;
        }
    }
}
