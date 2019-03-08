#region Using directives

using System.Collections.Generic;
using System.IO;
using System.Text;
using System;
using System.Windows.Forms;
using MongoDB.Driver;
using MongoDB.Bson;

#endregion

namespace ExcelApplication1
{
    public class Worker
    {

        public static bool merge(string selectedPath, List<ExcelFile> excelFileList,Settings settings ,System.ComponentModel.BackgroundWorker backgroundWorker)
        {
            // Start the search for primes and wait.
            UTF8Encoding utf8 = new UTF8Encoding(false);
            var targetPath = settings.TargetPath;
            var writer = new StreamWriter(targetPath, false, utf8);

            try
            {
                int i = 0;
                foreach (var f in excelFileList)
                {
                    i++;
                    if (backgroundWorker.CancellationPending)
                    {
                        // Return without doing any more work.
                        throw new Exception("用户取消了操作");
                    }

                    if (backgroundWorker.WorkerReportsProgress)
                    {
                        backgroundWorker.ReportProgress(i);
                    }

                    if (!f.IsSelected) continue;
                    f.Status = "处理中";
                    var stream = new FileStream(f.Path, FileMode.Open);
                    StreamReader txtReader = new StreamReader(stream, utf8);
                    try
                    {

                        string nextLine;                        
                        while ((nextLine = txtReader.ReadLine()) != null)
                        {
                            //虚拟卡编号,虚拟卡密码,开始有效期,截止有效期,活动名称,奖品名称,是否使用,活动编号
                            //var _id =ObjectId.GenerateNewId().ToString();// 序号
                            var code = nextLine.Trim();
                            //var pwd = "";
                            var start_time = settings.StartTime;
                            var end_time = settings.EndTime;
                            var prize_id = settings.PrizeId;
                            var is_used = settings.IsUsed ? 1 : 0;

                            DateTime now = DateTime.Now;

                            //按yyyy-MM-dd HH:mm:ss格式输出s
                            var __CREATE_TIME__ = now.ToString("yyyy-MM-dd HH:mm:ss");
                            var __MODIFY_TIME__ = now.ToString("yyyy-MM-dd HH:mm:ss");

                            if (!string.IsNullOrEmpty(code))
                            {
                                //writer.WriteLine("insert  into `iprize_code`(`_id`,`prize_id`,`code`,`pwd`,`is_used`,`start_time`,`end_time`,`__CREATE_TIME__`,`__MODIFY_TIME__`,`__REMOVED__`) values ('{0}','{1}','{2}','{3}',{4},'{5}','{6}','{7}','{8}',0);", _id, prize_id, code, pwd, is_used, start_time, end_time, __CREATE_TIME__, __MODIFY_TIME__);
                                writer.WriteLine("insert  into `luckydraw_virtual_code`(`gift_id`,`code`,`is_activited`,`is_used`,`start_at`,`end_at`,`use_start_at`,`use_end_at`,`created_at`,`updated_at`) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}');", prize_id, code, 1, is_used, start_time, end_time, start_time, end_time, __CREATE_TIME__, __MODIFY_TIME__);
                            }
                        }

                        f.Status = "处理完成";
                    }
                    catch (Exception e1)
                    {
                        f.Status = "处理失败";
                        throw e1;
                    }
                    finally
                    {;
                        txtReader.Close();
                    }
                }
                
                return true;
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                throw e;
                //return false;
            }
            finally
            {
                writer.Flush();
                writer.Close();
            }
        }


        public static bool randomCreate(Settings settings, System.ComponentModel.BackgroundWorker backgroundWorker)
        {
            // Start the search for primes and wait.
            UTF8Encoding utf8 = new UTF8Encoding(false);
            var targetPath = settings.TargetPath;
            var writer = new StreamWriter(targetPath, false, utf8);
            
           
            try
            {
                int i = 0;
                for (int s = 0; s < settings.Quantity; s++)
                {
                    i++;
                    if (backgroundWorker.CancellationPending)
                    {
                        // Return without doing any more work.
                        throw new Exception("用户取消了操作");
                    }

                    if (backgroundWorker.WorkerReportsProgress)
                    {
                        backgroundWorker.ReportProgress(i);
                    }
                    
                    try
                    {
                        //虚拟卡编号,虚拟卡密码,开始有效期,截止有效期,活动名称,奖品名称,是否使用,活动编号
                        //var _id =ObjectId.GenerateNewId().ToString();// 序号
                        var code = GenerateCheckCode(settings.CodeLength);
                        //var pwd = "";
                        var start_time = settings.StartTime;
                        var end_time = settings.EndTime;
                        var prize_id = settings.PrizeId;
                        var is_used = settings.IsUsed ? 1 : 0;

                        DateTime now = DateTime.Now;

                        //按yyyy-MM-dd HH:mm:ss格式输出s
                        var __CREATE_TIME__ = now.ToString("yyyy-MM-dd HH:mm:ss");
                        var __MODIFY_TIME__ = now.ToString("yyyy-MM-dd HH:mm:ss");

                        if (!string.IsNullOrEmpty(code))
                        {
                            //writer.WriteLine("insert  into `iprize_code`(`_id`,`prize_id`,`code`,`pwd`,`is_used`,`start_time`,`end_time`,`__CREATE_TIME__`,`__MODIFY_TIME__`,`__REMOVED__`) values ('{0}','{1}','{2}','{3}',{4},'{5}','{6}','{7}','{8}',0);", _id, prize_id, code, pwd, is_used, start_time, end_time, __CREATE_TIME__, __MODIFY_TIME__);
                            writer.WriteLine("insert  into `luckydraw_virtual_code`(`gift_id`,`code`,`is_activited`,`is_used`,`start_at`,`end_at`,`use_start_at`,`use_end_at`,`created_at`,`updated_at`) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}');", prize_id, code, 1, is_used, start_time, end_time, start_time, end_time, __CREATE_TIME__, __MODIFY_TIME__);
                        }
                    }
                    catch (Exception e1)
                    {
                        throw e1;
                    }
                    finally
                    {
                    }
                }

                return true;
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                throw e;
                //return false;
            }
            finally
            {
                writer.Flush();
                writer.Close();
            }
        }



        //方法一：随机生成不重复数字字符串
        private static int rep = 0;

        /// 
        /// 生成随机数字字符串
        /// 
        /// 待生成的位数
        /// 生成的数字字符串
        private static string GenerateCheckCodeNum(int codeCount)
        {
            string str = string.Empty;
            long num2 = DateTime.Now.Ticks + Worker.rep;
            Worker.rep++;
            Random random = new Random(((int)(((ulong)num2) & 0xffffffffL)) | ((int)(num2 >> Worker.rep)));
            for (int i = 0; i < codeCount; i++)
            {
                int num = random.Next();
                str = str + ((char)(0x30 + ((ushort)(num % 10)))).ToString();
            }
            return str;
        }

        //方法二：随机生成字符串（数字和字母混和） 
 /// 
       /// 生成随机字母字符串(数字字母混和)
       /// 
       /// 待生成的位数
       /// 生成的字母字符串
       private static string GenerateCheckCode(int codeCount)
        {
            string str = string.Empty;
            long num2 = DateTime.Now.Ticks + Worker.rep;
            Worker.rep++;
            Random random = new Random(((int)(((ulong)num2) & 0xffffffffL)) | ((int)(num2 >> Worker.rep)));
            for (int i = 0; i < codeCount; i++)
            {
                char ch;
                int num = random.Next();
                if ((num % 2) == 0)
                {
                    ch = (char)(0x30 + ((ushort)(num % 10)));
                }
                else
                {
                    ch = (char)(0x41 + ((ushort)(num % 0x1a)));
                }
                str = str + ch.ToString();
            }
            return str;
        }

        //方法三、
         private static char[] constant =
        {
        '0','1','2','3','4','5','6','7','8','9',
        'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
        'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
        };

        public static string GenerateRandomNumber(int Length)
        {
            System.Text.StringBuilder newRandom = new System.Text.StringBuilder(62);
            Random rd = new Random();
            for (int i = 0; i < Length; i++)
            {
                newRandom.Append(constant[rd.Next(62)]);
            }
            return newRandom.ToString();
        }


    }
}
