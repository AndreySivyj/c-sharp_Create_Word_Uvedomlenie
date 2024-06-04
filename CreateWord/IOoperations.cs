using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace CreateWord
{
    static class IOoperations
    {
        public const string katalogIn = @"C:\_CMD_\_In_WordDot_";                     //каталог для обрабатываемых файлов
        //public const string katalogInUPP = @"C:\_CMD_\_In_UPP_";                     //каталог для обрабатываемых файлов

        //public const string katalogIn = @"_In_";                                //каталог для обрабатываемых файлов
        //public const string katalogOut = @"_Out_";                              //каталог для результирующих файлов

        private static string errorLog = @"errorLog.txt";      //лог с ошибками обработки         

        //------------------------------------------------------------------------------------------
        //создаем каталог
        public static void DirectoryCreater(string createDirectoryName)
        {
            //Создаем пустой каталог
            if (!Directory.Exists(createDirectoryName))
                Directory.CreateDirectory(createDirectoryName);
        }

        //------------------------------------------------------------------------------------------
        //удаляем каталог
        private static void DirectoryDelete(string deleteDirectoryName)
        {
            try
            {
                //Удаляем каталог со всем содержимым 
                if (Directory.Exists(deleteDirectoryName))
                    Directory.Delete(deleteDirectoryName, true);
            }
            catch (IOException ex)
            {
                WriteLogError(ex.ToString());

                Console.WriteLine();
                Console.WriteLine(new string('-', 17));
                Console.WriteLine("Внимание! Ошибка достаупа к каталогу \"{0}\" .", deleteDirectoryName);
                Console.WriteLine(ex.ToString());
                Console.WriteLine(new string('-', 17));
            }
            catch (Exception ex)
            {
                WriteLogError(ex.ToString());

                Console.WriteLine();
                Console.WriteLine(new string('-', 17));
                Console.WriteLine("Внимание! Ошибка достаупа к каталогу \"{0}\" .", deleteDirectoryName);
                Console.WriteLine(ex.ToString());
                Console.WriteLine(new string('-', 17));
            }
        }

        //------------------------------------------------------------------------------------------
        //Создаем каталоги по умолчанию, очищаем временные каталоги
        public static void BasicDirectoryAndFileCreate()
        {            
            DirectoryCreater(katalogIn);
            //DirectoryCreater(katalogInUPP);
            
            //DirectoryDelete(katalogOut);
            //DirectoryCreater(katalogOut);
        }

        //------------------------------------------------------------------------------------------       
        //Пишем ошибки в лог-файл, по умолчанию @"errorLog.txt"
        public static void WriteLogError(string errormessage)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(errorLog, true, Encoding.GetEncoding(1251)))
                {
                    writer.WriteLine(new string('-', 17));
                    writer.WriteLine(DateTime.Now);
                    writer.WriteLine(new string('-', 17));
                    writer.WriteLine(errormessage);
                    writer.WriteLine(new string('-', 17));
                }

            }
            catch (IOException ex)
            {
                Console.WriteLine();
                Console.WriteLine(new string('-', 17));
                Console.WriteLine("Внимание! Ошибка достаупа к лог-файлу \"errorLog.txt\"");
                Console.WriteLine(ex.ToString());
                Console.WriteLine(new string('-', 17));
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine(new string('-', 17));
                Console.WriteLine("Внимание! Ошибка достаупа к лог-файлу \"errorLog.txt\"");
                Console.WriteLine(ex.ToString());
                Console.WriteLine(new string('-', 17));
            }
        }


        //------------------------------------------------------------------------------------------        
        //Формируем результирующий файл из результатов запросов к БД
        public static void CreateExportFile(string zagolovok, IEnumerable<string> listData, string nameFile)
        {
            try
            {
                //Добавляем в файл данные                
                using (StreamWriter writer = new StreamWriter(nameFile, true, Encoding.GetEncoding(1251)))
                {
                    writer.WriteLine(zagolovok);

                    foreach (string item in listData)
                    {
                        writer.WriteLine(item.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogError(ex.ToString());
            }
        }

    }
}
