using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Threading;
using System.Configuration;
using System.Collections.Specialized;

namespace CreateWord
{
    class Program
    {
        public static string regNum;

        public static string fioKurator;
        public static string phoneKurator;

        private static DateTime start;

        static void Main(string[] args)
        {


            NameValueCollection allAppSettings = ConfigurationManager.AppSettings;              //формируем массив настроек приложения

            Program.fioKurator = allAppSettings["fioKurator"];         //fioKurator
            Program.phoneKurator = allAppSettings["phoneKurator"];  //phoneKurator 

            //время начала обработки
            start = DateTime.Now;

            try
            {
                //Console.SetWindowSize(125, 55);  //Устанавливаем размер окна консоли
               
                //------------------------------------------------------------------------------------------
                //1. Создаем каталоги по умолчанию
                IOoperations.BasicDirectoryAndFileCreate();



                //------------------------------------------------------------------------------------------
                //2. Создаем каталоги по умолчанию
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(new string('-', 71));
                Console.WriteLine("Введите необходимые параметры:");
                Console.WriteLine(new string('-', 71));
                Console.ForegroundColor = ConsoleColor.Gray;

                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.Write("Введите рег. номер страхователя (42001000000): ");
                Console.ForegroundColor = ConsoleColor.Gray;

                string tmp = Console.ReadLine();
                if (tmp.Count() == 11)
                {
                    Program.regNum = tmp;
                }
                else
                {
                    Program.regNum = "42000000000";
                }



                //------------------------------------------------------------------------------------------
                //3. Выбираем данные для формирования уведомления

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(new string('-', 71));
                Console.WriteLine("Выбираем данные из БД для формирования уведомления.");
                Console.WriteLine(new string('-', 71));
                Console.ForegroundColor = ConsoleColor.Gray;

                //выбираем данные для формирования уведомления
                SelectDataFromDB.SelectDataFromRKASV(Program.regNum);

                                

                //------------------------------------------------------------------------------------------
                //4. Создаем файлы уведомлений на основании шаблона Word

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(new string('-', 71));
                Console.WriteLine("Создаем файл уведомления на основании шаблона Word.");
                Console.WriteLine(new string('-', 71));
                Console.ForegroundColor = ConsoleColor.Gray;


                // создаем путь к файлу шаблона Word
                Object templatePathObj = IOoperations.katalogIn + @"\" + @"Уведомление об ошибках.dotx";
                
                CreateWord.FindAndReplase(templatePathObj, SelectDataFromDB.dictionaryDataUPFR[Program.regNum]);


                //string strToFind1 = "nameStrah"; //строка для поиска
                //string replaceStr2 = "OOO TEST"; //строка для замены

                //CreateWord.FindAndReplase(templatePathObj, strToFind1, replaceStr2);

                Console.WriteLine();
            }
            catch (Exception ex)
            {
                IOoperations.WriteLogError(ex.ToString());

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(
                        Environment.NewLine
                        + "Внимание! При обработке файлов возникли ошибки."
                        + Environment.NewLine + Environment.NewLine
                        + "Дополнительня информация отражена в файле errorLog.txt");
                Console.ForegroundColor = ConsoleColor.Gray;

                //throw ex;

                //Задержка экрана
                Thread.Sleep(TimeSpan.FromSeconds(3));
            }


            //вычисляем время затраченное на обработку
            TimeSpan stop = DateTime.Now - start;

            //Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(new string('-', 71));
            Console.WriteLine("Обработка выполнилась за " + stop.Seconds + " сек.");
            Console.ForegroundColor = ConsoleColor.Gray;

            //Console.ReadKey();

            //Задержка экрана
            //Thread.Sleep(TimeSpan.FromSeconds(5));



        }




    }
}
