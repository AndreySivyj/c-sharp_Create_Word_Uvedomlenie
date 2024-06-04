using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Threading;

namespace CreateWord
{
    static class CreateWord
    {
        /// <summary>
        /// Поиск и замена строки в шаблоне Word 
        /// </summary>
        /// <param name="templatePathObj">путь к файлу</param>
        /// <param name="strToFind">строка для поиска</param>
        /// <param name="replaceStr">строка для замены</param>
        public static void FindAndReplase(Object templatePathObj, DataUPFR tmpData)
        {
            try
            {
                Word._Application application;
                Word._Document document = new Word.Document();

                Object missingObj = System.Reflection.Missing.Value;
                Object trueObj = true;
                Object falseObj = false;

                //создаем обьект приложения word
                application = new Word.Application();


                // если вылетим не этом этапе, приложение останется открытым
                try
                {
                    document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine();
                    Console.WriteLine("Ошибка доступа к файлу шаблона Word.");
                    Console.ForegroundColor = ConsoleColor.Gray;

                    document.Close(ref falseObj, ref missingObj, ref missingObj);
                    application.Quit(ref missingObj, ref missingObj, ref missingObj);
                    document = null;
                    application = null;
                    IOoperations.WriteLogError(ex.ToString());
                    //throw ex;
                }
                application.Visible = true;

                //меняем значения в шаблоне Word
                ReplaseText(document, "namePFR", tmpData.namePFR);
                ReplaseText(document, "adressPFR", tmpData.adressPFR);
                ReplaseText(document, "phonePFR", tmpData.phonePFR);
                ReplaseText(document, "nameStrah", tmpData.nameStrah);
                ReplaseText(document, "regNumStrah", tmpData.regNumStrah);
                ReplaseText(document, "innStrah", tmpData.innStrah);
                ReplaseText(document, "kppStrah", tmpData.kppStrah);
                ReplaseText(document, "adressStrah", tmpData.adressStrah);
                ReplaseText(document, "dataUvedoml", tmpData.dataUvedoml);
                //ReplaseText(document, "fioKurator", tmpData.fioKurator);
                //ReplaseText(document, "phoneKurator", tmpData.phoneKurator);
            }
            catch (Exception ex)
            {
                IOoperations.WriteLogError(ex.ToString());

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
                Console.ForegroundColor = ConsoleColor.Gray;

                //Задержка экрана
                Thread.Sleep(TimeSpan.FromSeconds(3));
            }
        }

        private static void ReplaseText(Word._Document document, string strToFind, string replaceStr)
        {
            try
            {
                Object missingObj = System.Reflection.Missing.Value;
                Object trueObj = true;
                Object falseObj = false;

                // обьектные строки для Word
                object strToFindObj = strToFind;
                object replaceStrObj = replaceStr;

                // диапазон документа Word
                Word.Range wordRange;

                //тип поиска и замены
                object replaceTypeObj;
                replaceTypeObj = Word.WdReplace.wdReplaceAll;

                // обходим все разделы документа
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    // берем всю секцию диапазоном
                    wordRange = document.Sections[i].Range;

                    /*
                    Обходим редкий глюк в Find, ПРИЗНАННЫЙ MICROSOFT, метод Execute на некоторых машинах вылетает с ошибкой "Заглушке переданы неправильные данные / Stub received bad data"  Подробности: http://support.microsoft.com/default.aspx?scid=kb;en-us;313104
                    // выполняем метод поиска и  замены обьекта диапазона ворд
                    wordRange.Find.Execute(ref strToFindObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref replaceStrObj, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                    */

                    Word.Find wordFindObj = wordRange.Find;
                    object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };

                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
                }
            }
            catch (Exception ex)
            {
                IOoperations.WriteLogError(ex.ToString());

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
                Console.ForegroundColor = ConsoleColor.Gray;
                
                //Задержка экрана
                Thread.Sleep(TimeSpan.FromSeconds(3));
            }
        }

        
    }
}
