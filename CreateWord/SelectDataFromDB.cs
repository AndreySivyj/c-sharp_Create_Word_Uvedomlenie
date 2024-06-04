using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Data.Common;
using IBM.Data.DB2;

namespace CreateWord
{

    #region считанная из файла информация
    public class DataFromRKASVDB
    {

        public string insurer_reg_num;
        public string insurer_short_name;
        public string insurer_last_name;
        public string insurer_first_name;
        public string insurer_middle_name;
        public string insurer_inn;
        public string insurer_kpp;

        public string adressStrah;

        public string insurer_jur_zipcode;
        public string k3_kladr_name;
        public string k2_kladr_name;
        public string k1_kladr_name;
        public string k_kladr_name;
        public string insurer_jur_house;
        public string insurer_jur_building;
        public string insurer_jur_flat;


        public DataFromRKASVDB(string insurer_reg_num = "", string insurer_short_name = "",
                                string insurer_last_name = "", string insurer_first_name = "", string insurer_middle_name = "",
                                string insurer_inn = "", string kpp = "",
                                string insurer_jur_zipcode = "", string k3_kladr_name = "", string k2_kladr_name = "", string k1_kladr_name = "",
                                string k_kladr_name = "", string insurer_jur_house = "", string insurer_jur_building = "", string insurer_jur_flat = "")
        {

            this.insurer_reg_num = insurer_reg_num;
            this.insurer_short_name = insurer_short_name;
            this.insurer_last_name = insurer_last_name;
            this.insurer_first_name = insurer_first_name;
            this.insurer_middle_name = insurer_middle_name;
            this.insurer_inn = insurer_inn;
            this.insurer_kpp = kpp;

            if (insurer_last_name != "" || insurer_first_name != "" || insurer_middle_name != "")
            {
                this.insurer_short_name = insurer_last_name + " " + insurer_first_name + " " + insurer_middle_name;
            }

            this.insurer_jur_zipcode = insurer_jur_zipcode;
            this.k3_kladr_name = k3_kladr_name;
            this.k2_kladr_name = k2_kladr_name;
            this.k1_kladr_name = k1_kladr_name;
            this.k_kladr_name = k_kladr_name;
            this.insurer_jur_house = insurer_jur_house;
            this.insurer_jur_building = insurer_jur_building;
            this.insurer_jur_flat = insurer_jur_flat;

            this.adressStrah =
                this.insurer_jur_zipcode + "," +
                this.k3_kladr_name + "," +
                this.k2_kladr_name + "," +
                this.k1_kladr_name + "," +
                this.k_kladr_name + "," +
                this.insurer_jur_house + "," +
                this.insurer_jur_building + "," +
                this.insurer_jur_flat;

        }

        public override string ToString()
        {
            return 
                insurer_reg_num + ";"
                + insurer_short_name + ";"
                + insurer_inn + ";"
                + insurer_kpp + ";"
                + adressStrah + ";";
        }
    }





    public class DataUPFR
    {
        public string namePFR;
        public string adressPFR;
        public string phonePFR;
        public string nameStrah;
        public string regNumStrah;
        public string innStrah;
        public string kppStrah;
        public string adressStrah;
        public string dataUvedoml;
        public string fioKurator;
        public string phoneKurator;

        public DataUPFR(string namePFR = "", string adresPFR = "", string phonePFR = "", string nameStrah = "",
                            string regNumStrah = "", string innStrah = "", string kppStrah = "", string adressStrah = "", string dataUvedoml = "",
                            string fioKurator = "", string phoneKurator = "")
        {
            this.namePFR = namePFR;
            this.adressPFR = adresPFR;
            this.phonePFR = phonePFR;
            this.nameStrah = nameStrah;
            this.regNumStrah = regNumStrah;
            this.innStrah = innStrah;
            this.kppStrah = kppStrah;
            this.adressStrah = adressStrah;
            this.dataUvedoml = dataUvedoml;
            this.fioKurator = fioKurator;
            this.phoneKurator = phoneKurator;
        }

        public override string ToString()
        {
            return namePFR + ";" + adressPFR + ";" + phonePFR + ";" + nameStrah + ";"
                + ";" + regNumStrah + ";" + innStrah + ";" + kppStrah + ";" + adressStrah + ";"
                + ";" + dataUvedoml + ";" + fioKurator + ";" + phoneKurator + ";";
        }
    }

    #endregion

    //------------------------------------------------------------------------------------------
    #region Выбор данных из файла
    static class SelectDataFromDB
    {
        public static Dictionary<string, DataUPFR> dictionaryDataUPFR = new Dictionary<string, DataUPFR>();
        public static Dictionary<string, DataFromRKASVDB> dictionaryDataFromRKASV = new Dictionary<string, DataFromRKASVDB>();


        async public static void SelectDataFromPerso_UPFR(string regNum)
        {
            try
            {
                string query = @"SELECT* FROM ASV_TO where to_code = '" + SelectRaion(regNum) + @"' order by TO_ID;";

                using (DB2Connection connection = new DB2Connection("Server=1.1.1.1:50000;Database=asv;UID=db2inst;PWD=password;"))
                {

                    //открываем соединение
                    await connection.OpenAsync();

                    DB2Command command = connection.CreateCommand();
                    command.CommandText = query;

                    //Устанавливаем значение таймаута
                    command.CommandTimeout = 570;

                    DbDataReader reader = await command.ExecuteReaderAsync();



                    while (await reader.ReadAsync())
                    {
                        //public DataFromFile(string namePFR = "", string adresPFR = "", string phonePFR = "", string nameStrah = "",
                        //    string regNumStrah = "", string innStrah = "", string kppStrah = "", string adressStrah = "", string dataUvedoml = "",
                        //    string fioKurator = "", string phoneKurator = "")

                        dictionaryDataUPFR[regNum] = new DataUPFR(
                            reader[3].ToString(), reader[6].ToString(), reader[7].ToString(),
                            dictionaryDataFromRKASV[regNum].insurer_short_name,
                            dictionaryDataFromRKASV[regNum].insurer_reg_num,
                            dictionaryDataFromRKASV[regNum].insurer_inn,
                            dictionaryDataFromRKASV[regNum].insurer_kpp,
                            dictionaryDataFromRKASV[regNum].adressStrah,
                            DateTime.Now.ToShortDateString(),
                            Program.fioKurator,
                            Program.phoneKurator
                            );

                    }
                    reader.Close();



                }
            }
            catch (Exception ex)
            {
                IOoperations.WriteLogError(ex.ToString());

                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }



        async public static void SelectDataFromRKASV(string regNum)
        {


            string query =
                @"select a.insurer_reg_num, a.insurer_short_name, a.insurer_last_name, a.insurer_first_name, a.insurer_middle_name, " +
                @"a.insurer_inn, a.insurer_kpp, " +
                @"a.insurer_jur_zipcode, k3.kladr_name, k2.kladr_name, k1.kladr_name, k.kladr_name, a.insurer_jur_house, a.insurer_jur_building, a.insurer_jur_flat " +
                @"from(select* FROM asv_insurer) a " +
                @"left join(select category_id, category_code from asv_category) b on a.category_id = b.category_id left join (select ro_id, ro_code from asv_ro) d on a.ro_id = d.ro_id " +
                @"left join (select reg_start_id, reg_start_code from asv_reg_start) e on a.reg_start_id = e.reg_start_id " +
                @"left join (select reg_finish_id, reg_finish_code from asv_reg_finish) r on a.reg_finish_id = r.reg_finish_id " +
                @"left outer join db2inst.asv_kladr k on (" +
                @"(a.KLADR_JUR_STREET_ID=k.kladr_id and k.kladr_type_id=5) " +
                @"or (a.KLADR_JUR_TOWN_ID=k.kladr_id and k.kladr_type_id=4) " +
                @"or (a.KLADR_JUR_CITY_ID=k.kladr_id and k.kladr_type_id=3)) " +
                @"left outer join db2inst.asv_kladr k1 on k.KLADR_PARENT_ID=k1.kladr_id " +
                @"left outer join db2inst.asv_kladr k2 on k1.KLADR_PARENT_ID=k2.kladr_id " +
                @"left outer join db2inst.asv_kladr k3 on k2.KLADR_PARENT_ID=k3.kladr_id " +
                @"where a.insurer_reg_num in (" + regNum + @") order by a.insurer_reg_num";



            //Подключаемся к БД и выполняем запрос
            using (DB2Connection connection = new DB2Connection("Server=1.1.1.1:50000;Database=asv;UID=db2inst;PWD=password;"))
            {
                try
                {
                    //открываем соединение
                    await connection.OpenAsync();
                    //Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.DarkCyan;
                    Console.Write("Соединение с БД: ");
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.WriteLine(connection.State);
                    //Console.WriteLine();

                    DB2Command command = connection.CreateCommand();
                    command.CommandText = query;

                    //Устанавливаем значение таймаута
                    command.CommandTimeout = 570;

                    DbDataReader reader = await command.ExecuteReaderAsync();

                    while (await reader.ReadAsync())
                    {
                        //@"select a.insurer_reg_num, a.insurer_short_name, a.insurer_last_name, a.insurer_first_name, a.insurer_middle_name, " +
                        //@"a.insurer_inn, a.insurer_kpp, " +
                        //@"a.insurer_jur_zipcode, k3.kladr_name, k2.kladr_name, k1.kladr_name, k.kladr_name, a.insurer_jur_house, a.insurer_jur_building, a.insurer_jur_flat " +

                        //public DataFromRKASVDB(string insurer_reg_num = "", string insurer_short_name = "",
                        //        string insurer_last_name = "", string insurer_first_name = "", string insurer_middle_name = "",
                        //        string insurer_inn = "", string kpp = "",
                        //        string insurer_jur_zipcode = "", string k3_kladr_name = "", string k2_kladr_name = "", string k1_kladr_name = "",
                        //        string k_kladr_name = "", string insurer_jur_house = "", string insurer_jur_building = "", string insurer_jur_flat = "")




                        dictionaryDataFromRKASV[regNum] = new DataFromRKASVDB(
                                                                ConvertRegNom(reader[0].ToString()), reader[1].ToString(),
                                                                reader[2].ToString(), reader[3].ToString(), reader[4].ToString(),
                                                                reader[5].ToString(), reader[6].ToString(),
                                                                reader[7].ToString(), reader[8].ToString(), reader[9].ToString(), reader[10].ToString(),
                                                                reader[11].ToString(), reader[12].ToString(), reader[13].ToString(), reader[14].ToString()
                                                                              );
                    }
                    reader.Close();


                    //выбираем данные по УПФР
                    SelectDataFromPerso_UPFR(regNum);

                }
                catch (Exception ex)
                {
                    IOoperations.WriteLogError(ex.ToString());

                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(ex.Message);
                    Console.ForegroundColor = ConsoleColor.Gray;
                }
            }
        }



        //------------------------------------------------------------------------------------------
        private static string SelectRaion(string regNum)
        {
            try
            {
                if (regNum.Count() == 11)
                {
                    return regNum[2].ToString() + regNum[3] + regNum[4];
                }
                else
                {
                    return "";
                }

            }
            catch (Exception ex)
            {
                IOoperations.WriteLogError(ex.ToString());

                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;

                return "";
            }
        }

        //------------------------------------------------------------------------------------------
        private static string ConvertRegNom(string regNom)
        {
            try
            {
                char[] regNomOld = regNom.ToCharArray();
                string regNomConvert = "0" + regNomOld[1].ToString() + regNomOld[2].ToString() + "-" + regNomOld[3].ToString() + regNomOld[4] + regNomOld[5] + "-" + regNomOld[6] + regNomOld[7] + regNomOld[8] + regNomOld[9] + regNomOld[10] + regNomOld[11];


                return regNomConvert;
            }
            catch (Exception ex)
            {
                IOoperations.WriteLogError(ex.ToString());

                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = ConsoleColor.Gray;

                return "";
            }
        }




    }

    #endregion
}
