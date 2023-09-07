using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Windows;

namespace SeminarCheck.Models
{
    public static class ReadFile
    {
        public static void GetConfig(ref Config config)
        {
            var xmlDoc = new XmlDocument();
            string path = System.AppDomain.CurrentDomain.BaseDirectory;

            xmlDoc.Load(@"C:\Users\kamizugu\Source\Repos\KM0426\MedicalInformationDepartment\config.xml");
            //            xmlDoc.Load(path + @"\config.xml");
            XmlNodeList? seminarconfs = xmlDoc.SelectNodes("/config/seminars");
            XmlNodeList? jinjiconfs = xmlDoc.SelectNodes("/config/jinjis");
            XmlNodeList? jinjigaiconfs = xmlDoc.SelectNodes("/config/jinjigais");

            config.seminar.sheet_name = seminarconfs[0].Attributes["sheetname"].InnerText.Split(',');
            config.jinji.sheet_name = jinjiconfs[0].Attributes["sheetname"].InnerText.Split(',');
            config.jinjigai.sheet_name = jinjigaiconfs[0].Attributes["sheetname"].InnerText.Split(','); ;
            config.seminar.first_row = int.Parse(seminarconfs[0].Attributes["first_row"].InnerText);
            config.jinji.first_row = int.Parse(jinjiconfs[0].Attributes["first_row"].InnerText);
            config.jinjigai.first_row = int.Parse(jinjigaiconfs[0].Attributes["first_row"].InnerText);

            foreach (XmlNode jinjiconf in seminarconfs[0].SelectNodes("seminar"))
            {
                var title = jinjiconf.SelectSingleNode("title").InnerText;
                var column = int.Parse(jinjiconf.SelectSingleNode("column").InnerText);
                switch (title)
                {
                    case "職員番号(セミナーリスト)":
                        config.seminar.syokuin_no = column;
                        break;
                    case "施設(セミナーリスト)":
                        config.seminar.site = column;
                        break;
                    case "部署(セミナーリスト)":
                        config.seminar.dep = column;
                        break;
                    case "職名(セミナーリスト)":
                        config.seminar.job = column;
                        break;
                    case "氏名(セミナーリスト)":
                        config.seminar.name = column;
                        break;
                    case "メール(セミナーリスト)":
                        config.seminar.mail = column;
                        break;
                    case "退職(セミナーリスト)":
                        config.seminar.taisyoku = column;
                        break;
                    case "セミナー受講(セミナーリスト)":
                        config.seminar.seminar_jukou = column;
                        break;
                    case "動画受講(セミナーリスト)":
                        config.seminar.movie_jukou = column;
                        break;
                    case "小テスト受講(セミナーリスト)":
                        config.seminar.min_test = column;
                        break;
                    case "受講番号(セミナーリスト)":
                        config.seminar.jukou_no = column;
                        break;
                    case "職員番号(対象外の受講ユーザー)":
                        config.seminar.syokuin_no_gai = column;
                        break;
                    case "施設(対象外の受講ユーザー)":
                        config.seminar.site_gai = column;
                        break;
                    case "部署(対象外の受講ユーザー)":
                        config.seminar.dep_gai = column;
                        break;
                    case "職名(対象外の受講ユーザー)":
                        config.seminar.job_gai = column;
                        break;
                    case "氏名(対象外の受講ユーザー)":
                        config.seminar.name_gai = column;
                        break;
                    case "メール(対象外の受講ユーザー)":
                        config.seminar.mail_gai = column;
                        break;
                    case "退職(対象外の受講ユーザー)":
                        config.seminar.taisyoku_gai = column;
                        break;
                    case "受講番号(対象外の受講ユーザー)":
                        config.seminar.jukou_no_gai = column;
                        break;
                    default:
                        break;
                }
            }
            foreach (XmlNode jinjiconf in jinjiconfs[0].SelectNodes("jinji"))
            {

                var title = jinjiconf.SelectSingleNode("title").InnerText;
                var column = int.Parse(jinjiconf.SelectSingleNode("column").InnerText);
                switch (title)
                {
                    case "職員番号":
                        config.jinji.syokuin_no = column;
                        break;
                    case "ステータス":
                        config.jinji.state = column;
                        break;
                    case "地区区分":
                        config.jinji.site = column;
                        break;
                    case "所属":
                        config.jinji.dep = column;
                        break;
                    case "職名":
                        config.jinji.job = column;
                        break;
                    case "氏名":
                        config.jinji.name = column;
                        break;
                    default:
                        break;
                }
            }
            foreach (XmlNode jinjiconf in jinjigaiconfs[0].SelectNodes("jinjigai"))
            {

                var title = jinjiconf.SelectSingleNode("title").InnerText;
                var column = int.Parse(jinjiconf.SelectSingleNode("column").InnerText);
                switch (title)
                {
                    case "会社名":
                        config.jinjigai.company = column;
                        break;
                    case "氏名":
                        config.jinjigai.name = column;
                        break;
                    case "勤務場所":
                        config.jinjigai.dep = column;
                        break;
                    default:
                        break;
                }
            }
        }

        public static DataSet AsDataSetSample(Config config)
        {
            DataSet ds = new DataSet();
            for (int i = 0; i < 3; i++)
            {
                string filepath = String.Empty;
                string[] sheetNames = { };
                switch (i)
                {
                    case 0:
                        filepath = config.seminar.file_name;
                        sheetNames = config.seminar.sheet_name;
                        break;
                    case 1:
                        filepath = config.jinji.file_name;
                        sheetNames = config.jinji.sheet_name;
                        break;
                    case 2:
                        filepath = config.jinjigai.file_name;
                        sheetNames = config.jinjigai.sheet_name;
                        break;
                }

                //ファイルの読み取り開始
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (FileStream stream = File.Open(filepath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;

                    //ファイルの拡張子を確認
                    if (filepath.EndsWith(".xls") || filepath.EndsWith(".xlsx") || filepath.EndsWith(".xlsb"))
                    {
                        reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("Shift_JIS")
                        });
                    }
                    else if (filepath.EndsWith(".csv"))
                    {
                        reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("Shift_JIS")
                        });
                    }
                    else
                    {
                        MessageBox.Show("サポート対象外の拡張子です。");
                        return null;
                    }

                    //全シート全セルを読み取り
                    var dataset = reader.AsDataSet();
                    foreach (var sheetName in sheetNames)
                    {
                        var worksheeta = dataset.Tables[sheetName];
                        if (worksheeta is null)
                        {
                            Debug.WriteLine($"指定されたワークシート名 '{sheetName}' が存在しません。");
                        }
                        else
                        {
                            DataTable modifiedTable = RemoveNewlinesFromTable(worksheeta);
                            ds.Tables.Add(modifiedTable.Copy());
                        }
                    }
                    reader.Close();
                }

            }


            return ds;
        }
        static DataTable RemoveNewlinesFromTable(DataTable table)
        {
            DataTable modifiedTable = table.Clone();

            foreach (DataRow row in table.Rows)
            {
                DataRow newRow = modifiedTable.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    string cellValue = row[i].ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        // セル内の改行を削除
                        cellValue = cellValue.Replace("\n", " ").Replace("\r", "");
                    }
                    newRow[i] = cellValue;
                }
                modifiedTable.Rows.Add(newRow);
            }

            return modifiedTable;
        }
    }
}
