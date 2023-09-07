using Microsoft.VisualBasic;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeminarCheck.Models
{
    public static class SetModels
    {
        public static List<Seminar> SetSeminars(DataSet ds, Config config)
        {
            List<Seminar> Seminars = new List<Seminar>();
            foreach (var sheet in config.seminar.sheet_name)
            {
                DataTable dt = ds.Tables[sheet];
                if (sheet == "セミナーリスト")
                {
                    for (int i = config.seminar.first_row - 1; i < dt.Rows.Count; i++)
                    {
                        var dr = dt.Rows[i];
                        int op;
                        int.TryParse(dr[config.seminar.syokuin_no - 1].ToString(), out op);
                        Seminars.Add(new Seminar
                        {
                            syokuin_no = op == 0 ? null : op,
                            dep = dr[config.seminar.dep - 1].ToString(),
                            job = dr[config.seminar.job - 1].ToString(),
                            name = dr[config.seminar.name - 1].ToString(),
                            jukou_no = dr[config.seminar.jukou_no - 1].ToString().Replace("\n", "").Replace("\r", "").Replace("<br />", ""),
                            mail = dr[config.seminar.mail - 1].ToString(),
                            min_test = dr[config.seminar.min_test - 1].ToString(),
                            movie_jukou = dr[config.seminar.movie_jukou - 1].ToString(),
                            site = dr[config.seminar.site - 1].ToString(),
                            taisyoku = dr[config.seminar.taisyoku - 1].ToString(),
                            seminar_jukou = dr[config.seminar.seminar_jukou - 1].ToString(),

                        });
                    }
                }
                if (sheet == "対象外の受講ユーザー")
                {
                    for (int i = config.seminar.first_row - 1; i < dt.Rows.Count; i++)
                    {
                        var dr = dt.Rows[i];
                        int op;
                        int.TryParse(dr[config.seminar.syokuin_no_gai - 1].ToString(), out op);
                        Seminars.Add(new Seminar
                        {
                            syokuin_no = op == 0 ? null : op,
                            dep = dr[config.seminar.dep_gai - 1].ToString(),
                            job = dr[config.seminar.job_gai - 1].ToString(),
                            name = dr[config.seminar.name_gai - 1].ToString(),
                            jukou_no = dr[config.seminar.jukou_no_gai - 1].ToString().Replace("\n", "").Replace("\r", "").Replace("<br />", ""),
                            mail = dr[config.seminar.mail_gai - 1].ToString(),
                            site = dr[config.seminar.site_gai - 1].ToString(),
                            taisyoku = dr[config.seminar.taisyoku_gai - 1].ToString(),
                        });
                    }
                }
            }
            return Seminars;
        }
        public static List<Staff_Jinji> SetJinji(DataSet ds, Config config)
        {
            List<Staff_Jinji> Staff_Jinji = new List<Staff_Jinji>();
            foreach (var sheet in config.jinji.sheet_name)
            {
                DataTable dt = ds.Tables[sheet];
                for (int i = config.jinji.first_row - 1; i < dt.Rows.Count; i++)
                {
                    var dr = dt.Rows[i];
                    if (dr[config.jinji.name - 1].ToString() == "") continue;
                    int op;
                    int.TryParse(dr[config.jinji.syokuin_no - 1].ToString(), out op);
                    Staff_Jinji.Add(new Staff_Jinji
                    {
                        Staff_No = op == 0 ? null : op,
                        Department = dr[config.jinji.dep - 1].ToString(),
                        Job = dr[config.jinji.job - 1].ToString(),
                        Name = dr[config.jinji.name - 1].ToString(),
                        State = dr[config.jinji.state - 1].ToString(),
                        Site = dr[config.jinji.site - 1].ToString(),

                    });
                }
            }
            return Staff_Jinji;
        }

        public static List<Staff_Jinjigai> SetJinji_gai(DataSet ds, Config config)
        {
            List<Staff_Jinjigai> Staff_Jinjigai = new List<Staff_Jinjigai>();
            foreach (var sheet in config.jinjigai.sheet_name)
            {
                DataTable dt = ds.Tables[sheet];
                for (int i = config.jinji.first_row - 1; i < dt.Rows.Count; i++)
                {
                    var dr = dt.Rows[i];
                    if (dr[config.jinjigai.name - 1].ToString() == "" || dr[config.jinjigai.company - 1].ToString() == "") continue;
                    Staff_Jinjigai.Add(new Staff_Jinjigai
                    {
                        Department = dr[config.jinjigai.dep - 1].ToString(),
                        Company = dr[config.jinjigai.company - 1].ToString(),
                        Name = dr[config.jinjigai.name - 1].ToString(),

                    });
                }
            }
            return Staff_Jinjigai;
        }
    }

    public class Config
    {

        public Seminar seminar = new Seminar();
        public Jinji jinji = new Jinji();
        public Jinjigai jinjigai = new Jinjigai();
        public class Seminar
        {

            public string file_name { get; set; }
            public string[] sheet_name { get; set; }
            public int first_row { get; set; }
            public int syokuin_no { get; set; }
            public int site { get; set; }
            public int dep { get; set; }
            public int job { get; set; }
            public int name { get; set; }
            public int mail { get; set; }
            public int taisyoku { get; set; }
            public int seminar_jukou { get; set; }
            public int movie_jukou { get; set; }
            public int min_test { get; set; }
            public int jukou_no { get; set; }

            public string file_name_gai { get; set; }
            public string[] sheet_name_gai { get; set; }
            public int first_row_gai { get; set; }
            public int syokuin_no_gai { get; set; }
            public int site_gai { get; set; }
            public int dep_gai { get; set; }
            public int job_gai { get; set; }
            public int name_gai { get; set; }
            public int mail_gai { get; set; }
            public int taisyoku_gai { get; set; }
            public int jukou_no_gai { get; set; }

        }
        public class Jinji
        {
            public string file_name { get; set; }
            public string[] sheet_name { get; set; }
            public int first_row { get; set; }
            public int syokuin_no { get; set; }
            public int state { get; set; }
            public int site { get; set; }
            public int dep { get; set; }
            public int job { get; set; }
            public int name { get; set; }
        }
        public class Jinjigai
        {
            public string file_name { get; set; }
            public string[] sheet_name { get; set; }
            public int first_row { get; set; }
            public int company { get; set; }
            public int name { get; set; }
            public int dep { get; set; }
        }
    }

    public static class NormalizeName
    {
        static private Dictionary<string, string> iseitai = new Dictionary<string, string>() {
            {"邉","辺" },
            {"邊","辺" },
            {"髙","高" },
            {"槗","橋" },
            {"𣘺","橋" },
            {"﨑","崎" },
            {"碕","橋" },
            {"嵜","橋" },
            {"斎","斉" },
            {"齊","斉" },
            {"齋","斉" },
            {"籐","藤" },
            {"嶌","島" },
            {"桒","桑" },
            {"瀨","瀬" }
        };
        static public string ConvertName(string name)
        {
            foreach (var item in iseitai)
            {
                if (name.Contains(item.Key)) name = name.Replace(item.Key, item.Value);
            }
            var s1 = System.Text.RegularExpressions.Regex.Replace(name, @"\s", "");
            var s2 = s1.Replace("・", "").ToLower();
            var s3 = Strings.StrConv(s2.ToLower(), VbStrConv.Wide, 0);
            return s3;
        }
    }
    public class Seminar : BindableBase
    {
        public string Normalize_name
        {
            get
            {
                return NormalizeName.ConvertName(name);
            }
        }
        public int? syokuin_no { get; set; }
        public string site { get; set; }
        public string dep { get; set; }
        public string job { get; set; }
        public string name { get; set; }
        public string mail { get; set; }
        public string taisyoku { get; set; }
        public string seminar_jukou { get; set; }
        public string movie_jukou { get; set; }
        public string min_test { get; set; }
        public string jukou_no { get; set; }
    }
    public class Staff_Jinji : BindableBase
    {
        public string Normalize_name
        {
            get
            {
                return NormalizeName.ConvertName(name);
            }
        }
        private int? staff_no;
        public int? Staff_No
        {
            get { return staff_no; }
            set { SetProperty(ref staff_no, value); }
        }
        private string name;
        public string Name
        {
            get { return name; }
            set { SetProperty(ref name, value); }
        }
        private string state;
        public string State
        {
            get { return state; }
            set { SetProperty(ref state, value); }
        }
        private string site;
        public string Site
        {
            get { return site; }
            set { SetProperty(ref site, value); }
        }
        private string department;
        public string Department
        {
            get { return department; }
            set { SetProperty(ref department, value); }
        }
        private string job;
        public string Job
        {
            get { return job; }
            set { SetProperty(ref job, value); }
        }

    }
    public class Staff_Jinjigai : BindableBase
    {
        public string Normalize_name
        {
            get
            {
                return NormalizeName.ConvertName(name);
            }
        }
        private string name;
        public string Name
        {
            get { return name; }
            set { SetProperty(ref name, value); }
        }
        private string company;
        public string Company
        {
            get { return company; }
            set { SetProperty(ref company, value); }
        }
        private string department;
        public string Department
        {
            get { return department; }
            set { SetProperty(ref department, value); }
        }

    }

    public class ViewList_Jinji : BindableBase
    {
        private string name;
        public string Name
        {
            get { return name; }
            set { SetProperty(ref name, value); }
        }
        private string site;
        public string Site
        {
            get { return site; }
            set { SetProperty(ref site, value); }
        }
        private string department;
        public string Department
        {
            get { return department; }
            set { SetProperty(ref department, value); }
        }
        private string job;
        public string Job
        {
            get { return job; }
            set { SetProperty(ref job, value); }
        }
        private int? staff_no;
        public int? Staff_No
        {
            get { return staff_no; }
            set { SetProperty(ref staff_no, value); }
        }
        private string jukou_no;
        public string Jukou_no
        {
            get { return jukou_no; }
            set { SetProperty(ref jukou_no, value); }
        }
        private bool isMatch;
        public bool IsMatch
        {
            get { return isMatch; }
            set { SetProperty(ref isMatch, value); }
        }

        public string IsTaishoku;

        //public bool IsJukou
        //{
        //    get { return jukou_no != "済み"; }
        //}
    }
    public class ViewList_Jinjigai : BindableBase
    {
        private string name;
        public string Name
        {
            get { return name; }
            set { SetProperty(ref name, value); }
        }
        private string company;
        public string Company
        {
            get { return company; }
            set { SetProperty(ref company, value); }
        }
        private string department;
        public string Department
        {
            get { return department; }
            set { SetProperty(ref department, value); }
        }
        private string jukou_no;
        public string Jukou_no
        {
            get { return jukou_no; }
            set { SetProperty(ref jukou_no, value); }
        }
        private bool isMatch;
        public bool IsMatch
        {
            get { return isMatch; }
            set { SetProperty(ref isMatch, value); }
        }
        public string IsTaishoku;
    }


}
