using Microsoft.WindowsAPICodePack.Dialogs;
using Prism.Commands;
using Prism.Mvvm;
using SeminarCheck.Models;
using System.Collections.Generic;
using System.Data;
using System;
using System.Windows;
using System.Linq;

namespace SeminarCheck.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        private List<ViewList_Jinji> viewList_Jinjis;
        public List<ViewList_Jinji> ViewList_Jinjis
        {
            get { return viewList_Jinjis; }
            set { SetProperty(ref viewList_Jinjis, value); }
        }
        private List<ViewList_Jinjigai> viewList_Jinjigais;
        public List<ViewList_Jinjigai> ViewList_Jinjigais
        {
            get { return viewList_Jinjigais; }
            set { SetProperty(ref viewList_Jinjigais, value); }
        }
        public Config config = new Config();
        private List<Seminar> seminars;
        public List<Seminar> Seminars
        {
            get { return seminars; }
            set { SetProperty(ref seminars, value); }
        }
        private List<Staff_Jinji> jinjis;
        public List<Staff_Jinji> Jinjis
        {
            get { return jinjis; }
            set { SetProperty(ref jinjis, value); }
        }
        private List<Staff_Jinjigai> jinjigais;
        public List<Staff_Jinjigai> Jinjigais
        {
            get { return jinjigais; }
            set { SetProperty(ref jinjigais, value); }
        }
        private string _title = "Seminar Checker";
        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }
        public DelegateCommand Click_FileBtn { get; private set; }

        public MainWindowViewModel()
        {
            Click_FileBtn = new DelegateCommand(StartCheck);
        }

        private string FileSelect(string title, string ext)
        {
            using (var cofd = new CommonOpenFileDialog()
            {
                Title = title,
                IsFolderPicker = false,
            })
            {
                cofd.Filters.Add(new CommonFileDialogFilter("*." + ext, "*." + ext));
                cofd.Filters.Add(new CommonFileDialogFilter("すべて", "*.*"));
                if (cofd.ShowDialog() != CommonFileDialogResult.Ok)
                {
                    MessageBox.Show($"キャンセル");
                    return string.Empty;
                }
                else
                {
                    return cofd.FileName;
                }
            }
        }

        private void StartCheck()
        {
            config = new Config();
            ReadFile.GetConfig(ref config);
            config.seminar.file_name = FileSelect("セミナー受講状況ファイルを選択してください", "xls");
            if (config.seminar.file_name == "") return;
            config.jinji.file_name = FileSelect("人事雇用の職員名簿ファイルを選択してください", "xlsx");
            if (config.jinji.file_name == "") return;
            config.jinjigai.file_name = FileSelect("人事雇用外の職員名簿ファイルを選択してください", "xlsx");
            if (config.jinjigai.file_name == "") return;



            DataSet ds = ReadFile.AsDataSetSample(config);
            if (ds == null) return;
            Seminars = SetModels.SetSeminars(ds, config);
            Jinjis = SetModels.SetJinji(ds, config);

            Jinjigais = SetModels.SetJinji_gai(ds, config);
            Jinjis = Jinjis.Where(w => w.Site == "柏" && w.State == "在籍").ToList();

            var j = from jinji in jinjis
                    join seminar in Seminars on jinji.Normalize_name equals seminar.Normalize_name into gj
                    from subpet in gj.DefaultIfEmpty()
                    select new ViewList_Jinji
                    {
                        Name = jinji.Name,
                        Site = jinji.Site,
                        Department = jinji.Department,
                        Job = jinji.Job,
                        Staff_No = jinji.Staff_No,
                        Jukou_no = subpet?.jukou_no ?? String.Empty,
                        IsMatch = subpet != null,
                        IsTaishoku = subpet?.taisyoku ?? String.Empty
                    };
            var jj = from jinjigai in jinjigais
                     join seminar in Seminars on jinjigai.Normalize_name equals seminar.Normalize_name into gj
                     from subpet in gj.DefaultIfEmpty()
                     select new ViewList_Jinjigai
                     {
                         Name = jinjigai.Name,
                         Company = jinjigai.Company,
                         Department = jinjigai.Department,
                         Jukou_no = subpet?.jukou_no ?? String.Empty,
                         IsMatch = subpet != null,
                         IsTaishoku = subpet?.taisyoku ?? String.Empty
                     };

            List<ViewList_Jinjigai> tmpjinjigai = new List<ViewList_Jinjigai>();
            var jjname_G = jj.Where(w => w.IsTaishoku == "").GroupBy(g => g.Name);
            foreach (var jjname in jjname_G)
            {
                if (jjname.Count() > 1)
                {
                    var jukouzumi = jjname.Select(s => s).Where(w => w.Jukou_no != "");
                    if (jukouzumi.Count() > 0)
                    {
                        tmpjinjigai.Add(new ViewList_Jinjigai
                        {
                            Name = jukouzumi.First().Name,
                            Company = jukouzumi.First().Company,
                            Department = jukouzumi.First().Department,
                            IsMatch = jukouzumi.First().IsMatch,
                            Jukou_no = jukouzumi.First().Jukou_no
                        });
                    }
                    else
                    {
                        tmpjinjigai.Add(new ViewList_Jinjigai
                        {
                            Name = jjname.First().Name,
                            Company = jjname.First().Company,
                            Department = jjname.First().Department,
                            IsMatch = jjname.First().IsMatch,
                            Jukou_no = jjname.First().Jukou_no
                        });
                    }
                }
                else
                {
                    tmpjinjigai.Add(new ViewList_Jinjigai
                    {
                        Name = jjname.First().Name,
                        Company = jjname.First().Company,
                        Department = jjname.First().Department,
                        IsMatch = jjname.First().IsMatch,
                        Jukou_no = jjname.First().Jukou_no
                    });
                }
            }

            var jg = j.Where(w => w.IsTaishoku == "").GroupBy(g => g.Staff_No);
            List<ViewList_Jinji> tmpjinji = new List<ViewList_Jinji>();

            foreach (var item in jg)
            {
                if (item.Count() > 1)
                {
                    var v = item.Where(w => w.Jukou_no != "");
                    if (v.Count() == 0)
                    {
                        tmpjinji.Add(item.First());
                    }
                    else
                    {
                        if (v.Count() == 1)
                        {
                            foreach (var vv in item)
                            {
                                if (vv.Jukou_no != "")
                                {
                                    tmpjinji.Add(vv);
                                }
                            }

                        }
                        else
                        {
                            string jukoutst = String.Empty;
                            foreach (var juko in v)
                            {
                                jukoutst += "_" + juko.Jukou_no;
                            }
                            tmpjinji.Add(new ViewList_Jinji()
                            {
                                Staff_No = v.First().Staff_No,
                                Site = v.First().Site,
                                Jukou_no = jukoutst,
                                Department = v.First().Department,
                                IsMatch = v.First() != null,
                                Job = v.First().Job,
                                Name = v.First().Name,
                            });
                        }

                    }
                }
                else
                {
                    tmpjinji.Add(item.First());
                }
            }
            ViewList_Jinjigais = tmpjinjigai;
            ViewList_Jinjis = tmpjinji;
        }
    }
}
