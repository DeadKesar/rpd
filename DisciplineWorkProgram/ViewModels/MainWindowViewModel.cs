using Avalonia.Controls;
using DisciplineWorkProgram.Models;
using DisciplineWorkProgram.Models.Sections;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System;
using DisciplineWorkProgram.Excel;
using System.Text;
using MsBox.Avalonia.Enums;
using MsBox.Avalonia;
using MsBox.Avalonia.Dto;
using MsBox.Avalonia.Models;

using NPOI.HSSF.UserModel;    // Для .xls (HSSF)
using NPOI.SS.UserModel;     // Общий интерфейс
using NPOI.XSSF.UserModel;   // Для .xlsx (XSSF)

namespace DisciplineWorkProgram.ViewModels
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        public string DWPTemplatePath = Directory.GetCurrentDirectory() + "\\DWP_TemplateBookmarks.docx";
        public const string DwpDir = "dwp/";
        public string FOSTemplatePath = Directory.GetCurrentDirectory() + "\\FOS_TemplateBookmarks.docx";
        public const string fosDir = "fos/";
        public string PRACTemplatePath = Directory.GetCurrentDirectory() + "\\PRAC_TemplateBookmarks.docx";
        public const string pracDir = "prac/";
        public string VKRTemplatePath = Directory.GetCurrentDirectory() + "\\VKR_TemplateBookmarks.docx";
        public const string vkrDir = "vkr/";
        public bool isHasDate = false;
        public string PathToDolz = Directory.GetCurrentDirectory() + "\\dolznosti.xlsx";
        public Employee employes = new Employee(Directory.GetCurrentDirectory() + "\\dolznosti.xlsx");

        public event PropertyChangedEventHandler PropertyChanged;

        //[NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string planPath;
        private string compListPath;
        private string compMatrixPath;

        public string PlanPath
        {
            get => planPath;
            set
            {
                planPath = value;
                OnPropertyChanged(nameof(PlanPath));
            }
        }

        public string CompListPath
        {
            get => compListPath;
            set
            {
                compListPath = value;
                OnPropertyChanged(nameof(CompListPath));
            }
        }

        public string CompMatrixPath
        {
            get => compMatrixPath;
            set
            {
                compMatrixPath = value;
                OnPropertyChanged(nameof(CompMatrixPath));
            }
        }

        public static readonly ObservableCollection<SectionsByWay> SectionsByWayName = new ObservableCollection<SectionsByWay>();
        public ObservableCollection<SectionsByWay> Items { get; set; }


        public async Task ChangePlanPath()
        {
            PlanPath = await ChangeFilePath();
        }

        public async Task ChangeCompListPath()
        {
            CompListPath = await ChangeFilePath();
        }

        public async Task ChangeCompMatrixPath()
        {
            CompMatrixPath = await ChangeFilePath();
        }

        public void LoadDataButton()
        {
            //var temp = employes.Employees["Финансовый менеджмент"]["position"];
            
            if (string.IsNullOrWhiteSpace(PlanPath) || string.IsNullOrWhiteSpace(CompListPath) || string.IsNullOrWhiteSpace(CompMatrixPath))
            {
                var messageBoxCustomWindow = MessageBoxManager.GetMessageBoxCustom(
                    new MessageBoxCustomParams
                    {
                        ContentMessage = "Выберите все документы",
                        ButtonDefinitions = new[] {
                        new ButtonDefinition {Name = "Ok"}
                        },
                        WindowStartupLocation = WindowStartupLocation.CenterOwner
                    });

                messageBoxCustomWindow.ShowWindowAsync();

                return;
            }
            LoadData();
            UpdateSource();

        }

        public void MakeDwps()
        {
            String TemplatePath;
            String dir;
#if DEBUG
            var section = SectionsByWayName.Single().Sections.Single();
            //new Dwp(section)
            //		.MakeDwp(TemplatePath, DwpDir, section.Disciplines.First().Key);
            foreach (var discipline in section.GetCheckedDisciplinesNames)
            {
                    if (discipline.Contains("Б2") 
                        //||section.Disciplines[discipline].Props["Name"].ToLower().Contains("практика") 
                        //|| section.Disciplines[discipline].Props["Name"].ToLower().Contains("аттестация"))
                        )
                    {
                        TemplatePath = PRACTemplatePath;
                        dir = pracDir;
                    }
                    else if (discipline.Contains("Б3"))
                    {
                        TemplatePath = VKRTemplatePath;
                        dir = vkrDir;
                    }
                    else
                    {
                        TemplatePath = DWPTemplatePath;
                        dir = DwpDir;
                    }
                new Dwp(section)
                    .MakeDwp(TemplatePath, dir, discipline, employes);
            }
#else
            try
            {
				var section = SectionsByWayName.Single().Sections.Single();
				//new Dwp(section)
				//		.MakeDwp(TemplatePath, DwpDir, section.Disciplines.First().Key);
				foreach (var discipline in section.GetCheckedDisciplinesNames)
				{
                    if (discipline.Contains("Б2") 
                        //||section.Disciplines[discipline].Props["Name"].ToLower().Contains("практика") 
                        //|| section.Disciplines[discipline].Props["Name"].ToLower().Contains("аттестация"))
                        )
                    {
                        TemplatePath = PRACTemplatePath;
                        dir = pracDir;
                    }
                    else if (discipline.Contains("Б3"))
                    {
                        TemplatePath = VKRTemplatePath;
                        dir = vkrDir;
                    }
                    else
                    {
                        TemplatePath = DWPTemplatePath;
                        dir = DwpDir;
                    }
					new Dwp(section)
						.MakeDwp(TemplatePath, dir, discipline, employes);
				}
			}
            catch (Exception ex)
            {
                ShowErrorAsync(ex.Message); // Отображаем ошибку пользователю
            }
#endif
        }
        public void MakeFoses()
        {

#if DEBUG
            var section = SectionsByWayName.Single().Sections.Single();

            foreach (var discipline in section.GetCheckedDisciplinesNames)
            {
                new Fos(section)
                    .MakeFos(FOSTemplatePath, fosDir, discipline, employes);
            }
#else
            try
            {
				var section = SectionsByWayName.Single().Sections.Single();
				foreach (var discipline in section.GetCheckedDisciplinesNames)
				{
					new Fos(section)
                    .MakeFos(FOSTemplatePath, fosDir, discipline, employes);
				}
			}
            catch (Exception ex)
            {
                ShowErrorAsync(ex.Message); // Отображаем ошибку пользователю
            }
#endif
        }

        public async Task ShowErrorAsync(string message)
        {
            var messageBox = MessageBoxManager.GetMessageBoxCustom(
                new MessageBoxCustomParams
                {
                    ContentTitle = "Ошибка",
                    ContentMessage = message,
                    Icon = Icon.Error,
                    ButtonDefinitions = new[] { new ButtonDefinition { Name = "OK" } },
                    WindowStartupLocation = WindowStartupLocation.CenterOwner
                });

            await messageBox.ShowAsync();
        }


        private async static Task<string> ChangeFilePath()
        {
            var dialog = new OpenFileDialog();
            string[] result = null;
            //dialog.Filters.Add(new FileDialogFilter() { Name = "Text", Extensions = { "txt" } });
            result = await dialog.ShowAsync(new Window());

            return result == null ? null : result[0];

            //var dialog = new OpenFileDialog();

            //return dialog.ShowDialog() is true
            //	? dialog.FileName
            //	: string.Empty;
        }

        private void UpdateSource()
        {
            Items = SectionsByWayName;
            OnPropertyChanged(nameof(Items));
        }

        private void LoadData()
        {
            var section = new Section(CompListPath, CompMatrixPath);

            var ext = Path.GetExtension(PlanPath);


         /*   if (ext != ".xlsx")
            {
                Excel.Converter.Convert2(PlanPath);
                //PlanPath = outPath;
            }*/
            using var plan = (ext != ".xlsx")? Excel.Converter.Convert2(PlanPath) : Excel.Converter.Convert(PlanPath);
            if (ext != ".xlsx")
            {
                var name = "\\temp.xlsx";

                if (!Directory.Exists(DwpDir))
                    Directory.CreateDirectory(DwpDir);

                using var fileOut = File.Create(DwpDir + name);
                plan.CopyTo(fileOut);
                planPath = DwpDir + name;
                fileOut.Close();
            }

#if DEBUG
            section.LoadDataFromPlan(PlanPath);
            section.LoadCompetenciesData();

            // пережиток прошлого, легаси
            // переход от списка планов к одному плану довольно трудоёмок, поэтому оставляю
            // логику списка с всегда одним элементом

            SectionsByWayName.Clear();
            SectionsByWayName.Add(
                new SectionsByWay(section)
                {
                    Name = section.SectionDictionary["WayName"]
                });

#else
			try
            {
				// section.LoadDataFromPlan(plan);
				section.LoadDataFromPlan(PlanPath);
				section.LoadCompetenciesData();

				// пережиток прошлого, легаси
				// переход от списка планов к одному плану довольно трудоёмок, поэтому оставляю
				// логику списка с всегда одним элементом

				SectionsByWayName.Clear();
				SectionsByWayName.Add(
					new SectionsByWay(section)
					{
						Name = section.SectionDictionary["WayName"]
					});
			}
            catch (Exception ex)
            {
                ShowErrorAsync(ex.Message); // Отображаем ошибку пользователю
            }
#endif
            isHasDate = true;
        }


        public void CheckDate()
        {
            if (!isHasDate)
            {
                var messageBoxCustomWindow = MessageBoxManager
                   .GetMessageBoxCustom(new MessageBoxCustomParams
                   {
                       ContentMessage = "Сначала загрузите данные",
                       ButtonDefinitions = new[] {
                                            new ButtonDefinition {Name = "Ok"}
                       },
                       WindowStartupLocation = WindowStartupLocation.CenterOwner
                   });
                messageBoxCustomWindow.ShowWindowAsync();
                return;
            }
            var section = SectionsByWayName.Single().Sections.Single();
            List<string> problems1 = new List<string>();//проверяем что дисциплины содежат дисциплины... бред, но эта проверка была изначально и я её оставил
            List<string> problems2 = new List<string>();//сверяем что дисциплины полученные из плана соответствуют дисциплинам из матрицы компетенции
            List<string> problems3 = new List<string>();// сверяем что все кафедры из плана есть в нашем плане
            foreach (var discipline in section.GetAnyDisciplinesNames)
            {
                if (!section.Disciplines.ContainsKey(discipline))
                {
                    problems1.Add(section.Disciplines[discipline].Name);
                }
                if (!section.DisciplineCompetencies.ContainsKey(section.Disciplines[discipline].Name))
                {
                    problems2.Add(section.Disciplines[discipline].Name);
                }
                if (!employes.Employees.ContainsKey(section.Disciplines[discipline].Props["Department"]))
                {
                    problems3.Add(section.Disciplines[discipline].Props["Department"]);
                }
            }
            if (problems1.Count > 0 || problems2.Count > 0)
            {
                StringBuilder strTemp = new StringBuilder();
                strTemp.Append("проблемы по учебному плану:\n");
                foreach (var problem in problems1)
                {
                    strTemp.Append(problem);
                    strTemp.Append("\n");
                }
                strTemp.Append("\nпроблемы по матрице компетенций:\n");
                foreach (var problem in problems2)
                {
                    strTemp.Append(problem);
                    strTemp.Append("\n");
                }
                strTemp.Append("\nпроблемы по файлу с должностями:\nНе обнаружены кафедры:\n");
                foreach (var problem in problems3)
                {
                    strTemp.Append(problem);
                    strTemp.Append("\n");
                }

                var messageBoxCustomWindow = MessageBoxManager
                   .GetMessageBoxCustom(new MessageBoxCustomParams
                   {

                       ButtonDefinitions = new[] {
                                            new ButtonDefinition {Name = "Ok"}
                       },
                       ContentMessage = strTemp.ToString(),
                       WindowStartupLocation = WindowStartupLocation.CenterOwner,
                       CanResize = true,
                       MinHeight = 300,
                       MinWidth = 400,
                       MaxWidth = 1000,
                       MaxHeight = 1500,
                       SizeToContent = SizeToContent.WidthAndHeight,
                       ShowInCenter = true,
                       Topmost = true
                   });

                messageBoxCustomWindow.ShowWindowAsync();
                return;
            }
            var messageBoxCustomWindow2 = MessageBoxManager
                   .GetMessageBoxCustom(new MessageBoxCustomParams
                   {
                       ContentMessage = "Проблем не найдено.",
                       ButtonDefinitions = new[] {
                                            new ButtonDefinition {Name = "Ok"}
                       },
                       WindowStartupLocation = WindowStartupLocation.CenterOwner
                   });
            messageBoxCustomWindow2.ShowWindowAsync();
        }


        /// <summary>
        /// legacy :D
        /// </summary>
        public class SectionsByWay : HierarchicalCheckableElement
        {
            public readonly IList<Section> Sections = new List<Section>();

            protected override IEnumerable<HierarchicalCheckableElement> GetNodes() => Sections;

            public string Name { get; set; }

            public SectionsByWay(Section section)
            {
                Sections.Add(section);
            }

            public void Add(Section section)
            {
                Sections.Add(section);
            }
        }




    }
}