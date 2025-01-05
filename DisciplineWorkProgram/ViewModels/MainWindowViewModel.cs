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

namespace DisciplineWorkProgram.ViewModels
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        public string TemplatePath = Directory.GetCurrentDirectory() + "\\DWP_TemplateBookmarks.docx";
        public const string DwpDir = "dwp/";
        public bool isHasDate = false;

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

#if DEBUG
            var section = SectionsByWayName.Single().Sections.Single();
            //new Dwp(section)
            //		.MakeDwp(TemplatePath, DwpDir, section.Disciplines.First().Key);
            foreach (var discipline in section.GetCheckedDisciplinesNames)
            {
                new Dwp(section)
                    .MakeDwp(TemplatePath, DwpDir, discipline);
            }
#else
			try
			{
				var section = SectionsByWayName.Single().Sections.Single();
				//new Dwp(section)
				//		.MakeDwp(TemplatePath, DwpDir, section.Disciplines.First().Key);
				foreach (var discipline in section.GetCheckedDisciplinesNames)
				{
					new Dwp(section)
						.MakeDwp(TemplatePath, DwpDir, discipline);
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
            var outPath = "";

            if (ext != ".xlsx")
            {
                Excel.Converter.Convert(PlanPath, out outPath);
                PlanPath = outPath;
            }
            using var plan = Excel.Converter.Convert(PlanPath);
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
            List<string> problems1 = new List<string>();
            List<string> problems2 = new List<string>();
            foreach (var discipline in section.GetCheckedDisciplinesNames)
            {
                if (!section.Disciplines.ContainsKey(discipline))
                {
                    problems1.Add(section.Disciplines[discipline].Name);
                }
                if (!section.DisciplineCompetencies.ContainsKey(section.Disciplines[discipline].Name))
                {
                    problems2.Add(section.Disciplines[discipline].Name);
                }
            }
            if (problems1.Count > 0 || problems2.Count > 0)
            {
                StringBuilder strTemp = new StringBuilder();
                strTemp.Append("\n проблемы по первой стадии:\n");
                foreach (var problem in problems1)
                {
                    strTemp.Append(problem);
                    strTemp.Append(", ");
                }
                strTemp.Append("\n проблемы по второй стадии:\n");
                foreach (var problem in problems2)
                {
                    strTemp.Append(problem);
                    strTemp.Append(", ");
                }

                var messageBoxCustomWindow = MessageBoxManager
                   .GetMessageBoxCustom(new MessageBoxCustomParams
                   {

                       ButtonDefinitions = new[] {
                                            new ButtonDefinition {Name = "Ok"}
                       },
                       ContentMessage = "sdasdsad:" + strTemp.ToString(),
                       WindowStartupLocation = WindowStartupLocation.CenterOwner,
                       CanResize = true,
                       MinHeight = 1000,
                       MinWidth = 1000,
                       MaxWidth = 2500,
                       MaxHeight = 2800,
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