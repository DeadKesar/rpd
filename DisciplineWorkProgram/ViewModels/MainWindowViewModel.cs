using Avalonia.Controls;
using DisciplineWorkProgram.Models;
using DisciplineWorkProgram.Models.Sections;
using JetBrains.Annotations;
using MessageBox.Avalonia.DTO;
using MessageBox.Avalonia.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using MessageBox.Avalonia.Enums;
using System;
using DisciplineWorkProgram.Excel;

namespace DisciplineWorkProgram.ViewModels
{
	public class MainWindowViewModel : INotifyPropertyChanged
	{
		public string TemplatePath = Directory.GetCurrentDirectory() + "\\DWP_TemplateBookmarks.docx";
		public const string DwpDir = "dwp/";

		public event PropertyChangedEventHandler PropertyChanged;

		[NotifyPropertyChangedInvocator]
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


		public async Task  ChangePlanPath()
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
				var messageBoxCustomWindow = MessageBox.Avalonia.MessageBoxManager
				   .GetMessageBoxCustomWindow(new MessageBoxCustomParams
				   {
					   ContentMessage = "Выберите все документы",
					   ButtonDefinitions = new[] {
							new ButtonDefinition {Name = "Ok"}
					   },
					   WindowStartupLocation = WindowStartupLocation.CenterOwner
				   });

				messageBoxCustomWindow.Show();

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
            var messageBox = MessageBox.Avalonia.MessageBoxManager
                .GetMessageBoxCustomWindow(new MessageBoxCustomParams
                {
                    ContentTitle = "Ошибка",
                    ContentMessage = message,
                    Icon = Icon.Error,
                    ButtonDefinitions = new[] { new ButtonDefinition { Name = "OK" } },
                    WindowStartupLocation = WindowStartupLocation.CenterOwner
                });

            await messageBox.Show();
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