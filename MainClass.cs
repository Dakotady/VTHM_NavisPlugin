using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using Autodesk.Navisworks.Api.Plugins;

namespace VthmNavisPlugin
{
	[Plugin("VTHalter", "VTRefresh", DisplayName = "VTHalter")]
	[RibbonLayout("VTHalter.xaml")]
	[RibbonTab("VTRefreshTab")]
	[Command("VTHM_RefreshButton", Icon = "Images\\VthmLogo16.png", LargeIcon = "Images\\VthmLogo32.png", CallCanExecute = CallCanExecute.Always, CanToggle = true, Shortcut = "Alt+R", ToolTip = "Refresh a selected files without needing to refresh entire document.", DisplayName = "Re-Merge File")]
	public class MainClass : CommandHandlerPlugin
	{
		public override int ExecuteCommand(string name, params string[] parameters)
		{
			if (name == "VTHM_RefreshButton")
			{
				ReMergeFile();
			}
			return 0;
		}

		public void ReMergeFile()
		{
			if (IsExpired())
			{
				return;
			}
			Document activeDocument = Autodesk.Navisworks.Api.Application.ActiveDocument;
			if (activeDocument.CurrentSelection.SelectedItems.Count <= 0)
			{
				return;
			}
			List<string> list = new List<string>();
			List<ModelItem> list2 = new List<ModelItem>();
			foreach (ModelItem selectedItem in activeDocument.CurrentSelection.SelectedItems)
			{
				list2.Add(selectedItem);
				while (list2[list2.Count - 1].Parent != null)
				{
					list2[list2.Count - 1] = list2[list2.Count - 1].Parent;
				}
			}
			activeDocument.CurrentSelection.Clear();
			foreach (ModelItem item in list2)
			{
				if (!list.Contains(item.Model.FileName))
				{
					activeDocument.CurrentSelection.Add(item);
					list.Add(item.Model.FileName);
				}
			}
			if (list.Count >= activeDocument.Models.Count)
			{
				MessageBox.Show("Use \"F5\" instead to refresh entire model.");
				activeDocument.CurrentSelection.Clear();
			}
			else
			{
				InwOpState10 state = ComApiBridge.State;
				state.DeleteSelectedFiles();
				activeDocument.TryMergeFiles(list);
			}
		}

		private bool IsExpired(string expiryDate = "Dec 31, 2022")
		{
			DateTime dateTime = DateTime.Parse(expiryDate, CultureInfo.InvariantCulture);
			DateTime now = DateTime.Now;
			//return dateTime < now; modified by dakota dyche 20220509
			return false;
		}
	}
}
