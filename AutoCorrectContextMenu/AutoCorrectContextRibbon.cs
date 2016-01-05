using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new AutoCorrectContextRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace AutoCorrectContextMenu
{
	[ComVisible(true)]
	public class AutoCorrectContextRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;
		private object missing = Type.Missing;

		public AutoCorrectContextRibbon()
		{
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("AutoCorrectContextMenu.AutoCorrectContextRibbon.xml");
		}

		#endregion

		#region Ribbon Callbacks
		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		public string AutoCorrectContextMenu_GetContent(Office.IRibbonControl control)
		{
			string word = GetCurrentWord();
			if (word == null)
				return String.Empty;

			SpellingSuggestions suggestions = GetSpellingSuggestions(word);

			return BuildMenu(word, suggestions);
		}

		public void AutoCorrectContextMenu_Correct(Office.IRibbonControl control)
		{
			var correction = SpellingCorrection.Decode(control.Tag);
			try
			{
				Globals.ThisAddIn.Application.AutoCorrect.Entries.Add(correction.Word, correction.Correction);
				Range range = Globals.ThisAddIn.Application.Selection.Words.First;
				range = GetTrimmedRange(range);
				range.Text = correction.Correction;
			}
			catch (Exception e)
			{
				MessageBox.Show($"Could not add AutoCorrect entry {correction}: {e.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		#endregion

		#region Helpers

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion


		private string GetCurrentWord()
		{
			var application = Globals.ThisAddIn.Application;

			var selection = application.Selection;
			if (selection == null)
				return null;

			string word = application.Selection.Words.First.Text;
			if (String.IsNullOrWhiteSpace(word))
				return null;

			return word;
		}

		private SpellingSuggestions GetSpellingSuggestions(string word)
		{
			return Globals.ThisAddIn.Application.GetSpellingSuggestions(
				word, ref missing, ref missing, ref missing,
				ref missing, ref missing, ref missing, ref missing,
				ref missing, ref missing, ref missing, ref missing,
				ref missing, ref missing);
		}

		private string BuildMenu(string word, SpellingSuggestions suggestions)
		{
			StringBuilder sb = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >");

			int index = 0;
			foreach (SpellingSuggestion suggestion in suggestions)
			{
				var correction = new SpellingCorrection(word, suggestion.Name);
				sb.Append($"<button id=\"AutoCorrectSuggestion{index}\" label=\"{suggestion.Name}\" onAction=\"AutoCorrectContextMenu_Correct\" tag=\"{SpellingCorrection.Encode(correction)}\" />");
				index += 1;
			}

			sb.Append(@"<menuSeparator id=""AutoCorrectSeparator"" />");
			sb.Append(@"<button idMso=""AutoCorrect"" />");
			sb.Append(@"</menu>");
			return sb.ToString();
		}

		private Range GetTrimmedRange(Range range)
		{
			string rangeText = range.Text;
			int start = range.Start;
			int end = range.End;
			while (rangeText[0] == ' ')
			{
				start++;
				rangeText = rangeText.Substring(1);
			}
			while (rangeText[rangeText.Length - 1] == ' ')
			{
				end--;
				rangeText = rangeText.Substring(0, rangeText.Length - 2);
			}
			range.Start = start;
			range.End = end;
			return range;
		}
	}
}
