using System;
using System.Windows;
using Interop.UIAutomationClient;
using System.Drawing;
using System.Collections.ObjectModel;

namespace dDeltaSolutions.Spy
{
    /// <summary>
    /// Interaction logic for MoreAttributes.xaml
    /// </summary>
    public partial class MoreAttributes : Window
    {
		private ObservableCollection<InternalAttribute> attributes = new ObservableCollection<InternalAttribute>();
		//private IUIAutomationTextRange TextRange = null;
	
        public MoreAttributes(IUIAutomationTextRange textRange)
        {
            InitializeComponent();
			
			try
			{
				lvAttributes.ItemsSource = attributes;
				GetMoreAttributes(textRange);
				lvAttributes.Items.Refresh();
				//TextRange = textRange;
			}
			catch {}
        }
		
		public class InternalAttribute
		{
			public string name;
			public string val;
		
			public string Name 
			{ 
				get
				{
					return name;
				}
				set
				{
					name = value;
				}
			}
			public string Val 
			{ 
				get
				{
					return val;
				}
				set
				{
					val = value;
				}
			}
			
			public InternalAttribute(string name, string val)
			{
				this.name = name;
				this.val = val;
			}
		}
		
		private void GetMoreAttributes(IUIAutomationTextRange textPatternRange)
		{
			object value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_AnimationStyleAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((AnimationStyle)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("AnimationStyle: ", 
					(value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_BackgroundColorAttributeId);
			}
			catch { }

			if (value != null)
			{
				if (value.GetType() == typeof(int))
				{
					//Color color = GetColorFromArgb((int)value);
					Color color = ColorTranslator.FromWin32((int)value);

					if (color.IsNamedColor == false)
					{
						attributes.Add(new InternalAttribute("BackgroundColor: ", "ARGB = (" + color.A.ToString() + ", " + 
							color.R.ToString() + ", " + color.G.ToString() + ", " + color.B.ToString() + ")"));
					}
					else
					{
						// is named
						attributes.Add(new InternalAttribute("BackgroundColor: ", color.Name));
					}
				}
				else
				{
					attributes.Add(new InternalAttribute("BackgroundColor: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
				}
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_BulletStyleAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((BulletStyle)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("BulletStyle: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_CapStyleAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((CapStyle)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("CapStyle: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_CultureAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("Culture: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_FontNameAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("FontName: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_FontSizeAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("FontSize: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_FontWeightAttributeId);
			}
			catch { }

			if (value != null)
			{
				try
				{
					int nValue = (int)value;
					string valueDesc = null;
					switch (nValue)
					{
						case 0:
							valueDesc = "DontCare";
							break;
						case 100:
							valueDesc = "Thin";
							break;
						case 200:
							valueDesc = "ExtraLight";
							break;
						case 300:
							valueDesc = "Light";
							break;
						case 400:
							valueDesc = "Normal";
							break;
						case 500:
							valueDesc = "Medium";
							break;
						case 600:
							valueDesc = "SemiBold";
							break;
						case 700:
							valueDesc = "Bold";
							break;
						case 800:
							valueDesc = "ExtraBold";
							break;
						case 900:
							valueDesc = "Heavy";
							break;
						default:
							break;
					}
					
					string sValue = value.ToString();
					if (valueDesc != null)
					{
						sValue += " (" + valueDesc + ")";
					}
					
					attributes.Add(new InternalAttribute("FontWeight: ", sValue));
				}
				catch
				{
					attributes.Add(new InternalAttribute("FontWeight: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
				}
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_ForegroundColorAttributeId);
			}
			catch { }

			if (value != null)
			{
				if (value.GetType() == typeof(int))
				{
					//Color color = GetColorFromArgb((int)value);
					Color color = ColorTranslator.FromWin32((int)value);

					if (color.IsNamedColor == false)
					{
						attributes.Add(new InternalAttribute("ForegroundColor: ", "ARGB = (" + color.A.ToString() + ", " +
							color.R.ToString() + ", " + color.G.ToString() + ", " + color.B.ToString() + ")"));
					}
					else
					{
						// is named
						attributes.Add(new InternalAttribute("ForegroundColor: ", color.Name));
					}
				}
				else
				{
					attributes.Add(new InternalAttribute("ForegroundColor: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
				}
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_HorizontalTextAlignmentAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((HorizontalTextAlignmentEnum)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("HorizontalTextAlignment: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IndentationFirstLineAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IndentationFirstLine: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IndentationLeadingAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IndentationLeading: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IndentationTrailingAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IndentationTrailing: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IsHiddenAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IsHidden: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IsItalicAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IsItalic: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IsReadOnlyAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IsReadOnly: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IsSubscriptAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IsSubscript: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_IsSuperscriptAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IsSuperscript: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_MarginBottomAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("MarginBottom: ", //value.ToString()));
					(value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_MarginLeadingAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("MarginLeading: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_MarginTopAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("MarginTop: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_MarginTrailingAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("MarginTrailing: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_OutlineStylesAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((OutlineStyles)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("OutlineStyles: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_OverlineColorAttributeId);
			}
			catch { }

			if (value != null)
			{
				if (value.GetType() == typeof(int))
				{
					//Color color = GetColorFromArgb((int)value);
					Color color = ColorTranslator.FromWin32((int)value);

					if (color.IsNamedColor == false)
					{
						attributes.Add(new InternalAttribute("OverlineColor: ", "ARGB = (" + color.A.ToString() + ", " +
							color.R.ToString() + ", " + color.G.ToString() + ", " + color.B.ToString() + ")"));
					}
					else
					{
						// is named
						attributes.Add(new InternalAttribute("OverlineColor: ", color.Name));
					}
				}
				else
				{
					attributes.Add(new InternalAttribute("OverlineColor: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
				}
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_OverlineStyleAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((TextDecorationLineStyle)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("OverlineStyle: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_StrikethroughColorAttributeId);
			}
			catch { }

			if (value != null)
			{
				if (value.GetType() == typeof(int))
				{
					//Color color = GetColorFromArgb((int)value);
					Color color = ColorTranslator.FromWin32((int)value);

					if (color.IsNamedColor == false)
					{
						attributes.Add(new InternalAttribute("StrikethroughColor: ", "ARGB = (" + color.A.ToString() + ", " +
							color.R.ToString() + ", " + color.G.ToString() + ", " + color.B.ToString() + ")"));
					}
					else
					{
						// is named
						attributes.Add(new InternalAttribute("StrikethroughColor: ", color.Name));
					}
				}
				else
				{
					attributes.Add(new InternalAttribute("StrikethroughColor: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
				}
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_StrikethroughStyleAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((TextDecorationLineStyle)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("StrikethroughStyle: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_TabsAttributeId);
			}
			catch { }

			if (value != null)
			{
				if (value.GetType() == typeof(double[]))
				{
					double[] doubleArray = (double[])value;
					if (doubleArray.Length == 0)
					{
						attributes.Add(new InternalAttribute("Tabs: ", "-"));
					}
					else
					{
						attributes.Add(new InternalAttribute("Tabs: ", MainWindow.DoubleArrayToString(doubleArray)));
					}
				}
				else
				{
					attributes.Add(new InternalAttribute("Tabs: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
				}
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_TextFlowDirectionsAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((FlowDirections)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}			
				attributes.Add(new InternalAttribute("TextFlowDirections: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_UnderlineColorAttributeId);
			}
			catch { }

			if (value != null)
			{
				if (value.GetType() == typeof(int))
				{
					//Color color = GetColorFromArgb((int)value);
					Color color = ColorTranslator.FromWin32((int)value);

					if (color.IsNamedColor == false)
					{
						attributes.Add(new InternalAttribute("UnderlineColor: ", "ARGB = (" + color.A.ToString() + ", " +
							color.R.ToString() + ", " + color.G.ToString() + ", " + color.B.ToString() + ")"));
					}
					else
					{
						// is named
						attributes.Add(new InternalAttribute("UnderlineColor: ", color.Name));
					}
				}
				else
				{
					attributes.Add(new InternalAttribute("UnderlineColor: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
				}
			}

			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(
					UIA_TextAttributeIds.UIA_UnderlineStyleAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((TextDecorationLineStyle)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("UnderlineStyle: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_AnnotationTypesAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("AnnotationTypes: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_AnnotationObjectsAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("AnnotationObjects: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_StyleNameAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("StyleName: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_StyleIdAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((StyleId)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
				attributes.Add(new InternalAttribute("StyleId: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_LinkAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("Link: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_IsActiveAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("IsActive: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_SelectionActiveEndAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((ActiveEnd)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}
			
				attributes.Add(new InternalAttribute("SelectionActiveEnd: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_CaretPositionAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((CaretPosition)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}			
				attributes.Add(new InternalAttribute("CaretPosition: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_CaretBidiModeAttributeId);
			}
			catch { }

			if (value != null)
			{
				string strValue = null;
				try
				{
					strValue = ((CaretBidiMode)value).ToString() + " (" + value + ")";
				}
				catch
				{
					strValue = value.ToString();
				}			
				attributes.Add(new InternalAttribute("CaretBidiMode: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : strValue)));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_LineSpacingAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("LineSpacing: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_BeforeParagraphSpacingAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("BeforeParagraphSpacing: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_AfterParagraphSpacingAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("AfterParagraphSpacing: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
			
			value = null;
			try
			{
				value = textPatternRange.GetAttributeValue(UIA_TextAttributeIds.UIA_SayAsInterpretAsAttributeId);
			}
			catch { }

			if (value != null)
			{
				attributes.Add(new InternalAttribute("SayAsInterpretAs: ", (value == MainWindow.uiAutomation.ReservedNotSupportedValue) ? "ReservedNotSupportedValue" : (value == MainWindow.uiAutomation.ReservedMixedAttributeValue ? "ReservedMixedAttributeValue" : value.ToString())));
			}
		}
		
		internal enum AnimationStyle
		{
			AnimationStyle_None = 0,
			AnimationStyle_LasVegasLights = 1,
			AnimationStyle_BlinkingBackground = 2,
			AnimationStyle_SparkleText = 3,
			AnimationStyle_MarchingBlackAnts = 4,
			AnimationStyle_MarchingRedAnts = 5,
			AnimationStyle_Shimmer = 6,
			AnimationStyle_Other = -1
		}
		
		internal enum BulletStyle
		{
			BulletStyle_None = 0,
			BulletStyle_HollowRoundBullet = 1,
			BulletStyle_FilledRoundBullet = 2,
			BulletStyle_HollowSquareBullet = 3,
			BulletStyle_FilledSquareBullet = 4,
			BulletStyle_DashBullet = 5,
			BulletStyle_Other = -1
		}
		
		internal enum CapStyle
		{
			CapStyle_None = 0,
			CapStyle_SmallCap = 1,
			CapStyle_AllCap = 2,
			CapStyle_AllPetiteCaps = 3,
			CapStyle_PetiteCaps = 4,
			CapStyle_Unicase = 5,
			CapStyle_Titling = 6,
			CapStyle_Other = -1
		}
		
		internal enum CaretBidiMode
		{
			CaretBidiMode_LTR = 0,
			CaretBidiMode_RTL = 1
		}
		
		internal enum CaretPosition
		{
			CaretPosition_Unknown = 0,
			CaretPosition_EndOfLine = 1,
			CaretPosition_BeginningOfLine = 2
		}
		
		internal enum HorizontalTextAlignmentEnum
		{
			HorizontalTextAlignment_Left, 
			HorizontalTextAlignment_Centered, 
			HorizontalTextAlignment_Right, 
			HorizontalTextAlignment_Justified 
		}
		
		internal enum OutlineStyles
		{
			OutlineStyles_None = 0,
			OutlineStyles_Outline = 1,
			OutlineStyles_Shadow = 2,
			OutlineStyles_Engraved = 4,
			OutlineStyles_Embossed = 8
		}
		
		internal enum TextDecorationLineStyle
		{
			TextDecorationLineStyle_None = 0,
			TextDecorationLineStyle_Single = 1,
			TextDecorationLineStyle_WordsOnly = 2,
			TextDecorationLineStyle_Double = 3,
			TextDecorationLineStyle_Dot = 4,
			TextDecorationLineStyle_Dash = 5,
			TextDecorationLineStyle_DashDot = 6,
			TextDecorationLineStyle_DashDotDot = 7,
			TextDecorationLineStyle_Wavy = 8,
			TextDecorationLineStyle_ThickSingle = 9,
			TextDecorationLineStyle_DoubleWavy = 11,
			TextDecorationLineStyle_ThickWavy = 12,
			TextDecorationLineStyle_LongDash = 13,
			TextDecorationLineStyle_ThickDash = 14,
			TextDecorationLineStyle_ThickDashDot = 15,
			TextDecorationLineStyle_ThickDashDotDot = 16,
			TextDecorationLineStyle_ThickDot = 17,
			TextDecorationLineStyle_ThickLongDash = 18,
			TextDecorationLineStyle_Other = -1
		}
		
		internal enum ActiveEnd
		{
			ActiveEnd_None = 0,
			ActiveEnd_Start = 1,
			ActiveEnd_End = 2
		}
		
		internal enum StyleId
		{
			StyleId_BulletedList = 70015,
			StyleId_Custom = 70000,
			StyleId_Emphasis = 70013,
			StyleId_Heading1 = 70001,
			StyleId_Heading2 = 70002,
			StyleId_Heading3 = 70003,
			StyleId_Heading4 = 70004,
			StyleId_Heading5 = 70005,
			StyleId_Heading6 = 70006,
			StyleId_Heading7 = 70007,
			StyleId_Heading8 = 70008,
			StyleId_Heading9 = 70009,
			StyleId_Normal = 70012,
			StyleId_NumberedList = 70016,
			StyleId_Quote = 70014,
			StyleId_Subtitle = 70011,
			StyleId_Title = 70010
		}
		
		internal enum FlowDirections
		{
			FlowDirections_Default = 0,
			FlowDirections_RightToLeft = 0x1,
			FlowDirections_BottomToTop = 0x2,
			FlowDirections_Vertical = 0x4
		}
    }
}
