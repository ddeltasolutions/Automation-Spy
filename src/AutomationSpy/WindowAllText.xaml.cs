using System;
using System.Windows;
using Interop.UIAutomationClient;
using System.Windows.Threading;

namespace dDeltaSolutions.Spy
{
    /// <summary>
    /// Interaction logic for WindowAllText.xaml
    /// </summary>
    public partial class WindowAllText : Window
    {
		//public string AllText { get; set; }
	
        public WindowAllText(MainWindow.Attribute selectedAttribute)
        {
            InitializeComponent();
			
			if (selectedAttribute.Group == "UIA_ValuePattern" && selectedAttribute.Name == "Value:")
			{
				IUIAutomationValuePattern valuePattern = selectedAttribute.Pattern as IUIAutomationValuePattern;
				if (valuePattern != null)
				{
					txbTitle.Text = "IUIAutomationValuePattern.CurrentValue";
					txtAllText.Text = valuePattern.CurrentValue;
				}
			}
			else if (selectedAttribute.Group == "UIA_LegacyIAccessiblePattern" && selectedAttribute.Name == "Value:")
			{
				IUIAutomationLegacyIAccessiblePattern legacyIAccessiblePattern = selectedAttribute.Pattern as IUIAutomationLegacyIAccessiblePattern;
				if (legacyIAccessiblePattern != null)
				{
					txbTitle.Text = "IUIAutomationLegacyIAccessiblePattern.CurrentValue";
					txtAllText.Text = legacyIAccessiblePattern.CurrentValue;
				}
			}
			else if (selectedAttribute.Name == "Text:" && selectedAttribute.TextPatternRange != null)
			{
				txbTitle.Text = selectedAttribute.Group.Replace("UIA_", "IUIAutomation") + ".GetText(-1)";
				txtAllText.Text = selectedAttribute.TextPatternRange.GetText(-1);
			}
        }
		
		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			//txtAllText.Text = AllText;
		}
		
		/*public void SetTitle(string title)
		{
			txbTitle.Text = title;
		}*/
		
		private void OnCopyClick(object sender, RoutedEventArgs e)
		{
			Clipboard.SetText(txtAllText.Text);
			txtAllText.SelectAll();
			txtAllText.Focus();
			
			txbLabel.Visibility = Visibility.Visible;
			DispatcherTimer dispatcherTimer = new DispatcherTimer();
			dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
			dispatcherTimer.Interval = TimeSpan.FromSeconds(5);
			dispatcherTimer.Start();
		}
		
		private void dispatcherTimer_Tick(object sender, EventArgs e)
		{
			txbLabel.Visibility = Visibility.Hidden;
		
			DispatcherTimer dispatcherTimer = sender as DispatcherTimer;
			if (dispatcherTimer != null)
			{
				dispatcherTimer.Stop();
			}
		}
	}
}
