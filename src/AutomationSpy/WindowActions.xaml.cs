using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using Interop.UIAutomationClient;

namespace dDeltaSolutions.Spy
{
    /// <summary>
    /// Interaction logic for WindowActions.xaml
    /// </summary>
    public partial class WindowActions : Window
    {
        public WindowActions(TreeNode node)
        {
            InitializeComponent();
			
			this.node = node;
        }
		
		private TreeNode node = null;
		
		private void OnKeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Escape)
			{
				Close();
			}
		}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
			object pattern = null;
			
            IUIAutomationDockPattern dockPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_DockPatternId) as IUIAutomationDockPattern;
			if (dockPattern != null)
			{
				dockTab.Visibility = Visibility.Visible;
				dockTab.Tag = pattern;
				
                try
                {
                    DockPosition pos = dockPattern.CurrentDockPosition;
                    if (pos == DockPosition.DockPosition_None)
                    {
                        cmbDocks.SelectedIndex = 0;
                    }
                    else if (pos == DockPosition.DockPosition_Left)
                    {
                        cmbDocks.SelectedIndex = 1;
                    }
                    else if (pos == DockPosition.DockPosition_Right)
                    {
                        cmbDocks.SelectedIndex = 2;
                    }
                    else if (pos == DockPosition.DockPosition_Top)
                    {
                        cmbDocks.SelectedIndex = 3;
                    }
                    else if (pos == DockPosition.DockPosition_Bottom)
                    {
                        cmbDocks.SelectedIndex = 4;
                    }
                    else if (pos == DockPosition.DockPosition_Fill)
                    {
                        cmbDocks.SelectedIndex = 5;
                    }
                }
                catch { }
			}
			
            IUIAutomationExpandCollapsePattern expandCollapsePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId) as IUIAutomationExpandCollapsePattern;
            if (expandCollapsePattern != null)
			{
				expandCollapseTab.Visibility = Visibility.Visible;
				expandCollapseTab.Tag = expandCollapsePattern;
			}
            
            IUIAutomationGridPattern gridPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_GridPatternId) as IUIAutomationGridPattern;
			if (gridPattern != null)
			{
				gridTab.Visibility = Visibility.Visible;
				gridTab.Tag = gridPattern;
			}
			
            IUIAutomationInvokePattern invokePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId) as IUIAutomationInvokePattern;
			if (invokePattern != null)
			{
				invokeTab.Visibility = Visibility.Visible;
				invokeTab.Tag = invokePattern;
			}
			
            IUIAutomationMultipleViewPattern multipleViewPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_MultipleViewPatternId) as IUIAutomationMultipleViewPattern;
			if (multipleViewPattern != null)
			{
				mvTab.Visibility = Visibility.Visible;
				mvTab.Tag = multipleViewPattern;
				
                try
                {
                    Array ids = multipleViewPattern.GetCurrentSupportedViews();
                    int selectedId = multipleViewPattern.CurrentCurrentView;
                    int selectedIndex = -1;
                    
                    for (int i = 0; i < ids.Length; i++)
                    {
                        int id = (int)ids.GetValue(i);
                        if (id == selectedId)
                        {
                            selectedIndex = i;
                        }
                        
                        string viewName = multipleViewPattern.GetViewName(id);
                        ViewItem viewItem = new ViewItem(id, viewName);
                        cmbViews.Items.Add(viewItem);
                    }
                    
                    if (selectedIndex >= 0)
                    {
                        cmbViews.SelectedIndex = selectedIndex;
                    }
                }
                catch { }
			}
			
            IUIAutomationRangeValuePattern rangeValuePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId) as IUIAutomationRangeValuePattern;
			if (rangeValuePattern != null)
			{
				rvTab.Visibility = Visibility.Visible;
				rvTab.Tag = rangeValuePattern;
				
                try
                {
                    txtRValue.Text = rangeValuePattern.CurrentValue.ToString();
                }
                catch { }
			}
            
            IUIAutomationScrollPattern scrollPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ScrollPatternId) as IUIAutomationScrollPattern;
			if (scrollPattern != null)
			{
				scrollTab.Visibility = Visibility.Visible;
				scrollTab.Tag = scrollPattern;
			}
			
            IUIAutomationScrollItemPattern scrollItemPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ScrollItemPatternId) as IUIAutomationScrollItemPattern;
			if (scrollItemPattern != null)
			{
				siTab.Visibility = Visibility.Visible;
				siTab.Tag = scrollItemPattern;
			}
			
            IUIAutomationSelectionItemPattern selectionItemPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) as IUIAutomationSelectionItemPattern;
			if (selectionItemPattern != null)
			{
				selITab.Visibility = Visibility.Visible;
				selITab.Tag = selectionItemPattern;
			}
			
            IUIAutomationTextPattern textPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId) as IUIAutomationTextPattern;
			if (textPattern != null)
			{
				this.Width += 125;
                this.Height += 40;
				
				textTab.Visibility = Visibility.Visible;
				textTab.Tag = textPattern;
				
                try
                {
                    TPRange range = new TPRange("DocumentRange", textPattern.DocumentRange);
                    cmbRanges.Items.Add(range);
                    
                    IUIAutomationTextRangeArray selRanges = textPattern.GetSelection();
                    for (int i = 0; i < selRanges.Length; i++)
                    {
                        IUIAutomationTextRange selRange = selRanges.GetElement(i);
                        range = new TPRange("Selection[" + i + "]", selRange);
                        cmbRanges.Items.Add(range);
                    }
                }
                catch { }
				
				cmbTextUnit.Items.Add(TextUnit.TextUnit_Character);
				cmbTextUnit.Items.Add(TextUnit.TextUnit_Format);
				cmbTextUnit.Items.Add(TextUnit.TextUnit_Word);
				cmbTextUnit.Items.Add(TextUnit.TextUnit_Line);
				cmbTextUnit.Items.Add(TextUnit.TextUnit_Paragraph);
				cmbTextUnit.Items.Add(TextUnit.TextUnit_Page);
				cmbTextUnit.Items.Add(TextUnit.TextUnit_Document);
				
				cmbTextUnitM.Items.Add(TextUnit.TextUnit_Character);
				cmbTextUnitM.Items.Add(TextUnit.TextUnit_Format);
				cmbTextUnitM.Items.Add(TextUnit.TextUnit_Word);
				cmbTextUnitM.Items.Add(TextUnit.TextUnit_Line);
				cmbTextUnitM.Items.Add(TextUnit.TextUnit_Paragraph);
				cmbTextUnitM.Items.Add(TextUnit.TextUnit_Page);
				cmbTextUnitM.Items.Add(TextUnit.TextUnit_Document);
				
				cmbTextUnitME.Items.Add(TextUnit.TextUnit_Character);
				cmbTextUnitME.Items.Add(TextUnit.TextUnit_Format);
				cmbTextUnitME.Items.Add(TextUnit.TextUnit_Word);
				cmbTextUnitME.Items.Add(TextUnit.TextUnit_Line);
				cmbTextUnitME.Items.Add(TextUnit.TextUnit_Paragraph);
				cmbTextUnitME.Items.Add(TextUnit.TextUnit_Page);
				cmbTextUnitME.Items.Add(TextUnit.TextUnit_Document);
				
				cmbEndpoint.Items.Add(TextPatternRangeEndpoint.TextPatternRangeEndpoint_Start);
				cmbEndpoint.Items.Add(TextPatternRangeEndpoint.TextPatternRangeEndpoint_End);
				
				pickBtn.ToolTip = "Pick a TextPatternRange from screen using the mouse cursor." + Environment.NewLine + "Press this button and go with the mouse over the desired text you want to pick and there hold Ctrl key pressed a little while." + Environment.NewLine + "If you change your mind and you don't want to pick anything then press again this button.";
			}
			
            IUIAutomationTogglePattern togglePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;
			if (togglePattern != null)
			{
				toggleTab.Visibility = Visibility.Visible;
				toggleTab.Tag = togglePattern;
			}
			
            IUIAutomationTransformPattern transformPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TransformPatternId) as IUIAutomationTransformPattern;
			if (transformPattern != null)
			{
				transformTab.Visibility = Visibility.Visible;
				transformTab.Tag = transformPattern;
				
				try
				{
					RECT rect;
					if (GetWindowRect(node.Element.CurrentNativeWindowHandle, out rect) == true)
					{
						txtX.Text = rect.Left.ToString();
						txtY.Text = rect.Top.ToString();
						
						txtWidth.Text = (rect.Right - rect.Left).ToString();
						txtHeight.Text = (rect.Bottom - rect.Top).ToString();
					}
				}
				catch { }
			}
			
            IUIAutomationValuePattern valuePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;
			if (valuePattern != null)
			{
				valueTab.Visibility = Visibility.Visible;
				valueTab.Tag = valuePattern;
				
                try
                {
                    txtValue.Text = valuePattern.CurrentValue;
                }
                catch { }
			}
			
            IUIAutomationWindowPattern windowPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId) as IUIAutomationWindowPattern;
			if (windowPattern != null)
			{
				windowTab.Visibility = Visibility.Visible;
				windowTab.Tag = windowPattern;
				
                try
                {
                    if (windowPattern.CurrentWindowVisualState == WindowVisualState.WindowVisualState_Normal)
                    {
                        cmbVisualStates.SelectedIndex = 0;
                    }
                    else if (windowPattern.CurrentWindowVisualState == WindowVisualState.WindowVisualState_Minimized)
                    {
                        cmbVisualStates.SelectedIndex = 1;
                    }
                    else if (windowPattern.CurrentWindowVisualState == WindowVisualState.WindowVisualState_Maximized)
                    {
                        cmbVisualStates.SelectedIndex = 2;
                    }
                }
                catch { }
			}
            
            IUIAutomationVirtualizedItemPattern virtualizedItemPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_VirtualizedItemPatternId) as IUIAutomationVirtualizedItemPattern;
			if (virtualizedItemPattern != null)
			{
				virtualizeTab.Visibility = Visibility.Visible;
				virtualizeTab.Tag = virtualizedItemPattern;
			}
            
            IUIAutomationLegacyIAccessiblePattern legacyIAccPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) as IUIAutomationLegacyIAccessiblePattern;
			if (legacyIAccPattern != null)
			{
				legacyIAccTab.Visibility = Visibility.Visible;
				legacyIAccTab.Tag = legacyIAccPattern;
                
                try
                {
                    txtIAccValue.Text = legacyIAccPattern.CurrentValue;
                }
                catch {}
			}
			
			List<TabItem> itemsToRemove = new List<TabItem>();
			foreach (TabItem item in tabCtrl.Items)
			{
				if (item.Visibility == Visibility.Hidden)
				{
					itemsToRemove.Add(item);
				}
			}
			foreach (TabItem item in itemsToRemove)
			{
				tabCtrl.Items.Remove(item);
			}
			
			if (tabCtrl.Items.Count == 0)
			{
				grid.Children.Remove(tabCtrl);
				Label label = new Label();
				label.Content = "No actions available";
				label.Margin = new Thickness(50, 30, 0, 0);
				grid.Children.Add(label);
			}
			else
			{
				tabCtrl.SelectedIndex = 0;
			}
        }
		
		private void OnInvoke(object sender, RoutedEventArgs e)
		{
			IUIAutomationInvokePattern invokePattern = invokeTab.Tag as IUIAutomationInvokePattern;
			if (invokePattern == null)
			{
				return;
			}
			
			try
			{
				invokePattern.Invoke();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnExpand(object sender, RoutedEventArgs e)
		{
			IUIAutomationExpandCollapsePattern ecPattern = expandCollapseTab.Tag as IUIAutomationExpandCollapsePattern;
			if (ecPattern == null)
			{
				return;
			}
			
			try
			{
				ecPattern.Expand();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnCollapse(object sender, RoutedEventArgs e)
		{
			IUIAutomationExpandCollapsePattern ecPattern = expandCollapseTab.Tag as IUIAutomationExpandCollapsePattern;
			if (ecPattern == null)
			{
				return;
			}
			
			try
			{
				ecPattern.Collapse();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnDock(object sender, RoutedEventArgs e)
		{
			IUIAutomationDockPattern dockPattern = dockTab.Tag as IUIAutomationDockPattern;
			if (dockPattern == null)
			{
				return;
			}
			
			DockPosition pos = DockPosition.DockPosition_None;
			if (cmbDocks.SelectedIndex == 0)
			{
				pos = DockPosition.DockPosition_None;
			}
			else if (cmbDocks.SelectedIndex == 1)
			{
				pos = DockPosition.DockPosition_Left;
			}
			else if (cmbDocks.SelectedIndex == 2)
			{
				pos = DockPosition.DockPosition_Right;
			}
			else if (cmbDocks.SelectedIndex == 3)
			{
				pos = DockPosition.DockPosition_Top;
			}
			else if (cmbDocks.SelectedIndex == 4)
			{
				pos = DockPosition.DockPosition_Bottom;
			}
			else if (cmbDocks.SelectedIndex == 5)
			{
				pos = DockPosition.DockPosition_Fill;
			}
			
			try
			{
				dockPattern.SetDockPosition(pos);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnSetCurrentView(object sender, RoutedEventArgs e)
		{
			IUIAutomationMultipleViewPattern mvPattern = mvTab.Tag as IUIAutomationMultipleViewPattern;
			if (mvPattern == null)
			{
				return;
			}
			
			ViewItem viewItem = cmbViews.SelectedItem as ViewItem;
			if (viewItem == null)
			{
				return;
			}
			
			try
			{
				mvPattern.SetCurrentView(viewItem.id);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnSetRValue(object sender, RoutedEventArgs e)
		{
			IUIAutomationRangeValuePattern rvPattern = rvTab.Tag as IUIAutomationRangeValuePattern;
			if (rvPattern == null)
			{
				return;
			}
			
			try
			{
				double val = double.Parse(txtRValue.Text);
				rvPattern.SetValue(val);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
        
        private void OnScroll(object sender, RoutedEventArgs e)
        {
            IUIAutomationScrollPattern scrollPattern = scrollTab.Tag as IUIAutomationScrollPattern;
			if (scrollPattern == null)
			{
				return;
			}
            
            ScrollAmount horizontally = ScrollAmount.ScrollAmount_NoAmount;
            if (cmbHorizontalAmt.SelectedIndex == 0)
            {
                horizontally = ScrollAmount.ScrollAmount_LargeDecrement;
            }
            else if (cmbHorizontalAmt.SelectedIndex == 1)
            {
                horizontally = ScrollAmount.ScrollAmount_SmallDecrement;
            }
            else if (cmbHorizontalAmt.SelectedIndex == 3)
            {
                horizontally = ScrollAmount.ScrollAmount_LargeIncrement;
            }
            else if (cmbHorizontalAmt.SelectedIndex == 4)
            {
                horizontally = ScrollAmount.ScrollAmount_SmallIncrement;
            }
            
            ScrollAmount vertically = ScrollAmount.ScrollAmount_NoAmount;
            if (cmbVerticalAmt.SelectedIndex == 0)
            {
                vertically = ScrollAmount.ScrollAmount_LargeDecrement;
            }
            else if (cmbVerticalAmt.SelectedIndex == 1)
            {
                vertically = ScrollAmount.ScrollAmount_SmallDecrement;
            }
            else if (cmbVerticalAmt.SelectedIndex == 3)
            {
                vertically = ScrollAmount.ScrollAmount_LargeIncrement;
            }
            else if (cmbVerticalAmt.SelectedIndex == 4)
            {
                vertically = ScrollAmount.ScrollAmount_SmallIncrement;
            }
            
            try
            {
                scrollPattern.Scroll(horizontally, vertically);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, ex.Message);
            }
        }
        
        private void OnSetScrollPercent(object sender, RoutedEventArgs e)
        {
            IUIAutomationScrollPattern scrollPattern = scrollTab.Tag as IUIAutomationScrollPattern;
			if (scrollPattern == null)
			{
				return;
			}
            
            double horizontalPercent = UIA_ScrollPattern.UIA_ScrollPatternNoScroll;
			string horizontalText = txtHorizontalPercent.Text.Trim();
			
            if (horizontalText != "" && double.TryParse(horizontalText, out horizontalPercent) == false)
            {
                System.Windows.MessageBox.Show(this, "Horizontal Percent is not a valid number");
				txtHorizontalPercent.Focus();
                return;
            }
            
            double verticalPercent = UIA_ScrollPattern.UIA_ScrollPatternNoScroll;
			string verticalText = txtVerticalPercent.Text.Trim();
			
            if (verticalText != "" && double.TryParse(verticalText, out verticalPercent) == false)
            {
                System.Windows.MessageBox.Show(this, "Vertical Percent is not a valid number");
				txtVerticalPercent.Focus();
                return;
            }
            
            if (horizontalPercent != UIA_ScrollPattern.UIA_ScrollPatternNoScroll && 
				(horizontalPercent < 0 || horizontalPercent > 100))
            {
                System.Windows.MessageBox.Show(this, "Horizontal Percent must be between 0 and 100");
				txtHorizontalPercent.Focus();
                return;
            }
            
            if (verticalPercent != UIA_ScrollPattern.UIA_ScrollPatternNoScroll && 
				(verticalPercent < 0 || verticalPercent > 100))
            {
                System.Windows.MessageBox.Show(this, "Vertical Percent must be between 0 and 100");
				txtVerticalPercent.Focus();
                return;
            }
            
            try
            {
                scrollPattern.SetScrollPercent(horizontalPercent, verticalPercent);
            }
            catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
        }
		
		private void OnScrollIntoView(object sender, RoutedEventArgs e)
		{
			IUIAutomationScrollItemPattern siPattern = siTab.Tag as IUIAutomationScrollItemPattern;
			if (siPattern == null)
			{
				return;
			}
			
			try
			{
				siPattern.ScrollIntoView();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnAddToSelection(object sender, RoutedEventArgs e)
		{
			IUIAutomationSelectionItemPattern selIPattern = selITab.Tag as IUIAutomationSelectionItemPattern;
			if (selIPattern == null)
			{
				return;
			}
			
			try
			{
				selIPattern.AddToSelection();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnRemoveFromSelection(object sender, RoutedEventArgs e)
		{
			IUIAutomationSelectionItemPattern selIPattern = selITab.Tag as IUIAutomationSelectionItemPattern;
			if (selIPattern == null)
			{
				return;
			}
			
			try
			{
				selIPattern.RemoveFromSelection();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnSelect(object sender, RoutedEventArgs e)
		{
			IUIAutomationSelectionItemPattern selIPattern = selITab.Tag as IUIAutomationSelectionItemPattern;
			if (selIPattern == null)
			{
				return;
			}
			
			try
			{
				selIPattern.Select();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnToggle(object sender, RoutedEventArgs e)
		{
			IUIAutomationTogglePattern togglePattern = toggleTab.Tag as IUIAutomationTogglePattern;
			if (togglePattern == null)
			{
				return;
			}
			
			try
			{
				togglePattern.Toggle();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnMove(object sender, RoutedEventArgs e)
		{
			IUIAutomationTransformPattern transformPattern = transformTab.Tag as IUIAutomationTransformPattern;
			if (transformPattern == null)
			{
				return;
			}
			
			try
			{
				double x = 0, y = 0;
				x = double.Parse(txtX.Text);
				y = double.Parse(txtY.Text);
				transformPattern.Move(x, y);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnResize(object sender, RoutedEventArgs e)
		{
			IUIAutomationTransformPattern transformPattern = transformTab.Tag as IUIAutomationTransformPattern;
			if (transformPattern == null)
			{
				return;
			}
			
			try
			{
				double width = 0, height = 0;
				width = double.Parse(txtWidth.Text);
				height = double.Parse(txtHeight.Text);
				transformPattern.Resize(width, height);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnRotate(object sender, RoutedEventArgs e)
		{
			IUIAutomationTransformPattern transformPattern = transformTab.Tag as IUIAutomationTransformPattern;
			if (transformPattern == null)
			{
				return;
			}
			
			try
			{
				double degrees = 0;
				degrees = double.Parse(txtDegrees.Text);
				transformPattern.Rotate(degrees);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnSetValue(object sender, RoutedEventArgs e)
		{
			IUIAutomationValuePattern valuePattern = valueTab.Tag as IUIAutomationValuePattern;
			if (valuePattern == null)
			{
				return;
			}
			
			try
			{
				valuePattern.SetValue(txtValue.Text);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnSetWindowVisualState(object sender, RoutedEventArgs e)
		{
			IUIAutomationWindowPattern windowPattern = windowTab.Tag as IUIAutomationWindowPattern;
			if (windowPattern == null)
			{
				return;
			}
			
			try
			{
				WindowVisualState visualState = WindowVisualState.WindowVisualState_Normal;
				if (cmbVisualStates.SelectedIndex == 1)
				{
					visualState = WindowVisualState.WindowVisualState_Minimized;
				}
				else if (cmbVisualStates.SelectedIndex == 2)
				{
					visualState = WindowVisualState.WindowVisualState_Maximized;
				}
				
				windowPattern.SetWindowVisualState(visualState);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnTextAddToSelection(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				rangeItem.range.AddToSelection();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnTextRemoveFromSelection(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				rangeItem.range.RemoveFromSelection();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnTextScrollIntoView(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
                int alignToTop = 0;
                if (chkAlign.IsChecked.Value)
                {
                    alignToTop = 1;
                }
				rangeItem.range.ScrollIntoView(alignToTop);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnTextSelect(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				rangeItem.range.Select();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnExpandToEnclosingUnit(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				TextUnit unit = (TextUnit)cmbTextUnit.SelectedItem;
				rangeItem.range.ExpandToEnclosingUnit(unit);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
				return;
			}
			
			System.Windows.MessageBox.Show(this, "Succeeded");
		}
		
		private void OnTextMove(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				TextUnit unit = (TextUnit)cmbTextUnitM.SelectedItem;
				int count = 0;
				if (int.TryParse(txtCount.Text, out count) == false)
				{
					System.Windows.MessageBox.Show(this, "Count is not a valid integer");
					txtCount.Focus();
					txtCount.SelectAll();
					return;
				}
				
				int moved = rangeItem.range.Move(unit, count);
				System.Windows.MessageBox.Show(this, "This method returned: " + moved);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnMoveEndpointByUnit(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				TextPatternRangeEndpoint endpoint = (TextPatternRangeEndpoint)cmbEndpoint.SelectedItem;
				TextUnit unit = (TextUnit)cmbTextUnitME.SelectedItem;
				int count = 0;
				if (int.TryParse(txtCountME.Text, out count) == false)
				{
					System.Windows.MessageBox.Show(this, "Count is not a valid integer");
					txtCountME.Focus();
					txtCountME.SelectAll();
					return;
				}
				
				int moved = rangeItem.range.MoveEndpointByUnit(endpoint, unit, count);
				System.Windows.MessageBox.Show(this, "This method returned: " + moved);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private Brush defaultBrush;
		private System.Windows.Forms.Timer timer = null;
		private void PickFromScreen(object sender, RoutedEventArgs e)
		{
			try
			{
				if (pickBtn.IsChecked == true)
				{
					defaultBrush = pickBtn.Background;
					pickBtn.Foreground = Brushes.Red;
					
					if (timer == null)
					{
						timer = new System.Windows.Forms.Timer();
						timer.Interval = 1000;
						timer.Tick += timer_Tick;
					}
					
					timer.Start();
				}
				else
				{
					pickBtn.Foreground = Brushes.Black;
					pickBtn.Background = defaultBrush;
					
					timer.Stop();
				}
			}
			catch (Exception ex)
			{ 
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private int rangeCount = 0;
		void timer_Tick(object sender, EventArgs e)
		{
			try
			{
				bool leftCtrlIsPressed = ((int)Keyboard.GetKeyStates(Key.LeftCtrl) & (int)KeyStates.Down) != 0;
				bool rightCtrlIsPressed = ((int)Keyboard.GetKeyStates(Key.RightCtrl) & (int)KeyStates.Down) != 0;

				if (leftCtrlIsPressed || rightCtrlIsPressed)
				{
					IUIAutomationTextPattern textPattern = textTab.Tag as IUIAutomationTextPattern;
					if (textPattern != null)
					{
						System.Drawing.Point ptCursor = System.Windows.Forms.Cursor.Position;
						tagPOINT pt;
                        pt.x = ptCursor.X;
                        pt.y = ptCursor.Y;
						IUIAutomationTextRange range = textPattern.RangeFromPoint(pt);
						pickBtn.IsChecked = false;
						PickFromScreen(null, null);
						
						rangeCount++;
						System.Windows.MessageBox.Show(this, "You picked a TextPatternRange named Range" + rangeCount + ". For this range you should call method ExpandToEnclosingUnit using TextUnit: Character, Word, Line, Paragraph or anything else to actually get something. Otherwise, you will get an empty TextPatternRange.");
						
						TPRange tpRange = new TPRange("Range" + rangeCount, range);
						cmbRanges.Items.Add(tpRange);
						cmbRanges.SelectedItem = tpRange;
					}
				}
			}
			catch { }
		}
		
		private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (timer != null)
			{
				try
				{
					timer.Stop();
				}
				catch { }
			}
		}
		
		private void OnGetText(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				System.Windows.MessageBox.Show(this, rangeItem.range.GetText(3000));
			}
			catch (Exception ex)
			{ 
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnGetEnclosingElement(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				TreeNode node = new TreeNode(rangeItem.range.GetEnclosingElement());
				System.Windows.MessageBox.Show(this, node.ToString());
			}
			catch (Exception ex)
			{ 
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnHighlightRange(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			try
			{
				Array boundingRectangles = rangeItem.range.GetBoundingRectangles();
				if (boundingRectangles != null && boundingRectangles.Length >= 4)
				{
					tagRECT boundingRectangle;
					boundingRectangle.left = (int)(double)boundingRectangles.GetValue(0);
					boundingRectangle.top = (int)(double)boundingRectangles.GetValue(1);
					int width = (int)(double)boundingRectangles.GetValue(2);
					boundingRectangle.right = boundingRectangle.left + width;
					int height = (int)(double)boundingRectangles.GetValue(3);
					boundingRectangle.bottom = boundingRectangle.top + height;

					MainWindow.HighlightRect(boundingRectangle);
				}
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(this, ex.Message);
			}
		}
		
		private void OnRangeProperties(object sender, RoutedEventArgs e)
		{
			TPRange rangeItem = cmbRanges.SelectedItem as TPRange;
			if (rangeItem == null)
			{
				return;
			}
			
			MoreAttributes moreAttributes = new MoreAttributes(rangeItem.range) { Owner = this };
			moreAttributes.ShowDialog();
		}
        
        private void OnRealize(object sender, RoutedEventArgs e)
        {
            IUIAutomationVirtualizedItemPattern virtualizedItemPattern = virtualizeTab.Tag as IUIAutomationVirtualizedItemPattern;
            if (virtualizedItemPattern == null)
            {
                return;
            }
            
            try
            {
                virtualizedItemPattern.Realize();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, ex.Message);
            }
        }
        
        private void OnIAccSelect(object sender, RoutedEventArgs e)
        {
            IUIAutomationLegacyIAccessiblePattern legacyIAccPattern = legacyIAccTab.Tag as IUIAutomationLegacyIAccessiblePattern;
            if (legacyIAccPattern == null)
            {
                return;
            }
            
            int flags = 0;
            try
            {
                flags = Convert.ToInt32(txtIAccFlags.Text);
            }
            catch
            {
                System.Windows.MessageBox.Show(this, "Flags field must be an integer");
                return;
            }
            
            try
            {
                legacyIAccPattern.Select(flags);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, ex.Message);
            }
        }
        
        private void OnIAccDoDefaultAction(object sender, RoutedEventArgs e)
        {
            IUIAutomationLegacyIAccessiblePattern legacyIAccPattern = legacyIAccTab.Tag as IUIAutomationLegacyIAccessiblePattern;
            if (legacyIAccPattern == null)
            {
                return;
            }
            
            try
            {
                legacyIAccPattern.DoDefaultAction();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, ex.Message);
            }
        }
        
        private void OnIAccSetValue(object sender, RoutedEventArgs e)
        {
            IUIAutomationLegacyIAccessiblePattern legacyIAccPattern = legacyIAccTab.Tag as IUIAutomationLegacyIAccessiblePattern;
            if (legacyIAccPattern == null)
            {
                return;
            }
            
            try
            {
                legacyIAccPattern.SetValue(txtIAccValue.Text);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, ex.Message);
            }
        }
        
        /*private void OnSelectionChanged(Object sender, SelectionChangedEventArgs args)
        {
            try
            {
                if (args.Source is TabControl)
                {
                    if (valueTab.IsSelected)
                    {
                        IUIAutomationValuePattern valuePattern = valueTab.Tag as IUIAutomationValuePattern;
                        if (valuePattern != null)
                        {
                            txtValue.Text = valuePattern.CurrentValue;
                        }
                    }
                    else if (legacyIAccTab.IsSelected)
                    {
                        IUIAutomationLegacyIAccessiblePattern legacyIAccPattern = legacyIAccTab.Tag as IUIAutomationLegacyIAccessiblePattern;
                        if (legacyIAccPattern != null)
                        {
                            txtIAccValue.Text = legacyIAccPattern.CurrentValue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {}
        }*/
        
        private IUIAutomationElement getItemElement = null;
        
        private void OnGetItem(object sender, RoutedEventArgs e)
        {
            IUIAutomationGridPattern gridPattern = gridTab.Tag as IUIAutomationGridPattern;
			if (gridPattern == null)
			{
				return;
			}
            
            int row = -1;
            if (int.TryParse(txtRow.Text, out row) == false)
            {
                System.Windows.MessageBox.Show(this, "Row is not a valid integer number");
                return;
            }
            
            int column = -1;
            if (int.TryParse(txtColumn.Text, out column) == false)
            {
                System.Windows.MessageBox.Show(this, "Column is not a valid integer number");
                return;
            }
            
            IUIAutomationElement element = null;
            try
            {
                element = gridPattern.GetItem(row, column);
            }
            catch (Exception ex)
            {
                txtInfo.Text = "Exception: " + ex.Message;
                getItemElement = null;
                return;
            }
            getItemElement = element;
            if (element == null)
            {
                txtInfo.Text = "(null)";
                return;
            }
            
            string content = "";
            try
            {
                content += "Class Name: \"" + element.CurrentClassName + "\"";
            }
            catch {}
            
            if (content != "")
            {
                content += Environment.NewLine;
            }
            
            try
            {
                content += "Control Type: " + Helper.ControlTypeIdToString(element.CurrentControlType);
            }
            catch {}
            
            if (content != "")
            {
                content += Environment.NewLine;
            }
            
            try
            {
                content += "Name: \"" + element.CurrentName + "\"";
            }
            catch {}
            
            object pattern = element.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId);
            IUIAutomationValuePattern valuePattern = pattern as IUIAutomationValuePattern;
        
            if (valuePattern != null)
            {
                if (content != "")
                {
                    content += Environment.NewLine;
                }
                
                try
                {
                    content += "ValuePattern.Value: \"" + valuePattern.CurrentValue + "\"";
                }
                catch {}
            }
            
            txtInfo.Text = content;
        }
        
        private void OnHighlight(object sender, RoutedEventArgs e)
        {
            if (getItemElement != null)
            {
                MainWindow.HighlightItemElement(getItemElement);
            }
        }
		
		[DllImport("user32.dll")]
		static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

		[StructLayout(LayoutKind.Sequential)]
		public struct RECT
		{
			public int Left;        // x position of upper-left corner
			public int Top;         // y position of upper-left corner
			public int Right;       // x position of lower-right corner
			public int Bottom;      // y position of lower-right corner
		}
    }
	
	public class ViewItem
	{
		public int id = 0;
		public string name = "";
		
		public ViewItem(int id, string name)
		{
			this.id = id;
			this.name = name;
		}
		
		public override string ToString()
		{
			return name + " (id:" + id + ")";
		}
	}
	
	public class TPRange
	{
		public string name = "";
		public IUIAutomationTextRange range = null;
		
		public TPRange(string name, IUIAutomationTextRange range)
		{
			this.name = name;
			this.range = range;
		}
		
		public override string ToString()
		{
			return name;
		}
	}
}
