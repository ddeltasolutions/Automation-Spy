using AutomationSpy;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Interop.UIAutomationClient;

namespace dDeltaSolutions.Spy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, System.Windows.Forms.IWin32Window
    {
        public MainWindow()
        {
            InitializeComponent();

            //autoupdate
            //CheckForUpdates();
			
			try
			{
				string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + 
					"\\dDeltaSolutions";
				if (!Directory.Exists(folder))
				{
					Directory.CreateDirectory(folder);
				}
				/*folder += "\\AutomationSpy";
				if (!Directory.Exists(folder))
				{
					Directory.CreateDirectory(folder);
				}*/
				prefFileName = folder + "\\AutomationSpyPref.xml";
			}
			catch {}

            try
            {
                LoadPreferences();
            }
            catch { }
            
            try
            {
                if (window_left > -10000 && window_top > -10000)
                {
                    this.Left = window_left;
                    this.Top = window_top;
                }
                if (window_width != 0 && window_height != 0)
                {
                    this.Width = window_width;
                    this.Height = window_height;
                }
                /*if (col_width1 > 0 && col_width2 > 0)
                {
                    mainGrid.ColumnDefinitions[0].Width = new GridLength(col_width1, GridUnitType.Star);
                    mainGrid.ColumnDefinitions[2].Width = new GridLength(col_width2, GridUnitType.Star);
                }*/
            }
            catch {}
			
			this.Title = MainWindow.AppName;
        }

        public static string Version = "4.5.0.0";
		public static string AppName = "Automation Spy";
        public static CUIAutomation8 uiAutomation = null;

        //private static string UpdateUrl = @"file:\\\C:\Projects\update";
        //private static string UpdateUrl = @"http://www.automationspy.com/update";

        private string prefFileName = System.IO.Path.GetTempPath() + "\\AutomationSpy_A5674D6E.xml";

		private static IUIAutomationCacheRequest cacheRequest = null;
		private Dictionary<int, string> properties = null;
        private double window_width = 0, window_height = 0, window_left = -10000, window_top = -10000;

        private void LoadPreferences()
        {
            //bool deleteOldVersions = true;

            if (File.Exists(prefFileName))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(prefFileName);
                XmlNodeList nodeList = null;

                nodeList = doc.GetElementsByTagName("AsyncContentLoadedEvent");
                if (nodeList.Count > 0)
                {
                    hasAsyncContentLoadedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("ElementAddedToSelectionEvent");
                if (nodeList.Count > 0)
                {
                    hasElementAddedToSelectionEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("ElementRemovedFromSelectionEvent");
                if (nodeList.Count > 0)
                {
                    hasElementRemovedFromSelectionEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("ElementSelectedEvent");
                if (nodeList.Count > 0)
                {
                    hasElementSelectedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("InvalidatedEvent");
                if (nodeList.Count > 0)
                {
                    hasInvalidatedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("InvokedEvent");
                if (nodeList.Count > 0)
                {
                    hasInvokedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("InputReachedTargetEvent");
                if (nodeList.Count > 0)
                {
                    hasInputReachedTargetEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("InputReachedOtherElementEvent");
                if (nodeList.Count > 0)
                {
                    hasInputReachedOtherElementEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("InputDiscardedEvent");
                if (nodeList.Count > 0)
                {
                    hasInputDiscardedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("LayoutInvalidatedEvent");
                if (nodeList.Count > 0)
                {
                    hasLayoutInvalidatedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("MenuOpenedEvent");
                if (nodeList.Count > 0)
                {
                    hasMenuOpenedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("MenuClosedEvent");
                if (nodeList.Count > 0)
                {
                    hasMenuClosedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("MenuModeStartEvent");
                if (nodeList.Count > 0)
                {
                    hasMenuModeStartEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("MenuModeEndEvent");
                if (nodeList.Count > 0)
                {
                    hasMenuModeEndEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("TextChangedEvent");
                if (nodeList.Count > 0)
                {
                    hasTextChangedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("TextSelectionChangedEvent");
                if (nodeList.Count > 0)
                {
                    hasTextSelectionChangedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("ToolTipClosedEvent");
                if (nodeList.Count > 0)
                {
                    hasToolTipClosedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("ToolTipOpenedEvent");
                if (nodeList.Count > 0)
                {
                    hasToolTipOpenedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("WindowOpenedEvent");
                if (nodeList.Count > 0)
                {
                    hasWindowOpenedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("AutomationFocusChangedEvent");
                if (nodeList.Count > 0)
                {
                    hasAutomationFocusChangedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("AutomationPropertyChangedEvent");
                if (nodeList.Count > 0)
                {
                    hasAutomationPropertyChangedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("StructureChangedEvent");
                if (nodeList.Count > 0)
                {
                    hasStructureChangedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("WindowClosedEvent");
                if (nodeList.Count > 0)
                {
                    hasWindowClosedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
				
				nodeList = doc.GetElementsByTagName("TextEditTextChangedEvent");
                if (nodeList.Count > 0)
                {
                    hasTextEditTextChangedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
				
				nodeList = doc.GetElementsByTagName("ChangesEvent");
                if (nodeList.Count > 0)
                {
                    hasChangesEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
				
				nodeList = doc.GetElementsByTagName("NotificationEvent");
                if (nodeList.Count > 0)
                {
                    hasNotificationEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
				
				nodeList = doc.GetElementsByTagName("ActiveTextPositionChangedEvent");
                if (nodeList.Count > 0)
                {
                    hasActiveTextPositionChangedEvent = Convert.ToBoolean(nodeList[0].InnerText);
                }
				

                nodeList = doc.GetElementsByTagName("Scope");
                if (nodeList.Count > 0)
                {
                    eventsScope = (TreeScope)Enum.Parse(typeof(TreeScope), nodeList[0].InnerText);
                }

                nodeList = doc.GetElementsByTagName("Width");
                if (nodeList.Count > 0)
                {
                    window_width = Convert.ToDouble(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("Height");
                if (nodeList.Count > 0)
                {
                    window_height = Convert.ToDouble(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("Left");
                if (nodeList.Count > 0)
                {
                    window_left = Convert.ToDouble(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("Top");
                if (nodeList.Count > 0)
                {
                    window_top = Convert.ToDouble(nodeList[0].InnerText);
                }
                
                /*nodeList = doc.GetElementsByTagName("ColWidth1");
                if (nodeList.Count > 0)
                {
                    col_width1 = Convert.ToDouble(nodeList[0].InnerText);
                }
                
                nodeList = doc.GetElementsByTagName("ColWidth2");
                if (nodeList.Count > 0)
                {
                    col_width2 = Convert.ToDouble(nodeList[0].InnerText);
                }*/

                /*nodeList = doc.GetElementsByTagName("Version");
                if (nodeList.Count > 0)
                {
                    if (nodeList[0].InnerText == MainWindow.Version)
                    {
                        deleteOldVersions = false;
                    }
                }*/
				
				nodeList = doc.GetElementsByTagName("AlwaysOnTop");
                if (nodeList.Count > 0)
                {
                    if (nodeList[0].InnerText == "True")
                    {
                        menuAlwaysOnTop.IsChecked = true;
						this.Topmost = true;
                    }
                }
            }

            #region comments
            /*if (deleteOldVersions)
            {
                string currentDirectory = Directory.GetCurrentDirectory();
                string[] files = Directory.GetFiles(currentDirectory, "*.exe", SearchOption.TopDirectoryOnly);

                foreach (string file in files)
                {
                    if (file.StartsWith("old_AutomationSpy_v"))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch { }
                    }
                }
            }*/
            #endregion
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            uiAutomation = new CUIAutomation8();
            
			cacheRequest = uiAutomation.CreateCacheRequest();
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_RuntimeIdPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_BoundingRectanglePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_ProcessIdPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_ControlTypePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_LocalizedControlTypePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_NamePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_AcceleratorKeyPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_AccessKeyPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_HasKeyboardFocusPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsKeyboardFocusablePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsEnabledPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_AutomationIdPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_ClassNamePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_HelpTextPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_ClickablePointPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_CulturePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsControlElementPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsContentElementPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_LabeledByPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsPasswordPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_NativeWindowHandlePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_ItemTypePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsOffscreenPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_OrientationPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_FrameworkIdPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsRequiredForFormPropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_ItemStatusPropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_AriaRolePropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_AriaPropertiesPropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsDataValidForFormPropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_ControllerForPropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_DescribedByPropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_FlowsToPropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_ProviderDescriptionPropertyId);
			
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_OptimizeForVisualContentPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_LiveSettingPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsPeripheralPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_PositionInSetPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_SizeOfSetPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_LevelPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_AnnotationTypesPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_LandmarkTypePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_LocalizedLandmarkTypePropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_FullDescriptionPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_HeadingLevelPropertyId);
			cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsDialogPropertyId);
			
			//Cache patterns
			if (properties == null)
			{
				properties = new Dictionary<int, string>();
				properties.Add(UIA_PropertyIds.UIA_IsDockPatternAvailablePropertyId, "DockPattern");
				properties.Add(UIA_PropertyIds.UIA_IsExpandCollapsePatternAvailablePropertyId, "ExpandCollapsePattern");
				properties.Add(UIA_PropertyIds.UIA_IsGridItemPatternAvailablePropertyId, "GridItemPattern");
				properties.Add(UIA_PropertyIds.UIA_IsGridPatternAvailablePropertyId, "GridPattern");
				properties.Add(UIA_PropertyIds.UIA_IsInvokePatternAvailablePropertyId, "InvokePattern");
				properties.Add(UIA_PropertyIds.UIA_IsMultipleViewPatternAvailablePropertyId, "MultipleViewPattern");
				properties.Add(UIA_PropertyIds.UIA_IsRangeValuePatternAvailablePropertyId, "RangeValuePattern");
				properties.Add(UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId, "ScrollPattern");
				properties.Add(UIA_PropertyIds.UIA_IsScrollItemPatternAvailablePropertyId, "ScrollItemPattern");
				properties.Add(UIA_PropertyIds.UIA_IsSelectionItemPatternAvailablePropertyId, "SelectionItemPattern");
				properties.Add(UIA_PropertyIds.UIA_IsSelectionPatternAvailablePropertyId, "SelectionPattern");
				properties.Add(UIA_PropertyIds.UIA_IsTablePatternAvailablePropertyId, "TablePattern");
				properties.Add(UIA_PropertyIds.UIA_IsTableItemPatternAvailablePropertyId, "TableItemPattern");
				properties.Add(UIA_PropertyIds.UIA_IsTextPatternAvailablePropertyId, "TextPattern");
				properties.Add(UIA_PropertyIds.UIA_IsTogglePatternAvailablePropertyId, "TogglePattern");
				properties.Add(UIA_PropertyIds.UIA_IsTransformPatternAvailablePropertyId, "TransformPattern");
				properties.Add(UIA_PropertyIds.UIA_IsValuePatternAvailablePropertyId, "ValuePattern");
				properties.Add(UIA_PropertyIds.UIA_IsWindowPatternAvailablePropertyId, "WindowPattern");
				properties.Add(UIA_PropertyIds.UIA_IsLegacyIAccessiblePatternAvailablePropertyId, "LegacyIAccessiblePattern");
				properties.Add(UIA_PropertyIds.UIA_IsItemContainerPatternAvailablePropertyId, "ItemContainerPattern");
				properties.Add(UIA_PropertyIds.UIA_IsVirtualizedItemPatternAvailablePropertyId, "VirtualizedItemPattern");
                properties.Add(UIA_PropertyIds.UIA_IsSynchronizedInputPatternAvailablePropertyId, "SynchronizedInputPattern");
			}
			
			foreach (int propertyid in properties.Keys)
			{
				cacheRequest.AddProperty(propertyid);
			}
            
            Traces.ClearTrace();
            RefreshTree();

            this.attributesListView.ItemsSource = this.listAttributes;
            this.patternsListView.ItemsSource = this.listPatterns;

            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(this.patternsListView.ItemsSource);
            PropertyGroupDescription groupDescription = new PropertyGroupDescription("Group");
            view.GroupDescriptions.Add(groupDescription);
            
            /*if (IsWindows7())
            {
                for (int id = 30000; id <= 30044; id++)
                {
                    WindowEventsOptions.evtProperties.Add(id);
                }
                WindowEventsOptions.evtProperties.Add(30090);
                for (int id = 30101; id <= 30110; id++)
                {
                    WindowEventsOptions.evtProperties.Add(id);
                }
            }
            else
            {*/
                // by default listen all properties
                for (int id = 30000; id <= 30110; id++)
                {
                    WindowEventsOptions.evtProperties.Add(id);
                }
            //}
        }
        
        private bool IsWindows7()
		{
			return (Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor == 1);
		}
        //private IUIAutomationElement root = null;
		
		private void RefreshTree()
		{
			try
			{
				TryRefreshTree();
			}
			catch { }
		}

        private void TryRefreshTree()
        {
            // keep the currently selected element
            IUIAutomationElement selectedElement = null;
            TreeViewItem selectedTreeViewItem = this.tvElements.SelectedItem as TreeViewItem;
            if (selectedTreeViewItem != null)
            {
                TreeNode selectedNode = selectedTreeViewItem.Tag as TreeNode;
                if (selectedNode != null && selectedNode.IsAlive)
                {
                    selectedElement = selectedNode.Element;
                }
            }
            
            MainWindow.lastAutomationElement = null;

            if (tvElements.Items.Count > 0)
            {
                tvElements.Items.Clear();
            }

            IUIAutomationElement root = uiAutomation.GetRootElement();
            TreeNode rootNode = new TreeNode(root, true);

            TreeViewItem rootItem = new TreeViewItem();
            rootItem.Expanded += treeviewItem_Expanded;

            rootItem.Header = rootNode.ToString();
            rootItem.Tag = rootNode;

            tvElements.Items.Add(rootItem);

            List<IUIAutomationElement> childrenElements = MainWindow.FindChildren(root);

            foreach (IUIAutomationElement element in childrenElements)
            {
                try
                {
                    string nodeLabel = element.CurrentName;
                    if (nodeLabel == "AutomationSpy_rect_top" || nodeLabel == "AutomationSpy_rect_left" ||
                        nodeLabel == "AutomationSpy_rect_bottom" || nodeLabel == "AutomationSpy_rect_right" ||
                        nodeLabel == "AutomationSpy_CP1" || nodeLabel == "AutomationSpy_CP2")
                    {
                        continue;
                    }
                }
                catch {}
                
                TreeViewItem currentItem = new TreeViewItem();

                TreeNode currentNode = new TreeNode(element);
                currentItem.Header = currentNode.ToString();
                currentItem.Tag = currentNode;
                rootItem.Items.Add(currentItem);
            }

            rootItem.IsExpanded = true;
            
            if (selectedElement != null)
            {
                if (CompareElements(selectedElement, uiAutomation.GetRootElement()) == true)
                {
                    rootItem.IsSelected = true;
                }
                else
                {
                    SelectElement(selectedElement, true);
                }
                tvElements.Focus();
            }
            else 
            {
                this.attributesListView.Tag = null;
                this.patternsListView.Tag = null;
                this.listAttributes.Clear();
                this.listPatterns.Clear();
                
                if (buttonHighlight.IsChecked == true)
                {
                    try
                    {
                        UnHighlight();
                    }
                    catch { }
                }
                if (showCPMenuItem.IsChecked == true)
                {
                    try
                    {
                        UnHighlightCP();
                    }
                    catch { }
                }
            }
        }

        void treeviewItem_Expanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem treeviewItem = e.OriginalSource as TreeViewItem;
            UpdateTreeViewItemDescendants(treeviewItem);
        }

        private void RefreshTreeViewItem(TreeViewItem treeviewItem)
        {
            if (treeviewItem == null)
            {
                return;
            }

            TreeNode node = treeviewItem.Tag as TreeNode;
            if (node == null)
            {
                return;
            }

            if (node.IsAlive == false)
            {
                RefreshTreeViewItem(treeviewItem.Parent as TreeViewItem);
                return;
            }

            treeviewItem.Header = node.ToString();
            treeviewItem.Items.Clear();

            List<IUIAutomationElement> children = MainWindow.FindChildren(node.Element);
            foreach (IUIAutomationElement element in children)
            {
                if (node.IsRoot == true)
                {
                    try
                    {
                        string nodeLabel = element.CurrentName;
                        if (nodeLabel == "AutomationSpy_rect_top" || nodeLabel == "AutomationSpy_rect_left" ||
                            nodeLabel == "AutomationSpy_rect_bottom" || nodeLabel == "AutomationSpy_rect_right" ||
                            nodeLabel == "AutomationSpy_CP1" || nodeLabel == "AutomationSpy_CP2")
                        {
                            continue;
                        }
                    }
                    catch {}
                }
                
                TreeViewItem currentItem = new TreeViewItem();
                TreeNode currentNode = new TreeNode(element);
                currentItem.Header = currentNode.ToString();
                currentItem.Tag = currentNode;
                treeviewItem.Items.Add(currentItem);
            }

            UpdateTreeViewItemDescendants(treeviewItem);
        }
        
        void RefreshControlPatterns(object sender, RoutedEventArgs e)
        {
            TreeViewItem treeviewItem = this.tvElements.SelectedItem as TreeViewItem;
			if (treeviewItem == null)
            {
                return;
            }
			
            try
			{
				SelectedItemChanged(treeviewItem);
			}
			catch { }
        }

        void RefreshItem(object sender, RoutedEventArgs e)
		{
			try
			{
				TryRefreshItem();
			}
			catch { }
		}
		
		private void TryRefreshItem()
        {
            TreeViewItem treeviewItem = this.tvElements.SelectedItem as TreeViewItem;
            if (treeviewItem == null)
            {
                return;
            }

            RefreshTreeViewItem(treeviewItem);

			treeviewItem = this.tvElements.SelectedItem as TreeViewItem;
			if (treeviewItem == null)
            {
                return;
            }
			
            TreeNode node = treeviewItem.Tag as TreeNode;
            if (node == null)
            {
                return;
            }
			
			bool isAlive = node.IsAlive;
			if ((buttonHighlight.IsChecked.HasValue) && (buttonHighlight.IsChecked.Value == true))
            {
                if (isAlive == false)
                {
                    UnHighlight();
                }
				else
				{
					this.Highlight(node.Element);
				}
            }
			
			bool showCP = showCPMenuItem.IsChecked;
			if (showCP)
			{
				try
				{
					UnHighlightCP();
				}
				catch {}
			}

            SelectedItemChanged(node, isAlive);
			
			if (showCP && isAlive)
			{
				try
				{
					HighlightCP(node.Element);
				}
				catch {}
			}
        }

        void HighlightItem(object sender, RoutedEventArgs e)
        {
            TreeViewItem tvItem = this.tvElements.SelectedItem as TreeViewItem;
            if (tvItem == null)
            {
                return;
            }

            TreeNode node = tvItem.Tag as TreeNode;
            if (node == null)
            {
                return;
            }
            
            HighlightItemElement(node.Element);
        }

        internal static void HighlightItemElement(IUIAutomationElement element)
        {
            tagRECT rect;

            try
            {
                rect = element.CurrentBoundingRectangle;
            }
            catch 
            {
                return;
            }

            HighlightRect(rect);
        }
		
		void OnCopyText(object sender, RoutedEventArgs e)
		{ 
			TreeViewItem tvItem = this.tvElements.SelectedItem as TreeViewItem;
            if (tvItem == null)
            {
                return;
            }

            TreeNode node = tvItem.Tag as TreeNode;
            if (node == null)
            {
                return;
            }
			
			string name = null;
			try
			{
				name = node.Element.CurrentName;
			}
			catch
			{
				try
				{
					name = node.Element.CachedName;
				}
				catch { }
			}
			
			if (name != null)
			{
				try
				{
					System.Windows.Forms.Clipboard.SetText(name);
				}
				catch { }
			}
		}
		
		void OnSetFocus(object sender, RoutedEventArgs e)
		{
			TreeViewItem tvItem = this.tvElements.SelectedItem as TreeViewItem;
            if (tvItem == null)
            {
                return;
            }
			
			TreeNode node = tvItem.Tag as TreeNode;
			if (node == null)
			{
				return;
			}
			
			try
			{
				node.Element.SetFocus();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message);
			}
		}

        private void UpdateTreeViewItemDescendants(TreeViewItem treeViewItem)
        {
            if (treeViewItem == null)
            {
                return;
            }

			try
			{
				foreach (TreeViewItem currentTreeViewItem in treeViewItem.Items)
				{
					if (currentTreeViewItem.IsExpanded)
					{
						UpdateTreeViewItemDescendants(currentTreeViewItem);
					}
					else
					{
						UpdateTreeViewItemChildren(currentTreeViewItem);
					}
				}
			}
			catch { }
        }

        private void UpdateTreeViewItemChildren(TreeViewItem treeviewItem)
        {
            treeviewItem.Items.Clear();

            TreeNode node = treeviewItem.Tag as TreeNode;
            if (node == null)
            {
                return;
            }

            List<TreeNode> children = node.Children;
			if (children == null)
			{
				return;
			}
            foreach (TreeNode currentNode in children)
            {
                TreeViewItem currentItem = new TreeViewItem();
                currentItem.Header = currentNode.ToString();
                currentItem.Tag = currentNode;
                treeviewItem.Items.Add(currentItem);
            }
        }

		private Form frmHighlightTop = null;
		private Form frmHighlightLeft = null;
		private Form frmHighlightBottom = null;
		private Form frmHighlightRight = null;
		private int width = 0;
        private int height = 0;
		private int thickness = 5;

        private void Highlight(IUIAutomationElement element)
        {
            if (this.buttonHighlight.IsChecked.HasValue == false)
            {
                return;
            }
                
            if (this.buttonHighlight.IsChecked.Value == false)
            {
                return;
            }

            if (element == null)
            {
                return;
            }

            tagRECT rect = new tagRECT {left=0, top=0, right=0, bottom=0};
            bool rectObtained = false;

            try
            {
                rect = element.CurrentBoundingRectangle;
                rectObtained = true;
            }
            catch (Exception ex)
            {
                rectObtained = false;
            }

            if ((rectObtained == false) || (rect.right <= 0 && rect.bottom <= 0)) // infinity
            {
				UnHighlight();
                return;
            }
			
			int left = 0;
            int top = 0;
			bool firstHighlight = false;
			
			try
			{
				left = rect.left - thickness;
				top = rect.top - thickness;
				width = rect.right - rect.left;
				height = rect.bottom - rect.top;
			}
			catch
			{
				UnHighlight();
				return;
			}
			
			if (frmHighlightTop == null)
			{
				frmHighlightTop = new Form();
                frmHighlightTop.Text = "AutomationSpy_rect_top";
				frmHighlightTop.FormBorderStyle = FormBorderStyle.None;
				frmHighlightTop.StartPosition = FormStartPosition.Manual;
				
				frmHighlightTop.Location = new System.Drawing.Point(left, top);
				frmHighlightTop.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightTop.Size = new System.Drawing.Size(0, 0);
				frmHighlightTop.ShowInTaskbar = false;
				frmHighlightTop.TopMost = true;
				frmHighlightTop.ForeColor = System.Drawing.Color.Red;
				frmHighlightTop.BackColor = System.Drawing.Color.Red;
				
				frmHighlightTop.Load += highlightTop_Loaded;
				frmHighlightTop.Show();
				
				firstHighlight = true;
			}
			else
			{
				frmHighlightTop.Location = new System.Drawing.Point(left, top);
				frmHighlightTop.Size = new System.Drawing.Size(width + 2 * thickness, thickness);
			}
            
			if (frmHighlightLeft == null)
			{
				frmHighlightLeft = new Form();
                frmHighlightLeft.Text = "AutomationSpy_rect_left";
				frmHighlightLeft.FormBorderStyle = FormBorderStyle.None;
				frmHighlightLeft.StartPosition = FormStartPosition.Manual;
				
				frmHighlightLeft.Location = new System.Drawing.Point(left, top);
				frmHighlightLeft.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightLeft.Size = new System.Drawing.Size(0, 0);
				frmHighlightLeft.ShowInTaskbar = false;
				frmHighlightLeft.TopMost = true;
				frmHighlightLeft.ForeColor = System.Drawing.Color.Red;
				frmHighlightLeft.BackColor = System.Drawing.Color.Red;
				
				frmHighlightLeft.Load += highlightLeft_Loaded;
				frmHighlightLeft.Show();
				
				firstHighlight = true;
			}
			else
			{
				frmHighlightLeft.Location = new System.Drawing.Point(left, top);
				frmHighlightLeft.Size = new System.Drawing.Size(thickness, height + 2 * thickness);
			}
			
			if (frmHighlightBottom == null)
			{
				frmHighlightBottom = new Form();
                frmHighlightBottom.Text = "AutomationSpy_rect_bottom";
				frmHighlightBottom.FormBorderStyle = FormBorderStyle.None;
				frmHighlightBottom.StartPosition = FormStartPosition.Manual;
				
				frmHighlightBottom.Location = new System.Drawing.Point(left, top + height + thickness);
				frmHighlightBottom.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightBottom.Size = new System.Drawing.Size(0, 0);
				frmHighlightBottom.ShowInTaskbar = false;
				frmHighlightBottom.TopMost = true;
				frmHighlightBottom.ForeColor = System.Drawing.Color.Red;
				frmHighlightBottom.BackColor = System.Drawing.Color.Red;
				
				frmHighlightBottom.Load += highlightBottom_Loaded;
				frmHighlightBottom.Show();
				
				firstHighlight = true;
			}
			else
			{
				frmHighlightBottom.Location = new System.Drawing.Point(left, top + height + thickness);
				frmHighlightBottom.Size = new System.Drawing.Size(width + 2 * thickness, thickness);
			}
			
			if (frmHighlightRight == null)
			{
				frmHighlightRight = new Form();
                frmHighlightRight.Text = "AutomationSpy_rect_right";
				frmHighlightRight.FormBorderStyle = FormBorderStyle.None;
				frmHighlightRight.StartPosition = FormStartPosition.Manual;
				
				frmHighlightRight.Location = new System.Drawing.Point(left + thickness + width, top);
				frmHighlightRight.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightRight.Size = new System.Drawing.Size(0, 0);
				frmHighlightRight.ShowInTaskbar = false;
				frmHighlightRight.TopMost = true;
				frmHighlightRight.ForeColor = System.Drawing.Color.Red;
				frmHighlightRight.BackColor = System.Drawing.Color.Red;
				
				frmHighlightRight.Load += highlightRight_Loaded;
				frmHighlightRight.Show();
				
				firstHighlight = true;
			}
			else
			{
				frmHighlightRight.Location = new System.Drawing.Point(left + thickness + width, top);
				frmHighlightRight.Size = new System.Drawing.Size(thickness, height + 2 * thickness);
			}

			if (firstHighlight == true)
			{
				this.Focus();
			}
        }
		
		private void highlightTop_Loaded(object sender, System.EventArgs e)
		{
			frmHighlightTop.Size = new System.Drawing.Size(width + 2 * thickness, thickness);
		}
		
		private void highlightLeft_Loaded(object sender, System.EventArgs e)
		{
			frmHighlightLeft.Size = new System.Drawing.Size(thickness, height + 2 * thickness);
		}
		
		private void highlightBottom_Loaded(object sender, System.EventArgs e)
		{
			frmHighlightBottom.Size = new System.Drawing.Size(width + 2 * thickness, thickness);
		}
		
		private void highlightRight_Loaded(object sender, System.EventArgs e)
		{
			frmHighlightRight.Size = new System.Drawing.Size(thickness, height + 2 * thickness);
		}

        [DllImport("User32.dll")]
        public static extern Int32 SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32.dll")]
        public static extern IntPtr GetForegroundWindow();

        public IntPtr Handle
        {
            get
            {
                var interopHelper = new WindowInteropHelper(this);
                return interopHelper.Handle;
            }
        }

        private void tvElements_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeViewItem selectedItem = e.NewValue as TreeViewItem;
            if (selectedItem == null)
            {
                return;
            }
			
			try
			{
				SelectedItemChanged(selectedItem);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message);
			}
		}
			
		private void SelectedItemChanged(TreeViewItem selectedItem)
		{
            TreeNode selectedNode = selectedItem.Tag as TreeNode;
            if (selectedNode == null)
            {
                return;
            }
			
			bool isAlive = selectedNode.IsAlive;
			
			// draw text with italic if element is not available anymore
			if (!isAlive)
			{
				try
				{
					/*TreeViewItem parent = null;
					try
					{
						parent = selectedItem.Parent as TreeViewItem;
					}
					catch { }
					
					bool isItalic = false;
					if (parent != null)
					{
						TreeNode parentNode = parent.Tag as TreeNode;
						
						if (parentNode != null && parentNode.IsAlive == false)
						{
							if (parent.FontStyle != FontStyles.Italic)
							{
								parent.FontStyle = FontStyles.Italic;
							}
							isItalic = true;
						}
					}
					
					if (isItalic == false && selectedItem.FontStyle != FontStyles.Italic)
					{
						selectedItem.FontStyle = FontStyles.Italic;
					}*/
					
					if (selectedItem.FontStyle != FontStyles.Italic)
					{
						bool isParentAlive = false;
						TreeViewItem italicElement = selectedItem;
						TreeViewItem parentElement = selectedItem;
						
						while (isParentAlive == false)
						{
							isParentAlive = true;
							
							try
							{
								parentElement = parentElement.Parent as TreeViewItem;
							}
							catch { }
							
							if (parentElement == null)
							{
								break;
							}
							
							TreeNode parentNode = parentElement.Tag as TreeNode;
							
							if (parentNode == null)
							{
								break;
							}

							isParentAlive = parentNode.IsAlive;
							if (isParentAlive == false)
							{
								italicElement = parentElement;
							}
						}
						
						if (italicElement.FontStyle != FontStyles.Italic)
						{
							italicElement.FontStyle = FontStyles.Italic;
						}
					}
				}
				catch { }
			}

            if ((buttonHighlight.IsChecked.HasValue) && (buttonHighlight.IsChecked.Value == true))
            {
                if (isAlive == false)
                {
                    UnHighlight();
                }
				else
				{
					Highlight(selectedNode.Element);
				}
            }
			
            bool showCP = showCPMenuItem.IsChecked;
			if (showCP)
			{
                try
                {
                    UnHighlightCP();
                }
                catch {}
			}

            SelectedItemChanged(selectedNode, isAlive);
            
            if (showCP && isAlive)
            {
                try
                {
                    HighlightCP(selectedNode.Element);
                }
                catch {}
            }
        }

        private void UnHighlight()
        {
			bool firstHighlight = false;
			
			if (frmHighlightTop != null)
			{
				frmHighlightTop.Location = new System.Drawing.Point(-10000, -10000);
			}
			else
			{
				frmHighlightTop = new Form();
                frmHighlightTop.Text = "AutomationSpy_rect_top";
				frmHighlightTop.FormBorderStyle = FormBorderStyle.None;
				frmHighlightTop.StartPosition = FormStartPosition.Manual;
				
				frmHighlightTop.Location = new System.Drawing.Point(-10000, -10000);
				frmHighlightTop.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightTop.Size = new System.Drawing.Size(0, 0);
				frmHighlightTop.ShowInTaskbar = false;
				frmHighlightTop.TopMost = true;
				frmHighlightTop.ForeColor = System.Drawing.Color.Red;
				frmHighlightTop.BackColor = System.Drawing.Color.Red;
				
				frmHighlightTop.Load += highlightTop_Loaded;
				frmHighlightTop.Show();
				
				firstHighlight = true;
			}
			
			if (frmHighlightLeft != null)
			{
				frmHighlightLeft.Location = new System.Drawing.Point(-10000, -10000);
			}
			else
			{
				frmHighlightLeft = new Form();
                frmHighlightLeft.Text = "AutomationSpy_rect_left";
				frmHighlightLeft.FormBorderStyle = FormBorderStyle.None;
				frmHighlightLeft.StartPosition = FormStartPosition.Manual;
				
				frmHighlightLeft.Location = new System.Drawing.Point(-10000, -10000);
				frmHighlightLeft.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightLeft.Size = new System.Drawing.Size(0, 0);
				frmHighlightLeft.ShowInTaskbar = false;
				frmHighlightLeft.TopMost = true;
				frmHighlightLeft.ForeColor = System.Drawing.Color.Red;
				frmHighlightLeft.BackColor = System.Drawing.Color.Red;
				
				frmHighlightLeft.Load += highlightLeft_Loaded;
				frmHighlightLeft.Show();
				
				firstHighlight = true;
			}
			
			if (frmHighlightBottom != null)
			{
				frmHighlightBottom.Location = new System.Drawing.Point(-10000, -10000);
			}
			else
			{
				frmHighlightBottom = new Form();
                frmHighlightBottom.Text = "AutomationSpy_rect_bottom";
				frmHighlightBottom.FormBorderStyle = FormBorderStyle.None;
				frmHighlightBottom.StartPosition = FormStartPosition.Manual;
				
				frmHighlightBottom.Location = new System.Drawing.Point(-10000, -10000);
				frmHighlightBottom.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightBottom.Size = new System.Drawing.Size(0, 0);
				frmHighlightBottom.ShowInTaskbar = false;
				frmHighlightBottom.TopMost = true;
				frmHighlightBottom.ForeColor = System.Drawing.Color.Red;
				frmHighlightBottom.BackColor = System.Drawing.Color.Red;
				
				frmHighlightBottom.Load += highlightBottom_Loaded;
				frmHighlightBottom.Show();
				
				firstHighlight = true;
			}
			
			if (frmHighlightRight != null)
			{
				frmHighlightRight.Location = new System.Drawing.Point(-10000, -10000);
			}
			else
			{
				frmHighlightRight = new Form();
                frmHighlightRight.Text = "AutomationSpy_rect_right";
				frmHighlightRight.FormBorderStyle = FormBorderStyle.None;
				frmHighlightRight.StartPosition = FormStartPosition.Manual;
				
				frmHighlightRight.Location = new System.Drawing.Point(-10000, -10000);
				frmHighlightRight.MinimumSize = new System.Drawing.Size(0, 0);
				frmHighlightRight.Size = new System.Drawing.Size(0, 0);
				frmHighlightRight.ShowInTaskbar = false;
				frmHighlightRight.TopMost = true;
				frmHighlightRight.ForeColor = System.Drawing.Color.Red;
				frmHighlightRight.BackColor = System.Drawing.Color.Red;
				
				frmHighlightRight.Load += highlightRight_Loaded;
				frmHighlightRight.Show();
				
				firstHighlight = true;
			}
			
			if (firstHighlight == true)
			{
				this.Focus();
			}
        }

        private void buttonRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshTree();
        }

        private void tvElements_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.F5)
            {
                RefreshTree();
            }
        }
		
		private void menuItemPick_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				menuItemPick.IsChecked = !menuItemPick.IsChecked;
				selectButton.IsChecked = menuItemPick.IsChecked;
			
				TryOnSelectElement();
			}
			catch {}
		}
		
		private void menuHighlightSelected_Click(object sender, RoutedEventArgs e)
		{
			menuHighlightSelected.IsChecked = !menuHighlightSelected.IsChecked;
			buttonHighlight.IsChecked = menuHighlightSelected.IsChecked;
			buttonHighlight_Click(sender, e);
		}
		
		private void menuClickablePoint_Click(object sender, RoutedEventArgs e)
		{
			menuClickablePoint.IsChecked = !menuClickablePoint.IsChecked;
			showCPMenuItem.IsChecked = menuClickablePoint.IsChecked;
			OnShowCP(sender, e);
		}

        private void buttonHighlight_Click(object sender, RoutedEventArgs e)
        {
			try
			{
				TryButtonHighlightClick();
			}
			catch { }
		}
		
		private void TryButtonHighlightClick()
		{
            if (!this.buttonHighlight.IsChecked.HasValue)
            {
                return;
            }
			
			menuHighlightSelected.IsChecked = buttonHighlight.IsChecked.Value;

            StackPanel panel = buttonHighlight.Content as StackPanel;
            Image img = null;

            if ((panel != null) && (panel.Children.Count > 0))
            {
                img = panel.Children[0] as Image;
            }

            if (this.buttonHighlight.IsChecked.Value == false)
            {
				UnHighlight();

                // display released button icon
                if (img != null)
                {
                    Uri sourceUri = new Uri("pack://application:,,,/AutomationSpy;component/Pictures/28x28xHighlight.png", UriKind.Absolute);
                    BitmapImage thisIcon = new BitmapImage(sourceUri);
                    img.Source = thisIcon;
                }
            }
            else
            {
                //display pressed button icon 
                if (img != null)
                {
                    Uri sourceUri = new Uri("pack://application:,,,/AutomationSpy;component/Pictures/28x28xHighlightPressed.png", UriKind.Absolute);
                    BitmapImage thisIcon = new BitmapImage(sourceUri);
                    img.Source = thisIcon;
                }

                TreeViewItem selectedItem = tvElements.SelectedItem as TreeViewItem;
                if (selectedItem == null)
                {
					UnHighlight();
                    return;
                }

                TreeNode selectedNode = selectedItem.Tag as TreeNode;
                if (selectedNode == null)
                {
                    return;
                }

                this.Highlight(selectedNode.Element);
            }
        }

        private Timer timer = null;

        private void OnSelectElement(object sender, RoutedEventArgs e)
		{
			try
			{
				if (!this.selectButton.IsChecked.HasValue)
				{
					return;
				}
				menuItemPick.IsChecked = this.selectButton.IsChecked.Value;
			
				TryOnSelectElement();
			}
			catch { }
		}
		
		private void TryOnSelectElement()
        {
			if (menuItemPick.IsChecked == true && menuItemTrack.IsChecked == true)
			{
				menuItemTrack.IsChecked = false;
				if (timerTrack != null)
				{
					timerTrack.Stop();
				}
			}
		
            StackPanel panel = this.selectButton.Content as StackPanel;
            Image img = null;

            if ((panel != null) && (panel.Children.Count > 0))
            {
                img = panel.Children[0] as Image;
            }

            if (this.selectButton.IsChecked.Value == false)
            {
                //Pointer button is released
                if (img != null)
                {
                    Uri sourceUri = new Uri("pack://application:,,,/AutomationSpy;component/Pictures/32x32xPointer.png", UriKind.Absolute);
                    BitmapImage thisIcon = new BitmapImage(sourceUri);
                    img.Source = thisIcon;
                }

                if (timer != null)
                {
                    timer.Stop();
                }
            }
            else
            {
                //Pointer button is pressed
                if (img != null)
                {
                    Uri sourceUri = new Uri("pack://application:,,,/AutomationSpy;component/Pictures/32x32xPointerPressed.png", UriKind.Absolute);
                    BitmapImage thisIcon = new BitmapImage(sourceUri);
                    img.Source = thisIcon;
                }

                if (timer == null)
                {
                    timer = new Timer();
                    timer.Interval = 1000;
                    timer.Tick += timer_Tick;
                }

                timer.Start();
            }
        }

        private static AutomationInfo lastAutomationElement = null;
        private static int currentProcessId = Process.GetCurrentProcess().Id;

        void timer_Tick(object sender, EventArgs e)
		{
			try
			{
				TryTimerTick();
			}
			catch { }
		}
		
		private void TryTimerTick()
        {
            bool leftCtrlIsPressed = ((int)Keyboard.GetKeyStates(Key.LeftCtrl) & (int)KeyStates.Down) != 0;
            bool rightCtrlIsPressed = ((int)Keyboard.GetKeyStates(Key.RightCtrl) & (int)KeyStates.Down) != 0;

            if (leftCtrlIsPressed || rightCtrlIsPressed)
            {
                int xDouble = Convert.ToInt32(System.Windows.Forms.Cursor.Position.X);
                int yDouble = Convert.ToInt32(System.Windows.Forms.Cursor.Position.Y);

                tagPOINT mousePoint = new tagPOINT {x=xDouble, y=yDouble};
                IUIAutomationElement element = uiAutomation.ElementFromPoint(mousePoint);

                SelectElement(element);
                if (tvElements.IsFocused == false)
                {
                    tvElements.Focus();
                }
            }
        }

        private bool SelectElement(IUIAutomationElement element, bool includeSpy = false)
        {
            try
            {
                if ((MainWindow.lastAutomationElement != null) &&
                    (MainWindow.lastAutomationElement.IsIdentic(element)))
                {
                    return false;
                }

                if (element.CurrentProcessId == MainWindow.currentProcessId && includeSpy == false)
                {
                    return false;
                }
            }
            catch { }

            if ((MainWindow.lastAutomationElement != null) &&
                (MainWindow.lastAutomationElement.IsSameElement(element)))
            {
                MainWindow.lastAutomationElement.UpdateBoundingRect(element);
                Highlight(element);
            }
            else
            {
                MainWindow.lastAutomationElement = new AutomationInfo(element);

                if (SelectAutomationElementInTree(element) == true)
                {
                    MainWindow.lastAutomationElement.UpdateBoundingRect(element);
                    Highlight(element);
                }
            }
			return true;
        }

        // return true if selected element was already selected
        private bool SelectAutomationElementInTree(IUIAutomationElement element)
        {
            if (element == null)
            {
                return false;
            }

            List<IUIAutomationElement> parentsList = new List<IUIAutomationElement>();
            IUIAutomationElement parent = element;

            IUIAutomationTreeWalker tw = uiAutomation.ControlViewWalker;
            if (MainWindow.mode == Modes.Control)
            {
                //tw = uiAutomation.ControlViewWalker;
            }
            else if (MainWindow.mode == Modes.Raw)
            {
                tw = uiAutomation.RawViewWalker;
            }
            else //if (MainWindow.mode == Modes.Content)
            {
                tw = uiAutomation.ContentViewWalker;
            }

            while (parent != null)
            {
                parentsList.Insert(0, parent);
                parent = tw.GetParentElement(parent);
            }

            if (tvElements.Items.Count == 0)
            {
                return false;
            }

            TreeViewItem rootItem = tvElements.Items[0] as TreeViewItem;
            TreeViewItem currentChildItem = rootItem;
			bool alreadySelected = false;

            for (int i = 1; i < parentsList.Count; i++ )
            {
                IUIAutomationElement currentParent = parentsList[i];
                currentChildItem = FindElementInTreeViewItem(currentChildItem, currentParent);

                if (currentChildItem == null)
                {
                    break;
                }

                if (i == (parentsList.Count - 1))
                {
					if (currentChildItem.IsSelected == true)
                    {
                        alreadySelected = true;
                    }
                    currentChildItem.IsSelected = true;
                    currentChildItem.BringIntoView();
                }
                else
                {
                    currentChildItem.IsExpanded = true;
                }
            }

            return alreadySelected;
        }

        private TreeViewItem FindElementInTreeViewItem(TreeViewItem rootItem, IUIAutomationElement currentParent)
        {
            if (currentParent == null)
            {
                return null;
            }

            if (rootItem.Items.Count == 0)
            {
                RefreshTreeViewItem(rootItem);

                if (rootItem.Items.Count == 0)
                {
                    return null;
                }
            }

			TreeViewItem resultTreeViewItem = FindAutomationElementInChildren(
                currentParent, rootItem);

            if (resultTreeViewItem == null)
            {
                RefreshTreeViewItem(rootItem);
				resultTreeViewItem = FindAutomationElementInChildren(currentParent, rootItem);
            }

            return resultTreeViewItem;
        }
        
        internal static int[] ArrayToIntArray(Array arr)
        {
            List<int> list = new List<int>();
            foreach (object obj in arr)
            {
                list.Add((int)obj);
            }
            return list.ToArray();
        }

        private TreeViewItem FindAutomationElementInChildren(IUIAutomationElement element, 
            TreeViewItem treeViewItem)
        {
			int[] runtimeIdArray = element.GetRuntimeId();
            //int[] runtimeIdArray = ArrayToIntArray(element.GetRuntimeId());
			//(int[])element.GetCurrentPropertyValue(AutomationElement.RuntimeIdProperty);

            tagRECT? rect = null;
            if (runtimeIdArray == null)
            {
                try
                {
                    rect = element.CurrentBoundingRectangle;
                }
                catch { }

                if (rect.HasValue == false)
                {
                    return null;
                }
            }

            TreeViewItem resultTreeViewItem = null;
            
            //////////// Also look in grandchildren ////////////
            List<TreeViewItem> itemsToSearch = new List<TreeViewItem>();
            itemsToSearch.AddRange(treeViewItem.Items.OfType<TreeViewItem>());
            foreach (TreeViewItem childTvItem in treeViewItem.Items)
            {
                itemsToSearch.AddRange(childTvItem.Items.OfType<TreeViewItem>());
            }
            ////////////////////////////////////////////////////
            
            foreach (TreeViewItem childTvItem in itemsToSearch)
            {
                TreeNode node = childTvItem.Tag as TreeNode;
                if (node == null)
                {
                    continue;
                }

                try
                {
                    if (runtimeIdArray != null)
                    {
						int[] currentRuntimeIdArray = node.Element.GetRuntimeId();
                        //int[] currentRuntimeIdArray = ArrayToIntArray(node.Element.GetRuntimeId());
                        //(int[])node.Element.GetCurrentPropertyValue(AutomationElement.RuntimeIdProperty);

                        if (ArraysEqual<int>(runtimeIdArray, currentRuntimeIdArray))
                        {
                            resultTreeViewItem = childTvItem;
                            break;
                        }
                    }
                    else
                    {
                        tagRECT? currentRect = null;
                        try
                        {
                            currentRect = node.Element.CurrentBoundingRectangle;
                        }
                        catch { }

                        if (currentRect != null)
                        {
                            if (rect.Value.Equals(currentRect.Value))
                            {
                                resultTreeViewItem = childTvItem;
                                break;
                            }
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            return resultTreeViewItem;
        }
		
		private void AlwaysOnTop_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				menuAlwaysOnTop.IsChecked = !menuAlwaysOnTop.IsChecked;
				this.Topmost = menuAlwaysOnTop.IsChecked;
			}
			catch {}
		}

        static bool ArraysEqual<T>(T[] a1, T[] a2)
        {
            if (ReferenceEquals(a1, a2))
            {
                return true;
            }

            if (a1 == null || a2 == null)
            {
                return false;
            }

            if (a1.Length != a2.Length)
            {
                return false;
            }

            EqualityComparer<T> comparer = EqualityComparer<T>.Default;
            for (int i = 0; i < a1.Length; i++)
            {
                if (!comparer.Equals(a1[i], a2[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            UnHighlight();
        }
        
        public static bool CompareElements(IUIAutomationElement element1, IUIAutomationElement element2)
		{
			int[] runtimeId1 = null;
			int[] runtimeId2 = null;
            
            if (element1 == null || element2 == null)
            {
                return false;
            }

			try
			{
				runtimeId1 = MainWindow.ArrayToIntArray(element1.GetRuntimeId());
				runtimeId2 = MainWindow.ArrayToIntArray(element2.GetRuntimeId());
				return MainWindow.ArraysEqual<int>(runtimeId1, runtimeId2);
			}
			catch
			{
				return false;
			}
		}

        internal class AutomationInfo
        {
            public int[] runtimeId = null;
            public tagRECT boundingRectangle;

            public AutomationInfo(IUIAutomationElement element)
            {
                try
                {
                    this.runtimeId = MainWindow.ArrayToIntArray(element.GetRuntimeId());
                    this.boundingRectangle = element.CurrentBoundingRectangle;
                }
                catch { }
            }

            public int[] RuntimeId
            {
                get
                {
                    return this.runtimeId;
                }
            }

            public tagRECT BoundingRectangle
            {
                get
                {
                    return this.boundingRectangle;
                }
                set
                {
                    this.boundingRectangle = value;
                }
            }

            public static bool CompareRectangles(tagRECT rect1, tagRECT rect2)
            {
                return (rect1.left == rect2.left) && (rect1.top == rect2.top) && 
                    (rect1.right == rect2.right) && (rect1.bottom == rect2.bottom);
            }

            // returns true if is the same element and position is the same
            public bool IsIdentic(IUIAutomationElement element)
            {
                if (element == null)
                {
                    return false;
                }

                if ((this.runtimeId == null) || (this.runtimeId.Length == 0))
                {
                    return false;
                }

                int[] currentRuntimeId = null;

                try
                {
                    currentRuntimeId = MainWindow.ArrayToIntArray(element.GetRuntimeId());
                    tagRECT currentBoundingRectangle = element.CurrentBoundingRectangle;

                    return (MainWindow.ArraysEqual<int>(this.runtimeId, currentRuntimeId) &&
                        AutomationInfo.CompareRectangles(this.boundingRectangle, currentBoundingRectangle));
                }
                catch
                {
                    return false;
                }
            }

            // returns true if is the same automation element no matter the position
            public bool IsSameElement(IUIAutomationElement element)
            {
                if (element == null)
                {
                    return false;
                }

                if ((this.runtimeId == null) || (this.runtimeId.Length == 0))
                {
                    return false;
                }

                int[] currentRuntimeId = null;

                try
                {
                    currentRuntimeId = MainWindow.ArrayToIntArray(element.GetRuntimeId());
                    return MainWindow.ArraysEqual<int>(this.runtimeId, currentRuntimeId);
                }
                catch
                {
                    return false;
                }
            }

            public void UpdateBoundingRect(IUIAutomationElement element)
            {
                try
                {
                    this.boundingRectangle = element.CurrentBoundingRectangle;
                }
                catch { }
            }
        }

        public static List<IUIAutomationElement> FindChildren(IUIAutomationElement automationElement)
        {
            List<IUIAutomationElement> results = new List<IUIAutomationElement>();

            if (MainWindow.mode == Modes.Control)
            {
                IUIAutomationElementArray children = automationElement.FindAllBuildCache(
                    TreeScope.TreeScope_Children, uiAutomation.CreateTrueCondition(), cacheRequest);
					
				if (children == null)
				{
					return results;
				}

                for (int i = 0; i < children.Length; i++)
                {
                    results.Add(children.GetElement(i));
                }
                return results;
            }
            else
            {
                IUIAutomationTreeWalker tw = uiAutomation.ControlViewWalker;
                if (MainWindow.mode == Modes.Content)
                {
                    tw = uiAutomation.ContentViewWalker;
                }
                else
                {
                    //Raw
                    tw = uiAutomation.RawViewWalker;
                }

                IUIAutomationElement child = tw.GetFirstChildElementBuildCache(automationElement, cacheRequest);
                while (child != null)
                {
                    results.Add(child);
                    child = tw.GetNextSiblingElementBuildCache(child, cacheRequest);
                }
            }

            return results;
        }
        
        private void OnCapture(object sender, RoutedEventArgs e)
        {
            try
            {
                Capture();
            }
            catch { }
        }
        
        private void Capture()
        {
            TreeViewItem selectedItem = this.tvElements.SelectedItem as TreeViewItem;
            if (selectedItem == null)
            {
                return;
            }

            TreeNode selectedNode = selectedItem.Tag as TreeNode;
            if (selectedNode == null)
            {
                return;
            }
            
            //System.Drawing.Bitmap bitmap = Helper.CaptureElementToBitmap(selectedNode.Element);
            /*if (bitmap == null)
            {
                System.Windows.MessageBox.Show("Cannot capture element");
                return;
            }*/
			BringToForeground(selectedNode.Element);
            
            string defaultName = selectedNode.Element.CurrentName;
            if (defaultName == null || defaultName == "")
            {
                defaultName = Helper.ControlTypeIdToString(selectedNode.Element.CurrentControlType);
                defaultName = defaultName.Replace("UIA_", "").Replace("ControlTypeId", "");
            }
            defaultName = defaultName.Replace("<", "").Replace(">", "").Replace(":", "").Replace("\"", "").Replace("/", "").Replace("\\", "").Replace("|", "").Replace("?", "").Replace("*", "").Replace(".", "");
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = defaultName;
            dlg.DefaultExt = ".bmp";
            dlg.Filter = "Jpeg picture (.jpg)|*.jpg|Bitmap picture (.bmp)|*.bmp|PNG picture (.png)|*.png";
            
            // Show save file dialog box
            bool? result = dlg.ShowDialog();
            if (result != true)
            {
                return;
            }
            
            string filename = dlg.FileName; // "C:\\Temp\\element.bmp"
			
			// Make the capture into a bitmap
			BringToForeground(selectedNode.Element);
			System.Threading.Thread.Sleep(500);
            System.Drawing.Bitmap bitmap = Helper.CaptureElementToBitmap(selectedNode.Element);
            if (bitmap == null)
            {
                System.Windows.MessageBox.Show("Cannot capture element");
                return;
            }
            
            //Helper.CaptureElementToFile(selectedNode.Element, filename);
            Helper.CaptureElementToFile(bitmap, filename);
            //this.Activate();
            //this.tvElements.Focus();
        }
		
		private void BringToForeground(IUIAutomationElement element)
		{
			IntPtr hwnd = GetWindow(element);
            if (hwnd != IntPtr.Zero)
            {
                SetForegroundWindow(hwnd);
            }
		}
		
		private IntPtr GetWindow(IUIAutomationElement element)
        {
            IUIAutomationTreeWalker tw = MainWindow.uiAutomation.ControlViewWalker;

            IntPtr hwnd = IntPtr.Zero;
            IUIAutomationElement currentElement = element;

            while (true)
            {
                try
                {
                    hwnd = currentElement.CurrentNativeWindowHandle;
                }
                catch (Exception ex)
                {
                    break;
                }

                if (hwnd != IntPtr.Zero)
                {
                    break;
                }

                currentElement = tw.GetParentElement(currentElement);
                if (currentElement == null)
                {
                    break;
                }
            }

            return hwnd;
        }

        private static Modes mode = Modes.Control;

        private bool hasAsyncContentLoadedEvent = true;
        private bool hasElementAddedToSelectionEvent = true;
        private bool hasElementRemovedFromSelectionEvent = true;
        private bool hasElementSelectedEvent = true;
        private bool hasInvalidatedEvent = true;
        private bool hasInvokedEvent = true;
        private bool hasInputReachedTargetEvent = true;
        private bool hasInputReachedOtherElementEvent = true;
        private bool hasInputDiscardedEvent = true;
        private bool hasLayoutInvalidatedEvent = true;
        private bool hasMenuOpenedEvent = true;
        private bool hasMenuClosedEvent = true;
        private bool hasMenuModeStartEvent = true;
        private bool hasMenuModeEndEvent = true;
        private bool hasTextChangedEvent = true;
        private bool hasTextSelectionChangedEvent = true;
        private bool hasToolTipClosedEvent = true;
        private bool hasToolTipOpenedEvent = true;
        private bool hasWindowOpenedEvent = true;
        private bool hasAutomationFocusChangedEvent = false; //true;
        private bool hasAutomationPropertyChangedEvent = true;
        private bool hasStructureChangedEvent = true;
        private bool hasWindowClosedEvent = true;
		
		private bool hasTextEditTextChangedEvent = false;
		private bool hasChangesEvent = false;
		private bool hasNotificationEvent = false;
		private bool hasActiveTextPositionChangedEvent = false;

        private TreeScope eventsScope = TreeScope.TreeScope_Subtree;
		private TextEditChangeType TextEditChangeType = TextEditChangeType.TextEditChangeType_None;
		private int ChangesCount = 1;

        private void OnEventsSettings(object sender, RoutedEventArgs e)
        {
            WindowEventsOptions wndEventsOptions = new WindowEventsOptions();
            wndEventsOptions.Owner = this;

            wndEventsOptions.hasAsyncContentLoadedEvent = hasAsyncContentLoadedEvent;
            wndEventsOptions.hasElementAddedToSelectionEvent = hasElementAddedToSelectionEvent;
            wndEventsOptions.hasElementRemovedFromSelectionEvent = hasElementRemovedFromSelectionEvent;
            wndEventsOptions.hasElementSelectedEvent = hasElementSelectedEvent;
            wndEventsOptions.hasInvalidatedEvent = hasInvalidatedEvent;
            wndEventsOptions.hasInvokedEvent = hasInvokedEvent;
            wndEventsOptions.hasInputReachedTargetEvent = hasInputReachedTargetEvent;
            wndEventsOptions.hasInputReachedOtherElementEvent = hasInputReachedOtherElementEvent;
            wndEventsOptions.hasInputDiscardedEvent = hasInputDiscardedEvent;
            wndEventsOptions.hasLayoutInvalidatedEvent = hasLayoutInvalidatedEvent;
            wndEventsOptions.hasMenuOpenedEvent = hasMenuOpenedEvent;
            wndEventsOptions.hasMenuClosedEvent = hasMenuClosedEvent;
            wndEventsOptions.hasMenuModeStartEvent = hasMenuModeStartEvent;
            wndEventsOptions.hasMenuModeEndEvent = hasMenuModeEndEvent;
            wndEventsOptions.hasTextChangedEvent = hasTextChangedEvent;
            wndEventsOptions.hasTextSelectionChangedEvent = hasTextSelectionChangedEvent;
            wndEventsOptions.hasToolTipClosedEvent = hasToolTipClosedEvent;
            wndEventsOptions.hasToolTipOpenedEvent = hasToolTipOpenedEvent;
            wndEventsOptions.hasWindowOpenedEvent = hasWindowOpenedEvent;
            wndEventsOptions.hasAutomationFocusChangedEvent = hasAutomationFocusChangedEvent;
            wndEventsOptions.hasAutomationPropertyChangedEvent = hasAutomationPropertyChangedEvent;
            wndEventsOptions.hasStructureChangedEvent = hasStructureChangedEvent;
            wndEventsOptions.hasWindowClosedEvent = hasWindowClosedEvent;
			
			wndEventsOptions.hasTextEditTextChangedEvent = hasTextEditTextChangedEvent;
			wndEventsOptions.hasChangesEvent = hasChangesEvent;
			wndEventsOptions.hasNotificationEvent = hasNotificationEvent;
			wndEventsOptions.hasActiveTextPositionChangedEvent = hasActiveTextPositionChangedEvent;
            wndEventsOptions.eventsScope = eventsScope;
			
			wndEventsOptions.TextEditChangeType = TextEditChangeType;
			wndEventsOptions.ChangesCount = ChangesCount;

            if (wndEventsOptions.ShowDialog() == true)
            {
                hasAsyncContentLoadedEvent = wndEventsOptions.hasAsyncContentLoadedEvent;
                hasElementAddedToSelectionEvent = wndEventsOptions.hasElementAddedToSelectionEvent;
                hasElementRemovedFromSelectionEvent = wndEventsOptions.hasElementRemovedFromSelectionEvent;
                hasElementSelectedEvent = wndEventsOptions.hasElementSelectedEvent;
                hasInvalidatedEvent = wndEventsOptions.hasInvalidatedEvent;
                hasInvokedEvent = wndEventsOptions.hasInvokedEvent;
                hasInputReachedTargetEvent = wndEventsOptions.hasInputReachedTargetEvent;
                hasInputReachedOtherElementEvent = wndEventsOptions.hasInputReachedOtherElementEvent;
                hasInputDiscardedEvent = wndEventsOptions.hasInputDiscardedEvent;
                hasLayoutInvalidatedEvent = wndEventsOptions.hasLayoutInvalidatedEvent;
                hasMenuOpenedEvent = wndEventsOptions.hasMenuOpenedEvent;
                hasMenuClosedEvent = wndEventsOptions.hasMenuClosedEvent;
                hasMenuModeStartEvent = wndEventsOptions.hasMenuModeStartEvent;
                hasMenuModeEndEvent = wndEventsOptions.hasMenuModeEndEvent;
                hasTextChangedEvent = wndEventsOptions.hasTextChangedEvent;
                hasTextSelectionChangedEvent = wndEventsOptions.hasTextSelectionChangedEvent;
                hasToolTipClosedEvent = wndEventsOptions.hasToolTipClosedEvent;
                hasToolTipOpenedEvent = wndEventsOptions.hasToolTipOpenedEvent;
                hasWindowOpenedEvent = wndEventsOptions.hasWindowOpenedEvent;
                hasAutomationFocusChangedEvent = wndEventsOptions.hasAutomationFocusChangedEvent;
                hasAutomationPropertyChangedEvent = wndEventsOptions.hasAutomationPropertyChangedEvent;
                hasStructureChangedEvent = wndEventsOptions.hasStructureChangedEvent;
                hasWindowClosedEvent = wndEventsOptions.hasWindowClosedEvent;
				
				hasTextEditTextChangedEvent = wndEventsOptions.hasTextEditTextChangedEvent;
				hasChangesEvent = wndEventsOptions.hasChangesEvent;
				hasNotificationEvent = wndEventsOptions.hasNotificationEvent;
				hasActiveTextPositionChangedEvent = wndEventsOptions.hasActiveTextPositionChangedEvent;
                eventsScope = wndEventsOptions.eventsScope;
				
				TextEditChangeType = wndEventsOptions.TextEditChangeType;
				ChangesCount = wndEventsOptions.ChangesCount;
            }
        }

        private void MenuItemClear_Click(object sender, RoutedEventArgs e)
        {
            eventsCollection.Clear();
        }
		
		private void MenuItemClearSelected_Click(object sender, RoutedEventArgs e)
		{
			if (eventsListView.SelectedItems.Count == 0)
            {
                return;
            }
			
			try
			{
				while (eventsListView.SelectedItems.Count > 0)
				{
					int idx = eventsListView.Items.IndexOf(eventsListView.SelectedItem);
					eventsCollection.RemoveAt(idx);
				}
			}
			catch { }
		}

        private void MenuItemGoTo_Click(object sender, RoutedEventArgs e)
        {
            if (eventsListView.SelectedItems.Count == 0)
            {
                return;
            }

            Event ev = eventsListView.SelectedItems[0] as Event;
            if (ev == null || ev.Element == null)
            {
                return;
            }

            if ((new TreeNode(ev.Element)).IsAlive)
            {
                SelectAutomationElementInTree(ev.Element);
                this.tvElements.Focus();
            }
            else
            {
                System.Windows.MessageBox.Show("Element not available anymore");
            }
        }
        
        private void MenuItemHighlight_Click(object sender, RoutedEventArgs e)
        {
            if (eventsListView.SelectedItems.Count == 0)
            {
                return;
            }

            Event ev = eventsListView.SelectedItems[0] as Event;
            if (ev == null)
            {
                return;
            }

            HighlightItemElement(ev.Element);
        }

        private void CmbViewMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbViewMode.SelectedIndex == 0)
            {
                MainWindow.mode = Modes.Control;
            }
            else if (cmbViewMode.SelectedIndex == 1)
            {
                MainWindow.mode = Modes.Raw;
            }
            else
            {
                MainWindow.mode = Modes.Content;
            }

            //try
            //{
                RefreshTree();
            //}
            //catch { }
        }

		private void OnShowCP(object sender, RoutedEventArgs e)
		{
			bool showCP = showCPMenuItem.IsChecked;
			menuClickablePoint.IsChecked = showCP;
			
			if (showCP)
			{
				TreeViewItem selectedItem = tvElements.SelectedItem as TreeViewItem;
				if (selectedItem == null)
				{
					return;
				}

				TreeNode selectedNode = selectedItem.Tag as TreeNode;
				if (selectedNode == null)
				{
					return;
				}
				
				if (selectedNode.IsAlive)
				{
					try
					{
						HighlightCP(selectedNode.Element);
					}
					catch {}
				}
			}
			else
			{
				try
				{
					UnHighlightCP();
				}
				catch {}
			}
		}
		
		private Form frmCP1 = null;
		private Form frmCP2 = null;
		private System.Drawing.Color cpColor = System.Drawing.Color.DodgerBlue;
		int cpLen = 5;
		
		private void UnHighlightCP()
		{	
			if (frmCP1 != null && frmCP2 != null)
			{
				frmCP1.Location = new System.Drawing.Point(-100, -100);
				frmCP2.Location = new System.Drawing.Point(-100, -100);
			}
			else
			{
				// vertical
				frmCP1 = new Form();
                frmCP1.Text = "AutomationSpy_CP1";
				frmCP1.FormBorderStyle = FormBorderStyle.None;
				frmCP1.StartPosition = FormStartPosition.Manual;
				
				frmCP1.Location = new System.Drawing.Point(-100, -100);
				frmCP1.MinimumSize = new System.Drawing.Size(0, 0);
				frmCP1.Size = new System.Drawing.Size(0, 0);
				frmCP1.ShowInTaskbar = false;
				frmCP1.TopMost = true;
				frmCP1.ForeColor = cpColor;
				frmCP1.BackColor = cpColor;
				
				frmCP1.Load += window1_Loaded;
				frmCP1.Show();
				
				// horizontal
				frmCP2 = new Form();
                frmCP2.Text = "AutomationSpy_CP2";
				frmCP2.FormBorderStyle = FormBorderStyle.None;
				frmCP2.StartPosition = FormStartPosition.Manual;
				
				frmCP2.Location = new System.Drawing.Point(-100, -100);
				frmCP2.MinimumSize = new System.Drawing.Size(0, 0);
				frmCP2.Size = new System.Drawing.Size(0, 0);
				frmCP2.ShowInTaskbar = false;
				frmCP2.TopMost = true;
				frmCP2.ForeColor = cpColor;
				frmCP2.BackColor = cpColor;
				
				frmCP2.Load += window2_Loaded;
				frmCP2.Show();
				
				this.Focus();
			}
		}
		
		private void HighlightCP(IUIAutomationElement element)
		{
            tagPOINT pt;
            /*element.GetClickablePoint(out pt);
            if (pt.x == 0 && pt.y == 0)
            {
                return;
            }*/
			
			if (CompareElements(element, uiAutomation.GetRootElement()) == true)
			{
				return;
			}
			
            if (element.GetClickablePoint(out pt) == 0)
            {
                return;
            }
            
            int left = pt.x - 1;
            int top = pt.y - cpLen;
            if (frmCP1 != null)
            {
                frmCP1.Location = new System.Drawing.Point(left, top);
                frmCP1.BringToFront();
            }
            else
            {
                frmCP1 = new Form();
                frmCP1.Text = "AutomationSpy_CP1";
                frmCP1.FormBorderStyle = FormBorderStyle.None;
                frmCP1.StartPosition = FormStartPosition.Manual;
                
                frmCP1.Location = new System.Drawing.Point(left, top);
                frmCP1.MinimumSize = new System.Drawing.Size(0, 0);
                frmCP1.Size = new System.Drawing.Size(0, 0);
                frmCP1.ShowInTaskbar = false;
                frmCP1.TopMost = true;
                frmCP1.ForeColor = cpColor;
                frmCP1.BackColor = cpColor;
                
                frmCP1.Load += window1_Loaded;
                frmCP1.Show();
            }
            
            left = pt.x - cpLen;
            top = pt.y - 1;
            if (frmCP2 != null)
            {
                frmCP2.Location = new System.Drawing.Point(left, top);
                frmCP2.BringToFront();
            }
            else
            {
                frmCP2 = new Form();
                frmCP2.Text = "AutomationSpy_CP2";
                frmCP2.FormBorderStyle = FormBorderStyle.None;
                frmCP2.StartPosition = FormStartPosition.Manual;
                
                frmCP2.Location = new System.Drawing.Point(left, top);
                frmCP2.MinimumSize = new System.Drawing.Size(0, 0);
                frmCP2.Size = new System.Drawing.Size(0, 0);
                frmCP2.ShowInTaskbar = false;
                frmCP2.TopMost = true;
                frmCP2.ForeColor = cpColor;
                frmCP2.BackColor = cpColor;
                
                frmCP2.Load += window2_Loaded;
                frmCP2.Show();
            }
            
            this.Focus();
		}
		
		private void window1_Loaded(object sender, System.EventArgs e)
		{
			frmCP1.Size = new System.Drawing.Size(3, 2 * cpLen + 1);
		}
		
		private void window2_Loaded(object sender, System.EventArgs e)
		{
			frmCP2.Size = new System.Drawing.Size(2 * cpLen + 1, 3);
		}
        
        private void SavePreferences()
        {
            XmlDocument doc = new XmlDocument();

            XmlNode root = doc.CreateNode(XmlNodeType.Element, "preferences", null);
            doc.AppendChild(root);

            /*XmlNode node = doc.CreateNode(XmlNodeType.Element, "Version", null);
            node.InnerText = MainWindow.Version.ToString();
            root.AppendChild(node);*/
            
            XmlNode node = doc.CreateNode(XmlNodeType.Element, "AsyncContentLoadedEvent", null);
            node.InnerText = hasAsyncContentLoadedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "ElementAddedToSelectionEvent", null);
            node.InnerText = hasElementAddedToSelectionEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "ElementRemovedFromSelectionEvent", null);
            node.InnerText = hasElementRemovedFromSelectionEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "ElementSelectedEvent", null);
            node.InnerText = hasElementSelectedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "InvalidatedEvent", null);
            node.InnerText = hasInvalidatedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "InvokedEvent", null);
            node.InnerText = hasInvokedEvent.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "InputReachedTargetEvent", null);
            node.InnerText = hasInputReachedTargetEvent.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "InputReachedOtherElementEvent", null);
            node.InnerText = hasInputReachedOtherElementEvent.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "InputDiscardedEvent", null);
            node.InnerText = hasInputDiscardedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "LayoutInvalidatedEvent", null);
            node.InnerText = hasLayoutInvalidatedEvent.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "MenuOpenedEvent", null);
            node.InnerText = hasMenuOpenedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "MenuClosedEvent", null);
            node.InnerText = hasMenuClosedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "MenuModeStartEvent", null);
            node.InnerText = hasMenuModeStartEvent.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "MenuModeEndEvent", null);
            node.InnerText = hasMenuModeEndEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "TextChangedEvent", null);
            node.InnerText = hasTextChangedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "TextSelectionChangedEvent", null);
            node.InnerText = hasTextSelectionChangedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "ToolTipClosedEvent", null);
            node.InnerText = hasToolTipClosedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "ToolTipOpenedEvent", null);
            node.InnerText = hasToolTipOpenedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "WindowOpenedEvent", null);
            node.InnerText = hasWindowOpenedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "AutomationFocusChangedEvent", null);
            node.InnerText = hasAutomationFocusChangedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "AutomationPropertyChangedEvent", null);
            node.InnerText = hasAutomationPropertyChangedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "StructureChangedEvent", null);
            node.InnerText = hasStructureChangedEvent.ToString();
            root.AppendChild(node);

            node = doc.CreateNode(XmlNodeType.Element, "WindowClosedEvent", null);
            node.InnerText = hasWindowClosedEvent.ToString();
            root.AppendChild(node);
			
			if (hasTextEditTextChangedEvent == true)
			{
				node = doc.CreateNode(XmlNodeType.Element, "TextEditTextChangedEvent", null);
				node.InnerText = hasTextEditTextChangedEvent.ToString();
				root.AppendChild(node);
			}
			
			if (hasChangesEvent == true)
			{
				node = doc.CreateNode(XmlNodeType.Element, "ChangesEvent", null);
				node.InnerText = hasChangesEvent.ToString();
				root.AppendChild(node);
			}
			
			if (hasNotificationEvent == true)
			{
				node = doc.CreateNode(XmlNodeType.Element, "NotificationEvent", null);
				node.InnerText = hasNotificationEvent.ToString();
				root.AppendChild(node);
			}
			
			if (hasActiveTextPositionChangedEvent == true)
			{
				node = doc.CreateNode(XmlNodeType.Element, "ActiveTextPositionChangedEvent", null);
				node.InnerText = hasActiveTextPositionChangedEvent.ToString();
				root.AppendChild(node);
			}
			

            node = doc.CreateNode(XmlNodeType.Element, "Scope", null);
            node.InnerText = eventsScope.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "Width", null);
            node.InnerText = this.Width.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "Height", null);
            node.InnerText = this.Height.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "Left", null);
            node.InnerText = this.Left.ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "Top", null);
            node.InnerText = this.Top.ToString();
            root.AppendChild(node);
            
            /*node = doc.CreateNode(XmlNodeType.Element, "ColWidth1", null);
            node.InnerText = Math.Round(mainGrid.ColumnDefinitions[0].Width.Value, 2).ToString();
            root.AppendChild(node);
            
            node = doc.CreateNode(XmlNodeType.Element, "ColWidth2", null);
            node.InnerText = Math.Round(mainGrid.ColumnDefinitions[2].Width.Value, 2).ToString();
            root.AppendChild(node);*/
			
			node = doc.CreateNode(XmlNodeType.Element, "AlwaysOnTop", null);
            node.InnerText = menuAlwaysOnTop.IsChecked.ToString();
            root.AppendChild(node);

            doc.Save(prefFileName);
        }
        
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                SavePreferences();
            }
            catch { }
			
			if (listenerInstalled != null)
			{
				System.Threading.Tasks.Task.Run(() => UninstallListener()).Wait(1000);
			}
        }
		
		private void optionsButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				PropertySettings dlgPropSettings = new PropertySettings() { Owner = this };
				if (dlgPropSettings.ShowDialog() == true)
				{
					RefreshTree();
				}
			}
			catch {}
		}
		
		Timer timerTrack = null;
		private void menuItemTrack_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				menuItemTrack.IsChecked = !menuItemTrack.IsChecked;
				if (menuItemTrack.IsChecked)
				{
					if (menuItemPick.IsChecked == true)
					{
						menuItemPick.IsChecked = false;
						selectButton.IsChecked = false;
						try
						{
							TryOnSelectElement();
						}
						catch {}
					}
				
					if (timerTrack == null)
					{
						timerTrack = new Timer();
						timerTrack.Interval = 1000;
						timerTrack.Tick += timerTrack_Tick;
					}
					timerTrack.Start();
				}
				else
				{
					if (timerTrack != null)
					{
						timerTrack.Stop();
					}
				}
			}
			catch {}
		}
		
		private void timerTrack_Tick(object sender, EventArgs e)
		{
			try
			{
				IUIAutomationElement element = uiAutomation.GetFocusedElement();
				if (element != null)
				{
					if (SelectElement(element) == true && tvElements.IsFocused == false)
					{
						tvElements.Focus();
					}
				}
			}
			catch {}
		}
		
		private void OnListViewDoubleClick(object sender, MouseButtonEventArgs e)
		{
			Attribute selectedAttribute = patternsListView.SelectedItem as Attribute;
			if (selectedAttribute == null)
			{
				return;
			}
			
			if (selectedAttribute.Name == "Double click here...")
			{
				MoreAttributes moreAttributes = new MoreAttributes(selectedAttribute.TextPatternRange) { Owner = this };
				moreAttributes.ShowDialog();
			}
			else if (selectedAttribute.Value.EndsWith(" -> (Double-click to see all text)"))
			{
				try
				{
					WindowAllText windowAllText = new WindowAllText(selectedAttribute) { Owner = this };
					windowAllText.ShowDialog();
				}
				catch 
				{
					System.Windows.MessageBox.Show("Cannot get text. The element may not be available anymore.");
				}
			}
		}
		
		private void OnShowContextMenu(object sender, RoutedEventArgs e)
		{
			TreeViewItem treeviewItem = this.tvElements.SelectedItem as TreeViewItem;
            if (treeviewItem == null)
            {
				System.Windows.MessageBox.Show("Select an element in the tree");
                return;
            }
			
			TreeNode node = treeviewItem.Tag as TreeNode;
            if (node == null)
            {
                return;
            }
			
			IUIAutomationElement3 element3 = node.Element as IUIAutomationElement3;
			if (element3 == null)
			{
				System.Windows.MessageBox.Show("IUIAutomationElement3 not supported");
			}
			
			try
			{
				element3.ShowContextMenu();
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message);
			}
		}
    }
}
