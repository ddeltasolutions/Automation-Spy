using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using AutomationSpy;
using System.Diagnostics;
using System.Xml;
using System.Threading;
using System.Threading.Tasks;
using Interop.UIAutomationClient;

namespace dDeltaSolutions.Spy
{
    public class Event
    {
        private int eventId;
        private string eventIdString = "";
        private IUIAutomationElement element = null;
        private string tooltip = null;
        private string details = "";
		private string elementString = "";

        public Event(int eventId, IUIAutomationElement element)
        {
            this.eventId = eventId;
            this.element = element;
            
            try
            {
            	eventIdString = GetEventName(eventId); //ProgrammaticName;
            }
            catch { }
            tooltip = eventIdString;
			
			try
			{
				elementString = GetStringFromElement(this.element);
			}
			catch { }
        }
		
		public Event(string eventIdString, IUIAutomationElement element)
		{
			this.element = element;
			this.eventIdString = eventIdString;
			tooltip = eventIdString;
			
			try
			{
				elementString = GetStringFromElement(this.element);
			}
			catch { }
		}
        
        public string GetEventName(int eventId)
        {
            if (eventId == UIA_EventIds.UIA_ToolTipOpenedEventId)
            {
                return "UIA_ToolTipOpenedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_ToolTipClosedEventId)
            {
                return "UIA_ToolTipClosedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_StructureChangedEventId)
            {
                return "UIA_StructureChangedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_MenuOpenedEventId)
            {
                return "UIA_MenuOpenedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_AutomationPropertyChangedEventId)
            {
                return "UIA_AutomationPropertyChangedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_AutomationFocusChangedEventId)
            {
                return "UIA_AutomationFocusChangedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_AsyncContentLoadedEventId)
            {
                return "UIA_AsyncContentLoadedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_MenuClosedEventId)
            {
                return "UIA_MenuClosedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_LayoutInvalidatedEventId)
            {
                return "UIA_LayoutInvalidatedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_Invoke_InvokedEventId)
            {
                return "UIA_Invoke_InvokedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_SelectionItem_ElementAddedToSelectionEventId)
            {
                return "UIA_SelectionItem_ElementAddedToSelectionEvent";
            }
            else if (eventId == UIA_EventIds.UIA_SelectionItem_ElementRemovedFromSelectionEventId)
            {
                return "UIA_SelectionItem_ElementRemovedFromSelectionEvent";
            }
            else if (eventId == UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId)
            {
                return "UIA_SelectionItem_ElementSelectedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_Selection_InvalidatedEventId)
            {
                return "UIA_Selection_InvalidatedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_Text_TextSelectionChangedEventId)
            {
                return "UIA_Text_TextSelectionChangedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_Text_TextChangedEventId)
            {
                return "UIA_Text_TextChangedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_Window_WindowOpenedEventId)
            {
                return "UIA_Window_WindowOpenedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_Window_WindowClosedEventId)
            {
                return "UIA_Window_WindowClosedEvent";
            }
            else if (eventId == UIA_EventIds.UIA_MenuModeStartEventId)
            {
                return "UIA_MenuModeStartEvent";
            }
            else if (eventId == UIA_EventIds.UIA_MenuModeEndEventId)
            {
                return "UIA_MenuModeEndEvent";
            }
            else if (eventId == UIA_EventIds.UIA_InputReachedTargetEventId)
            {
                return "UIA_InputReachedTargetEvent";
            }
            else if (eventId == UIA_EventIds.UIA_InputReachedOtherElementEventId)
            {
                return "UIA_InputReachedOtherElementEvent";
            }
            else if (eventId == UIA_EventIds.UIA_InputDiscardedEventId)
            {
                return "UIA_InputDiscardedEvent";
            }
            return eventId.ToString();

        }

        public int EventId
        {
            get
            {
                return this.eventId;
            }
        }

        public string EventIdString
        {
            get
            {
                return eventIdString;
            }
        }

        public IUIAutomationElement Element
        {
            get
            {
                return this.element;
            }
        }

        public string ElementString
        {
            get
            {
                return this.elementString;
            }
        }

        public static string GetStringFromElement(IUIAutomationElement el)
        {
            string name = null;
			if (el == null)
			{
				return "(null)";
			}

            try
            {
                name = el.CurrentName;
            }
            catch { }

            if (name == null)
            {
                return "(null)";
            }

            if (name.Length > 30)
            {
                name = name.Substring(0, 30) + "...";
            }
			
			int controlType = el.CurrentControlType;
			if (controlType == UIA_ControlTypeIds.UIA_DataItemControlTypeId)
			{
				string firstText = GetFirstText(el);
				if (firstText != null)
				{
					firstText = firstText.Trim();
					if (firstText != "")
					{
						if (firstText.Length > 20)
						{
							firstText = firstText.Substring(0, 20);
						}
						name += (" <" + firstText + "...>");
					}
				}
			}

            name = "\"" + name + "\"";

            try
            {
				string controlTypeName = null;
				if (cacheTypeNames.ContainsKey(controlType))
				{
					cacheTypeNames.TryGetValue(controlType, out controlTypeName);
				}
				else
				{
					controlTypeName = Helper.ControlTypeIdToString(controlType);
					if (controlTypeName != "")
					{
						controlTypeName = controlTypeName.Remove(0, 4);
						controlTypeName = controlTypeName.Remove(controlTypeName.Length - 13);
					}
					
					cacheTypeNames.TryAdd(controlType, controlTypeName);
				}
			
				name = name + " (" + controlTypeName + ")";
            }
            catch { }

            return name;
        }
		
		private static System.Collections.Concurrent.ConcurrentDictionary<int, string> cacheTypeNames = 
			new System.Collections.Concurrent.ConcurrentDictionary<int, string>();
		
		private static string GetFirstText(IUIAutomationElement el)
		{
			try
			{
				IUIAutomationCondition typeCondition = MainWindow.uiAutomation.CreatePropertyCondition(
					UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_TextControlTypeId);
				TreeScope scope = TreeScope.TreeScope_Descendants; //TreeScope_Children

				IUIAutomationElementArray collection = el.FindAll(scope, typeCondition);
				if (collection != null && collection.Length > 0)
				{
					IUIAutomationElement firstText = collection.GetElement(0);
					if (firstText != null)
					{
						return firstText.CurrentName;
					}
				}
				
				return null;
			}
			catch 
			{
				return null;
			}
		}

        public string Tooltip
        {
            get
            {
                return EventIdString;
            }
        }

        public string Details
        {
            get
            {
                return details;
            }
            set
            {
                details = value;
            }
        }
    }

    public partial class MainWindow
    {
        public ObservableCollection<Event> eventsCollection = new ObservableCollection<Event>();

        private void OnEvents(object sender, RoutedEventArgs e)
        {
			try
			{
				OnEvents();
			}
			catch {}
		}
		
		private void OnEvents()
		{
            TreeViewItem selectedItem = tvElements.SelectedItem as TreeViewItem;
            TreeNode node = null;

            if (this.eventsPropButton.IsEnabled == true)
            {
                if (selectedItem == null)
                {
                    MessageBox.Show(this, "Select an element in the tree");
                    return;
                }

                node = selectedItem.Tag as TreeNode;
                if (node == null)
                {
                    return;
                }

                if (node.IsAlive == false)
                {
                    MessageBox.Show(this, "The selected UI element is not available anymore");
                    return;
                }
				
				try
				{
					if (node.Element.CurrentProcessId == Process.GetCurrentProcess().Id)
					{
						MessageBox.Show(this, "You shouldn't monitor events on Automation Spy window");
						return;
					}
				}
				catch { }
            }

            if (hasAsyncContentLoadedEvent == false && hasElementAddedToSelectionEvent == false &&
                hasElementRemovedFromSelectionEvent == false && hasElementSelectedEvent == false &&
                hasInvalidatedEvent == false && hasInvokedEvent == false &&
                hasInputReachedTargetEvent == false && hasInputReachedOtherElementEvent == false && 
                hasInputDiscardedEvent == false &&
                hasLayoutInvalidatedEvent == false && hasMenuClosedEvent == false &&
                hasMenuOpenedEvent == false && hasMenuModeStartEvent == false &&
                hasMenuModeEndEvent == false && hasTextChangedEvent == false &&
                hasTextSelectionChangedEvent == false && hasToolTipClosedEvent == false &&
                hasToolTipOpenedEvent == false && hasWindowOpenedEvent == false &&
                hasAutomationFocusChangedEvent == false && hasAutomationPropertyChangedEvent == false &&
                hasStructureChangedEvent == false && hasWindowClosedEvent == false && hasTextEditTextChangedEvent == false &&
				hasChangesEvent == false && hasNotificationEvent == false && hasActiveTextPositionChangedEvent == false)
            {
                MessageBox.Show(this, "No event type is selected in the Events Settings dialog");
                return;
            }

            this.eventsListView.ItemsSource = eventsCollection;

            // install watchers
            TabItem selectedTab = this.tabCtrl.SelectedItem as TabItem;
            if (selectedTab == null)
            {
                return;
            }

			if (selectedTab != this.tabEvents && this.eventsPropButton.IsEnabled == true)
            {
                this.tabCtrl.SelectedItem = tabEvents;
            }

            this.eventsPropButton.IsEnabled = !this.eventsPropButton.IsEnabled;
			menuEventsSettings.IsEnabled = this.eventsPropButton.IsEnabled;

            if (selectButton.IsChecked == true)
            {
                selectButton.IsChecked = false;
                OnSelectElement(selectButton, null);
            }
            this.selectButton.IsEnabled = this.eventsPropButton.IsEnabled;
			this.menuItemPick.IsEnabled = this.eventsPropButton.IsEnabled;
			
			if (menuItemTrack.IsChecked == true)
			{
				menuItemTrack.IsChecked = false;
				if (timerTrack != null)
				{
					timerTrack.Stop();
				}
			}
			this.menuItemTrack.IsEnabled = this.eventsPropButton.IsEnabled;

            StackPanel panel = this.eventsButton.Content as StackPanel;
            Image img = null;

            if ((panel != null) && (panel.Children.Count > 0))
            {
                img = panel.Children[0] as Image;
            }

            if (this.eventsPropButton.IsEnabled == false)
            {
                // Capturing
                //Task.Run(() => InstallListener(node.Element));
				InstallListener(node.Element);

                this.eventsButton.ToolTip = "Stop capturing UI Automation events";
				eventsMenu.Header = "Stop Capturing Events";
				menuStartCapturing.Header = "Stop Capturing Events";
				menuStartCapturing.ToolTip = "Stop capturing UI Automation events for the selected element";

                if (img != null)
                {
                    Uri sourceUri = new Uri("pack://application:,,,/AutomationSpy;component/Pictures/Stop.png", UriKind.Absolute);
                    BitmapImage thisIcon = new BitmapImage(sourceUri);
                    img.Source = thisIcon;
                }
            }
            else
            { 
                // Not capturing
                Task.Run(() => UninstallListener());

                this.eventsButton.ToolTip = "Start capturing UI Automation events for the selected element";
				eventsMenu.Header = "Start Capturing Events";
				menuStartCapturing.Header = "Start Capturing Events";
				menuStartCapturing.ToolTip = "Start capturing UI Automation events for the selected element";

                if (img != null)
                {
                    Uri sourceUri = new Uri("pack://application:,,,/AutomationSpy;component/Pictures/Play.png", UriKind.Absolute);
                    BitmapImage thisIcon = new BitmapImage(sourceUri);
                    img.Source = thisIcon;
                }
            }
        }
        
        internal class UIA_AutomationEventHandler: IUIAutomationEventHandler
        {
            private MainWindow mainWindow = null;
            
            public UIA_AutomationEventHandler(MainWindow mainWindow)
            {
                this.mainWindow = mainWindow;
            }
            
            public void HandleAutomationEvent(IUIAutomationElement sender, int eventId)
            {
                try
                {
                    if (sender != null && sender.CurrentProcessId == MainWindow.crtProcessId)
                    {
                        return;
                    }
                }
                catch {}
                
                Event ev = new Event(eventId, sender);
                mainWindow.Dispatcher.Invoke(new Action(() =>
                    {
                        mainWindow.eventsCollection.Add(ev);
						mainWindow.eventsListView.ScrollIntoView(ev);
                    }));
            }
        }
        
        internal class UIA_AutomationFocusChangedEventHandler: IUIAutomationFocusChangedEventHandler
        {
            private MainWindow mainWindow = null;
            
            public UIA_AutomationFocusChangedEventHandler(MainWindow mainWindow)
            {
                this.mainWindow = mainWindow;
            }
            
            public void HandleFocusChangedEvent(IUIAutomationElement sender)
            {
                try
                {
                    if (sender != null && sender.CurrentProcessId == MainWindow.crtProcessId)
                    {
                        return;
                    }
                }
                catch {}
                
                Event ev = new Event(UIA_EventIds.UIA_AutomationFocusChangedEventId, sender);
                mainWindow.Dispatcher.Invoke(new Action(() =>
                    {
                        mainWindow.eventsCollection.Add(ev);
						mainWindow.eventsListView.ScrollIntoView(ev);
                    }));
            }
        }
        
        internal class UIA_AutomationStructureChangedEventHandler: IUIAutomationStructureChangedEventHandler
        {
            private MainWindow mainWindow = null;
            
            public UIA_AutomationStructureChangedEventHandler(MainWindow mainWindow)
            {
                this.mainWindow = mainWindow;
            }
            
            public void HandleStructureChangedEvent(IUIAutomationElement sender, StructureChangeType changeType, int[] runtimeId)
            {
                try
                {
                    if (sender != null && sender.CurrentProcessId == MainWindow.crtProcessId)
                    {
                        return;
                    }
                }
                catch {}
                
                Event ev = new Event(UIA_EventIds.UIA_StructureChangedEventId, sender);
                ev.Details = "StructureChangeType: " + changeType.ToString() + ", RuntimeId: [" +
                    GetStringFromArray(runtimeId) + "]";

                mainWindow.Dispatcher.Invoke(new Action(() =>
                    {
                        mainWindow.eventsCollection.Add(ev);
                        mainWindow.eventsListView.ScrollIntoView(ev);
                    }));
            }
        }
        
        internal class UIA_AutomationPropertyChangedEventHandler: IUIAutomationPropertyChangedEventHandler
        {
            private MainWindow mainWindow = null;
            
            public UIA_AutomationPropertyChangedEventHandler(MainWindow mainWindow)
            {
                this.mainWindow = mainWindow;
            }
            
            public void HandlePropertyChangedEvent(IUIAutomationElement sender, int propertyId, object newValue)
            {
                try
                {
                    if (sender != null && sender.CurrentProcessId == MainWindow.crtProcessId)
                    {
                        return;
                    }
                }
                catch {}
                
                Event ev = new Event(UIA_EventIds.UIA_AutomationPropertyChangedEventId, sender);
                
                string newValueStr = "";
                if (newValue == null)
                {
                    newValueStr = "(null)";
                }
                else
                {
                    Type type = newValue.GetType();
                    if (type == typeof(tagRECT))
                    {
                        tagRECT rect = (tagRECT)newValue;
                        string rectangleString = "Left:" + rect.left +
                            ", Top:" + rect.top +
                            ", Right:" + rect.right +
                            ", Bottom:" + rect.bottom;
                        //newValue = "{" + MainWindow.RectangleToString((tagRECT)newValue) + "}";
                        newValue = "{" + rectangleString + "}";
                    }
                    else if ((propertyId == UIA_PropertyIds.UIA_BoundingRectanglePropertyId) &&
                             (type == typeof(double[])))
                    {
                        double[] doubleArr = (double[])newValue;
                        if (doubleArr.Length == 4)
                        {
                            newValueStr = "{Left: " + doubleArr[0].ToString() + ", Top: " + doubleArr[1] +
                                ", Width: " + doubleArr[2] + ", Height: " + doubleArr[3] + "}";
                        }
                    }
                    else if (type == typeof(IUIAutomationElement))
                    {
                        newValueStr = Event.GetStringFromElement((IUIAutomationElement)newValue);
                    }
                    else if (type == typeof(tagPOINT))
                    {
                        tagPOINT pt = (tagPOINT)newValue;
                        string pointString = "X: " + pt.x + ", Y: " + pt.y;
                        //newValueStr = "{" + GetStringFromPoint((tagPOINT)newValue) + "}";
                        newValueStr = "{" + pointString + "}";
                    }
                    else if (propertyId == UIA_PropertyIds.UIA_ControlTypePropertyId)
                    {
                        newValueStr = Helper.ControlTypeIdToString((int)newValue);
                    }
                    else if (propertyId == UIA_PropertyIds.UIA_WindowWindowVisualStatePropertyId)
                    {
                        try
                        {
                            newValueStr = ((WindowVisualState)newValue).ToString();
                        }
                        catch
                        {
                            newValueStr = newValue.ToString();
                        }
                    }
                    else if (propertyId == UIA_PropertyIds.UIA_WindowWindowInteractionStatePropertyId)
                    {
                        try
                        {
                            newValueStr = ((WindowInteractionState)newValue).ToString();
                        }
                        catch
                        {
                            newValueStr = newValue.ToString();
                        }
                    }
                    else if (propertyId == UIA_PropertyIds.UIA_ToggleToggleStatePropertyId)
                    {
                        try
                        {
                            newValueStr = ((ToggleState)newValue).ToString();
                        }
                        catch
                        {
                            newValueStr = newValue.ToString();
                        }
                    }
                    else if (propertyId == UIA_PropertyIds.UIA_ExpandCollapseExpandCollapseStatePropertyId)
                    {
                        try
                        {
                            newValueStr = ((ExpandCollapseState)newValue).ToString();
                        }
                        catch
                        {
                            newValueStr = newValue.ToString();
                        }
                    }
                    /*else if (propertyId == UIA_PropertyIds.UIA_DockDockPositionPropertyId)
                    {
                        try
                        {
                            newValueStr = ((DockPosition)newValue).ToString();
                        }
                        catch
                        {
                            newValueStr = newValue.ToString();
                        }
                    }*/
                    /*else if (propertyId == UIA_PropertyIds.UIA_OrientationPropertyId)
                    {
                        try
                        {
                            newValueStr = ((OrientationType)newValue).ToString();
                        }
                        catch
                        {
                            newValueStr = newValue.ToString();
                        }
                    }*/
                    else
                    {
                        if (type == typeof(string))
                        {
							if (propertyId == UIA_PropertyIds.UIA_ValueValuePropertyId)
							{
								newValueStr = newValue.ToString();
								if (newValueStr.Length > MAX_VALUE_LENGTH)
								{
									newValueStr = newValueStr.Substring(0, MAX_VALUE_LENGTH) + "...";
								}
								newValueStr = "\"" + newValueStr + "\"";
							}
							else
							{
								newValueStr = "\"" + newValue.ToString() + "\"";
							}
                        }
                        /*else if (type == typeof(double))
                        {
                            newValueStr = Math.Round((double)newValue, 2).ToString();
                        }*/
                        else
                        {
                            newValueStr = newValue.ToString();
                        }
                    }
                    
                    //newValueStr += ", Type: " + type.ToString();
                }
                
                string propertyName = uiAutomation.GetPropertyProgrammaticName(propertyId);
                
                ev.Details = "Property: " + propertyName + ", New Value: " + newValueStr;

                mainWindow.Dispatcher.Invoke(new Action(() =>
                    {
                        mainWindow.eventsCollection.Add(ev);
                        mainWindow.eventsListView.ScrollIntoView(ev);
                    }));
            }
        }

        private UIA_AutomationEventHandler UIAeventHandler = null;
        private UIA_AutomationFocusChangedEventHandler UIAFocusChangedEventHandler = null;
        private UIA_AutomationPropertyChangedEventHandler UIAPropChangedEventHandler = null;
        private UIA_AutomationStructureChangedEventHandler UIAStructureChangedEventHandler = null;
        private bool listenerInstalled = false;

        private void InstallListener(IUIAutomationElement element)
        {
			crtProcessId = Process.GetCurrentProcess().Id;

            if ((hasInvokedEvent == true) || (hasMenuOpenedEvent == true) ||
                (hasMenuClosedEvent == true) || (hasElementAddedToSelectionEvent == true) ||
                (hasElementRemovedFromSelectionEvent == true) || (hasElementSelectedEvent == true) ||
                (hasInvalidatedEvent == true) || (hasLayoutInvalidatedEvent == true) ||
                (hasTextChangedEvent == true) || (hasTextSelectionChangedEvent == true) ||
                (hasToolTipClosedEvent == true) || (hasToolTipOpenedEvent == true) ||
                (hasWindowOpenedEvent == true) || (hasAsyncContentLoadedEvent == true) || 
                (hasWindowClosedEvent == true) || (hasInputReachedTargetEvent == true) ||
                (hasInputReachedOtherElementEvent == true) || (hasInputDiscardedEvent == true) ||
                (hasMenuModeStartEvent == true) || (hasMenuModeEndEvent == true))
            {
                UIAeventHandler = new UIA_AutomationEventHandler(this);
            }

            #region install standard event handlers
            if (hasInvokedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Invoke_InvokedEventId,
                             element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasMenuOpenedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_MenuOpenedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasMenuClosedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_MenuClosedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasElementAddedToSelectionEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_SelectionItem_ElementAddedToSelectionEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasElementRemovedFromSelectionEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_SelectionItem_ElementRemovedFromSelectionEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasElementSelectedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_SelectionItem_ElementSelectedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasInvalidatedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Selection_InvalidatedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasLayoutInvalidatedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_LayoutInvalidatedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasTextChangedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Text_TextChangedEventId, 
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasTextSelectionChangedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Text_TextSelectionChangedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasToolTipClosedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_ToolTipClosedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasToolTipOpenedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_ToolTipOpenedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasWindowOpenedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Window_WindowOpenedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasWindowClosedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_Window_WindowClosedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasAsyncContentLoadedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_AsyncContentLoadedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasInputReachedTargetEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_InputReachedTargetEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasInputReachedOtherElementEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_InputReachedOtherElementEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasInputDiscardedEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_InputDiscardedEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasMenuModeStartEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_MenuModeStartEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            if (hasMenuModeEndEvent)
            {
                try
                {
                    uiAutomation.AddAutomationEventHandler(UIA_EventIds.UIA_MenuModeEndEventId,
                        element, eventsScope, null, UIAeventHandler);
                }
                catch { }
            }
            #endregion

            #region install event handlers using specific functions
            if (hasAutomationFocusChangedEvent)
            {
                try
                {
                    UIAFocusChangedEventHandler = new UIA_AutomationFocusChangedEventHandler(this);
                    uiAutomation.AddFocusChangedEventHandler(null, UIAFocusChangedEventHandler);
                }
                catch { }
            }

            if (hasAutomationPropertyChangedEvent)
            {
                try
                {
                    UIAPropChangedEventHandler = new UIA_AutomationPropertyChangedEventHandler(this);
                    uiAutomation.AddPropertyChangedEventHandler(element, eventsScope, null,
                        UIAPropChangedEventHandler, WindowEventsOptions.evtProperties.ToArray());
                }
                catch { }
            }

            if (hasStructureChangedEvent)
            {
                try
                {
                    UIAStructureChangedEventHandler = new UIA_AutomationStructureChangedEventHandler(this);
                    uiAutomation.AddStructureChangedEventHandler(element, eventsScope,
                        null, UIAStructureChangedEventHandler);
                }
                catch { }

            }
            #endregion
			
			if (hasTextEditTextChangedEvent)
            {
                try
                {
					IUIAutomation3 automation3 = uiAutomation as IUIAutomation3;
					if (automation3 != null)
					{
						var uiAutomationTextEditTextChangedEventHandler = new UIAutomationTextEditTextChangedEventHandler(this);
						automation3.AddTextEditTextChangedEventHandler(element, eventsScope, TextEditChangeType, null, 
							uiAutomationTextEditTextChangedEventHandler);
						//MessageBox.Show("TextEditTextChangedEventHandler installed");
					}
                }
                catch { }
            }
			
			if (hasChangesEvent)
            {
                try
                {
					IUIAutomation4 automation4 = uiAutomation as IUIAutomation4;
					if (automation4 != null)
					{
						var uiAutomationChangesEventHandler = new UIAutomationChangesEventHandler(this);
						int changeTypes = UIA_ChangeIds.UIA_SummaryChangeId;
						automation4.AddChangesEventHandler(element, eventsScope, ref changeTypes, ChangesCount, null, 
							uiAutomationChangesEventHandler);
						//MessageBox.Show("ChangesEventHandler installed");
					}
                }
                catch { }
            }
			
			if (hasNotificationEvent)
            {
                try
                {
					IUIAutomation5 automation5 = uiAutomation as IUIAutomation5;
					if (automation5 != null)
					{
						var uiAutomationNotificationEventHandler = new UIAutomationNotificationEventHandler(this);
						automation5.AddNotificationEventHandler(element, eventsScope, null, uiAutomationNotificationEventHandler);
						//MessageBox.Show("NotificationEventHandler installed");
					}
                }
                catch { }
            }
			
			if (hasActiveTextPositionChangedEvent)
            {
                try
                {
					IUIAutomation6 automation6 = uiAutomation as IUIAutomation6;
					if (automation6 != null)
					{
						var uiAutomationActiveTextPositionChangedEventHandler = new UIAutomationActiveTextPositionChangedEventHandler(this);
						automation6.AddActiveTextPositionChangedEventHandler(element, eventsScope, null, uiAutomationActiveTextPositionChangedEventHandler);
						//MessageBox.Show("ActiveTextPositionChangedHandler installed");
					}
                }
                catch { }
            }

            listenerInstalled = true;
        }

        private void UninstallListener()
        {
			listenerInstalled = false;
            UIAeventHandler = null;
            UIAFocusChangedEventHandler = null;
            UIAPropChangedEventHandler = null;
            UIAStructureChangedEventHandler = null;
            try
            {
            	uiAutomation.RemoveAllEventHandlers();
            }
            catch { }
            return;
        }
        
		public static int crtProcessId = 0;

        private static string GetStringFromArray(Array array)
        {
            string s = "";
            foreach (object o in array)
            {
                if (s != "")
                {
                    s += " ";
                }
                s += o.ToString();
            }

            return s;
        }
		
		/*private void MenuOpening(object sender, RoutedEventArgs e)
		{
			if (this.eventsPropButton.IsEnabled == true)
			{
				eventsMenu.Header = "Start Capturing Events";
			}
			else
			{
				eventsMenu.Header = "Stop Capturing Events";
			}
		}*/
    }
	
	public class UIAutomationTextEditTextChangedEventHandler: IUIAutomationTextEditTextChangedEventHandler
	{
		private MainWindow mainWindow = null;
		public UIAutomationTextEditTextChangedEventHandler(MainWindow mainWindow)
		{
			this.mainWindow = mainWindow;
		}
	
		public void HandleTextEditTextChangedEvent(IUIAutomationElement sender, TextEditChangeType TextEditChangeType, string[] eventStrings)
		{
			Event ev = new Event("TextEditTextChangedEvent", sender);
			string eventStringsConcat = null;
			foreach (string eventString in eventStrings)
			{
				if (eventStringsConcat != null)
				{
					eventStringsConcat += ",";
				}
				eventStringsConcat += ("\"" + eventString + "\"");
			}
			ev.Details = "TextEditChangeType: " + TextEditChangeType + ", eventStrings: [" + eventStringsConcat + "]";
		
			mainWindow.Dispatcher.Invoke(new Action(() =>
			{
				//MessageBox.Show("HandleTextEditTextChangedEvent");
				mainWindow.eventsCollection.Add(ev);
				mainWindow.eventsListView.ScrollIntoView(ev);
			}));
		}
	}
	
	public class UIAutomationChangesEventHandler: IUIAutomationChangesEventHandler
	{
		private MainWindow mainWindow = null;
		public UIAutomationChangesEventHandler(MainWindow mainWindow)
		{
			this.mainWindow = mainWindow;
		}
	
		public void HandleChangesEvent(IUIAutomationElement sender, ref UiaChangeInfo uiaChanges, int changesCount)
		{
			Event ev = new Event("ChangesEvent", sender);
			ev.Details = "UiaChangeInfo.uiaId: " + uiaChanges.uiaId + ", UiaChangeInfo.payload: " + uiaChanges.payload +
				", UiaChangeInfo.extraInfo: " + uiaChanges.extraInfo + ", changesCount: " + changesCount;
			
			mainWindow.Dispatcher.Invoke(new Action(() =>
			{
				mainWindow.eventsCollection.Add(ev);
				mainWindow.eventsListView.ScrollIntoView(ev);
			}));
		}
	}
	
	public class UIAutomationNotificationEventHandler: IUIAutomationNotificationEventHandler
	{
		private MainWindow mainWindow = null;
		public UIAutomationNotificationEventHandler(MainWindow mainWindow)
		{
			this.mainWindow = mainWindow;
		}
	
		public void HandleNotificationEvent(IUIAutomationElement sender, NotificationKind NotificationKind, 
			NotificationProcessing NotificationProcessing, string displayString, string activityId)
		{
			Event ev = new Event("NotificationEvent", sender);
			ev.Details = "NotificationKind: " + NotificationKind + ", NotificationProcessing: " + NotificationProcessing +
				", displayString: \"" + displayString + "\", activityId: \"" + activityId + "\"";
		
			mainWindow.Dispatcher.Invoke(new Action(() =>
			{
				mainWindow.eventsCollection.Add(ev);
				mainWindow.eventsListView.ScrollIntoView(ev);
			}));
		}
	}
	
	public class UIAutomationActiveTextPositionChangedEventHandler: IUIAutomationActiveTextPositionChangedEventHandler
	{
		private MainWindow mainWindow = null;
		public UIAutomationActiveTextPositionChangedEventHandler(MainWindow mainWindow)
		{
			this.mainWindow = mainWindow;
		}
	
		public void HandleActiveTextPositionChangedEvent(IUIAutomationElement sender, IUIAutomationTextRange range)
		{
			Event ev = new Event("ActiveTextPositionChangedEvent", sender);
			
			mainWindow.Dispatcher.Invoke(new Action(() =>
			{
				mainWindow.eventsCollection.Add(ev);
				mainWindow.eventsListView.ScrollIntoView(ev);
			}));
		}
	}
}
