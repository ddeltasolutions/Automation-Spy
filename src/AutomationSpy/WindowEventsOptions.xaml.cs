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
using Interop.UIAutomationClient;

namespace dDeltaSolutions.Spy
{
    /// <summary>
    /// Interaction logic for WindowEventsOptions.xaml
    /// </summary>
    public partial class WindowEventsOptions : Window
    {
        public WindowEventsOptions()
        {
            InitializeComponent();
        }

        public bool hasAsyncContentLoadedEvent = true;
        public bool hasElementAddedToSelectionEvent = true;
        public bool hasElementRemovedFromSelectionEvent = true;
        public bool hasElementSelectedEvent = true;
        public bool hasInvalidatedEvent = true;
        public bool hasInvokedEvent = true;
        public bool hasInputReachedTargetEvent = true;
        public bool hasInputReachedOtherElementEvent = true;
        public bool hasInputDiscardedEvent = true;
        public bool hasLayoutInvalidatedEvent = true;
        public bool hasMenuOpenedEvent = true;
        public bool hasMenuClosedEvent = true;
        public bool hasMenuModeStartEvent = true;
        public bool hasMenuModeEndEvent = true;
        public bool hasTextChangedEvent = true;
        public bool hasTextSelectionChangedEvent = true;
        public bool hasToolTipClosedEvent = true;
        public bool hasToolTipOpenedEvent = true;
        public bool hasWindowOpenedEvent = true;
        public bool hasAutomationFocusChangedEvent = true;
        public bool hasAutomationPropertyChangedEvent = true;
        public bool hasStructureChangedEvent = true;
        public bool hasWindowClosedEvent = true;
		
		public bool hasTextEditTextChangedEvent = false;
		public bool hasChangesEvent = false;
		public bool hasNotificationEvent = false;
		public bool hasActiveTextPositionChangedEvent = false;

        public TreeScope eventsScope = TreeScope.TreeScope_Subtree;
        private List<CheckBox> allCheckBoxes = new List<CheckBox>();
        private bool ignoreEvent = false;
        internal static List<int> evtProperties = new List<int>();
		
		public TextEditChangeType TextEditChangeType = TextEditChangeType.TextEditChangeType_None;
		public int ChangesCount = 1;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            chkAsyncContentLoadedEv.IsChecked = hasAsyncContentLoadedEvent;
            chkElementAddedToSelectionEv.IsChecked = hasElementAddedToSelectionEvent;
            chkElementRemovedFromSelectionEv.IsChecked = hasElementRemovedFromSelectionEvent;
            chkElementSelectedEv.IsChecked = hasElementSelectedEvent;
            chkInvalidatedEv.IsChecked = hasInvalidatedEvent;
            chkInvokedEvent.IsChecked = hasInvokedEvent;
            chkInputReachedTargetEvent.IsChecked = hasInputReachedTargetEvent;
            chkInputReachedOtherElementEvent.IsChecked = hasInputReachedOtherElementEvent;
            chkInputDiscardedEvent.IsChecked = hasInputDiscardedEvent;
            chkLayoutInvalidatedEv.IsChecked = hasLayoutInvalidatedEvent;
            chkMenuOpenedEv.IsChecked = hasMenuOpenedEvent;
            chkMenuClosedEv.IsChecked = hasMenuClosedEvent;
            chkMenuModeStartEv.IsChecked = hasMenuModeStartEvent;
            chkMenuModeEndEv.IsChecked = hasMenuModeEndEvent;
            chkTextChangedEv.IsChecked = hasTextChangedEvent;
            chkTextSelectionChangedEv.IsChecked = hasTextSelectionChangedEvent;
            chkToolTipClosedEv.IsChecked = hasToolTipClosedEvent;
            chkToolTipOpenedEv.IsChecked = hasToolTipOpenedEvent;
            chkWindowOpenedEv.IsChecked = hasWindowOpenedEvent;
            chkAutomationFocusChangedEv.IsChecked = hasAutomationFocusChangedEvent;
            chkAutomationPropertyChangedEv.IsChecked = hasAutomationPropertyChangedEvent;
            chkStructureChangedEv.IsChecked = hasStructureChangedEvent;
            chkWindowClosedEv.IsChecked = hasWindowClosedEvent;
			
			chkTextEditTextChangedEv.IsChecked = hasTextEditTextChangedEvent;
			chkChangesEv.IsChecked = hasChangesEvent;
			chkNotificationEv.IsChecked = hasNotificationEvent;
			chkActiveTextPositionChangedEv.IsChecked = hasActiveTextPositionChangedEvent;
            
            cmbScope.SelectedItem = eventsScope;

            allCheckBoxes.Add(chkAsyncContentLoadedEv);
            allCheckBoxes.Add(chkElementAddedToSelectionEv);
            allCheckBoxes.Add(chkElementRemovedFromSelectionEv);
            allCheckBoxes.Add(chkElementSelectedEv);
            allCheckBoxes.Add(chkInvalidatedEv);
            allCheckBoxes.Add(chkInvokedEvent);
            allCheckBoxes.Add(chkInputReachedTargetEvent);
            allCheckBoxes.Add(chkInputReachedOtherElementEvent);
            allCheckBoxes.Add(chkInputDiscardedEvent);
            allCheckBoxes.Add(chkLayoutInvalidatedEv);
            allCheckBoxes.Add(chkMenuOpenedEv);
            allCheckBoxes.Add(chkMenuClosedEv);
            allCheckBoxes.Add(chkMenuModeStartEv);
            allCheckBoxes.Add(chkMenuModeEndEv);
            allCheckBoxes.Add(chkTextChangedEv);
            allCheckBoxes.Add(chkTextSelectionChangedEv);
            allCheckBoxes.Add(chkToolTipClosedEv);
            allCheckBoxes.Add(chkToolTipOpenedEv);
            allCheckBoxes.Add(chkWindowOpenedEv);
            allCheckBoxes.Add(chkAutomationFocusChangedEv);
            allCheckBoxes.Add(chkAutomationPropertyChangedEv);
            allCheckBoxes.Add(chkStructureChangedEv);
            allCheckBoxes.Add(chkWindowClosedEv);
			
			//allCheckBoxes.Add(chkTextEditTextChangedEv);
			//allCheckBoxes.Add(chkChangesEv);
			//allCheckBoxes.Add(chkNotificationEv);
			//allCheckBoxes.Add(chkActiveTextPositionChangedEv);
			
			if (TextEditChangeType == TextEditChangeType.TextEditChangeType_AutoCorrect)
			{
				cmbChangeType.SelectedIndex = 1;
			}
			else if (TextEditChangeType == TextEditChangeType.TextEditChangeType_Composition)
			{
				cmbChangeType.SelectedIndex = 2;
			}
			else if (TextEditChangeType == TextEditChangeType.TextEditChangeType_CompositionFinalized)
			{
				cmbChangeType.SelectedIndex = 3;
			}
			else if (TextEditChangeType == TextEditChangeType.TextEditChangeType_AutoComplete)
			{
				cmbChangeType.SelectedIndex = 4;
			}
			txtChangesCount.Text = ChangesCount.ToString();

            TestCheckAll();
            int prop = 0;
            ListBoxItem item = null;
            
            prop = UIA_PropertyIds.UIA_AcceleratorKeyPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_AccessKeyPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_AriaRolePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_AriaPropertiesPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_AutomationIdPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_BoundingRectanglePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ClassNamePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ClickablePointPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ControllerForPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ControlTypePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_CulturePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_DescribedByPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_FlowsToPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_FrameworkIdPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_HasKeyboardFocusPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_HelpTextPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsContentElementPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsControlElementPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsDataValidForFormPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsEnabledPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsKeyboardFocusablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsOffscreenPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsPasswordPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsRequiredForFormPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ItemStatusPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ItemTypePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LabeledByPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LocalizedControlTypePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_NamePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_NativeWindowHandlePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_OrientationPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ProcessIdPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ProviderDescriptionPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_RuntimeIdPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsDockPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsExpandCollapsePatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsGridPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsGridItemPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsInvokePatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsMultipleViewPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsRangeValuePatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsScrollItemPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsSelectionItemPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsSelectionPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsTablePatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsTableItemPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsTextPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsTogglePatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsTransformPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsValuePatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsWindowPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsLegacyIAccessiblePatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsItemContainerPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsVirtualizedItemPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_IsSynchronizedInputPatternAvailablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_DockDockPositionPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ExpandCollapseExpandCollapseStatePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_GridColumnCountPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_GridRowCountPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_GridItemColumnPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_GridItemColumnSpanPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_GridItemContainingGridPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_GridItemRowPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_GridItemRowSpanPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_MultipleViewCurrentViewPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_MultipleViewSupportedViewsPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_RangeValueIsReadOnlyPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_RangeValueLargeChangePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_RangeValueMaximumPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_RangeValueMinimumPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_RangeValueSmallChangePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_RangeValueValuePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ScrollHorizontallyScrollablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ScrollHorizontalScrollPercentPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ScrollHorizontalViewSizePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ScrollVerticallyScrollablePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ScrollVerticalScrollPercentPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ScrollVerticalViewSizePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_SelectionCanSelectMultiplePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_SelectionIsSelectionRequiredPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_SelectionSelectionPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_SelectionItemIsSelectedPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_SelectionItemSelectionContainerPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TableColumnHeadersPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TableRowHeadersPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TableRowOrColumnMajorPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TableItemColumnHeaderItemsPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TableItemRowHeaderItemsPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ToggleToggleStatePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TransformCanMovePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TransformCanResizePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_TransformCanRotatePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ValueIsReadOnlyPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_ValueValuePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_WindowCanMaximizePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_WindowCanMinimizePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_WindowIsModalPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_WindowIsTopmostPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_WindowWindowInteractionStatePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_WindowWindowVisualStatePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleChildIdPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleNamePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleValuePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleDescriptionPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleRolePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleStatePropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleHelpPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleKeyboardShortcutPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleSelectionPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            prop = UIA_PropertyIds.UIA_LegacyIAccessibleDefaultActionPropertyId;
            item = new ListBoxItem();
            item.Content = PropertyToString(prop);
            item.Tag = prop;
            propList.Items.Add(item);
            
            try
            {
                ignoreCheck = true;
                foreach (ListBoxItem crtItem in propList.Items)
                {
                    int crtProperty = (int)crtItem.Tag;
                    if (evtProperties.Contains(crtProperty))
                    {
                        crtItem.IsSelected = true;
                    }
                }
                ignoreCheck = false;
                
                CheckSelectedItems();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private bool ignoreCheck = false;

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            hasAsyncContentLoadedEvent = chkAsyncContentLoadedEv.IsChecked.Value;
            hasElementAddedToSelectionEvent = chkElementAddedToSelectionEv.IsChecked.Value;
            hasElementRemovedFromSelectionEvent = chkElementRemovedFromSelectionEv.IsChecked.Value;
            hasElementSelectedEvent = chkElementSelectedEv.IsChecked.Value;
            hasInvalidatedEvent = chkInvalidatedEv.IsChecked.Value;
            hasInvokedEvent = chkInvokedEvent.IsChecked.Value;
            hasInputReachedTargetEvent = chkInputReachedTargetEvent.IsChecked.Value;
            hasInputReachedOtherElementEvent = chkInputReachedOtherElementEvent.IsChecked.Value;
            hasInputDiscardedEvent = chkInputDiscardedEvent.IsChecked.Value;
            hasLayoutInvalidatedEvent = chkLayoutInvalidatedEv.IsChecked.Value;
            hasMenuOpenedEvent = chkMenuOpenedEv.IsChecked.Value;
            hasMenuClosedEvent = chkMenuClosedEv.IsChecked.Value;
            hasMenuModeStartEvent = chkMenuModeStartEv.IsChecked.Value;
            hasMenuModeEndEvent = chkMenuModeEndEv.IsChecked.Value;
            hasTextChangedEvent = chkTextChangedEv.IsChecked.Value;
            hasTextSelectionChangedEvent = chkTextSelectionChangedEv.IsChecked.Value;
            hasToolTipClosedEvent = chkToolTipClosedEv.IsChecked.Value;
            hasToolTipOpenedEvent = chkToolTipOpenedEv.IsChecked.Value;
            hasWindowOpenedEvent = chkWindowOpenedEv.IsChecked.Value;
            hasAutomationFocusChangedEvent = chkAutomationFocusChangedEv.IsChecked.Value;
            hasAutomationPropertyChangedEvent = chkAutomationPropertyChangedEv.IsChecked.Value;
            hasStructureChangedEvent = chkStructureChangedEv.IsChecked.Value;
            hasWindowClosedEvent = chkWindowClosedEv.IsChecked.Value;
			
			hasTextEditTextChangedEvent = chkTextEditTextChangedEv.IsChecked.Value;
			hasChangesEvent = chkChangesEv.IsChecked.Value;
			hasNotificationEvent = chkNotificationEv.IsChecked.Value;
			hasActiveTextPositionChangedEvent = chkActiveTextPositionChangedEv.IsChecked.Value;

            eventsScope = (TreeScope)cmbScope.SelectedItem;
            
            evtProperties.Clear();
            foreach (ListBoxItem crtItem in propList.Items)
            {
                if (crtItem.IsSelected == false)
                {
                    continue;
                }
                int crtProperty = (int)crtItem.Tag;
                evtProperties.Add(crtProperty);
            }
			
			int selectedIndex = cmbChangeType.SelectedIndex;
			if (selectedIndex == 0)
			{
				TextEditChangeType = TextEditChangeType.TextEditChangeType_None;
			}
			else if (selectedIndex == 1)
			{
				TextEditChangeType = TextEditChangeType.TextEditChangeType_AutoCorrect;
			}
			else if (selectedIndex == 2)
			{
				TextEditChangeType = TextEditChangeType.TextEditChangeType_Composition;
			}
			else if (selectedIndex == 3)
			{
				TextEditChangeType = TextEditChangeType.TextEditChangeType_CompositionFinalized;
			}
			else if (selectedIndex == 4)
			{
				TextEditChangeType = TextEditChangeType.TextEditChangeType_AutoComplete;
			}
			if (int.TryParse(txtChangesCount.Text, out ChangesCount) == false)
			{
				MessageBox.Show("Changes Count has to be an integer number");
				return;
			}
			if (ChangesCount <= 0)
			{
				MessageBox.Show("Changes Count has to be positive and non-zero");
				return;
			}

            this.DialogResult = true;
        }

        private void chkCheckAll_Checked(object sender, RoutedEventArgs e)
        {
            if (ignoreEvent)
            {
                return;
            }

            foreach (CheckBox checkBox in allCheckBoxes)
            {
                ignoreEvent = true;
                checkBox.IsChecked = chkCheckAll.IsChecked;
                ignoreEvent = false;
            }
        }

        private void CheckBoxChecked(object sender, RoutedEventArgs e)
        {
			if (sender == chkTextEditTextChangedEv && cmbChangeType != null)
			{
				if (chkTextEditTextChangedEv.IsChecked == true)
				{
					txbLabel1.Foreground = Brushes.Black;
					cmbChangeType.Foreground = Brushes.Black;
					cmbChangeType.IsEnabled = true;
				}
				else
				{
					txbLabel1.Foreground = Brushes.Gray;
					cmbChangeType.Foreground = Brushes.Gray;
					cmbChangeType.IsEnabled = false;
				}
			}
			else if (sender == chkChangesEv && txtChangesCount != null)
			{
				if (chkChangesEv.IsChecked == true)
				{
					txbLabel2.Foreground = Brushes.Black;
					txtChangesCount.IsEnabled = true;
				}
				else
				{
					txbLabel2.Foreground = Brushes.Gray;
					txtChangesCount.IsEnabled = false;
				}
			}
		
            TestCheckAll();
        }

        private void TestCheckAll()
        {
            if (ignoreEvent)
            {
                return;
            }

            if (chkCheckAll == null)
            {
                return;
            }

            bool allChecked = true;
			bool checkedAtLeastOne = false;

            foreach (CheckBox checkBox in allCheckBoxes)
            {
                if (checkBox.IsChecked.HasValue == false)
                {
                    continue;
                }

                if (checkBox.IsChecked == false)
                {
                    allChecked = false;
					if (checkedAtLeastOne == true)
					{
						break;
					}
                }
				else
				{
					checkedAtLeastOne = true;
				}
            }

            if (allChecked == true)
            {
                ignoreEvent = true;
                chkCheckAll.IsChecked = true;
                ignoreEvent = false;
            }
			else if (checkedAtLeastOne == true)
			{
				ignoreEvent = true;
                chkCheckAll.IsChecked = null;
                ignoreEvent = false;
			}
            else
            {
                ignoreEvent = true;
                chkCheckAll.IsChecked = false;
                ignoreEvent = false;
            }
        }

        private void chkCheckAll_Unchecked(object sender, RoutedEventArgs e)
        {
            chkCheckAll_Checked(sender, e);
        }

        private void CheckBoxUnchecked(object sender, RoutedEventArgs e)
        {
            CheckBoxChecked(sender, e);
        }
        
        double offset = 140;

        private void OnMore(object sender, RoutedEventArgs e)
        {
            this.Width += (2*offset);
            if (this.Left - offset < 0)
            {
                this.Left = 0;
            }
            else
            {
                this.Left -= offset;
            }
            btnMore.Visibility = Visibility.Hidden;
        }
        
        private void OnLess(object sender, RoutedEventArgs e)
        {
            this.Width -= (2*offset);
            this.Left += offset;
            btnMore.Visibility = Visibility.Visible;
        }
        
        private void chkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            if (ignoreCheck)
            {
                return;
            }
            
            try
            {
                ignoreCheck = true;
                foreach (ListBoxItem crtItem in propList.Items)
                {
                    crtItem.IsSelected = true;
                }
                ignoreCheck = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void chkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            if (ignoreCheck)
            {
                return;
            }
            
            try
            {
                ignoreCheck = true;
                foreach (ListBoxItem crtItem in propList.Items)
                {
                    crtItem.IsSelected = false;
                }
                ignoreCheck = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        void OnSelChanged(object sender, SelectionChangedEventArgs args)
        {
            if (ignoreCheck)
            {
                return;
            }
            
            CheckSelectedItems();
        }
        
        private void CheckSelectedItems()
        {
            bool allSelected = true;
			bool selectedAtLeastOne = false;
			
            foreach (ListBoxItem crtItem in propList.Items)
            {
                if (crtItem.IsSelected == false)
                {
                    allSelected = false;
					if (selectedAtLeastOne == true)
					{
						break;
					}
                }
				else
				{
					selectedAtLeastOne = true;
				}
            }
            
            ignoreCheck = true;
            if (allSelected)
            {
                chkSelectAll.IsChecked = true;
            }
			else if (selectedAtLeastOne == true)
			{
				chkSelectAll.IsChecked = null;
			}
            else
            {
                chkSelectAll.IsChecked = false;
            }
            ignoreCheck = false;
        }
        
        public static string PropertyToString(int property)
        {
            string progname = dDeltaSolutions.Spy.MainWindow.uiAutomation.GetPropertyProgrammaticName(property);
            //progname.Replace("Identifiers", "");
            //progname.Replace("Property", "");
            return progname;
        }
    }
}
