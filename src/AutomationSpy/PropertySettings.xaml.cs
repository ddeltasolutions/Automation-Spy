using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace dDeltaSolutions.Spy
{
    /// <summary>
    /// Interaction logic for PropertySettings.xaml
    /// </summary>
    public partial class PropertySettings : Window
    {
        public PropertySettings()
        {
            InitializeComponent();
        }

        public static bool hasAcceleratorKey = true;
        public static bool hasAccessKey = true;
        public static bool hasAriaProperties = true;
        public static bool hasAriaRole = true;
        public static bool hasAutomationId = true;
        public static bool hasBoundingRectangle = true;
        public static bool hasClassName = true;
        public static bool hasClickablePoint = true;
        public static bool hasControllerFor = true;
        public static bool hasControlType = true;
        public static bool hasCulture = true;
        public static bool hasDescribedBy = true;
        public static bool hasFlowsTo = true;
        public static bool hasFrameworkId = true;
        public static bool hasHasKeyboardFocus = true;
        public static bool hasHelpText = true;
        public static bool hasIsContentElement = true;
        public static bool hasIsControlElement = true;
        public static bool hasIsDataValidForForm = true;
        public static bool hasIsEnabled = true;
        public static bool hasIsKeyboardFocusable = true;
        public static bool hasIsOffscreen = true;
        public static bool hasIsPassword = true;
		public static bool hasIsRequiredForForm = true;
		public static bool hasItemStatus = true;
		public static bool hasItemType = true;
		public static bool hasLabeledBy = true;
		public static bool hasLocalizedControlType = true;
		public static bool hasName = true;
		public static bool hasNativeWindowHandle = true;
		public static bool hasOrientation = true;
		public static bool hasProcessId = true;
		public static bool hasProviderDescription = true;
		public static bool hasRuntimeId = true;

        private List<CheckBox> allCheckBoxes = new List<CheckBox>();
        private bool ignoreEvent = false;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            chkAcceleratorKey.IsChecked = hasAcceleratorKey;
            chkAccessKey.IsChecked = hasAccessKey;
            chkAriaProperties.IsChecked = hasAriaProperties;
            chkAriaRole.IsChecked = hasAriaRole;
            chkAutomationId.IsChecked = hasAutomationId;
            chkBoundingRectangle.IsChecked = hasBoundingRectangle;
            chkClassName.IsChecked = hasClassName;
            chkClickablePoint.IsChecked = hasClickablePoint;
            chkControllerFor.IsChecked = hasControllerFor;
            chkControlType.IsChecked = hasControlType;
            chkCulture.IsChecked = hasCulture;
            chkDescribedBy.IsChecked = hasDescribedBy;
            chkFlowsTo.IsChecked = hasFlowsTo;
            chkFrameworkId.IsChecked = hasFrameworkId;
            chkHasKeyboardFocus.IsChecked = hasHasKeyboardFocus;
            chkHelpText.IsChecked = hasHelpText;
            chkIsContentElement.IsChecked = hasIsContentElement;
            chkIsControlElement.IsChecked = hasIsControlElement;
            chkIsDataValidForForm.IsChecked = hasIsDataValidForForm;
            chkIsEnabled.IsChecked = hasIsEnabled;
            chkIsKeyboardFocusable.IsChecked = hasIsKeyboardFocusable;
            chkIsOffscreen.IsChecked = hasIsOffscreen;
            chkIsPassword.IsChecked = hasIsPassword;
			chkIsRequiredForForm.IsChecked = hasIsRequiredForForm;
			chkItemStatus.IsChecked = hasItemStatus;
			chkItemType.IsChecked = hasItemType;
			chkLabeledBy.IsChecked = hasLabeledBy;
			chkLocalizedControlType.IsChecked = hasLocalizedControlType;
			chkName.IsChecked = hasName;
			chkNativeWindowHandle.IsChecked = hasNativeWindowHandle;
			chkOrientation.IsChecked = hasOrientation;
			chkProcessId.IsChecked = hasProcessId;
			chkProviderDescription.IsChecked = hasProviderDescription;
			chkRuntimeId.IsChecked = hasRuntimeId;

            allCheckBoxes.Add(chkAcceleratorKey);
            allCheckBoxes.Add(chkAccessKey);
            allCheckBoxes.Add(chkAriaProperties);
            allCheckBoxes.Add(chkAriaRole);
            allCheckBoxes.Add(chkAutomationId);
            allCheckBoxes.Add(chkBoundingRectangle);
            allCheckBoxes.Add(chkClassName);
            allCheckBoxes.Add(chkClickablePoint);
            allCheckBoxes.Add(chkControllerFor);
            allCheckBoxes.Add(chkControlType);
            allCheckBoxes.Add(chkCulture);
            allCheckBoxes.Add(chkDescribedBy);
            allCheckBoxes.Add(chkFlowsTo);
            allCheckBoxes.Add(chkFrameworkId);
            allCheckBoxes.Add(chkHasKeyboardFocus);
            allCheckBoxes.Add(chkHelpText);
            allCheckBoxes.Add(chkIsContentElement);
            allCheckBoxes.Add(chkIsControlElement);
            allCheckBoxes.Add(chkIsDataValidForForm);
            allCheckBoxes.Add(chkIsEnabled);
            allCheckBoxes.Add(chkIsKeyboardFocusable);
            allCheckBoxes.Add(chkIsOffscreen);
            allCheckBoxes.Add(chkIsPassword);
			allCheckBoxes.Add(chkIsRequiredForForm);
			allCheckBoxes.Add(chkItemStatus);
			allCheckBoxes.Add(chkItemType);
			allCheckBoxes.Add(chkLabeledBy);
			allCheckBoxes.Add(chkLocalizedControlType);
			allCheckBoxes.Add(chkName);
			allCheckBoxes.Add(chkNativeWindowHandle);
			allCheckBoxes.Add(chkOrientation);
			allCheckBoxes.Add(chkProcessId);
			allCheckBoxes.Add(chkProviderDescription);
			allCheckBoxes.Add(chkRuntimeId);

            TestCheckAll();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            hasAcceleratorKey = chkAcceleratorKey.IsChecked.Value;
            hasAccessKey = chkAccessKey.IsChecked.Value;
            hasAriaProperties = chkAriaProperties.IsChecked.Value;
            hasAriaRole = chkAriaRole.IsChecked.Value;
            hasAutomationId = chkAutomationId.IsChecked.Value;
            hasBoundingRectangle = chkBoundingRectangle.IsChecked.Value;
            hasClassName = chkClassName.IsChecked.Value;
            hasClickablePoint = chkClickablePoint.IsChecked.Value;
            hasControllerFor = chkControllerFor.IsChecked.Value;
            hasControlType = chkControlType.IsChecked.Value;
            hasCulture = chkCulture.IsChecked.Value;
            hasDescribedBy = chkDescribedBy.IsChecked.Value;
            hasFlowsTo = chkFlowsTo.IsChecked.Value;
            hasFrameworkId = chkFrameworkId.IsChecked.Value;
            hasHasKeyboardFocus = chkHasKeyboardFocus.IsChecked.Value;
            hasHelpText = chkHelpText.IsChecked.Value;
            hasIsContentElement = chkIsContentElement.IsChecked.Value;
            hasIsControlElement = chkIsControlElement.IsChecked.Value;
            hasIsDataValidForForm = chkIsDataValidForForm.IsChecked.Value;
            hasIsEnabled = chkIsEnabled.IsChecked.Value;
            hasIsKeyboardFocusable = chkIsKeyboardFocusable.IsChecked.Value;
            hasIsOffscreen = chkIsOffscreen.IsChecked.Value;
            hasIsPassword = chkIsPassword.IsChecked.Value;
			hasIsRequiredForForm = chkIsRequiredForForm.IsChecked.Value;
			hasItemStatus = chkItemStatus.IsChecked.Value;
			hasItemType = chkItemType.IsChecked.Value;
			hasLabeledBy = chkLabeledBy.IsChecked.Value;
			hasLocalizedControlType = chkLocalizedControlType.IsChecked.Value;
			hasName = chkName.IsChecked.Value;
			hasNativeWindowHandle = chkNativeWindowHandle.IsChecked.Value;
			hasOrientation = chkOrientation.IsChecked.Value;
			hasProcessId = chkProcessId.IsChecked.Value;
			hasProviderDescription = chkProviderDescription.IsChecked.Value;
			hasRuntimeId = chkRuntimeId.IsChecked.Value;
            
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
    }
}
