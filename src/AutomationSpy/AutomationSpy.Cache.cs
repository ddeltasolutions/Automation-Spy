using System;
using System.Windows;
using System.Diagnostics;
using Interop.UIAutomationClient;

namespace dDeltaSolutions.Spy
{
    public partial class MainWindow
    {
		private void GetCachedInformation(TreeNode node)
		{
			if (PropertySettings.hasAcceleratorKey == true)
			{
				try
				{
					string acceleratorKey = "\"" + node.Element.CachedAcceleratorKey + "\"";
					Attribute attribute = new Attribute("Accelerator Key:", acceleratorKey);
					attribute.Tooltip = "String containing the accelerator key combinations for the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
			if (PropertySettings.hasAccessKey == true)
			{
				try
				{
					string accessKey = "\"" + node.Element.CachedAccessKey + "\"";
					Attribute attribute = new Attribute("Access Key:", accessKey);
					attribute.Tooltip = "String containing the access key character for the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasAriaProperties == true)
			{
				try
				{
					string ariaProperties = "\"" + node.Element.CachedAriaProperties + "\"";
					Attribute attribute = new Attribute("Aria Properties:", ariaProperties);
					attribute.Tooltip = "Accessible Rich Internet Applications (ARIA) properties of the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasAriaRole == true)
			{
				try
				{
					string ariaRole = "\"" + node.Element.CachedAriaRole + "\"";
					Attribute attribute = new Attribute("Aria Role:", ariaRole);
					attribute.Tooltip = "Accessible Rich Internet Applications (ARIA) role of the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
            //get automation id
			if (PropertySettings.hasAutomationId == true)
			{
				try
				{
					string automationIdString = "\"" + node.Element.CachedAutomationId + "\"";
					Attribute attribute = new Attribute("AutomationId:", automationIdString);
					attribute.Tooltip = "String containing the UI Automation identifier (ID) for the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
            //get bounding rectangle
			if (PropertySettings.hasBoundingRectangle == true)
			{
				try
				{
					tagRECT boundingRectangle = node.Element.CachedBoundingRectangle;
					string rectangleString = RectangleToString(boundingRectangle);
					Attribute attribute = new Attribute("Bounding rectangle:", rectangleString);
					attribute.Tooltip = "Coordinates of the rectangle that completely encloses the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
            //get window class
			if (PropertySettings.hasClassName == true)
			{
				try
				{
					string windowClass = "\"" + node.Element.CachedClassName + "\"";
					Attribute attribute = new Attribute("Class Name:", windowClass);
					attribute.Tooltip = "String containing the class name of the element as assigned by the control developer";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
			// clickable point
			if (PropertySettings.hasClickablePoint == true)
			{
				try
				{
					object clickablePointObj = node.Element.GetCachedPropertyValue(UIA_PropertyIds.UIA_ClickablePointPropertyId);
					if (clickablePointObj != null)
					{
						tagPOINT clickablePoint = (tagPOINT)clickablePointObj;
						string pointString = GetStringFromPoint(clickablePoint);
						Attribute attribute = new Attribute("ClickablePoint:", pointString);
						attribute.Tooltip = "Point on the Automation Element that could be clicked";
						this.listAttributes.Add(attribute);
					}
				}
				catch (Exception ex)
				{ }
			}
            
			if (PropertySettings.hasControllerFor == true)
			{
				try
				{
					IUIAutomationElementArray array = node.Element.CachedControllerFor;
					for (int i = 0; i < array.Length; i++)
					{
						IUIAutomationElement element = array.GetElement(i);
						TreeNode elementNode = new TreeNode(element);
						string controllerForString = elementNode.ToString();
						Attribute attribute = new Attribute("ControllerFor[" + i + "]:", controllerForString);
						attribute.Tooltip = "Array of elements for which this element served as the controller";
						attribute.UnderneathElement = element;
						this.listAttributes.Add(attribute);
					}
				}
				catch {}
			}
            
            //get Control Type
			if (PropertySettings.hasControlType == true)
			{
				try
				{
					int controlTypeId = node.Element.CachedControlType;
					string controlTypeString = Helper.ControlTypeIdToString(controlTypeId) + " (" + controlTypeId + ")";
					Attribute attribute = new Attribute("Control type:", controlTypeString);
					attribute.Tooltip = "Control type of the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
			if (PropertySettings.hasCulture == true)
			{
				try
				{
					int culture = node.Element.CachedCulture;
					Attribute attribute = new Attribute("Culture:", culture.ToString());
					attribute.Tooltip = "Culture identifier";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasDescribedBy == true)
			{
				try
				{
					IUIAutomationElementArray array = node.Element.CachedDescribedBy;
					for (int i = 0; i < array.Length; i++)
					{
						IUIAutomationElement element = array.GetElement(i);
						TreeNode elementNode = new TreeNode(element);
						string describedByString = elementNode.ToString();
						Attribute attribute = new Attribute("DescribedBy[" + i + "]:", describedByString);
						attribute.Tooltip = "Array of elements that describe this element";
						attribute.UnderneathElement = element;
						this.listAttributes.Add(attribute);
					}
				}
				catch {}
			}
            
			if (PropertySettings.hasFlowsTo == true)
			{
				try
				{
					IUIAutomationElementArray array = node.Element.CachedFlowsTo;
					for (int i = 0; i < array.Length; i++)
					{
						IUIAutomationElement element = array.GetElement(i);
						TreeNode elementNode = new TreeNode(element);
						string flowsToString = elementNode.ToString();
						Attribute attribute = new Attribute("FlowsTo[" + i + "]:", flowsToString);
						attribute.Tooltip = "Array of elements that indicates the reading order after the current element";
						attribute.UnderneathElement = element;
						this.listAttributes.Add(attribute);
					}
				}
				catch {}
			}
			
			if (PropertySettings.hasFrameworkId == true)
			{
				try
				{
					string frameworkIdString = node.Element.CachedFrameworkId;
					Attribute attribute = new Attribute("FrameworkId:", "\"" + frameworkIdString + "\"");
					attribute.Tooltip = "Name of the underlying UI framework";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasHasKeyboardFocus == true)
			{
				try
				{
					int hasKeyboboardFocus = node.Element.CachedHasKeyboardFocus;
					Attribute attribute = new Attribute("HasKeyboardFocus:", hasKeyboboardFocus.ToString());
					attribute.Tooltip = "Value that indicates whether the element had keyboard focus";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasHelpText == true)
			{
				try
				{
					string helpText = "\"" + node.Element.CachedHelpText + "\"";
					Attribute attribute = new Attribute("Help Text:", helpText);
					attribute.Tooltip = "Help text associated with the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasIsContentElement == true)
			{
				try
				{
					string isContentString = node.Element.CachedIsContentElement.ToString();
					Attribute attribute = new Attribute("IsContentElement:", isContentString);
					attribute.Tooltip = "Value that specifies whether the element is a content element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasIsControlElement == true)
			{
				try
				{
					string isControlString = node.Element.CachedIsControlElement.ToString();
					Attribute attribute = new Attribute("IsControlElement:", isControlString);
					attribute.Tooltip = "Value that indicates whether the element is viewed as a control";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasIsDataValidForForm == true)
			{
				try
				{
					string isDataValidForFormString = node.Element.CachedIsDataValidForForm.ToString();
					Attribute attribute = new Attribute("IsDataValidForForm:", isDataValidForFormString);
					attribute.Tooltip = "Indicates whether the element contains valid data for a form";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasIsEnabled == true)
			{
				try
				{
					string isEnabledString = node.Element.CachedIsEnabled.ToString();
					Attribute attribute = new Attribute("IsEnabled:", isEnabledString);
					attribute.Tooltip = "Value that indicates whether the element is enabled";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
			if (PropertySettings.hasIsKeyboardFocusable == true)
			{
				try
				{
					string isKeyboardFocusableString = node.Element.CachedIsKeyboardFocusable.ToString();
					Attribute attribute = new Attribute("IsKeyboardFocusable:", isKeyboardFocusableString);
					attribute.Tooltip = "Value that indicates whether the UI Automation element can accept keyboard focus";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasIsOffscreen == true)
			{
				try
				{
					string isOffscreenString = node.Element.CachedIsOffscreen.ToString();
					Attribute attribute = new Attribute("IsOffscreen:", isOffscreenString);
					attribute.Tooltip = "Value that indicates whether the UI Automation element is visible on the screen";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasIsPassword == true)
			{
				try
				{
					string isPasswordString = node.Element.CachedIsPassword.ToString();
					Attribute attribute = new Attribute("IsPassword:", isPasswordString);
					attribute.Tooltip = "Value that indicates whether the UI Automation element contains protected content";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasIsRequiredForForm == true)
			{
				try
				{
					string isRequiredForFormString = node.Element.CachedIsRequiredForForm.ToString();
					Attribute attribute = new Attribute("IsRequiredForForm:", isRequiredForFormString);
					attribute.Tooltip = "Indicates whether the UI Automation element is required to be filled out on a form";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
			if (PropertySettings.hasItemStatus == true)
			{
				try
				{
					string itemStatusString = "\"" + node.Element.CachedItemStatus + "\"";
					Attribute attribute = new Attribute("ItemStatus:", itemStatusString);
					attribute.Tooltip = "Description of the status of an item within an element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasItemType == true)
			{
				try
				{
					string itemTypeString = "\"" + node.Element.CachedItemType + "\"";
					Attribute attribute = new Attribute("ItemType:", itemTypeString);
					attribute.Tooltip = "Description of the type of an item";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasLabeledBy == true)
			{
				try
				{
					IUIAutomationElement labeledBy = node.Element.CachedLabeledBy;
					if (labeledBy != null)
					{
						TreeNode nodeLabeledBy = new TreeNode(labeledBy);
						string labeledByString = nodeLabeledBy.ToString();
						Attribute attribute = new Attribute("Labeled by:", labeledByString);
						attribute.Tooltip = "Element that contains the text label for this element";
						attribute.UnderneathElement = labeledBy;
						this.listAttributes.Add(attribute);
					}
				}
				catch { }
			}
            
            //get control type
			if (PropertySettings.hasLocalizedControlType == true)
			{
				try
				{
					string localizedType = "\"" + node.Element.CachedLocalizedControlType + "\"";
					Attribute attribute = new Attribute("Localized control type:", localizedType);
					attribute.Tooltip = "Gets a description of the control type";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
            //get name
			if (PropertySettings.hasName == true)
			{
				try
				{
					string name = node.Element.CachedName;

					if (name != null && name.Length > maxNameLength)
					{
						name = name.Substring(0, maxNameLength);
						name += "...";
					}

					name = "\"" + name + "\"";
					Attribute attribute = new Attribute("Name:", name);
					attribute.Tooltip = "Name of the element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
			//get window handle
			if (PropertySettings.hasNativeWindowHandle == true)
			{
				try
				{
					string windowHandleString = node.Element.CachedNativeWindowHandle.ToString("X");
					Attribute attribute = new Attribute("Native Window handle:", "0x" + windowHandleString);
					attribute.Tooltip = "Handle of the element's window";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasOrientation == true)
			{
				try
				{
					string orientationString = node.Element.CachedOrientation.ToString();
					Attribute attribute = new Attribute("Orientation:", orientationString);
					attribute.Tooltip = "Orientation of the control";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
			//get process id
			if (PropertySettings.hasProcessId == true)
			{
				try
				{
					int pid = node.Element.CachedProcessId;
					Process process = null;
					try
					{
						process = Process.GetProcessById(pid);
					}
					catch { }
					
					string processString = "";
					string attrName = "";
					if (process != null)
					{
						processString = process.ProcessName + ".exe (PID: " + pid.ToString() + ")";
						attrName = "Process:";
					}
					else
					{
						processString = pid.ToString() + " - not running";
						attrName = "ProcessId:";
					}

					Attribute attribute = new Attribute(attrName, processString);
					attribute.Tooltip = "Process identifier (ID) of this element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
            
			if (PropertySettings.hasProviderDescription == true)
			{
				try
				{
					string providerDescription = node.Element.CachedProviderDescription;
					Attribute attribute = new Attribute("Provider Description:", providerDescription);
					attribute.Tooltip = "Description of the provider for this element";
					this.listAttributes.Add(attribute);
				}
				catch { }
			}
			
            //get runtime id
			if (PropertySettings.hasRuntimeId == true)
			{
				try
				{
					Array runtimeId = node.Element.GetCachedPropertyValue(UIA_PropertyIds.UIA_RuntimeIdPropertyId) as Array;
					if (runtimeId != null)
					{
						string[] runtimeIdStrings = ArrayToStringArray(runtimeId);
						string runtimeIdString = "[" + string.Join(",", runtimeIdStrings) + "]";
						Attribute attribute = new Attribute("RuntimeId:", runtimeIdString);
						attribute.Tooltip = "Unique identifier assigned to the user interface (UI) item";
						this.listAttributes.Add(attribute);
					}
				}
				catch { }
			}
            
			//get supported patterns
			string supportedPatterns = "";
			foreach (int propertyid in properties.Keys)
			{
				bool isSupported = false;
				try
				{
					isSupported = (bool)node.Element.GetCachedPropertyValue(propertyid);
				}
				catch { }
				
				if (isSupported == true)
				{
					if (supportedPatterns != "")
					{
						//supportedPatterns += ", ";
                        supportedPatterns += Environment.NewLine;
					}
					
					supportedPatterns += properties[propertyid];
				}
			}
			
			try
			{
				Attribute attr = new Attribute("Supported Patterns:", supportedPatterns);
				attr.Tooltip = "Control patterns that this Automation Element supported";
				this.listAttributes.Add(attr);
			}
			catch { }
			
			Attribute emptyAttribute = new Attribute("", "");
			this.listAttributes.Add(emptyAttribute);
			
			IUIAutomationElement2 element2 = node.Element as IUIAutomationElement2;
			if (element2 != null)
			{
				Attribute attribute = new Attribute("============", "IUIAutomationElement2 ========");
				attribute.Tooltip = "Properties added in IUIAutomationElement2 interface";
				this.listAttributes.Add(attribute);
			
				try
				{
					int optimizeForVisualContent = element2.CachedOptimizeForVisualContent;
					attribute = new Attribute("OptimizeForVisualContent:", optimizeForVisualContent.ToString());
					attribute.Tooltip = "Indicates whether the provider exposes only elements that are visible";
					this.listAttributes.Add(attribute);
				}
				catch {}
				
				try
				{
					LiveSetting liveSetting = element2.CachedLiveSetting;
					attribute = new Attribute("LiveSetting:", liveSetting.ToString());
					attribute.Tooltip = "Indicates the type of notifications, if any, that the element sends when the content of the element changes";
					this.listAttributes.Add(attribute);
				}
				catch {}
			}
			
			IUIAutomationElement3 element3 = node.Element as IUIAutomationElement3;
			if (element3 != null)
			{
				Attribute attribute = new Attribute("============", "IUIAutomationElement3 ========");
				attribute.Tooltip = "Properties added in IUIAutomationElement3 interface";
				this.listAttributes.Add(attribute);
				
				try
				{
					int isPeripheral = element3.CachedIsPeripheral;
					attribute = new Attribute("IsPeripheral:", isPeripheral.ToString());
					attribute.Tooltip = "The current peripheral UI indicator for the element";
					this.listAttributes.Add(attribute);
				}
				catch {}
			}
			
			IUIAutomationElement4 element4 = node.Element as IUIAutomationElement4;
			if (element4 != null)
			{
				Attribute attribute = new Attribute("============", "IUIAutomationElement4 ========");
				attribute.Tooltip = "Properties added in IUIAutomationElement4 interface";
				this.listAttributes.Add(attribute);
				
				try
				{
					int positionInSet = element4.CachedPositionInSet;
					attribute = new Attribute("PositionInSet:", positionInSet.ToString());
					attribute.Tooltip = "The current 1-based integer for the ordinal position in the set for the element";
					this.listAttributes.Add(attribute);
				}
				catch {}
				
				try
				{
					int sizeOfSet = element4.CachedSizeOfSet;
					attribute = new Attribute("SizeOfSet:", sizeOfSet.ToString());
					attribute.Tooltip = "The current 1-based integer for the size of the set where the element is located";
					this.listAttributes.Add(attribute);
				}
				catch {}
				
				try
				{
					int level = element4.CachedLevel;
					attribute = new Attribute("Level:", level.ToString());
					attribute.Tooltip = "The current 1-based integer for the level (hierarchy) for the element";
					this.listAttributes.Add(attribute);
				}
				catch {}
				
				try
				{
					int[] annotationTypes = element4.CachedAnnotationTypes;
					attribute = new Attribute("AnnotationTypes:", "[" + IntArrayToString(annotationTypes) + "]");
					attribute.Tooltip = "The current list of annotation types associated with this element, such as comment, header, footer, and so on";
					this.listAttributes.Add(attribute);
				}
				catch {}
			}
			
			IUIAutomationElement5 element5 = node.Element as IUIAutomationElement5;
			if (element5 != null)
			{
				Attribute attribute = new Attribute("============", "IUIAutomationElement5 ========");
				attribute.Tooltip = "Properties added in IUIAutomationElement5 interface";
				this.listAttributes.Add(attribute);
				
				try
				{
					int landmarkType = element5.CachedLandmarkType;
					string landmarkTypeString = null;
					if (LandmarkTypesDict.ContainsKey(landmarkType))
					{
						landmarkTypeString = LandmarkTypesDict[landmarkType];
					}
					else
					{
						landmarkTypeString = landmarkType.ToString();
					}
					attribute = new Attribute("LandmarkType:", landmarkTypeString);
					attribute.Tooltip = "The current landmark type ID for the automation element";
					this.listAttributes.Add(attribute);
				}
				catch {}
				
				try
				{
					string localizedLandmarkType = element5.CachedLocalizedLandmarkType;
					attribute = new Attribute("LocalizedLandmarkType:", "\"" + localizedLandmarkType + "\"");
					attribute.Tooltip = "A string containing the current localized landmark type for the automation element";
					this.listAttributes.Add(attribute);
				}
				catch {}
			}
			
			IUIAutomationElement6 element6 = node.Element as IUIAutomationElement6;
			if (element6 != null)
			{
				Attribute attribute = new Attribute("============", "IUIAutomationElement6 ========");
				attribute.Tooltip = "Properties added in IUIAutomationElement6 interface";
				this.listAttributes.Add(attribute);
				
				try
				{
					string fullDescription = element6.CachedFullDescription;
					attribute = new Attribute("FullDescription:", "\"" + fullDescription + "\"");
					attribute.Tooltip = "The current full description of the automation element";
					this.listAttributes.Add(attribute);
				}
				catch {}
			}
			
			IUIAutomationElement8 element8 = node.Element as IUIAutomationElement8;
			if (element8 != null)
			{
				Attribute attribute = new Attribute("============", "IUIAutomationElement8 ========");
				attribute.Tooltip = "Properties added in IUIAutomationElement8 interface";
				this.listAttributes.Add(attribute);
				
				try
				{
					int headingLevel = element8.CachedHeadingLevel;
					attribute = new Attribute("HeadingLevel:", HeadingLevelsDict[headingLevel]);
					attribute.Tooltip = "The current heading level of the automation element";
					this.listAttributes.Add(attribute);
				}
				catch {}
			}
			
			IUIAutomationElement9 element9 = node.Element as IUIAutomationElement9;
			if (element9 != null)
			{
				Attribute attribute = new Attribute("============", "IUIAutomationElement9 ========");
				attribute.Tooltip = "Properties added in IUIAutomationElement9 interface";
				this.listAttributes.Add(attribute);
				
				try
				{
					int isDialog = element9.CachedIsDialog;
					attribute = new Attribute("IsDialog:", isDialog.ToString());
					attribute.Tooltip = "The current is dialog window indicator for the element";
					this.listAttributes.Add(attribute);
				}
				catch {}
			}
		}
    }
}