using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;
using System.Linq;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Runtime.InteropServices;
using Interop.UIAutomationClient;

namespace dDeltaSolutions.Spy
{
    public partial class MainWindow
    {
		private int maxNameLength = 100;
		private const int MAX_VALUE_LENGTH = 200;
		
		public class PatternInfo
		{
			public string Name;
			public bool IsSupported = false;
			
			public PatternInfo(string name)
			{
				Name = name;
			}
			
			public override string ToString()
			{
				return Name;
			}
		}
        
        private Dictionary<int, PatternInfo> patternsIds = null;
        
        private void FillPatternsDictionary()
        {
            if (patternsIds != null)
            {
				foreach (int key in patternsIds.Keys)
				{
					patternsIds[key].IsSupported = false;
				}
			
                return;
            }
            
            patternsIds = new Dictionary<int, PatternInfo>();
            
            patternsIds.Add(UIA_PatternIds.UIA_InvokePatternId, new PatternInfo("UIA_InvokePattern"));
            patternsIds.Add(UIA_PatternIds.UIA_SelectionPatternId, new PatternInfo("UIA_SelectionPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_ValuePatternId, new PatternInfo("UIA_ValuePattern"));
            patternsIds.Add(UIA_PatternIds.UIA_RangeValuePatternId, new PatternInfo("UIA_RangeValuePattern"));
            patternsIds.Add(UIA_PatternIds.UIA_ScrollPatternId, new PatternInfo("UIA_ScrollPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_ExpandCollapsePatternId, new PatternInfo("UIA_ExpandCollapsePattern"));
            patternsIds.Add(UIA_PatternIds.UIA_GridPatternId, new PatternInfo("UIA_GridPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_GridItemPatternId, new PatternInfo("UIA_GridItemPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_MultipleViewPatternId, new PatternInfo("UIA_MultipleViewPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_WindowPatternId, new PatternInfo("UIA_WindowPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_SelectionItemPatternId, new PatternInfo("UIA_SelectionItemPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_DockPatternId, new PatternInfo("UIA_DockPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_TablePatternId, new PatternInfo("UIA_TablePattern"));
            patternsIds.Add(UIA_PatternIds.UIA_TableItemPatternId, new PatternInfo("UIA_TableItemPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_TextPatternId, new PatternInfo("UIA_TextPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_TogglePatternId, new PatternInfo("UIA_TogglePattern"));
            patternsIds.Add(UIA_PatternIds.UIA_TransformPatternId, new PatternInfo("UIA_TransformPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_ScrollItemPatternId, new PatternInfo("UIA_ScrollItemPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_LegacyIAccessiblePatternId, new PatternInfo("UIA_LegacyIAccessiblePattern"));
            patternsIds.Add(UIA_PatternIds.UIA_ItemContainerPatternId, new PatternInfo("UIA_ItemContainerPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_VirtualizedItemPatternId, new PatternInfo("UIA_VirtualizedItemPattern"));
            patternsIds.Add(UIA_PatternIds.UIA_SynchronizedInputPatternId, new PatternInfo("UIA_SynchronizedInputPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_ObjectModelPatternId, new PatternInfo("UIA_ObjectModelPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_AnnotationPatternId, new PatternInfo("UIA_AnnotationPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_TextPattern2Id, new PatternInfo("UIA_TextPattern2"));
			patternsIds.Add(UIA_PatternIds.UIA_StylesPatternId, new PatternInfo("UIA_StylesPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_SpreadsheetPatternId, new PatternInfo("UIA_SpreadsheetPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_SpreadsheetItemPatternId, new PatternInfo("UIA_SpreadsheetItemPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_TransformPattern2Id, new PatternInfo("UIA_TransformPattern2"));
			patternsIds.Add(UIA_PatternIds.UIA_TextChildPatternId, new PatternInfo("UIA_TextChildPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_DragPatternId, new PatternInfo("UIA_DragPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_DropTargetPatternId, new PatternInfo("UIA_DropTargetPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_TextEditPatternId, new PatternInfo("UIA_TextEditPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_CustomNavigationPatternId, new PatternInfo("UIA_CustomNavigationPattern"));
			patternsIds.Add(UIA_PatternIds.UIA_SelectionPattern2Id, new PatternInfo("UIA_SelectionPattern2"));
        }
        
        private string[] ArrayToStringArray(Array arr)
        {
            List<string> result = new List<string>();
            foreach (object item in arr)
            {
                result.Add(item.ToString());
            }
            return result.ToArray();
        }

        private void SelectedItemChanged(TreeNode node, bool isAlive)
        {
            if (!isAlive)
            {
                this.attributesListView.Tag = node;
                this.patternsListView.Tag = null;
                this.listAttributes.Clear();
                this.listPatterns.Clear();
				GetCachedInformation(node);
            }
			else
			{
				this.attributesListView.Tag = node;
				this.listAttributes.Clear();
				GetCurrentInformation(node);
			}

			try
			{
				if (isAlive)
				{
					RefreshPatterns(node);
				}
			}
			catch { }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private static string GetStringFromPoint(tagPOINT pt)
        {
            return "X: " + pt.x + ", Y: " + pt.y;
        }

        private string IntToString(int input)
        {
            return input.ToString();
        }

        public static string RectangleToString(tagRECT rect)
        {
            string rectangleString = "Left:" + rect.left +
                            ", Top:" + rect.top +
                            ", Right:" + rect.right +
                            ", Bottom:" + rect.bottom + 
							" (Width:" + (rect.right - rect.left) + 
							", Height:" + (rect.bottom - rect.top) + ")";
            return rectangleString;
        }

        private void RefreshPatterns(TreeNode node)
        {
            this.patternsListView.Tag = node;
            this.listPatterns.Clear();

            if (MainWindow.rangeFromPointTimer != null)
            {
                MainWindow.rangeFromPointTimer.Stop();
            }
			
			Attribute attribute = null;

            #region DockPattern
			if (patternsIds[UIA_PatternIds.UIA_DockPatternId].IsSupported)
			{
				IUIAutomationDockPattern dockPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_DockPatternId) as IUIAutomationDockPattern;
				if (dockPattern != null)
				{
					try
					{
						string dockPositionString = dockPattern.CurrentDockPosition.ToString();
						attribute = new Attribute("DockPosition:", dockPositionString, "UIA_DockPattern");
						attribute.Tooltip = "The DockPosition of an Automation Element within a docking container";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region ExpandCollapsePattern
			if (patternsIds[UIA_PatternIds.UIA_ExpandCollapsePatternId].IsSupported)
			{
				IUIAutomationExpandCollapsePattern expandCollapsePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId) as IUIAutomationExpandCollapsePattern;
				if (expandCollapsePattern != null)
				{
					try
					{
						string expandCollapseState = expandCollapsePattern.CurrentExpandCollapseState.ToString();
						attribute = new Attribute("ExpandCollapseState:", expandCollapseState, "UIA_ExpandCollapsePattern");
						attribute.Tooltip = "The ExpandCollapseState of the Automation Element";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
			
            #region WindowPattern
			if (patternsIds[UIA_PatternIds.UIA_WindowPatternId].IsSupported)
			{
				IUIAutomationWindowPattern windowPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId) as IUIAutomationWindowPattern;
				if (windowPattern != null)
				{
					try
					{
						string canMaximizeString = windowPattern.CurrentCanMaximize.ToString();
						attribute = new Attribute("CanMaximize:", canMaximizeString, "UIA_WindowPattern");
						attribute.Tooltip = "A value that specifies whether the AutomationElement can be maximized";
						this.listPatterns.Add(attribute);
					}
					catch {}
						
					try
					{
						string canMinimizeString = windowPattern.CurrentCanMinimize.ToString();
						attribute = new Attribute("CanMinimize:", canMinimizeString, "UIA_WindowPattern");
						attribute.Tooltip = "A value that specifies whether the current AutomationElement can be minimized";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string isModalString = windowPattern.CurrentIsModal.ToString();
						attribute = new Attribute("IsModal:", isModalString, "UIA_WindowPattern");
						attribute.Tooltip = "A value that specifies whether the AutomationElement is modal.";
						this.listPatterns.Add(attribute);
					}
					catch {}
						
					try
					{
						string isTopmostString = windowPattern.CurrentIsTopmost.ToString();
						attribute = new Attribute("IsTopmost:", isTopmostString, "UIA_WindowPattern");
						attribute.Tooltip = "A value that specifies whether the AutomationElement is the topmost element in the z-order.";
						this.listPatterns.Add(attribute);
					}
					catch {}
						
					try
					{
						string windowInteractionState = windowPattern.CurrentWindowInteractionState.ToString();
						attribute = new Attribute("WindowInteraction:", windowInteractionState, "UIA_WindowPattern");
						attribute.Tooltip = "The WindowInteractionState of the AutomationElement";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string windowVisualState = windowPattern.CurrentWindowVisualState.ToString();
						attribute = new Attribute("WindowVisualState:", windowVisualState, "UIA_WindowPattern");
						attribute.Tooltip = "The WindowVisualState of the AutomationElement";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region TransformPattern
			if (patternsIds[UIA_PatternIds.UIA_TransformPatternId].IsSupported)
			{
				IUIAutomationTransformPattern transformPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TransformPatternId) as IUIAutomationTransformPattern;
				if (transformPattern != null)
				{
					try
					{
						string canMove = transformPattern.CurrentCanMove.ToString();
						attribute = new Attribute("CanMove:", canMove, "UIA_TransformPattern");
						attribute.Tooltip = "A value that specifies whether the UI Automation element can be moved";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string canResize = transformPattern.CurrentCanResize.ToString();
						attribute = new Attribute("CanResize:", canResize, "UIA_TransformPattern");
						attribute.Tooltip = "A value that specifies whether the UI Automation element can be resized";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string canRotate = transformPattern.CurrentCanRotate.ToString();
						attribute = new Attribute("CanRotate:", canRotate, "UIA_TransformPattern");
						attribute.Tooltip = "A value that specifies whether the UI Automation element can be rotated";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region GridPattern
			if (patternsIds[UIA_PatternIds.UIA_GridPatternId].IsSupported)
			{
				IUIAutomationGridPattern gridPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_GridPatternId) as IUIAutomationGridPattern;
				if (gridPattern != null)
				{
					try
					{
						string columnCount = gridPattern.CurrentColumnCount.ToString();
						attribute = new Attribute("ColumnCount:", columnCount, "UIA_GridPattern");
						attribute.Tooltip = "The number of columns in a grid";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string rowCount = gridPattern.CurrentRowCount.ToString();
						attribute = new Attribute("RowCount:", rowCount, "UIA_GridPattern");
						attribute.Tooltip = "Total number of rows in a grid";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region GridItemPattern
			if (patternsIds[UIA_PatternIds.UIA_GridItemPatternId].IsSupported)
			{
				IUIAutomationGridItemPattern gridItemPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_GridItemPatternId) as IUIAutomationGridItemPattern;
				if (gridItemPattern != null)
				{
					try
					{
						string column = gridItemPattern.CurrentColumn.ToString();
						attribute = new Attribute("Column:", column, "UIA_GridItemPattern");
						attribute.Tooltip = "Ordinal number of the column that contains the cell or item";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string columnSpan = gridItemPattern.CurrentColumnSpan.ToString();
						attribute = new Attribute("ColumnSpan:", columnSpan, "UIA_GridItemPattern");
						attribute.Tooltip = "The number of columns spanned by a cell or item";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						TreeNode tempNode = new TreeNode(gridItemPattern.CurrentContainingGrid);
						attribute = new Attribute("ContainingGrid:", tempNode.ToString(), "UIA_GridItemPattern");
						attribute.UnderneathElement = gridItemPattern.CurrentContainingGrid;
						attribute.Tooltip = "UI Automation element that supports GridPattern and represents the container of the cell or item";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string row = gridItemPattern.CurrentRow.ToString();
						attribute = new Attribute("Row:", row, "UIA_GridItemPattern");
						attribute.Tooltip = "Ordinal number of the row that contains the cell or item";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string rowSpan = gridItemPattern.CurrentRowSpan.ToString();
						attribute = new Attribute("RowSpan:", rowSpan, "UIA_GridItemPattern");
						attribute.Tooltip = "The number of rows spanned by a cell or item";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region ValuePattern
			if (patternsIds[UIA_PatternIds.UIA_ValuePatternId].IsSupported)
			{
				IUIAutomationValuePattern valuePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;
				if (valuePattern != null)
				{
					try
					{
						string isReadOnly = valuePattern.CurrentIsReadOnly.ToString();
						attribute = new Attribute("IsReadOnly:", isReadOnly, "UIA_ValuePattern");
						attribute.Tooltip = "A value that specifies whether the value of an UI Automation element is read-only";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string value = valuePattern.CurrentValue;
						bool truncated = false;
						if (value != null && value.Length > MAX_VALUE_LENGTH)
						{
							value = value.Substring(0, MAX_VALUE_LENGTH) + "...";
							truncated = true;
						}

						value = "\"" + value + "\"";
						if (truncated)
						{
							value += " -> (Double-click to see all text)";
						}
						attribute = new Attribute("Value:", value, "UIA_ValuePattern") { Pattern = valuePattern };
						attribute.Tooltip = "The value of the UI Automation element";
						if (truncated)
						{
							attribute.Tooltip += " (Double-click to see all text)";
						}
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region MultipleViewPattern
			if (patternsIds[UIA_PatternIds.UIA_MultipleViewPatternId].IsSupported)
			{
				IUIAutomationMultipleViewPattern multipleViewPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_MultipleViewPatternId) as IUIAutomationMultipleViewPattern;
				if (multipleViewPattern != null)
				{
					try
					{
						int currentViewId = multipleViewPattern.CurrentCurrentView;
						string currentViewString = "\"" + multipleViewPattern.GetViewName(currentViewId) + "\"";
						attribute = new Attribute("CurrentView:", currentViewString, "UIA_MultipleViewPattern");
						attribute.Tooltip = "Current control-specific view";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						Array viewIds = multipleViewPattern.GetCurrentSupportedViews();
						int index = 0;

						foreach (int viewId in viewIds)
						{
							string viewName = "\"" + multipleViewPattern.GetViewName(viewId) + "\"";
							attribute = new Attribute("SupportedViews[" + index.ToString() + "]:",
								viewName, "UIA_MultipleViewPattern");
							this.listPatterns.Add(attribute);

							index++;
						}
					}
					catch {}
				}
			}
            #endregion
            
            #region RangeValuePattern
			if (patternsIds[UIA_PatternIds.UIA_RangeValuePatternId].IsSupported)
			{
				IUIAutomationRangeValuePattern rangeValuePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_RangeValuePatternId) as IUIAutomationRangeValuePattern;
				if (rangeValuePattern != null)
				{
					try
					{
						string isReadOnlyString = rangeValuePattern.CurrentIsReadOnly.ToString();
						attribute = new Attribute("IsReadOnly:", isReadOnlyString, "UIA_RangeValuePattern");
						attribute.Tooltip = "A value that specifies whether the value of an UI Automation element is read-only";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string largeChangeString = rangeValuePattern.CurrentLargeChange.ToString();
						attribute = new Attribute("LargeChange:", largeChangeString, "UIA_RangeValuePattern");
						attribute.Tooltip = "Control specific large-change value";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string smallChangeString = rangeValuePattern.CurrentSmallChange.ToString();
						attribute = new Attribute("SmallChange:", smallChangeString, "UIA_RangeValuePattern");
						attribute.Tooltip = "The small-change value";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string maximumString = rangeValuePattern.CurrentMaximum.ToString();
						attribute = new Attribute("Maximum:", maximumString, "UIA_RangeValuePattern");
						attribute.Tooltip = "The maximum range value supported by the UI Automation element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string minimumString = rangeValuePattern.CurrentMinimum.ToString();
						attribute = new Attribute("Minimum:", minimumString, "UIA_RangeValuePattern");
						attribute.Tooltip = "The minimum range value supported by the UI Automation element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string valueString = rangeValuePattern.CurrentValue.ToString();
						attribute = new Attribute("Value:", valueString, "UIA_RangeValuePattern");
						attribute.Tooltip = "The current value of the UI Automation element";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region ScrollPattern
			if (patternsIds[UIA_PatternIds.UIA_ScrollPatternId].IsSupported)
			{
				IUIAutomationScrollPattern scrollPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ScrollPatternId) as IUIAutomationScrollPattern;
				if (scrollPattern != null)
				{
					try
					{
						string horizontallyScrollable = scrollPattern.CurrentHorizontallyScrollable.ToString();
						attribute = new Attribute("HorizontallyScrollable:", horizontallyScrollable, "UIA_ScrollPattern");
						attribute.Tooltip = "A value that indicates whether the UI Automation element can scroll horizontally";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string horizontalScrollPercent = scrollPattern.CurrentHorizontalScrollPercent.ToString();
						attribute = new Attribute("HorizontalScrollPercent:", horizontalScrollPercent, "UIA_ScrollPattern");
						attribute.Tooltip = "The current horizontal scroll position";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string horizontalViewSize = scrollPattern.CurrentHorizontalViewSize.ToString();
						attribute = new Attribute("HorizontalViewSize:", horizontalViewSize, "UIA_ScrollPattern");
						attribute.Tooltip = "Current horizontal view size";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						string verticallyScrollable = scrollPattern.CurrentVerticallyScrollable.ToString();
						attribute = new Attribute("VerticallyScrollable:", verticallyScrollable, "UIA_ScrollPattern");
						attribute.Tooltip = "A value that indicates whether the UI Automation element can scroll vertically";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string verticalScrollPercent = scrollPattern.CurrentVerticalScrollPercent.ToString();
						attribute = new Attribute("VerticalScrollPercent:", verticalScrollPercent, "UIA_ScrollPattern");
						attribute.Tooltip = "Current vertical scroll position";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string verticalViewSize = scrollPattern.CurrentVerticalViewSize.ToString();
						attribute = new Attribute("VerticalViewSize:", verticalViewSize, "UIA_ScrollPattern");
						attribute.Tooltip = "Current vertical view size";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region SelectionPattern
			if (patternsIds[UIA_PatternIds.UIA_SelectionPatternId].IsSupported)
			{
				IUIAutomationSelectionPattern selectionPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_SelectionPatternId) as IUIAutomationSelectionPattern;
				if (selectionPattern != null)
				{
					try
					{
						string canSelectMultiple = selectionPattern.CurrentCanSelectMultiple.ToString();
						attribute = new Attribute("CanSelectMultiple:", canSelectMultiple, "UIA_SelectionPattern");
						attribute.Tooltip = "A value that specifies whether the container allows " +
							"more than one child element to be selected concurrently";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						IUIAutomationElementArray selectedElements = selectionPattern.GetCurrentSelection();
						for (int i = 0; i < selectedElements.Length; i++)
						{
							IUIAutomationElement selectedElement = selectedElements.GetElement(i);
							TreeNode tempNode = new TreeNode(selectedElement);

							attribute = new Attribute("Selection[" + i.ToString() + "]:",
								tempNode.ToString(), "UIA_SelectionPattern");
							attribute.UnderneathElement = selectedElement;
							
							this.listPatterns.Add(attribute);
						}
					}
					catch {}

					try
					{
						string isSelectionRequired = selectionPattern.CurrentIsSelectionRequired.ToString();
						attribute = new Attribute("IsSelectionRequired:", isSelectionRequired, "UIA_SelectionPattern");
						attribute.Tooltip = "A value that specifies whether the container " + 
							"requires at least one child item to be selected";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region SelectionItemPattern
			if (patternsIds[UIA_PatternIds.UIA_SelectionItemPatternId].IsSupported)
			{
				IUIAutomationSelectionItemPattern selectionItemPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) as IUIAutomationSelectionItemPattern;
				if (selectionItemPattern != null)
				{
					try
					{
						string isSelected = selectionItemPattern.CurrentIsSelected.ToString();
						attribute = new Attribute("IsSelected:", isSelected, "UIA_SelectionItemPattern");
						attribute.Tooltip = "A value that indicates whether an item is selected";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						TreeNode tempNode = new TreeNode(selectionItemPattern.CurrentSelectionContainer);
						attribute = new Attribute("SelectionContainer:", tempNode.ToString(), "UIA_SelectionItemPattern");
						attribute.UnderneathElement = selectionItemPattern.CurrentSelectionContainer;
						attribute.Tooltip = "The AutomationElement that supports the SelectionPattern " +
							"control pattern and acts as the container for the calling object";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region TablePattern
			if (patternsIds[UIA_PatternIds.UIA_TablePatternId].IsSupported)
			{
				IUIAutomationTablePattern tablePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TablePatternId) as IUIAutomationTablePattern;
				if (tablePattern != null)
				{
					try
					{
						string rowOrColumnMajor = tablePattern.CurrentRowOrColumnMajor.ToString();
						attribute = new Attribute("RowOrColumnMajor:", rowOrColumnMajor, "UIA_TablePattern");
						attribute.Tooltip = "The primary direction of traversal for the table";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					//string columnCount = tablePattern.CurrentColumnCount.ToString();
					//string rowCount = tablePattern.CurrentRowCount.ToString();

					/*Attribute attribute = new Attribute("ColumnCount:", columnCount, "UIA_TablePattern");
					attribute.Tooltip = "The total number of columns in a table";
					this.listPatterns.Add(attribute);*/

					/*attribute = new Attribute("RowCount:", rowCount, "UIA_TablePattern");
					attribute.Tooltip = "The total number of rows in a table";
					this.listPatterns.Add(attribute);*/

					try
					{
						IUIAutomationElementArray columnHeaders = tablePattern.GetCurrentColumnHeaders();
						for (int i = 0; i < columnHeaders.Length; i++)
						{
							IUIAutomationElement columnHeader = columnHeaders.GetElement(i);
							TreeNode tempNode = new TreeNode(columnHeader);

							attribute = new Attribute("ColumnHeader[" + i + "]:", tempNode.ToString(),
								"UIA_TablePattern");
							attribute.UnderneathElement = columnHeader;
							this.listPatterns.Add(attribute);
						}
					}
					catch {}

					try
					{
						IUIAutomationElementArray rowHeaders = tablePattern.GetCurrentRowHeaders();
						for (int i = 0; i < rowHeaders.Length; i++)
						{
							IUIAutomationElement rowHeader = rowHeaders.GetElement(i);
							TreeNode tempNode = new TreeNode(rowHeader);

							attribute = new Attribute("RowHeader[" + i + "]:", tempNode.ToString(),
								"UIA_TablePattern");
							attribute.UnderneathElement = rowHeader;
							this.listPatterns.Add(attribute);
						}
					}
					catch {}
				}
			}
            #endregion
            
            #region TableItemPattern
			if (patternsIds[UIA_PatternIds.UIA_TableItemPatternId].IsSupported)
			{
				IUIAutomationTableItemPattern tableItemPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TableItemPatternId) as IUIAutomationTableItemPattern;
				if (tableItemPattern != null)
				{
					//Attribute attribute = null;
					TreeNode tempNode = null;
					IUIAutomationElementArray headerItems = null;
					
					/*string column = tableItemPattern.CurrentColumn.ToString();
					Attribute attribute = new Attribute("Column:", column, "UIA_TableItemPattern");
					attribute.Tooltip = "The ordinal number of the column containing the table cell or item";
					this.listPatterns.Add(attribute);*/

					/*string columnSpan = tableItemPattern.CurrentColumnSpan.ToString();
					attribute = new Attribute("ColumnSpan:", columnSpan, "UIA_TableItemPattern");
					attribute.Tooltip = "The total number of columns spanned by a table cell or item";
					this.listPatterns.Add(attribute);*/

					/*TreeNode tempNode = new TreeNode(tableItemPattern.CurrentContainingGrid);
					attribute = new Attribute("ContainingGrid:", tempNode.ToString(), "UIA_TableItemPattern");
					attribute.UnderneathElement = tableItemPattern.Current.ContainingGrid;
					attribute.Tooltip = "A UI Automation element that supports the GridPattern control pattern " +
						"and represents the table cell or item container";
					this.listPatterns.Add(attribute);*/

					try
					{
						headerItems = tableItemPattern.GetCurrentColumnHeaderItems();
						for (int i = 0; i < headerItems.Length; i++)
						{
							IUIAutomationElement header = headerItems.GetElement(i);
							tempNode = new TreeNode(header);
							attribute = new Attribute("ColumnHeaderItems[" + i + "]:",
								tempNode.ToString(), "UIA_TableItemPattern");
							attribute.UnderneathElement = header;
							this.listPatterns.Add(attribute);
						}
					}
					catch {}

					try
					{
						headerItems = tableItemPattern.GetCurrentRowHeaderItems();
						for (int i = 0; i < headerItems.Length; i++)
						{
							IUIAutomationElement header = headerItems.GetElement(i);
							tempNode = new TreeNode(header);
							attribute = new Attribute("RowHeaderItems[" + i + "]:",
								tempNode.ToString(), "UIA_TableItemPattern");
							attribute.UnderneathElement = header;
							this.listPatterns.Add(attribute);
						}
					}
					catch {}

					/*string row = tableItemPattern.CurrentRow.ToString();
					attribute = new Attribute("Row:", row, "UIA_TableItemPattern");
					attribute.Tooltip = "The ordinal number of the row containing the table cell or item";
					this.listPatterns.Add(attribute);*/

					/*string rowSpan = tableItemPattern.CurrentRowSpan.ToString();
					attribute = new Attribute("RowSpan:", rowSpan, "UIA_TableItemPattern");
					attribute.Tooltip = "The number of rows spanned by a table cell or item";
					this.listPatterns.Add(attribute);*/
				}
			}
            #endregion
            
            #region TogglePattern
			if (patternsIds[UIA_PatternIds.UIA_TogglePatternId].IsSupported)
			{
				IUIAutomationTogglePattern togglePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;
				if (togglePattern != null)
				{
					try
					{
						string toggleState = togglePattern.CurrentToggleState.ToString();
						attribute = new Attribute("ToggleState:", toggleState, "UIA_TogglePattern");
						attribute.Tooltip = "The toggle state of the AutomationElement";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
            #endregion
            
            #region TextPattern
			if (patternsIds[UIA_PatternIds.UIA_TextPatternId].IsSupported)
			{
				IUIAutomationTextPattern textPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId) as IUIAutomationTextPattern;
				if (textPattern != null)
				{
					try
					{
						string supportedTextSelection = textPattern.SupportedTextSelection.ToString();
						attribute = new Attribute("SupportedTextSelection:", supportedTextSelection,
							"UIA_TextPattern");
						attribute.Tooltip = "A value that specifies whether a text provider supports selection " + 
							"and, if so, the type of selection supported";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						IUIAutomationTextRange textPatternRange = textPattern.DocumentRange;
						AddTextPatternRangeAttributes(textPatternRange, "UIA_TextPattern.DocumentRange");
					}
					catch {}

					try
					{
						IUIAutomationTextRangeArray selectionRanges = textPattern.GetSelection();
						if (selectionRanges != null)
						{
							for (int i = 0; i < selectionRanges.Length; i++)
							{
								IUIAutomationTextRange textRange = selectionRanges.GetElement(i);
								AddTextPatternRangeAttributes(textRange, "UIA_TextPattern.GetSelection(" + i + ")");
							}
						}
					}
					catch {}

					try
					{
						IUIAutomationTextRangeArray visibleRanges = textPattern.GetVisibleRanges();
						if (visibleRanges != null)
						{
							for (int i = 0; i < visibleRanges.Length; i++)
							{
								IUIAutomationTextRange textRange = visibleRanges.GetElement(i);
								AddTextPatternRangeAttributes(textRange, "UIA_TextPattern.GetVisibleRanges(" + i + ")");
							}
						}
					}
					catch {}

					//double x = (double)System.Windows.Forms.Cursor.Position.X;
					//double y = (double)System.Windows.Forms.Cursor.Position.Y;

					//System.Windows.Point ptCursor = new System.Windows.Point(x, y);
				}
			}
            #endregion
            
            #region LegacyIAccessiblePattern
			if (patternsIds[UIA_PatternIds.UIA_LegacyIAccessiblePatternId].IsSupported)
			{
				IUIAutomationLegacyIAccessiblePattern legacyIAccessiblePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) as IUIAutomationLegacyIAccessiblePattern;
				if (legacyIAccessiblePattern != null)
				{
					try
					{
						string childId = legacyIAccessiblePattern.CurrentChildId.ToString();
						attribute = new Attribute("ChildId:", childId, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility child identifier for the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string name = "\"" + legacyIAccessiblePattern.CurrentName + "\"";
						attribute = new Attribute("Name:", name, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility name property of the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string value = legacyIAccessiblePattern.CurrentValue;
						bool truncated = false;
						if (value.Length > MAX_VALUE_LENGTH)
						{
							value = value.Substring(0, MAX_VALUE_LENGTH) + "...";
							truncated = true;
						}
						value = "\"" + value + "\"";
						if (truncated)
						{
							value += " -> (Double-click to see all text)";
						}
						attribute = new Attribute("Value:", value, "UIA_LegacyIAccessiblePattern") { Pattern = legacyIAccessiblePattern };
						attribute.Tooltip = "Microsoft Active Accessibility value property";
						if (truncated)
						{
							attribute.Tooltip += " (Double-click to see all text)";
						}
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string description = "\"" + legacyIAccessiblePattern.CurrentDescription + "\"";
						attribute = new Attribute("Description:", description, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility description of the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						uint role = legacyIAccessiblePattern.CurrentRole;
						string roleStr = role.ToString() + " (" + ((AccRoles)role).ToString() + ")";
						attribute = new Attribute("Role:", roleStr, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility role identifier of the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						uint state = legacyIAccessiblePattern.CurrentState;
						//StringBuilder stateText = new StringBuilder(1024);
						//Helper.GetStateText(state, stateText, 1024);
						string stateStr = "0x" + state.ToString("X") + " (" + Helper.GetStatesAsText(state) + ")";
						attribute = new Attribute("State:", stateStr, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility state identifier for the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string help = "\"" + legacyIAccessiblePattern.CurrentHelp + "\"";
						attribute = new Attribute("Help:", help, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility help string for the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string keyboardShortcut = "\"" + legacyIAccessiblePattern.CurrentKeyboardShortcut + "\"";
						attribute = new Attribute("KeyboardShortcut:", keyboardShortcut, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility keyboard shortcut property for the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string defaultAction = "\"" + legacyIAccessiblePattern.CurrentDefaultAction + "\"";
						attribute = new Attribute("DefaultAction:", defaultAction, "UIA_LegacyIAccessiblePattern");
						attribute.Tooltip = "Microsoft Active Accessibility current default action for the element";
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					/*IUIAutomationElementArray selection = legacyIAccessiblePattern.GetCurrentSelection();
					if (selection != null)
					{
						TreeNode tempNode = null;
						//System.Windows.MessageBox.Show("" + selection.Length);
						for (int i = 0; i < selection.Length; i++)
						{
							IUIAutomationElement crtSelection = selection.GetElement(i);
							tempNode = new TreeNode(crtSelection);
							attribute = new Attribute("CurrentSelection[" + i + "]:",
								tempNode.ToString(), "UIA_LegacyIAccessiblePattern");
							attribute.UnderneathElement = crtSelection;
							this.listPatterns.Add(attribute);
						}
					}*/
				}
			}
			#endregion
			
			if (patternsIds[UIA_PatternIds.UIA_ObjectModelPatternId].IsSupported)
			{
				IUIAutomationObjectModelPattern objectModelPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ObjectModelPatternId) as IUIAutomationObjectModelPattern;
				if (objectModelPattern != null)
				{
					try
					{
						object underlyingObjectModel = objectModelPattern.GetUnderlyingObjectModel();
						if (underlyingObjectModel != null)
						{
							attribute = new Attribute("UnderlyingObjectModel:", underlyingObjectModel.ToString(), "UIA_ObjectModelPattern");
							this.listPatterns.Add(attribute);
						}
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_AnnotationPatternId].IsSupported)
			{
				IUIAutomationAnnotationPattern annotationPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_AnnotationPatternId) as IUIAutomationAnnotationPattern;
				if (annotationPattern != null)
				{
					try
					{
						string annotationTypeId = annotationPattern.CurrentAnnotationTypeId.ToString();
						attribute = new Attribute("AnnotationTypeId:", annotationTypeId, "UIA_AnnotationPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string annotationTypeName = annotationPattern.CurrentAnnotationTypeName;
						attribute = new Attribute("AnnotationTypeName:", annotationTypeName, "UIA_AnnotationPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string author = annotationPattern.CurrentAuthor;
						attribute = new Attribute("Author:", author, "UIA_AnnotationPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						string dateTime = annotationPattern.CurrentDateTime;
						attribute = new Attribute("DateTime:", dateTime, "UIA_AnnotationPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement target = annotationPattern.CurrentTarget;
						TreeNode tempNode = new TreeNode(target);
						attribute = new Attribute("Target:", tempNode.ToString(), "UIA_AnnotationPattern");
						attribute.UnderneathElement = target;
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_TextPattern2Id].IsSupported)
			{
				IUIAutomationTextPattern2 textPattern2 = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TextPattern2Id) as IUIAutomationTextPattern2;
				if (textPattern2 != null)
				{
					try
					{
						string supportedTextSelection = textPattern2.SupportedTextSelection.ToString();
						attribute = new Attribute("SupportedTextSelection:", supportedTextSelection, "UIA_TextPattern2");
						attribute.Tooltip = "A value that specifies whether a text provider supports selection " + 
							"and, if so, the type of selection supported";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						IUIAutomationTextRange textPattern2Range = textPattern2.DocumentRange;
						AddTextPatternRangeAttributes(textPattern2Range, "UIA_TextPattern2.DocumentRange");
					}
					catch {}

					try
					{
						IUIAutomationTextRangeArray selectionRanges = textPattern2.GetSelection();
						if (selectionRanges != null)
						{
							for (int i = 0; i < selectionRanges.Length; i++)
							{
								IUIAutomationTextRange textRange = selectionRanges.GetElement(i);
								AddTextPatternRangeAttributes(textRange, "UIA_TextPattern2.GetSelection(" + i + ")");
							}
						}
					}
					catch {}

					try
					{
						IUIAutomationTextRangeArray visibleRanges = textPattern2.GetVisibleRanges();
						if (visibleRanges != null)
						{
							for (int i = 0; i < visibleRanges.Length; i++)
							{
								IUIAutomationTextRange textRange = visibleRanges.GetElement(i);
								AddTextPatternRangeAttributes(textRange, "UIA_TextPattern2.GetVisibleRanges(" + i + ")");
							}
						}
					}
					catch {}

					try
					{
						int isActive = 0;
						IUIAutomationTextRange caretRange = textPattern2.GetCaretRange(out isActive);
						AddTextPatternRangeAttributes(caretRange, "UIA_TextPattern2.GetCaretRange");
						attribute = new Attribute("isActive:", isActive.ToString(), "UIA_TextPattern2.GetCaretRange");
						//attribute.Tooltip = "";
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_StylesPatternId].IsSupported)
			{
				IUIAutomationStylesPattern stylesPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_StylesPatternId) as IUIAutomationStylesPattern;
				if (stylesPattern != null)
				{
					try
					{
						attribute = new Attribute("StyleId:", stylesPattern.CurrentStyleId.ToString(), "UIA_StylesPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("StyleName:", stylesPattern.CurrentStyleName, "UIA_StylesPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("FillColor:", stylesPattern.CurrentFillColor.ToString(), "UIA_StylesPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("FillPatternStyle:", stylesPattern.CurrentFillPatternStyle, "UIA_StylesPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("Shape:", stylesPattern.CurrentShape, "UIA_StylesPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("FillPatternColor:", stylesPattern.CurrentFillPatternColor.ToString(), "UIA_StylesPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("ExtendedProperties:", stylesPattern.CurrentExtendedProperties, "UIA_StylesPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_SpreadsheetItemPatternId].IsSupported)
			{
				IUIAutomationSpreadsheetItemPattern spreadsheetItemPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_SpreadsheetItemPatternId) as IUIAutomationSpreadsheetItemPattern;
				if (spreadsheetItemPattern != null)
				{
					try
					{
						attribute = new Attribute("Formula:", spreadsheetItemPattern.CurrentFormula, "UIA_SpreadsheetItemPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElementArray annObjects = spreadsheetItemPattern.GetCurrentAnnotationObjects();
						for (int i = 0; i < annObjects.Length; i++)
						{
							IUIAutomationElement annObject = annObjects.GetElement(i);
							TreeNode tempNode = new TreeNode(annObject);
							attribute = new Attribute("AnnotationObjects[" + i + "]:", tempNode.ToString(), "UIA_SpreadsheetItemPattern");
							attribute.UnderneathElement = annObject;
							this.listPatterns.Add(attribute);
						}
					}
					catch {}
					
					try
					{
						Array annTypes = spreadsheetItemPattern.GetCurrentAnnotationTypes();
						if (annTypes.Length > 0 && annotationTypesDict == null)
						{
							annotationTypesDict = new Dictionary<int, string>();
						
							annotationTypesDict.Add(60020, "AnnotationType_AdvancedProofingIssue");
							annotationTypesDict.Add(60019, "AnnotationType_Author");
							annotationTypesDict.Add(60022, "AnnotationType_CircularReferenceError");
							annotationTypesDict.Add(60003, "AnnotationType_Comment");
							annotationTypesDict.Add(60018, "AnnotationType_ConflictingChange");
							annotationTypesDict.Add(60021, "AnnotationType_DataValidationError");
							annotationTypesDict.Add(60012, "AnnotationType_DeletionChange");
							annotationTypesDict.Add(60016, "AnnotationType_EditingLockedChange");
							annotationTypesDict.Add(60009, "AnnotationType_Endnote");
							annotationTypesDict.Add(60017, "AnnotationType_ExternalChange");
							annotationTypesDict.Add(60007, "AnnotationType_Footer");
							annotationTypesDict.Add(60010, "AnnotationType_Footnote");
							annotationTypesDict.Add(60014, "AnnotationType_FormatChange");
							annotationTypesDict.Add(60004, "AnnotationType_FormulaError");
							annotationTypesDict.Add(60002, "AnnotationType_GrammarError");
							annotationTypesDict.Add(60006, "AnnotationType_Header");
							annotationTypesDict.Add(60008, "AnnotationType_Highlighted");
							annotationTypesDict.Add(60011, "AnnotationType_InsertionChange");
							annotationTypesDict.Add(60023, "AnnotationType_Mathematics");
							annotationTypesDict.Add(60013, "AnnotationType_MoveChange");
							annotationTypesDict.Add(60001, "AnnotationType_SpellingError");
							annotationTypesDict.Add(60005, "AnnotationType_TrackChanges");
							annotationTypesDict.Add(60000, "AnnotationType_Unknown");
							annotationTypesDict.Add(60015, "AnnotationType_UnsyncedChange");
						}
						for (int i = 0; i < annTypes.Length; i++)
						{
							object annTypeObj = annTypes.GetValue(i);
							if (annTypeObj != null)
							{
								int annType = (int)annTypeObj;
								attribute = new Attribute("AnnotationTypes[" + i + "]:", annotationTypesDict[annType], "UIA_SpreadsheetItemPattern");
								this.listPatterns.Add(attribute);
							}
						}
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_TransformPattern2Id].IsSupported)
			{
				IUIAutomationTransformPattern2 transformPattern2 = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TransformPattern2Id) as IUIAutomationTransformPattern2;
				if (transformPattern2 != null)
				{
					try
					{
						attribute = new Attribute("CanMove:", transformPattern2.CurrentCanMove.ToString(), "UIA_TransformPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("CanResize:", transformPattern2.CurrentCanResize.ToString(), "UIA_TransformPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("CanRotate:", transformPattern2.CurrentCanRotate.ToString(), "UIA_TransformPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("CanZoom:", transformPattern2.CurrentCanZoom.ToString(), "UIA_TransformPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("ZoomLevel:", transformPattern2.CurrentZoomLevel.ToString(), "UIA_TransformPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("ZoomMinimum:", transformPattern2.CurrentZoomMinimum.ToString(), "UIA_TransformPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("ZoomMaximum:", transformPattern2.CurrentZoomMaximum.ToString(), "UIA_TransformPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_TextChildPatternId].IsSupported)
			{
				IUIAutomationTextChildPattern textChildPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TextChildPatternId) as IUIAutomationTextChildPattern;
				if (textChildPattern != null)
				{
					try
					{
						IUIAutomationElement textContainer = textChildPattern.TextContainer;
						TreeNode tempNode = new TreeNode(textContainer);
						attribute = new Attribute("TextContainer:", tempNode.ToString(), "UIA_TextChildPattern");
						attribute.UnderneathElement = textContainer;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationTextRange textRange = textChildPattern.TextRange;
						AddTextPatternRangeAttributes(textRange, "UIA_TextChildPattern.TextRange");
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_DragPatternId].IsSupported)
			{
				IUIAutomationDragPattern dragPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_DragPatternId) as IUIAutomationDragPattern;
				if (dragPattern != null)
				{
					try
					{
						attribute = new Attribute("IsGrabbed:", dragPattern.CurrentIsGrabbed.ToString(), "UIA_DragPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("DropEffect:", dragPattern.CurrentDropEffect, "UIA_DragPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						Array dropEffects = dragPattern.CurrentDropEffects;
						for (int i = 0; i < dropEffects.Length; i++)
						{
							attribute = new Attribute("DropEffects[" + i + "]:", dropEffects.GetValue(i).ToString(), "UIA_DragPattern");
							this.listPatterns.Add(attribute);
						}
					}
					catch {}
					
					try
					{
						IUIAutomationElementArray grabbedItems = dragPattern.GetCurrentGrabbedItems();
						for (int i = 0; i < grabbedItems.Length; i++)
						{
							IUIAutomationElement grabbedItem = grabbedItems.GetElement(i);
							TreeNode tempNode = new TreeNode(grabbedItem);
							attribute = new Attribute("GrabbedItems[" + i + "]:", tempNode.ToString(), "UIA_DragPattern");
							attribute.UnderneathElement = grabbedItem;
							this.listPatterns.Add(attribute);
						}
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_DropTargetPatternId].IsSupported)
			{
				IUIAutomationDropTargetPattern dropTargetPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_DropTargetPatternId) as IUIAutomationDropTargetPattern;
				if (dropTargetPattern != null)
				{
					try
					{
						attribute = new Attribute("DropTargetEffect:", dropTargetPattern.CurrentDropTargetEffect, "UIA_DropTargetPattern");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						Array dropTargetEffects = dropTargetPattern.CurrentDropTargetEffects;
						for (int i = 0; i < dropTargetEffects.Length; i++)
						{
							attribute = new Attribute("DropTargetEffects[" + i + "]:", dropTargetEffects.GetValue(i).ToString(), "UIA_DropTargetPattern");
							this.listPatterns.Add(attribute);
						}
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_TextEditPatternId].IsSupported)
			{
				IUIAutomationTextEditPattern textEditPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_TextEditPatternId) as IUIAutomationTextEditPattern;
				if (textEditPattern != null)
				{
					try
					{
						string supportedTextSelection = textEditPattern.SupportedTextSelection.ToString();
						attribute = new Attribute("SupportedTextSelection:", supportedTextSelection, "UIA_TextEditPattern");
						attribute.Tooltip = "A value that specifies whether a text provider supports selection " + 
							"and, if so, the type of selection supported";
						this.listPatterns.Add(attribute);
					}
					catch {}

					try
					{
						IUIAutomationTextRange textEditPatternRange = textEditPattern.DocumentRange;
						AddTextPatternRangeAttributes(textEditPatternRange, "UIA_TextEditPattern.DocumentRange");
					}
					catch {}

					try
					{
						IUIAutomationTextRangeArray selectionRanges = textEditPattern.GetSelection();
						if (selectionRanges != null)
						{
							for (int i = 0; i < selectionRanges.Length; i++)
							{
								IUIAutomationTextRange textRange = selectionRanges.GetElement(i);
								AddTextPatternRangeAttributes(textRange, "UIA_TextEditPattern.GetSelection(" + i + ")");
							}
						}
					}
					catch {}

					try
					{
						IUIAutomationTextRangeArray visibleRanges = textEditPattern.GetVisibleRanges();
						if (visibleRanges != null)
						{
							for (int i = 0; i < visibleRanges.Length; i++)
							{
								IUIAutomationTextRange textRange = visibleRanges.GetElement(i);
								AddTextPatternRangeAttributes(textRange, "UIA_TextEditPattern.GetVisibleRanges(" + i + ")");
							}
						}
					}
					catch {}
					
					try
					{
						IUIAutomationTextRange activeComposition = textEditPattern.GetActiveComposition();
						AddTextPatternRangeAttributes(activeComposition, "UIA_TextEditPattern.GetActiveComposition()");
					}
					catch {}
					
					try
					{
						IUIAutomationTextRange conversionTarget = textEditPattern.GetConversionTarget();
						AddTextPatternRangeAttributes(conversionTarget, "UIA_TextEditPattern.GetConversionTarget()");
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_CustomNavigationPatternId].IsSupported)
			{
				IUIAutomationCustomNavigationPattern custNavPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_CustomNavigationPatternId) as IUIAutomationCustomNavigationPattern;
				if (custNavPattern != null)
				{
					try
					{
						IUIAutomationElement parent = custNavPattern.Navigate(NavigateDirection.NavigateDirection_Parent);
						TreeNode tempNode = new TreeNode(parent);
						attribute = new Attribute("Navigate(NavigateDirection_Parent)", tempNode.ToString(), "UIA_CustomNavigationPatternId");
						attribute.UnderneathElement = parent;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement nextSibling = custNavPattern.Navigate(NavigateDirection.NavigateDirection_NextSibling);
						TreeNode tempNode = new TreeNode(nextSibling);
						attribute = new Attribute("Navigate(NavigateDirection_NextSibling)", tempNode.ToString(), "UIA_CustomNavigationPatternId");
						attribute.UnderneathElement = nextSibling;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement prevSibling = custNavPattern.Navigate(NavigateDirection.NavigateDirection_PreviousSibling);
						TreeNode tempNode = new TreeNode(prevSibling);
						attribute = new Attribute("Navigate(NavigateDirection_PreviousSibling)", tempNode.ToString(), "UIA_CustomNavigationPatternId");
						attribute.UnderneathElement = prevSibling;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement firstChild = custNavPattern.Navigate(NavigateDirection.NavigateDirection_FirstChild);
						TreeNode tempNode = new TreeNode(firstChild);
						attribute = new Attribute("Navigate(NavigateDirection_FirstChild)", tempNode.ToString(), "UIA_CustomNavigationPatternId");
						attribute.UnderneathElement = firstChild;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement lastChild = custNavPattern.Navigate(NavigateDirection.NavigateDirection_LastChild);
						TreeNode tempNode = new TreeNode(lastChild);
						attribute = new Attribute("Navigate(NavigateDirection_LastChild)", tempNode.ToString(), "UIA_CustomNavigationPatternId");
						attribute.UnderneathElement = lastChild;
						this.listPatterns.Add(attribute);
					}
					catch {}
				}
			}
			
			if (patternsIds[UIA_PatternIds.UIA_SelectionPattern2Id].IsSupported)
			{
				IUIAutomationSelectionPattern2 selectionPattern2 = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_SelectionPattern2Id) as IUIAutomationSelectionPattern2;
				if (selectionPattern2 != null)
				{
					try
					{
						attribute = new Attribute("CanSelectMultiple:", selectionPattern2.CurrentCanSelectMultiple.ToString(), "UIA_SelectionPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("IsSelectionRequired:", selectionPattern2.CurrentIsSelectionRequired.ToString(), "UIA_SelectionPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement firstSelectedItem = selectionPattern2.CurrentFirstSelectedItem;
						TreeNode tempNode = new TreeNode(firstSelectedItem);
						attribute = new Attribute("FirstSelectedItem:", tempNode.ToString(), "UIA_SelectionPattern2");
						attribute.UnderneathElement = firstSelectedItem;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement lastSelectedItem = selectionPattern2.CurrentLastSelectedItem;
						TreeNode tempNode = new TreeNode(lastSelectedItem);
						attribute = new Attribute("LastSelectedItem:", tempNode.ToString(), "UIA_SelectionPattern2");
						attribute.UnderneathElement = lastSelectedItem;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElement currentSelectedItem = selectionPattern2.CurrentCurrentSelectedItem;
						TreeNode tempNode = new TreeNode(currentSelectedItem);
						attribute = new Attribute("CurrentSelectedItem:", tempNode.ToString(), "UIA_SelectionPattern2");
						attribute.UnderneathElement = currentSelectedItem;
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						attribute = new Attribute("ItemCount:", selectionPattern2.CurrentItemCount.ToString(), "UIA_SelectionPattern2");
						this.listPatterns.Add(attribute);
					}
					catch {}
					
					try
					{
						IUIAutomationElementArray selection = selectionPattern2.GetCurrentSelection();
						for (int i = 0; i < selection.Length; i++)
						{
							IUIAutomationElement selectedItem = selection.GetElement(i);
							TreeNode tempNode = new TreeNode(selectedItem);
							attribute = new Attribute("GetCurrentSelection()[" + i + "]:", tempNode.ToString(), "UIA_SelectionPattern2");
							attribute.UnderneathElement = selectedItem;
							this.listPatterns.Add(attribute);
						}
					}
					catch {}
				}
			}
        }
		
		private Dictionary<int, string> annotationTypesDict = null;

        private static System.Windows.Forms.Timer rangeFromPointTimer = null;
        private static List<Attribute> rangeFromPointAttributes = new List<Attribute>();

        void rangeFromPointTimer_Tick(object sender, EventArgs e)
        {
            return;
        }

        private void AddTextPatternRangeAttributes(IUIAutomationTextRange textPatternRange, string rangeName)
        {
            if (textPatternRange == null)
            {
                return;
            }

            //range from point
            bool isRangeFromPoint = (rangeName == "TextPattern.RangeFromPoint");
            if (isRangeFromPoint)
            {
                if (MainWindow.rangeFromPointTimer == null)
                {
                    MainWindow.rangeFromPointTimer = new System.Windows.Forms.Timer();

                    MainWindow.rangeFromPointTimer.Interval = 100;
                    MainWindow.rangeFromPointTimer.Tick += rangeFromPointTimer_Tick;

                    MainWindow.rangeFromPointAttributes.Clear();
                }
            }
            /////////

			Attribute attribute = null;
			try
			{
				string documentRangeText = textPatternRange.GetText(-1);
				bool truncated = false;
				if (documentRangeText != null && documentRangeText.Length > 30)
				{
					documentRangeText = documentRangeText.Substring(0, 30) + "...";
					truncated = true;
				}
				documentRangeText = "\"" + documentRangeText + "\"";
				if (truncated)
				{
					documentRangeText += " -> (Double-click to see all text)";
				}
				//documentRangeText = (documentRangeText == null ? "null" : "\"" + documentRangeText + "\"");

				attribute = new Attribute("Text:", documentRangeText, rangeName);
				attribute.TextPatternRange = textPatternRange;
				attribute.Tooltip = "The plain text of the text range";
				if (truncated)
				{
					attribute.Tooltip += " (Double-click to see all text)";
				}
				this.listPatterns.Add(attribute);
				if (isRangeFromPoint)
				{
					MainWindow.rangeFromPointAttributes.Add(attribute);
				}
			}
			catch {}

            TreeNode tempNode = null;
			try
			{
				IUIAutomationElementArray children = textPatternRange.GetChildren();
				if (children != null)
				{
					for (int i = 0; i < children.Length; i++)
					{
						IUIAutomationElement child = children.GetElement(i);
						tempNode = new TreeNode(child);
						attribute = new Attribute("Children[" + i + "]:", tempNode.ToString(), rangeName);
						attribute.UnderneathElement = child;
						this.listPatterns.Add(attribute);

						if (isRangeFromPoint)
						{
							MainWindow.rangeFromPointAttributes.Add(attribute);
						}
					}
				}
			}
			catch {}

			try
			{
				IUIAutomationElement enclosingElement = textPatternRange.GetEnclosingElement();
				if (enclosingElement != null)
				{
					tempNode = new TreeNode(enclosingElement);
					attribute = new Attribute("EnclosingElement:", tempNode.ToString(), rangeName);
					attribute.UnderneathElement = enclosingElement;
					attribute.Tooltip = "The innermost Automation Element that encloses the text range";
					this.listPatterns.Add(attribute);
					if (isRangeFromPoint)
					{
						MainWindow.rangeFromPointAttributes.Add(attribute);
					}
				}
			}
			catch {}

			try
			{
				Array boundingRectangles = textPatternRange.GetBoundingRectangles();
				if (boundingRectangles != null)
				{
					try
					{
						for (int i = 0; i < boundingRectangles.Length/4; i++)
						{
							//System.Windows.MessageBox.Show(boundingRectangles.GetValue(4*i).ToString());
							//System.Windows.MessageBox.Show(boundingRectangles.GetValue(4*i+1).ToString());
							//System.Windows.MessageBox.Show(boundingRectangles.GetValue(4*i+2).ToString());
							//System.Windows.MessageBox.Show(boundingRectangles.GetValue(4*i+3).ToString());
							
							tagRECT boundingRectangle;
							boundingRectangle.left = (int)(double)boundingRectangles.GetValue(4*i);
							boundingRectangle.top = (int)(double)boundingRectangles.GetValue(4*i + 1);
							int width = (int)(double)boundingRectangles.GetValue(4*i + 2);
							boundingRectangle.right = boundingRectangle.left + width;
							int height = (int)(double)boundingRectangles.GetValue(4*i + 3);
							boundingRectangle.bottom = boundingRectangle.top + height;
							
							string boundingRectangleString = RectangleToString(boundingRectangle);
							attribute = new Attribute("BoundingRectangles[" + i + "]:",
								boundingRectangleString, rangeName);
							attribute.BoundingRectangle = boundingRectangle;
							this.listPatterns.Add(attribute);

							if (isRangeFromPoint)
							{
								MainWindow.rangeFromPointAttributes.Add(attribute);
							}
						}
					}
					catch {}
				}
			}
			catch {}

            //attribute = new Attribute("Mouse over for...", "more attributes", rangeName);
			attribute = new Attribute("Double click here...", "for TextRange attributes", rangeName);
            attribute.TextPatternRange = textPatternRange;
            this.listPatterns.Add(attribute);

            //start timer
            if (isRangeFromPoint)
            {
                MainWindow.rangeFromPointTimer.Start();
            }
        }

        private static string IntArrayToString(int[] array)
        {
            StringBuilder result = new StringBuilder();
            foreach (int n in array)
            {
                if (result.Length != 0)
                {
                    result.Append(",");
                }

                result.Append(n.ToString());
            }
            return result.ToString();
        }

        public static string DoubleArrayToString(double[] array)
        {
            string result = string.Empty;
            foreach (double d in array)
            {
                if (result != string.Empty)
                {
                    result += ",";
                }

                result += d.ToString("0.00");
            }
            return result;
        }

        private static string[] IntArrayToStringArray(int[] intArray)
        {
            List<string> result = new List<string>();
            foreach (int n in intArray)
            {
                result.Add(n.ToString());
            }
            return result.ToArray();
        }

        public class Attribute
        {
            private string name = string.Empty;
            private string value = string.Empty;
            private string group = string.Empty;
            private string tooltip = null;

            private IUIAutomationTextRange textPatternRange = null;
            private IUIAutomationElement underneathElement = null;
            private tagRECT? boundingRectangle = null;

            public Attribute(string name, string value)
            {
                this.name = name;
                this.value = value;
            }

            public Attribute(string name, string value, string group)
            {
                this.name = name;
                this.value = value;
                this.group = group;
            }

            public string Name
            {
                get
                {
                    return this.name;
                }
                set
                {
                    this.name = value;
                }
            }

            public string Value
            {
                get
                {
                    return this.value;
                }
                set
                {
                    this.value = value;
                }
            }

            public string Group
            {
                get
                {
                    return this.group;
                }
                set
                {
                    this.group = value;
                }
            }

            public string Tooltip
            {
                get
                {
                    /*if ((this.name == "Mouse over for...") && (this.textPatternRange != null))
                    {
                        string moreAttributes = GetMoreAttributes(this.textPatternRange);
                        this.tooltip = moreAttributes;
                    }*/

                    return this.tooltip;
                }
                set
                {
                    this.tooltip = value;
                }
            }

            public IUIAutomationTextRange TextPatternRange
            {
                get
                {
                    return this.textPatternRange;
                }
                set
                {
                    this.textPatternRange = value;
                }
            }

            public IUIAutomationElement UnderneathElement
            {
                get
                {
                    return this.underneathElement;
                }
                set
                {
                    this.underneathElement = value;
                }
            }

            public tagRECT? BoundingRectangle
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
			
			public object Pattern { get; set; }

            /*private static string GetAutomationTextAttributeValue(
                TextPatternRange textPatternRange, AutomationTextAttribute attribute)
            {
                object value = null;
                string retString = string.Empty;

                try
                {
                    value = textPatternRange.GetAttributeValue(attribute);
                }
                catch { }

                if (value != null)
                {
                    if (value == TextPattern.MixedAttributeValue)
                    {
                        retString = "Mixed values";
                    }
                    else
                    {
                        retString = value.ToString();
                    }
                }

                return retString;
            }*/

            /*private static Color GetColorFromArgb(int argb)
            {
                Color retColor = Color.FromArgb(argb);

                //check if it is a known color
                foreach (KnownColor knownColor in Enum.GetValues(typeof(KnownColor)))
                {
                    Color color = Color.FromKnownColor(knownColor);
                    if (color.ToArgb() == argb)
                    {
                        retColor = color;
                        break;
                    }
                }

                return retColor;
            }*/
        }

        private void OnClickCopyValue(object sender, RoutedEventArgs e)
        {
			try
			{
				Attribute attribute = this.attributesListView.SelectedItem as Attribute;
				if (attribute == null)
				{
					return;
				}

				string attrString = GetAttributeString(attribute);
				System.Windows.Clipboard.SetText(attrString);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message);
			}
        }

        private void OnClickCopyAllAttributes(object sender, RoutedEventArgs e)
        {
			try
			{
				string text = "";

				foreach (object attr in attributesListView.Items)
				{
					Attribute attribute = attr as Attribute;
					if (attribute != null)
					{
						text += (attribute.Name + " " + GetAttributeString(attribute, true) + Environment.NewLine);
					}
				}

				System.Windows.Clipboard.SetText(text);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message);
			}
        }
		
		private void OnClickCopyPropNameAndValue(object sender, RoutedEventArgs e)
        {
			try
			{
				Attribute attribute = this.attributesListView.SelectedItem as Attribute;
				if (attribute == null)
				{
					return;
				}

				string attrString = attribute.Name + " " + GetAttributeString(attribute, true);
				System.Windows.Clipboard.SetText(attrString);
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message);
			}
        }

        private string GetAttributeString(Attribute attribute, bool copyAll = false)
        {
            try
            {
                if (attribute.Name == "Name:")
                {
                    TreeNode node = this.attributesListView.Tag as TreeNode;

                    if (node != null)
                    {
                        string nameString = "";
						try
						{
							nameString = node.Element.CurrentName;
						}
						catch
						{
							nameString = node.Element.CachedName;
						}

                        //return ("Name: \"" + nameString + "\"");
						if (copyAll == true)
						{
							return ("\"" + nameString + "\"");
						}
						return nameString;
                    }
                }
                
            }
            catch { }

            try
            {
                //return (attribute.Name + " " + attribute.Value);
				if (copyAll == true)
				{
					return attribute.Value;
				}
				
				string value = attribute.Value;
				int length = value.Length;
				if (value.StartsWith("\"") && value.EndsWith("\"") && length >= 2)
				{
					return value.Remove(length - 1, 1).Remove(0, 1);
				}
				return value;
            }
            catch { }

            return "";
        }

        private void OnClickCopyValuePattern(object sender, RoutedEventArgs e)
        {
            Attribute attribute = this.patternsListView.SelectedItem as Attribute;
            if (attribute == null)
            {
                return;
            }

            try
            {
                if ((attribute.Name == "Value:") && (attribute.Group == "UIA_ValuePattern"))
                {
                    TreeNode node = this.patternsListView.Tag as TreeNode;
                    
                    if (node != null)
                    {
                        IUIAutomationValuePattern valuePattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;
                        string value = valuePattern.CurrentValue;
                        System.Windows.Clipboard.SetText(value);
                        return;
                    }
                }
                if ((attribute.Name == "Text:") && 
					(attribute.Group.StartsWith("UIA_TextPattern.") || attribute.Group.StartsWith("UIA_TextChildPattern.") || 
					attribute.Group.StartsWith("UIA_TextPattern2.") || attribute.Group.StartsWith("UIA_TextEditPattern.")))
                {
                    IUIAutomationTextRange textPatternRange = attribute.TextPatternRange;
                    if (textPatternRange != null)
                    {
                        string text = textPatternRange.GetText(-1);
                        System.Windows.Clipboard.SetText(text);
                        return;
                    }
                }
                if ((attribute.Name == "Value:") && (attribute.Group == "UIA_LegacyIAccessiblePattern"))
                {
                    TreeNode node = this.patternsListView.Tag as TreeNode;
                    
                    if (node != null)
                    {
                        var legacyPattern = node.Element.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) as IUIAutomationLegacyIAccessiblePattern;
                        string value = legacyPattern.CurrentValue;
                        System.Windows.Clipboard.SetText(value);
                        return;
                    }
                }
            }
            catch { }

            try
            {
                System.Windows.Clipboard.SetText(attribute.Value);
            }
            catch { }
        }

        private void aboutButton_Click(object sender, RoutedEventArgs e)
        {
			/*if (menuAlwaysOnTop.IsChecked == true)
			{
				return;
			}*/
            AboutForm aboutForm = new AboutForm() { TopMost = this.Topmost };
            aboutForm.ShowDialog();
        }

        private void helpButton_Click(object sender, RoutedEventArgs e)
        {
			ProcessStartInfo startInfo = new ProcessStartInfo("http://automationspy.freecluster.eu/Help/AutomationSpyHelp.html");
            Process.Start(startInfo);
        }

        private void OnPatternsContextOpened(object sender, RoutedEventArgs e)
        {
            return;
        }

        private void OnPatternsContextLoaded(object sender, RoutedEventArgs e)
        {
            Attribute attribute = this.patternsListView.SelectedItem as Attribute;
            if (attribute == null)
            {
                return;
            }

            try
            {
                if (attribute.BoundingRectangle.HasValue)
                {
                    System.Windows.Controls.MenuItem highlightMenuItem = new System.Windows.Controls.MenuItem();
                    highlightMenuItem.Header = "Highlight";
                    highlightMenuItem.Click += highlightMenuItem_Click;
                    highlightMenuItem.Tag = attribute.BoundingRectangle.Value;
                    this.patternsListView.ContextMenu.Items.Add(highlightMenuItem);
                    
                    return;
                }
                else if (attribute.UnderneathElement != null)
                {
                    System.Windows.Controls.MenuItem gotoMenuItem = new System.Windows.Controls.MenuItem();
                    gotoMenuItem.Header = "Jump to";
                    gotoMenuItem.Click += gotoMenuItem_Click;
                    gotoMenuItem.Tag = attribute.UnderneathElement;
                    this.patternsListView.ContextMenu.Items.Add(gotoMenuItem);

                    System.Windows.Controls.MenuItem highlightMenuItem = new System.Windows.Controls.MenuItem();
                    highlightMenuItem.Header = "Highlight";
                    highlightMenuItem.Click += highlightMenuItem_Click;
                    highlightMenuItem.Tag = attribute.UnderneathElement;
                    this.patternsListView.ContextMenu.Items.Add(highlightMenuItem);

                    return;
                }
            }
            catch { }

            return;
        }

        private void OnAttributesContextLoaded(object sender, RoutedEventArgs e)
        {
            Attribute attribute = this.attributesListView.SelectedItem as Attribute;
            if (attribute == null)
            {
                return;
            }

            try
            {
                if (attribute.UnderneathElement != null)
                {
                    System.Windows.Controls.MenuItem gotoMenuItem = new System.Windows.Controls.MenuItem();
                    gotoMenuItem.Header = "Jump to";
                    gotoMenuItem.Click += gotoMenuItem_Click;
                    gotoMenuItem.Tag = attribute.UnderneathElement;
                    this.attributesListView.ContextMenu.Items.Add(gotoMenuItem);
                    
                    System.Windows.Controls.MenuItem highlightMenuItem = new System.Windows.Controls.MenuItem();
                    highlightMenuItem.Header = "Highlight";
                    highlightMenuItem.Click += highlightMenuItem_Click;
                    highlightMenuItem.Tag = attribute.UnderneathElement;
                    this.attributesListView.ContextMenu.Items.Add(highlightMenuItem);
                    
                    return;
                }
            }
            catch { }

            return;
        }

        private void OnAttributesContextUnloaded(object sender, RoutedEventArgs e)
        {
            List<System.Windows.Controls.MenuItem> itemsToDelete =
                new List<System.Windows.Controls.MenuItem>();

            foreach (System.Windows.Controls.MenuItem menuitem in
                this.attributesListView.ContextMenu.Items)
            {
                string header = menuitem.Header.ToString();
                if (header == "Jump to" || header == "Highlight")
                {
                    itemsToDelete.Add(menuitem);
                }
            }

            foreach (System.Windows.Controls.MenuItem menuItem in itemsToDelete)
            {
                this.attributesListView.ContextMenu.Items.Remove(menuItem);
            }
        }

        private void OnPatternsContextUnloaded(object sender, RoutedEventArgs e)
        {
            List<System.Windows.Controls.MenuItem> itemsToDelete =
                new List<System.Windows.Controls.MenuItem>();

            foreach (System.Windows.Controls.MenuItem menuitem in 
                this.patternsListView.ContextMenu.Items)
            {
                if (menuitem.Header.ToString() == "Highlight")
                {
                    itemsToDelete.Add(menuitem);
                }
                else if (menuitem.Header.ToString() == "Jump to")
                {
                    itemsToDelete.Add(menuitem);
                }
            }

            foreach (System.Windows.Controls.MenuItem menuItem in itemsToDelete)
            {
                this.patternsListView.ContextMenu.Items.Remove(menuItem);
            }
        }

        void gotoMenuItem_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.MenuItem menuItem = 
                (e.OriginalSource as System.Windows.Controls.MenuItem);
            if (menuItem == null)
            {
                return;
            }

            IUIAutomationElement element = menuItem.Tag as IUIAutomationElement;
            if (element == null)
            {
                return;
            }

            SelectAutomationElementInTree(element);
            this.tvElements.Focus();
        }

        void highlightMenuItem_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.MenuItem menuItem = 
                (e.OriginalSource as System.Windows.Controls.MenuItem);
            if (menuItem == null)
            {
                return;
            }

            IUIAutomationElement underneathElement = menuItem.Tag as IUIAutomationElement;
            tagRECT rect;

            if (underneathElement != null)
            {
                try
                {
                    rect = underneathElement.CurrentBoundingRectangle;
                }
                catch 
                {
                    return;
                }
            }
            else
            {
                rect = (tagRECT)menuItem.Tag;
            }

            HighlightRect(rect);
        }

        public static void HighlightRect(tagRECT rect)
		{
			try
			{
				TryHighlightRect(rect);
			}
			catch { }
		}
		
		private static void TryHighlightRect(tagRECT rect)
        {
            if (rect.right <= 0 && rect.bottom <= 0) // infinity
            {
                return;
            }

            int left = 0;
            int top = 0;
            int width = 0;
            int height = 0;

            try
            {
                left = rect.left;
                top = rect.top;

                width = rect.right - rect.left;
                height = rect.bottom - rect.top;
            }
            catch { }

            int thickness = 4;

            GraphicsPath path = new GraphicsPath();

            path.AddRectangle(new System.Drawing.Rectangle(0, 0, thickness, 
                height)); //add left
            path.AddRectangle(new System.Drawing.Rectangle(thickness, 
                height - thickness, width - 2 * thickness, thickness)); //add bottom
            path.AddRectangle(new System.Drawing.Rectangle(width - thickness, 0,
                thickness, height)); //add right
            path.AddRectangle(new System.Drawing.Rectangle(thickness, 0,
                width - 2 * thickness, thickness)); //add top

            System.Drawing.Region region = new System.Drawing.Region(path);
            Form tempForm = new Form();

            tempForm.ForeColor = System.Drawing.Color.Green;
            tempForm.BackColor = System.Drawing.Color.Green;

            tempForm.FormBorderStyle = FormBorderStyle.None;
            tempForm.StartPosition = FormStartPosition.Manual;

            tempForm.ShowInTaskbar = false;
            tempForm.TopMost = true;

            tempForm.Left = left;
            tempForm.Top = top;

            tempForm.Width = width;
            tempForm.Height = height;

            tempForm.Region = region;

            // highlight one second
            tempForm.Show();
            Thread.Sleep(200);
            //tempForm.Visible = false;
			tempForm.Left = -10000;
            tempForm.Top = -10000;
            Thread.Sleep(200);
            
            //tempForm.Visible = true;
			tempForm.Left = left;
            tempForm.Top = top;
            Thread.Sleep(200);
            tempForm.Close();
        }

        private ObservableCollection<Attribute> listAttributes = new ObservableCollection<Attribute>();
        private ObservableCollection<Attribute> listPatterns = new ObservableCollection<Attribute>();
    }
}