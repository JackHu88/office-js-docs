# Please referece [here](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec) for Excel's upcoming features including 1.5 and beyond. 

## Excel 1.3 Release Features

### Pivot Table resource and refresh action

The pivot table collection is now available on the worksheet resource. To begin with, this will enable developers to know about the pivot tables are available on a given workbook. This will also enable applications to refresh an individual pivot table or all pivot tables associated with a worksheet from the source data. Applications can take advantage of this feature by refreshing the pivot table remotely based on timing or other business criteria.

* Get pivotTable: Read properties of pivotTable object; currently limited to `name`. 
* `refresh()`: Refreshes a given pivotTable. Example, GET /worksheets/{id or name}/pivotTables/{id}/refresh
* `refreshAll()`: Refresh all tables within a given worksheet. Note that this action is available only on the pivot table collection. 
* In the future, we'll look to add more pivot table functionality extending the same objects.

### Visible rows on a filtered range

We've added a new resource called `visibleView` that represents visible rows of a given `range`. When a filter is applied on a range, it is often useful to know what values are visible (or selected from filter criteria). Previously, this required going through the underlying range and checking `hidden` property to determine whether a row is visible or not. Such a procedure is cumbersome when the range is fairly large.

With the new feature being added, applications can now apply a filter on a large table based on desired criteria and get easy access to filtered data. This feature will reduce the complexity involved in determining visible rows and take advantage of Excel's filtering capability.

`range.visibleView`

The visible view resource comes with useful properties that represent the range such as cell addresses, column and row count, index, formulas, number format, values/text, and value types.

If the visible range happens to be large, getting the entire resource might not be performant. In such a case, iterating over the rows provides a better user experience. The `rows` navigation property on visibleView allows just that - iterating over the visible range rows.


### New range manipulation functions

The following new range functions have been added. These functions reduce the complexity involved in determining the target range that the applications wishes to operate on.

Gets a certain number of rows or columns based on relative position of a given range:

* `columnsAfter(count: int)` 
* `columnsBefore(count: int)`
* `rowsAbove(count: int)`
* `rowsBelow(count: int)` 

`resizedRange`: Gets a range object similar to the current range object, but with its bottom-right corner expanded or contracted by some number of rows and columns.

### New table resource properties

Gain insights into the structure of the Excel table by using the following new properties: 

* `highlightFirstColumn`
* `highlightLastColumn`
* `showBandedColumns`
* `showBandedRows`
* `showFilterButton`

### Add binding

[`binding`](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingcollection.md) object represents an office.js binding that is defined in the workbook. It can be used to create a binding on a range that can later be used for tracking the range and also to regreister for events. 

Previously, we only allowed binding read methods. Now, we offer: 
* `add()` Add a new binding based on a named item in the workbook.
