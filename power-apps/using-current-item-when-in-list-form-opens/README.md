# Using the Current List Item When an In-list Form Opens

There are many practical use cases where the developer needs to access the currently selected list item when the form opens to take some sort of action. This could be updating various fields, showing or hiding certain fields, or even securing the form away from a user if they should not be able to access it based on the currently selected list item.

In many cases, using the `SharePointIntegration.Selected` value is sufficient. However, for time-sensitive formulas that execute with an in-list form opens, use a different strategy: look up the currently selected list item based on `SharePointIntegration.SelectedListItemID`. This ensures that even if the currently selected list item has not been loaded into the `SharePointIntegration.Selected` property yet, the developer can still use data from the currently selected list item.

The following example shows the expression to use instead of `SharePointIntegration.Selected` given a data source of **Contoso Contacts**.

```
LookUp ( [@'Contoso Contacts'], ID = SharePointIntegration.SelectedListItemID )
```

*A special thanks to Dao Dinh Nguyen (dnguyen256@dxc.com) for developing this best practice!*
