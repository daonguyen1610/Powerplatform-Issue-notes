# Choice Fields: Enabling users to add values manually

In SharePoint lists, an option is available for choice fields where the user can add values manually. By default, when creating a Power App based on a SharePoint list with such a field where that option is selected, the Power App does not honor that option.

For single-valued choice fields, there is a very straightforward resolution that can allow a Power App to honor that option. In the choice field's card's (`crdChoiceField`) `Update` property, change the formula to the following.

```
Coalesce ( cmbChoiceField.Selected, { Value: cmbChoiceField.SearchText } )
```

For multi-valued choice fields, there is unfortunately no straightforward resolution to this issue. One could explore options of adding icons or buttons to trigger similar functionality.
