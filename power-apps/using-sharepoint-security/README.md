# Using SharePoint Security in Power Apps

SharePoint does not include any mechanism for column-level security. Often times, in customized SharePoint list forms, columns are shown or hidden based on security already present on the site; specifically via AD or SharePoint security group membership. For example, there is an Approvers group on the site. Certain columns are only available to edit by users who are members in that security group.

On the surface, Power Apps does not have a mechanism to read any security information from SharePoint sites; neither AD group membership nor SharePoint security group membership. However, Power Apps does respect SharePoint's security trimming when retrieving list data, including item-level security. We can exploit this to develop a solution that will enable a Power App to modify its functionality based on security groups.

If security groups are not currently used to restrict content by role, another solution will be presented that leverages a SharePoint configuration list to determine which users should be present in which role for the Power App.

# Determining security based on security groups

To implement this solution, we will create a SharePoint list to store in the information regarding security. That list will be populated with one item per SharePoint security group that the Power App needs to know about. Those items will be individually secured. Then, the Power App will be configured to use that new list and query it as needed. Follow the steps below to implement this solution.

1. **Set up a new Power Apps Security list in SharePoint.**
   1. Create a new custom list in the SharePoint site titled **Power Apps Security**. Even if there are multiple Power Apps for this particular site, only one of these lists need to be created. Leave the single default field of **Title** on the list.
   2. For each SharePoint security group that the Power Apps needs to be informed about,
      1. Create a new list item. Use the name of the security group for the **Title** field.
      2. Edit the new item's security.
         1. Stop it from inheriting permissions from its parent list.
         2. Add a **Read** permission for the security group pertaining to this item.
         3. Remove all other permissions *except* the **Full Control** permission for the site owners group.

2. **Set up a new data source in the Power App.**

   In the Power App, create a new data source for the **Power Apps Security** list.

3. **Retrieve the information from the new data source *once* when the Power App starts.**

   When the Power App starts, the app can load the information from the **Power Apps Security** list once, store it in a collection, and reuse it whenever a formula needs the information. This improves performance, especially with slower network connections.

   To do this, create a collection called `Security` and populate it with only the `Title` field from the **Power Apps Security** list. Add the formula below to the `App.OnStart` property of the Power App.

   ```
   ClearCollect (
       Security,
       ShowColumns (
           'Power Apps Security',
           "Title"
       )
   )
   ```

4. **Use the security collection information as needed.**
   
   In the Power App, when a formula is needed to determine if a user is a member of a particular group, use the following formula, replacing `SHAREPOINT GROUP NAME` with the name of the desired SharePoint group.

   ```
   Not (
       IsBlank (
           LookUp (
               Security,
               Title = "SHAREPOINT GROUP NAME"
           )
       )
   )
   ```

   This formula could be used with setting a variable, in the `Visible` property of controls, etc. It returns a boolean value: `true` if the user is a member of that SharePoint group, or `false` if the user is not a member of that SharePoint group.

# Determining security based on a SharePoint configuration list

TBD
