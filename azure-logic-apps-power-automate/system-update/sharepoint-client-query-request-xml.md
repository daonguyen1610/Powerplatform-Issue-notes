# SharePoint Client Query Request XML

In support of the primary article of this repository, this article reviews more in depth how the CSOM (client-side object model) API's Request XML body is built. Using that knowledge, this article will then use that knowledge to review extended scenarios of manipulating multiple list items in the same CSOM API request.

Microsoft has open sourced the protocols which power SharePoint's CSOM API and has several deep technical resources available on its implementation. Key resources and references include:
- The complete documentation of the [SharePoint Client Query Protocol](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csom/fd645da2-fa28-4daa-b3cd-8f4e506df117), published in both PDF and DOCX formats. This protocol outlines the client/server communication for the CSOM API. Section 3.1.1 is most relevant to this article.
- The full [Request XML Schema](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csom/2845342b-b15c-4870-baf8-3187870695b2) of the SharePoint Client Query Protocol.
- The complete documentation of the [SharePoint Client-Side Object Model Protocol](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/8ebba305-5af3-477d-bf4f-ad378a39eaba), published in both PDF and DOCX formats. This protocol outlines the representation of the CSOM API. Section 3.2.5 is an excellent reference regarding the objects, properties, and methods available in the CSOM API in the exact format that the Request XML of the SharePoint Client Query Protocol will require. This is close but not exactly similar to the .NET implementation of the CSOM API.
- The documentation for the [Microsoft.SharePoint.Client.ListItem object](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/4a06e0a3-2752-47b2-92f6-970b437d1e17) in the SharePoint Client-Side Object Model Protocol, which includes information regarding which types to use for which field types and the various properties and methods available on the ListItem object. (However, please note that this information isn't always proving to be enough to determine the exact format that should be used. Refer to the [SetFieldValue field value parameter article](setfieldvalue-field-value-parameter.md) for more details.)

When building a request to the CSOM API, it's most straightforward to consider the CSOM code you might write first. Then, translate that code to the request XML sent to the `client.svc` WCF service.

Note that since the action in the logic app or flow is calling the CSOM API WCF Service endpoint (`_vti_bin/client.svc/ProcessQuery`) from a particular web, the request is executed within the context of that web. It is possible to manipulate and utilize multiple objects from the same web or even from within the same site collection containing that web should permissions allow the caller to do so.

## Translating between the Request XML and CSOM API Code

To start, consider the example from the primary article. In this scenario, consider updating two fields in the list item that triggered the logic app or flow: a `Status` field to `In Progress` and a `Completed` field to `No`. The XML would be built as follows.

```xml
<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0"
         LibraryVersion="16.0.0.0" ApplicationName="Power Automate">
    <Actions>
        <Method Name="SetFieldValue" Id="3" ObjectPathId="2">
            <Parameters>
                <Parameter Type="String"><![CDATA[Status]]></Parameter>
                <Parameter Type="String"><![CDATA[In Progress]]></Parameter>
            </Parameters>
        </Method>
        <Method Name="SetFieldValue" Id="4" ObjectPathId="2">
            <Parameters>
                <Parameter Type="String"><![CDATA[Completed]]></Parameter>
                <Parameter Type="String"><![CDATA[No]]></Parameter>
            </Parameters>
        </Method>
        <Method Name="SystemUpdate" Id="5" ObjectPathId="2" />
    </Actions>
    <ObjectPaths>
        <StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />
        <Property Id="1" ParentId="0" Name="Web" />
        <Method Id="2" ParentId="1" Name="GetListItem">
            <Parameters>
                <Parameter Type="String">/sites/ModernTeam/@{triggerBody()?['{FullPath}']}</Parameter>
            </Parameters>
        </Method>
    </ObjectPaths>
</Request>
```

*Note: The `@{triggerBody()?['{FullPath}']}` translates to the web-relative URL to the list and list item, such as `Lists/Example/123_.000`, and is available on any triggered item.*

Consider the CSOM code that would be written in this scenario.

```csharp
ClientContext currentClientContext = ClientContext.Current;
Web web = currentClientContext.Web;
ListItem listItem = web.GetListItem("/sites/ModernTeam/Lists/Example/123_.000");

listItem.SetFieldValue("Status", "In Progress");
listItem.SetFieldValue("Completed", "No");
listItem.SystemUpdate();

currentClientContext.ExecuteQuery();
```

Start with the `ObjectPaths` section. Those are evaluated first in order based on their assigned Id number. Think of this section as defining variables that will be used later.

```csharp
// <StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />
ClientContext currentClientContext = ClientContext.Current;
```

- `Id="0"` now represents the object in the `currentClientContext` variable
- `TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}"` refers to the `ClientContext` class and it does not change no matter the tenant, site, or environment
- The `StaticProperty` element name refers to getting a static property from this class
- The `Name` attribute gives the specific name of the static property - in this example, `Current`

```csharp
// <Property Id="1" ParentId="0" Name="Web" />
Web web = currentClientContext.Web;
```

- `Id="1"` now represents the object in the `web` variable
- The `currentClientContext` object is held in `Id="0"`, which is why the `ParentId` attribute is set to `0`
- The `Property` element name refers to getting an instance property from the parent (`Id="0"` also known as `currentClientContext`)
- The `Name` attribute gives the specific name of the instance property - in this example, `Web`

```csharp
// <Method Id="2" ParentId="1" Name="GetListItem">
//     <Parameters>
//         <Parameter Type="String">/sites/ModernTeam/@{triggerBody()?['{FullPath}']}</Parameter>
//     </Parameters>
// </Method>
ListItem listItem = web.GetListItem("/sites/ModernTeam/Lists/Example/123_.000");
```

- `Id="2"` now represents the object in the `listItem` variable
- The `web` object is held in `Id="1"`, which is why the `ParentId` attribute is set to `1`
- The `Method` element name refers to calling an instance method from the parent (`Id="1"` also known as `web`)
- The `Name` attribute gives the specific name of the instance method - in this example, `GetListItem`
- The `Parameters` element name and its children specify the parameters of the method

Once all of the `ObjectPaths` are evaluated, a set of variables/objects is now available for use in the `Actions`.

Continue with the `Actions` section. Actions are evaluated in order based on their assigned Id number. Think of this section as calling methods that will perform an action.

```csharp
// <Method Name="SetFieldValue" Id="3" ObjectPathId="2">
//     <Parameters>
//         <Parameter Type="String"><![CDATA[Status]]></Parameter>
//         <Parameter Type="String"><![CDATA[In Progress]]></Parameter>
//     </Parameters>
// </Method>
listItem.SetFieldValue("Status", "In Progress");
```

- The `listItem` object is held in `Id="2"`, which is why the `ObjectPathId` attribute is set to `2`
- The `Method`  refers to calling an instance method from the object (`Id="2"` also known as `listItem`)
- The `Name` attribute gives the specific name of the instance method - in this example, `SetFieldValue`
- The `Parameters` element has one `Parameter` child element per parameter of the method. There are two parameters into this method: one for the field's internal name and one for the updated value. Based on the Status field's type in SharePoint, the second parameter is a `String` data type.

Note that the `SetFieldValue` method is *not* a method on the .NET implementation of the CSOM API. Rather, this method is defined as part of the [SharePoint Client-Side Object Model Protocol](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/8ebba305-5af3-477d-bf4f-ad378a39eaba). See [Section 3.2.5.87](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/4a06e0a3-2752-47b2-92f6-970b437d1e17) for information regarding the Microsoft.SharePoint.Client.ListItem object. Specifically, the `SetFieldValue` method is covered in [Section 3.2.5.87.2.1.6](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/84e7ddc7-dd63-4ba5-b184-1f7e550c5048) of the protocol documentation.

The determination of the format of the second parameter for the field value is reviewed at the higher-level [3.2.5.87 section](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/4a06e0a3-2752-47b2-92f6-970b437d1e17). However, this has not proved to provide enough documentation, so please refer to the [SetFieldValue field value parameter article](setfieldvalue-field-value-parameter.md) for further details and documentation on how to format this second parameter for the field value.

It is recommended to use the `<![CDATA[]]>` section around the field internal name and updated value to avoid needing to encode characters that have special meaning in XML documents: `<`, `>`, `&`, `'`, and `"`.

```csharp
// <Method Name="SetFieldValue" Id="4" ObjectPathId="2">
//     <Parameters>
//         <Parameter Type="String"><![CDATA[Completed]]></Parameter>
//         <Parameter Type="String"><![CDATA[No]]></Parameter>
//     </Parameters>
// </Method>
listItem.SetFieldValue("Completed", "No");
```

- The `listItem` object is held in `Id="2"`, which is why the `ObjectPathId` attribute is set to `2`
- The `Method` element name refers to calling an instance method from the object (`Id="2"` also known as `listItem`)
- The `Name` attribute gives the specific name of the instance method - in this example, `SetFieldValue`
- The `Parameters` element contains `Parameter` child elements per parameter of the method as described above

```csharp
// <Method Name="SystemUpdate" Id="5" ObjectPathId="2" />
listItem.SystemUpdate();
```

- The `listItem` object is held in `Id="2"`, which is why the `ObjectPathId` attribute is set to `2`
- The `Method` element name refers to calling an instance method from the object (`Id="2"` also known as `listItem`)
- The `Name` attribute gives the specific name of the instance method - in this example, `SystemUpdate`
- The `SystemUpdate` method has no parameters, so the `Parameters` element is omitted

Finally, now that both the `ObjectPaths` and `Actions` sections have been handled, any resulting objects are returned in the response from the WCF service. Neither the `SetFieldValue` method nor the `SystemUpdate` method return anything, so no data is returned from the WCF service; only a status and error/exception if one occurred. The call to the WCF service is akin to the call to the `ExecuteQuery` method.

```csharp
currentClientContext.ExecuteQuery();
```

## Example Scenario: Updating both the triggered item and an item in a different list in the same web

For this scenario, let's consider that we want to change two separate items. We do not want to trigger an update event and therefore we'll execute a `SystemUpdate()` for both items.

1. The first item is the item that originally triggered the logic app or flow. This is in the **Example** list with an ID of **123**. We need to change
   - its **Status** value to **Requested Approval**.
2. The second item is a item in a different list in the same web as the first item. This is in the **Activity** list with an ID of **42**. We need to change
   - its **LastActivity** field's value to **Requested Approval**, and its
   - its **ApprovalsRequested** field's value to **1**.

The CSOM code would be written as follows.

```csharp
/* ObjectPaths (think defining variables to be used later) */

ClientContext currentClientContext = ClientContext.Current; // Id="0"
Web web = currentClientContext.Web; // Id="1"

// For the first list item, the logic app or flow provides the {FullPath} value
// from the trigger's body. Use it to avoid extra work.
ListItem firstListItem = web.GetListItem("/sites/ModernTeam/Lists/Example/123_.000"); // Id="2"

// For the second list item, it may be more readable to use the list title and item ID.
ListCollection lists = web.Lists; // Id="3"
List secondList = lists.GetByTitle("Activity"); // Id="4"
ListItem secondListItem = secondList.GetItemById(42); // Id="5"

/* Actions (think calling methods that will perform an action) */

firstListItem.SetFieldValue("Status", "Requested Approval"); // Id="6"
firstListItem.SystemUpdate(); // Id="7"

secondListItem.SetFieldValue("LastActivity", "Requested Approval"); // Id="8"
secondListItem.SetFieldValue("ApprovalsRequested", 1); // Id="9"
secondListItem.SystemUpdate(); // Id="10"

currentClientContext.ExecuteQuery(); // This represents the call to the WCF Service
```

Based on the above CSOM code, the following XML fragments can be constructed.

`ObjectPaths` > `Id="0"`

```xml
<!-- ClientContext currentClientContext = ClientContext.Current; -->
<StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />
```

- The CSOM code is getting a static property. Therefore, use the `StaticProperty` element.
- Start with `Id="0"` and increment the `Id` number for each object path, then each action.
- The `ClientContext` class is referenced using `TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}"`. This does not ever change, regardless of the tenant, site, or environment being used.
- The `Name` attribute refers to the property being obtained, which is `Current` in this example

`ObjectPaths` > `Id="1"`

```xml
<!-- Web web = currentClientContext.Web; -->
<Property Id="1" ParentId="0" Name="Web" />
```

- The CSOM code is getting an instance property. Therefore, use the `Property` element.
- Continue with `Id="1"` (don't forget to increment!)
- The `currentClientContext` object is referenced via `Id="0"`. Therefore, set the `ParentId` attribute to `0`.
- The `Name` attribute refers to the property being obtained, which is `Web` in this example

`ObjectPaths` > `Id="2"`

```xml
<!-- ListItem firstListItem = web.GetListItem("/sites/ModernTeam/Lists/Example/123_.000"); -->
<Method Id="2" ParentId="1" Name="GetListItem">
    <Parameters>
        <Parameter Type="String">/sites/ModernTeam/@{triggerBody()?['{FullPath}']}</Parameter>
    </Parameters>
</Method>
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- Continue with `Id="2"` (don't forget to increment!)
- The `web` object is referenced via `Id="1"`. Therefore, set the `ParentId` attribute to `1`.
- The `Name` attribute refers to the method being called, which is `GetListItem` in this example
- The `Parameters` element refers to the parameters being passed into the method being called, which is a single `String` parameter containing a server-relative path to the list item.
- Recall that since the first item is using the item that triggered the logic app or flow, it's easier to use this syntax with the expression `@{triggerBody()?['{FullPath}']}` so that you can refer to that list item in one method call.

`ObjectPaths` > `Id="3"`

```xml
<!-- ListCollection lists = web.Lists; -->
<Property Id="3" ParentId="1" Name="Lists" />
```

- The CSOM code is getting an instance property. Therefore, use the `Property` element.
- Continue with `Id="3"` (don't forget to increment!)
- The `web` object is referenced via `Id="1"`. Therefore, set the `ParentId` attribute to `1`.
- The `Name` attribute refers to the property being obtained, which is `Lists` in this example

`ObjectPaths` > `Id="4"`

```xml
<!-- List secondList = lists.GetByTitle("Activity"); -->
<Method Id="4" ParentId="3" Name="GetByTitle">
    <Parameters>
        <Parameter Type="String">Activity</Parameter>
    </Parameters>
</Method>
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- Continue with `Id="4"` (don't forget to increment!)
- The `lists` object is referenced via `Id="3"`. Therefore, set the `ParentId` attribute to `3`.
- The `Name` attribute refers to the method being called, which is `GetByTitle` in this example
- The `Parameters` element refers to the parmeters being passed into the method being called, which is a single `String` parameter containing the list title.

`ObjectPaths` > `Id="5"`

```xml
<!-- ListItem secondListItem = secondList.GetItemById(42); -->
<Method Id="5" ParentId="4" Name="GetItemById">
    <Parameters>
        <Parameter Type="Int32">42</Parameter>
    </Parameters>
</Method>
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- Continue with `Id="5"` (don't forget to increment!)
- The `secondList` object is referenced via `Id="4"`. Therefore, set the `ParentId` attribute to `4`.
- The `Name` attribute refers to the method being called, which is `GetItemById` in this example
- The `Parameters` element refers to the parameters being passed into the method being called, which is a single `Int32` parameter containing the list item ID number.
- What data type should be used for a particular method? It is stated in the Microsoft documentation to be an `Int32` data type in [section 3.2.5.79.2.2.3 GetItemById method](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/41ffb409-2c60-44fa-9d58-b3e4a224b798) of the SharePoint Client-Side Object Model Protocol documentation. [Section 3.2.5](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/191beaea-4f63-47db-995a-d066326cc26c) of the SharePoint Client-Side Object Model Protocol documentation contains all of the objects with descriptions of their properties and methods. That section is the best reference to determine specific data types, method parameters, and more.

**Note: `Id="2"` represents the first list item and `Id="5"` represents the second list item in this example. Have the Id number(s) of the list item(s) quickly available since they will be used often in the Actions below.**

`Actions` > `Id="6"`

```xml
<!-- firstListItem.SetFieldValue("Status", "Requested Approval"); -->
<Method Name="SetFieldValue" Id="6" ObjectPathId="2">
    <Parameters>
        <Parameter Type="String"><![CDATA[Status]]></Parameter>
        <Parameter Type="String"><![CDATA[Requested Approval]]></Parameter>
    </Parameters>
</Method>
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- The `Name` attribute refers to the method being called, which is `SetFieldValue` in this example
- Continue with `Id="6"` (don't forget to increment!)
- The `firstListItem` object is referenced via `Id="2"`. Therefore, set the `ObjectPathId` attribute to `2`.
- The `Parameters` element refers to the parameters being passed into the method being called

Per the [3.2.5.87.2.1.6 SetFieldValue method](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/84e7ddc7-dd63-4ba5-b184-1f7e550c5048) Client-Side Object Model Protocol documentation, the `SetFieldValue` method accepts two parameters.
1. The first is a `String` representing the field's internal name.
2. The second is a `CSOM Object`, which could be nearly any available type.
   - The [3.2.5.87 Microsoft.SharePoint.Client.ListItem object](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/4a06e0a3-2752-47b2-92f6-970b437d1e17) Client-Side Object Model Protocol documentation details exactly which SharePoint field type maps to which CSOM value type.
   - The specifc syntax for the value types to use in the XML `Type` attribute (i.e. Boolean, Int32, UInt64, String, Guid, etc.) can be found in the [6.1 Request XML Schema](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csom/2845342b-b15c-4870-baf8-3187870695b2) Client Query Protocol documentation under `MethodParameterTypeType`.

`Actions` > `Id="7"`

```xml
<!-- firstListItem.SystemUpdate(); -->
<Method Name="SystemUpdate" Id="7" ObjectPathId="2" />
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- The `Name` attribute refers to the method being called, which is `SystemUpdate` in this example
- Continue with `Id="7"` (don't forget to increment!)
- The `firstListItem` object is referenced via `Id="2"`. Therefore, set the `ObjectPathId` attribute to `2`.
- The `Parameters` element is omitted since this method does not accept any parameters

`Actions` > `Id="8"`

```xml
<!-- secondListItem.SetFieldValue("LastActivity", "Requested Approval"); -->
<Method Name="SetFieldValue" Id="8" ObjectPathId="5">
    <Parameters>
        <Parameter Type="String"><![CDATA[LastActivity]]></Parameter>
        <Parameter Type="String"><![CDATA[Requested Approval]]></Parameter>
    </Parameters>
</Method>
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- The `Name` attribute refers to the method being called, which is `SetFieldValue` in this example
- Continue with `Id="8"` (don't forget to increment!)
- The `secondListItem` object is referenced via `Id="5"`. Therefore, set the `ObjectPathId` attribute to `5`.
- The `Parameters` element refers to the parameters being passed into the method being called

`Actions` > `Id="9"`

```xml
<!-- secondListItem.SetFieldValue("ApprovalsRequested", 1); -->
<Method Name="SetFieldValue" Id="9" ObjectPathId="5">
    <Parameters>
        <Parameter Type="String"><![CDATA[ApprovalsRequested]]></Parameter>
        <Parameter Type="Int32"><![CDATA[1]]></Parameter>
    </Parameters>
</Method>
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- The `Name` attribute refers to the method being called, which is `SetFieldValue` in this example
- Continue with `Id="9"` (don't forget to increment!)
- The `secondListItem` object is referenced via `Id="5"`. Therefore, set the `ObjectPathId` attribute to `5`.
- The `Parameters` element refers to the parameters being passed into the method being called

`Actions` > `Id="10"`

```xml
<!-- secondListItem.SystemUpdate(); -->
<Method Name="SystemUpdate" Id="10" ObjectPathId="5" />
```

- The CSOM code is calling an instance method. Therefore, use the `Method` element.
- The `Name` attribute refers to the method being called, which is `SystemUpdate` in this example
- Continue with `Id="10"` (don't forget to increment!)
- The `secondListItem` object is referenced via `Id="5"`. Therefore, set the `ObjectPathId` attribute to `5`.
- The `Parameters` element is omitted since this method does not accept any parameters

**At this point, a collection of XML fragments have been created representing `ObjectPaths` and `Actions`. Insert those fragments in order by Id number into the following XML document.**

```xml
<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0"
         LibraryVersion="16.0.0.0" ApplicationName="<!-- Either Logic Apps or Power Automate -->">
    <Actions>
        <!-- All Action fragments go here, in order by Id number -->
    </Actions>
    <ObjectPaths>
        <!-- All ObjectPath fragments go here, in order by Id number -->
    </ObjectPaths>
</Request>
```

The final request XML document should be as follows.

```xml
<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0"
         LibraryVersion="16.0.0.0" ApplicationName="Power Automate">
    <Actions>
        <Method Name="SetFieldValue" Id="6" ObjectPathId="2">
            <Parameters>
                <Parameter Type="String"><![CDATA[Status]]></Parameter>
                <Parameter Type="String"><![CDATA[Requested Approval]]></Parameter>
            </Parameters>
        </Method>
        <Method Name="SystemUpdate" Id="7" ObjectPathId="2" />
        <Method Name="SetFieldValue" Id="8" ObjectPathId="5">
            <Parameters>
                <Parameter Type="String"><![CDATA[LastActivity]]></Parameter>
                <Parameter Type="String"><![CDATA[Requested Approval]]></Parameter>
            </Parameters>
        </Method>
        <Method Name="SetFieldValue" Id="9" ObjectPathId="5">
            <Parameters>
                <Parameter Type="String"><![CDATA[ApprovalsRequested]]></Parameter>
                <Parameter Type="Int32"><![CDATA[1]]></Parameter>
            </Parameters>
        </Method>
        <Method Name="SystemUpdate" Id="10" ObjectPathId="5" />
    </Actions>
    <ObjectPaths>
        <StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />
        <Property Id="1" ParentId="0" Name="Web" />
        <Method Id="2" ParentId="1" Name="GetListItem">
            <Parameters>
                <Parameter Type="String">/sites/ModernTeam/@{triggerBody()?['{FullPath}']}</Parameter>
            </Parameters>
        </Method>
        <Property Id="3" ParentId="1" Name="Lists" />
        <Method Id="4" ParentId="3" Name="GetByTitle">
            <Parameters>
                <Parameter Type="String">Activity</Parameter>
            </Parameters>
        </Method>
        <Method Id="5" ParentId="4" Name="GetItemById">
            <Parameters>
                <Parameter Type="Int32">42</Parameter>
            </Parameters>
        </Method>
    </ObjectPaths>
</Request>
```
