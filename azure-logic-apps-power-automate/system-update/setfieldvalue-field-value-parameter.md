# `SetFieldValue` Field Value Parameter <!-- omit in toc -->

As reviewed in the main article of this repository and in the SharePoint Client Query Request XML article of this repository, the `SetFieldValue` method is the way that one can trigger a particular field's value to be updated.

The second parameter of this method accepts the value for the field. The format of the second parameter for the field value is reviewed in [Section 3.2.5.87 of the Microsoft SharePoint Client-Side Object Model Protocol](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/4a06e0a3-2752-47b2-92f6-970b437d1e17). Also, a listing of different data types can be found in Microsoft's [Request XML Schema](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csom/2845342b-b15c-4870-baf8-3187870695b2) documentation in the [SharePoint Client Query Protocol standard](https://docs.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csom/fd645da2-fa28-4daa-b3cd-8f4e506df117). The types are listed under the simple type `MethodParameterTypeType`.

However, these references have not proved to provide enough documentation to fully determine the proper format that should be used in the request XML. This article therefore reviews the individual field types available and the proper format for the `SetFieldValue` method call in the request XML.

# Table of Contents <!-- omit in toc -->

- [Review of the `SetFieldValue` method's Request XML](#review-of-the-setfieldvalue-methods-request-xml)
- [Null values](#null-values)
- [Single line of text field](#single-line-of-text-field)
- [Multiple lines of text field (plain text)](#multiple-lines-of-text-field-plain-text)
- [Multiple lines of text field (rich text)](#multiple-lines-of-text-field-rich-text)
- [Choice field (single selection)](#choice-field-single-selection)
- [Choice field (multiple selections)](#choice-field-multiple-selections)
- [Yes/No field](#yesno-field)
- [Date/Time field](#datetime-field)
- [Lookup field (single selection)](#lookup-field-single-selection)
- [Lookup field (multiple selections)](#lookup-field-multiple-selections)
- [Person/Group field (single selection)](#persongroup-field-single-selection)
- [Person/Group field (multiple selections)](#persongroup-field-multiple-selections)

# Review of the `SetFieldValue` method's Request XML

```xml
<Method Name="SetFieldValue" Id="<!--ID-->" ObjectPathId="<!--OBJECT_PATH_ID-->">
    <Parameters>
        <Parameter Type="String"><![CDATA[<!--FIELD_INTERNAL_NAME-->]]></Parameter>
        <!--FIELD_VALUE-->
    </Parameters>
</Method>
```

The above snippet is reviewed in this repository's other articles regarding how and when to call the `SetFieldValue` method. This article will not review that same information. Rather, the sections below will focus on the XML snippet that should replace `<!--FIELD_VALUE-->`.

# Null values

For any field that is not required, a null value can be specified which will remove any data currently stored in that field. To specify a null value, use the parameter type `Null` with no value.

```xml
<Parameter Type="Null" />
```

# Single line of text field

The single line of text field value is represented in the XML as a string value. The parameter type is a `String`.

It is recommended to use the `<![CDATA[]]>` section around the string value to avoid needing to encode characters that have special meaning in XML documents: `<`, `>`, `&`, `'`, and `"`.

```xml
<Parameter Type="String"><![CDATA[This is a sample value.]]></Parameter>
```

# Multiple lines of text field (plain text)

The multiple lines of text field value is represented in the XML as a string value. The parameter type is a `String`.

It is recommended to use the `<![CDATA[]]>` section around the string value to avoid needing to encode characters that have special meaning in XML documents: `<`, `>`, `&`, `'`, and `"`.

Another benefit of the `<![CDATA[]]>` section is that it preserves all characters found within it including white space and new line characters. To insert new lines into plain text fields, include them in the value and they will be preserved in SharePoint using the syntax below. Remember *not* to indent text after new line characters unless you want the text value in SharePoint to be indented also.

```xml
<Parameter Type="String"><![CDATA[This is the first line.
This is the second line.

This is the fourth line. There is a blank line above this one.]]></Parameter>
```

# Multiple lines of text field (rich text)

The rich text field value is represented in the XML as a string value. The parameter type is a `String`.

It is recommended to use the `<![CDATA[]]>` section around the string value to avoid needing to encode characters that have special meaning in XML documents: `<`, `>`, `&`, `'`, and `"`. This is especially important for rich text fields as they often include HTML tags with these special characters.

To insert new lines into rich text fields, leverage standard HTML syntax to do so such as a `<br />` tag or by using block elements such as `p` or `div`.

```xml
<Parameter Type="String"><![CDATA[This is the first line.<br />This is the second line.<br /><br />This is the fourth line. There is a blank line above this one.]]></Parameter>
```

# Choice field (single selection)

The choice field value is represented in the XML as a string value. The parameter type is a `String`.

It is recommended to use the `<![CDATA[]]>` section around the string value to avoid needing to encode characters that have special meaning in XML documents: `<`, `>`, `&`, `'`, and `"`.

```xml
<Parameter Type="String"><![CDATA[Choice 1]]></Parameter>
```

# Choice field (multiple selections)

The multi-select choice field value is represented in the XML as a single string value. The parameter type is a `String`.

It is recommended to use the `<![CDATA[]]>` section around the string value to avoid needing to encode characters that have special meaning in XML documents: `<`, `>`, `&`, `'`, and `"`.

Each selected choice value is combined together as a single string by separating each choice value with the two characters `;#`

It is also valid to leave this value blank to de-select all choices or only supply a single choice value if that is desired.

```xml
<Parameter Type="String"><![CDATA[Choice 1;#Choice 2;#Choice 3]]></Parameter>
```

# Yes/No field

The boolean Yes/No field value is represented in the XML as an integer value. The parameter type is an `Int32`.

A `No` value is represented in the XML with a `0` value.

```xml
<Parameter Type="Int32">0</Parameter>
```

A `Yes` value is represented in the XML with a `1` value.

```xml
<Parameter Type="Int32">1</Parameter>
```

# Date/Time field

The date/time field value is represented by using a date/time in UTC with the [ISO 8601](https://en.wikipedia.org/wiki/ISO_8601) compatible format of `yyyy-MM-ddTHH:mm:ss.fffffffZ` where
- `yyyy` represents the four-digit year,
- `MM` represents the two-digit month,
- `dd` represents the two-digit day,
- `HH` represents the two-digit hour based on a 24-hour clock,
- `mm` represents the two-digit minute,
- `ss` represents the two-digit second, and
- `fffffff` represents the seven-digit fractional second.

The parameter type is a `DateTime`.

Note that if using Power Automate or Logic Apps, the [`formatDateTime` function](https://docs.microsoft.com/en-us/azure/logic-apps/workflow-definition-language-functions-reference#formatDateTime) will automatically convert any date/time value to the format outlined above. Ensure that value is in the UTC time zone though prior to conversion. SharePoint cannot handle any date/time value that is not in the UTC time zone.

A value of 12/31/2020 at 11:46:57 PM UTC would be represented in the XML as follows.

```xml
<Parameter Type="DateTime">2020-12-31T23:46:57.0000000Z</Parameter>
```

# Lookup field (single selection)

The lookup field value is represented in the XML as an integer value. The parmeter type is an `Int32`, though a `String` type also appears to work.

The value is the lookup list item's ID number for the desired selection.

```xml
<Parameter Type="Int32">12</Parameter>
```

# Lookup field (multiple selections)

The multi-select lookup field value is represented in the XML as a string value. The parameter type is a `String`.

A single lookup field value is represented using both the lookup ID number and lookup value concatenated with the two characters `;#`, such as this example: `12;#Honda`. Providing the lookup value is optional since the site will refer back to the lookup list for the value anyway. Therefore, each value can be simplified to have only its ID number from the lookup list and an empty lookup value, such as: `12;#`

After creating each value, these values are concatenated also using the two characters `;#`. Since the lookup values are being left blank for this field type, combining multiple values will result in repeating the `;#` sequence of characters twice in a row.

Consider the following example.

| Lookup ID | Lookup Value | Value for the XML Request (with the lookup value omitted) |
| --- | --- | --- |
| 12 | Honda | `12;#` |
| 1 | Acura | `1;#` |
| 30 | Toyota | `30;#` |
| 19 | Lexus | `19;#` |

Combining all of these values together with the `;#` characters in between each value produces the following string: `12;#;#1;#;#30;#;#19;#`

It is also valid to leave this value blank to de-select all choices or only supply a single person or group (e.g. `12;#`) if that is desired.

# Person/Group field (single selection)

The person/group (or user) field value is represented in the XML as an integer value. The parmeter type is an `Int32`, though a `String` type also appears to work.

The value is the site collection's ID number for the particular user or group. In the trigger, SharePoint does not provide user ID numbers by default. To obtain user ID numbers, please refer to [How to obtain a SharePoint User ID number](how-to-obtain-a-sharepoint-user-id-number.md).

```xml
<Parameter Type="Int32">42</Parameter>
```

# Person/Group field (multiple selections)

The multi-select person/group (or user) field value is represented in the XML as a string value. The parameter type is a `String`.

Recall that person/group fields are actually represented internally in SharePoint as lookup fields against the User Information List. This means that it follows a similar syntax as lookup fields.

A single lookup field value is represented using both the lookup ID number and lookup value concatenated with the two characters `;#`, such as this example: `42;#Steve Mayes`. Providing the lookup value is optional since the site will refer back to the User Information List anyway. Therefore, each value can be simplified to have only its ID number from the site collection's User Information List and an empty lookup value, such as: `42;#` In the trigger, SharePoint does not provide user ID numbers by default. To obtain user ID numbers, please refer to [How to obtain a SharePoint User ID number](how-to-obtain-a-sharepoint-user-id-number.md).

After creating each value, these values are concatenated also using the two characters `;#`. Since the lookup values are being left blank for this field type, combining multiple values will result in repeating the `;#` sequence of characters twice in a row.

Consider the following example.

| User / Group ID (Lookup ID) | Name (Lookup Value) | Value for the XML Request (with the lookup value omitted) |
| --- | --- | --- |
| 5 | Robyn Cranor | `5;#` |
| 42 | Steve Mayes | `42;#` |
| 50 | Kapiljith Rajappan | `50;#` |

Combining all of these values together with the `;#` characters in between each value produces the following string: `5;#;#42;#;#50;#`

It is also valid to leave this value blank to de-select all choices or only supply a single person or group (e.g. `42;#`) if that is desired.

```xml
<Parameter Type="String">5;#;#42;#;#50;#</Parameter>
```
