# Modernize SharePoint Applications

This repository contains generally applicable solutions and patterns to modernize SharePoint on-premises customizations for hosting in Microsoft 365 and Azure.

# Note regarding Azure Logic Apps and Power Automate

Power Automate is built upon the Azure Logic Apps platform. Power Automate is targeted for Microsoft 365 users while Azure Logic Apps is targeted as a fully-featured Azure serverless platform-as-a-serivce (PaaS) offering. Connectors are typically available for both Azure Logic Apps and Power Automate, with few exceptions. Both platforms share a common set of expression syntax and functions. Therefore, most of the content in this repository will refer only to Azure Logic Apps for simplicity and brevity. When Azure Logic Apps is referenced, one can safely assume that the reference applies to both Azure Logic Apps and Power Automate interchangably. In cases where this is untrue, that will be noted as needed.

# Solutions and Patterns

- Azure Logic Apps and Power Automate
  - Approvals: Replacing SharePoint 2010 and 2013 approval workflows with modern approvals (applies to Power Automate only)
  - [Print to PDF: Enabling PDFs to be generated from list item data to replace custom-built printing functionality in on-premises SharePoint applications](azure-logic-apps-power-automate/print-to-pdf/README.md)
  - [Starting Instant Flows: Triggering Power Automate instant flows from SharePoint list views (applies to Power Automate only)](azure-logic-apps-power-automate/starting-instant-flows/README.md)
  - [`SystemUpdate()`: Using the `SystemUpdate()` CSOM method in Azure Logic Apps](azure-logic-apps-power-automate/system-update/README.md)
  - [Task Processes: Modernizing SharePoint 2010 Task Process workflows](azure-logic-apps-power-automate/task-processes/README.md)
- Power Apps
  - [Choice Fields: Enabling users to add values manually](power-apps/choice-fields/README.md)
  - Form Designs: Comparing multi-screen (with back/next buttons), single-page scrolling, and tab designs
  - Moving SharePoint list form based Power Apps between environments
  - [People Picker Best Practices](power-apps/people-picker-best-practices)
  - [Using the current list item when an in-list form opens](power-apps/using-current-item-when-in-list-form-opens/README.md)
  - [Using SharePoint Security: Obtaining security information from SharePoint sites to enable changing the logic of a Power App based on the current user's role or security group membership](power-apps/using-sharepoint-security/README.md)
- SharePoint Online List Forms
  - Cascading Drop Downs: Using Managed Metadata fields as a replacement for cascading drop down fields
  - Customization: Using JSON and formulas to configure custom headers, footers, field grouping, and conditional visibility
  -
# Load more than 10k items 

Clear(gloAllItems);
With(
    {
        wSets: With(
            {
                wLimits: With(
                    {
                        wLimit: Sort(Filter(ContractMasterData, Ticket_Contract = "Ticket", Status = "Review by Legal", LegalEmail = User().Email),SubID,Descending)
                    },
                    RoundDown((First(wLimit).SubID) / 2000,0) + 1
                )
            },
            Set(test, wLimits);
            AddColumns(
                RenameColumns(
                    Sequence(
                        wLimits,
                        0,
                        2000
                    ),
                    "Value",
                    "LowID"
                ),
                "HighID",
                LowID + 2000
            )
        )
    },
    Set(test1,wSets);
    ForAll(
        wSets As MaxMin,
        Collect(
            gloAllItems,
            Filter(
                ContractMasterData,
                SubID > MaxMin.LowID && SubID <= MaxMin.HighID,
                Ticket_Contract = "Ticket",
                Status = "Review by Legal",
                LegalEmail = User().Email
            )
        )
    )
);
ClearCollect(gloGalleryItems, gloAllItems)
# Define global variable for display & hide variable
Clear(gloAllItems);
With(
    {
        wSets: With(
            {
                wLimits: With(
                    {
                        wLimit: Sort(
                            Contacts,
                            SubID,
                            Descending
                        )
                    },
                    RoundDown(
                            (First(wLimit).SubID) / 2000,
                            0
                    ) + 1
                )
            },
            Set(test, wLimits);
            AddColumns(
                RenameColumns(
                    Sequence(
                        wLimits,
                        0,
                        2000
                    ),
                    "Value",
                    "LowID"
                ),
                "HighID",
                LowID + 2000
            )
        )
    },
    Set(test1,wSets);
    ForAll(
        wSets As MaxMin,
        Collect(
            gloAllItems,
            Filter(
                Contacts,
                SubID > MaxMin.LowID && SubID <= MaxMin.HighID
            )
        )
    )
);
ClearCollect(gloDisplayAllItems, gloAllItems);
# Load 10k items using created
ClearCollect(test, TestOL2);
UpdateContext({CountItems:  (First(SortByColumns(TestOL2, "ID", SortOrder.Descending)).ID - First(TestOL2).ID)/2000});
ForAll(
    Sequence(Round(CountItems, 0)),
    Collect(test, Filter(TestOL2, Created >= Last(test).Created))
);
ClearCollect(test1, ForAll(Distinct(test, LookUp(test, ID = ThisRecord.ID)), Value));
# Check blank for combobox
Filter(
    ContactsResend, 
    StartsWith(Name, TextInput5_6.Text) || StartsWith(Code, TextInput5_6.Text) || StartsWith(CustomerName, TextInput5_6.Text),
    ComboBox3_7.Selected.Result = Status || IsBlank(ComboBox3_7.Selected.Result)
)

thay vì 
If(
IsBlank(ComboBox3_7.Selected.Result),
ContactsResend,
Filter(
    ContactsResend, 
    StartsWith(Name, TextInput5_6.Text) || StartsWith(Code, TextInput5_6.Text) || StartsWith(CustomerName, TextInput5_6.Text),
    ComboBox3_7.Selected.Result = Status || IsBlank(ComboBox3_7.Selected.Result)
)
)
# RELOAD PAGE 
UpdateContext({gloRefresh:false});UpdateContext({gloRefresh:true}) ON onvisible 

Filter(
   ErrorCustomer,
   ComboBox3_4.Selected.Value = Title || IsBlank(ComboBox3_4.Selected.Value),
   StartsWith(CustomerName, TextInput5_3.Text) || StartsWith(Email, TextInput5_3.Text),
   Created >= DatePicker1.SelectedDate && Created <= DatePicker1_1.SelectedDate,
   **gloRefresh**
)
# UPDATE MUTIL ROW
Set(
    varSeq,
    RoundUp(
        CountRows(
            Filter(
                'Activity Driver Mappings',
                'Activity Driver ID'.'Activity Driver Name' = DataCardValue7.Text
            )
        ) / 2000,
        0
    )
);
ForAll(
    Sequence(varSeq),
    UpdateIf(
        'Activity Driver Mappings',
        'Activity Driver ID'.'Activity Driver Name' = DataCardValue7.Text,
        {'Activity Driver ID': ComboBox1.Selected}
    );
    Refresh('Activity Driver Mappings');
)
# GET USER FROM AD USER OR GROUP
UpdateContext({locGetMemberID: MicrosoftEntraID.GetUser(cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.UserPrincipalName).id});
If(
    !IsBlank(cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.SelectedItems) And IsBlank(
        LookUp(
            colAppUserPermission,
            mail = cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail
        ).mail
    ) And IsBlank(
        LookUp(
            colAddMember,
            mail = cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail
        ).mail
    ),
    IfError(
        MicrosoftEntraID.AddUserToGroup(
            gblSecurityGroup,
            locGetMemberID
        ),
        Notify(
            If(
                !IsBlank(
                    First(
                        Filter(
                            colDvLocalizations,
                            Key = "Error Adding Notify Msg" And Language = gblLanguage
                        )
                    ).Translation
                ),
                First(
                    Filter(
                        colDvLocalizations,
                        Key = "Error Adding Notify Msg" And Language = gblLanguage
                    )
                ).Translation,
                First(
                    Filter(
                        colDvDefaultLocalizations,
                        Key = "Error Adding Notify Msg"
                    )
                ).Translation
            ) & " " & cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail & ". " & If(
                !IsBlank(
                    First(
                        Filter(
                            colDvLocalizations,
                            Key = "Contact Site Super Admin Notify Msg" And Language = gblLanguage
                        )
                    ).Translation
                ),
                First(
                    Filter(
                        colDvLocalizations,
                        Key = "Contact Site Super Admin Notify Msg" And Language = gblLanguage
                    )
                ).Translation,
                First(
                    Filter(
                        colDvDefaultLocalizations,
                        Key = "Contact Site Super Admin Notify Msg"
                    )
                ).Translation
            ),
            NotificationType.Error,
            3000
        ),
        Notify(
            cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail & " " & If(
                !IsBlank(
                    First(
                        Filter(
                            colDvLocalizations,
                            Key = "Add Member Notify Msg" And Language = gblLanguage
                        )
                    ).Translation
                ),
                First(
                    Filter(
                        colDvLocalizations,
                        Key = "Add Member Notify Msg" And Language = gblLanguage
                    )
                ).Translation,
                First(
                    Filter(
                        colDvDefaultLocalizations,
                        Key = "Add Member Notify Msg"
                    )
                ).Translation
            ),
            NotificationType.Success,
            1000
        );
        Collect(
            colAddMember,
            {
                mail: cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail,
                displayName: cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.DisplayName,
                userPrincipalName: cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.UserPrincipalName
            }
        )
    ),
    If(
        !IsBlank(
            LookUp(
                colAppUserPermission,
                mail = cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail
            ).mail
        ) Or !IsBlank(
            LookUp(
                colAddMember,
                mail = cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail
            ).mail
        ),
        Notify(
            cmb_SiteSuperAdminSettings_PermissionsSearchNewMember.Selected.Mail & " " & If(
                !IsBlank(
                    First(
                        Filter(
                            colDvLocalizations,
                            Key = "Member Exist Notify Msg" And Language = gblLanguage
                        )
                    ).Translation
                ),
                First(
                    Filter(
                        colDvLocalizations,
                        Key = "Member Exist Notify Msg" And Language = gblLanguage
                    )
                ).Translation,
                First(
                    Filter(
                        colDvDefaultLocalizations,
                        Key = "Member Exist Notify Msg"
                    )
                ).Translation
            ),
            NotificationType.Warning,
            2000
        )
    )
);
Reset(cmb_SiteSuperAdminSettings_PermissionsSearchNewMember);
ClearCollect(
    colADGroupMembers,
    Office365Groups.ListGroupMembers(gblSecurityGroup).value
);
