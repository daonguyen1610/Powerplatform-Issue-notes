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
# CONCURRENT DYNAMIC
Concurrent(
    Set(glo_AccountsFirstItemID, First('2.1. Accounts').ID),
    Set(glo_AccountsLastItemID, First(SortByColumns('2.1. Accounts', "ID", SortOrder.Descending)).ID),
    Clear(col_AccountMaxMin)
);
ForAll(
    If(
        glo_AccountsLastItemID = glo_AccountsFirstItemID, 
        Sequence(1), 
        Sequence(RoundUp((glo_AccountsLastItemID -  glo_AccountsFirstItemID)/2000, 0))
    ),
    If(
        CountRows(col_AccountMaxMin) = 0,
        Collect(col_AccountMaxMin,{Min: glo_AccountsFirstItemID, Max: glo_AccountsFirstItemID-1+2000}),
        Collect(col_AccountMaxMin,{Min: Last(col_AccountMaxMin).Max+1, Max: 2000+Last(col_AccountMaxMin).Max})
    )
);
Concurrent(
    IfError(ClearCollect(col_Account1, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,1).Min && SubID <= Index(col_AccountMaxMin,1).Max)), Blank()),
    IfError(ClearCollect(col_Account2, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,2).Min && SubID <= Index(col_AccountMaxMin,2).Max)), Blank()),
    IfError(ClearCollect(col_Account3, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,3).Min && SubID <= Index(col_AccountMaxMin,3).Max)), Blank()),
    IfError(ClearCollect(col_Account4, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,4).Min && SubID <= Index(col_AccountMaxMin,4).Max)), Blank()),
    IfError(ClearCollect(col_Account5, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,5).Min && SubID <= Index(col_AccountMaxMin,5).Max)), Blank()),
    IfError(ClearCollect(col_Account6, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,6).Min && SubID <= Index(col_AccountMaxMin,6).Max)), Blank()),
    IfError(ClearCollect(col_Account7, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,7).Min && SubID <= Index(col_AccountMaxMin,7).Max)), Blank()),
    IfError(ClearCollect(col_Account8, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,8).Min && SubID <= Index(col_AccountMaxMin,8).Max)), Blank()),
    IfError(ClearCollect(col_Account9, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,9).Min && SubID <= Index(col_AccountMaxMin,9).Max)), Blank()),
    IfError(ClearCollect(col_Account10, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,10).Min && SubID <= Index(col_AccountMaxMin,10).Max)), Blank()),
    IfError(ClearCollect(col_Account11, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,11).Min && SubID <= Index(col_AccountMaxMin,11).Max)), Blank()),
    IfError(ClearCollect(col_Account12, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,12).Min && SubID <= Index(col_AccountMaxMin,12).Max)), Blank()),
    IfError(ClearCollect(col_Account13, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,13).Min && SubID <= Index(col_AccountMaxMin,13).Max)), Blank()),
    IfError(ClearCollect(col_Account14, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,14).Min && SubID <= Index(col_AccountMaxMin,14).Max)), Blank()),
    IfError(ClearCollect(col_Account15, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,15).Min && SubID <= Index(col_AccountMaxMin,15).Max)), Blank()),
    IfError(ClearCollect(col_Account16, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,16).Min && SubID <= Index(col_AccountMaxMin,16).Max)), Blank()),
    IfError(ClearCollect(col_Account17, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,17).Min && SubID <= Index(col_AccountMaxMin,17).Max)), Blank()),
    IfError(ClearCollect(col_Account18, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,18).Min && SubID <= Index(col_AccountMaxMin,18).Max)), Blank()),
    IfError(ClearCollect(col_Account19, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,19).Min && SubID <= Index(col_AccountMaxMin,19).Max)), Blank()),
    IfError(ClearCollect(col_Account20, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,20).Min && SubID <= Index(col_AccountMaxMin,20).Max)), Blank()),
    IfError(ClearCollect(col_Account21, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,21).Min && SubID <= Index(col_AccountMaxMin,21).Max)), Blank()),
    IfError(ClearCollect(col_Account22, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,22).Min && SubID <= Index(col_AccountMaxMin,22).Max)), Blank()),
    IfError(ClearCollect(col_Account23, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,23).Min && SubID <= Index(col_AccountMaxMin,23).Max)), Blank()),
    IfError(ClearCollect(col_Account24, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,24).Min && SubID <= Index(col_AccountMaxMin,24).Max)), Blank()),
    IfError(ClearCollect(col_Account25, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,25).Min && SubID <= Index(col_AccountMaxMin,25).Max)), Blank()),
    IfError(ClearCollect(col_Account26, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,26).Min && SubID <= Index(col_AccountMaxMin,26).Max)), Blank()),
    IfError(ClearCollect(col_Account27, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,27).Min && SubID <= Index(col_AccountMaxMin,27).Max)), Blank()),
    IfError(ClearCollect(col_Account28, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,28).Min && SubID <= Index(col_AccountMaxMin,28).Max)), Blank()),
    IfError(ClearCollect(col_Account29, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,29).Min && SubID <= Index(col_AccountMaxMin,29).Max)), Blank()),
    IfError(ClearCollect(col_Account30, Filter('2.1. Accounts', SubID >= Index(col_AccountMaxMin,30).Min && SubID <= Index(col_AccountMaxMin,30).Max)), Blank())
);
ClearCollect(
    col_AccountsAllItem, 
    col_Account1, col_Account2, col_Account3, col_Account4, col_Account5, col_Account6, col_Account7, col_Account8, col_Account9, col_Account10,
    col_Account11, col_Account12, col_Account13, col_Account14, col_Account15, col_Account16, col_Account17, col_Account18, col_Account19, col_Account20,
col_Account21, col_Account22, col_Account23, col_Account24, col_Account25, col_Account26, col_Account27, col_Account28, col_Account29, col_Account30
);
