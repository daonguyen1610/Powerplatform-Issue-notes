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
