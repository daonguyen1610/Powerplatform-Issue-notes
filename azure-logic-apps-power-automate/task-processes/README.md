# Modernize SharePoint 2010 Task Process Workflows to Azure Logic Apps and Power Automate

In SharePoint 2010 workflows, users can create task processes. These task processes are akin to workflows within a workflow. It allows users to create a process and specify actions that should happen regarding that process and each generated task within the process during its lifecycle. There is no direct functionality in Azure Logic Apps or Power Automate which replaces task processes. Rather, similar functionality can be achieved using different (non-SharePoint) actions or by using multiple Logic Apps or Flows against multiple lists.

Keep it simple. Often times, a user was only leveraging pieces and parts of the available functionality. Therefore, a simpler solution may exist for any given workflow. Do not default to using the one size fits all solution. Instead, consider the lowest complex available solution that meets the **business need** the workflow is solving. The technical implementation may vary widely.

# Approvals

Consider using the [Approvals connector](https://docs.microsoft.com/en-us/connectors/approvals/), which is a standard connector. Power Automate also adds functionality for users by giving them a separate section to see all pending approvals awaiting their review.

# Feedback

For feedback oriented workflows, consider if any of the following may fit the need:

- [Microsoft Forms](https://forms.microsoft.com), with a Logic App or Flow as needed to follow up on responses
- A simple SharePoint list created specifically for the purpose of soliciting feedback, with a Logic App or Flow as needed to follow up on responses

# Custom

Custom task processes can incorporate any number of actions based on multiple different events and stages depending on the configuration of the process. Again, **keep it simple**. Do not implement more than is required to meet the **business need** achieved with the current workflow.

Below, the entire task process will be reviewed with implementation notes for how to achieve each step of the process using Logic Apps and Flows. After that review, a real-world example will be provided that will not require all of the implementation steps outlined during the initial review to demonstrate how one should keep the implementation as tailored as possible to the actual need.

The task process is represented as a single action in a SharePoint 2010 workflow. It is a lot more than that however! It encompasses an entire process and workflow on its own.

When a task process is created, the following configuration is solicited from the workflow author:

- Participants (users)
- One at a time (serial) vs. All at the same time (parallel)
- CC (carbon copied users)
- Title
- Instructions
- Duration per Task
- Due Date for Task Process

The task process is then started. There are a total of four stages triggered by events in the task process lifecycle where any number of workflow actions can be added.

- **When the task process starts**: These actions run immediately after the main workflow process and before the task process.
- **When the task process is running**: These actions run before the task process assigns its first task.
- **When the task process is canceled**: These actions run if the task process is cancelled.
- **When the task process completes**: These actions run when the last individual task is complete or when the end task process action is run.

*Reference: [Use the task process editor for approval workflows](https://support.microsoft.com/en-us/office/use-the-task-process-editor-for-approval-workflows-8680b4a4-36b1-441c-b070-e515976078aa)*

Each task process generates one or more tasks, depending on the configuration set by the workflow author. Multiple tasks are generated if multiple participants are involved with a serial task process. Otherwise, only a single task is generated. Each task has five stages triggered by events in each task's lifecycle where any number of workflow actions can be added.

- **Before a task is assigned**: These actions run before every individual task is created.
- **When a task is pending**: These actions run after every individual task has been created.
- **When a task expires**: These actions run every time an individual task is incomplete after its due date.
- **When a task is deleted**: These actions run every time an individual task is deleted before it's completed.
- **When a task completes**: These actions run every time an individual task is complete.

*Reference: [Use the task process editor for approval workflows](https://support.microsoft.com/en-us/office/use-the-task-process-editor-for-approval-workflows-8680b4a4-36b1-441c-b070-e515976078aa)*

Two key factors will determine the necessary complexity of the solution and drive the technical implementation details.

1) The number of tasks the task process generates.
2) The number of stages which contain any workflow actions.

Let's review three scenarios. To simplify the terminology, the word **Flow** is used generically to refer to either a Logic App or a Flow.

## Scenario 1 - Highest complexity, multiple tasks, all stages have actions

For the first scenario, assume the task process generates multiple tasks and that there are actions in each of the nine stages. This scenario carries the highest level of complexity.

For ease of maintainence and design, leverage two separate lists - a "task process" list and a "task" list. The task process list will contain a single list item per task process. The task list will contain a single list item per task generated by each task process.

The main list should have one column added to it:
- Currently Running Task Process Item ID: A numeric column which will store the ID of the currently running task process from the task process list.

The task process list should contain the following columns:

- Main List Item ID: A numeric column which will store the ID of the item from the main list.
- Status: A choice column with the values Running, Canceled, and Completed. Defaults to Running.
- Participants: A multiple user column.
- Serial or Parallel: A choice column with the values Serial and Parallel.
- CC: A multiple user column.
- Title: A single line of text column.
- Instructions: A plain text, multiple lines of text column.
- Duration per Task: A numeric column representing the duration of the each task in any unit desired. (minutes, hours, days, weeks, etc.) The unit should make sense to the Flows developed for this need.
- Due Date for Task Process: A date and time column.

The task list should contain the following columns:

- Main List Item ID: A numeric column which will store the ID of the item from the main list.
- Task Process List Item ID: A numeric column which will store the ID of the item from the task process list.
- Outcome: A choice column with the desired outcomes. (e.g. Approve, Reject, etc.)
- Status: A choice column with the values Pending, Expired, Deleted, and Completed.
- Expiration Date: A date and time column derived from the task process using the Duration per Task and Due Date for Task Process.

Once the main list, the task process list, and the task list are created, configure Flows against these lists as follows.

- **Main Flow**: This is the Flow that is executing against the desired list item on the main list. This is the Flow that is replacing the SharePoint Designer workflow which contains the task process. This Flow should include a single action which creates the list item in the task process list as a replacement for the task process creation in SharePoint Designer. Once that item is created initially, the Flow should stop its execution. Upon the main list item being modified, the Flow should check to see if the task process item has finished (Status is anything other than Running) before continuing on with the remainder of the logic within the SharePoint Designer workflow. If the task process item still has a Status of Running, then exit the Flow without executing any actions.
- **Task Process Flow**: Create a Flow against the task process list that triggers whenever a list item is created or modified.
  - When a task process item is created,
    - Execute the actions from the **When the task process starts** phase.
    - Execute the actions from the **When the task process is running** phase.
    - Create the first task item in the task list for this process based on the configuration in the task process item.
  - When a task process item is modified,
    - If the Status is equal to Canceled,
      - Execute the actions from the **When the task process is canceled** phase.
    - If the Status is equal to Completed,
      - Execute the actions from the **When the task process completes** phase.
    - If the Status is equal to Canceled or Completed,
      - Update the main list item without changing any field values, so that the Main Flow will be triggered to run again.
- **Task Flow**: Create a Flow against the task list that triggers whenever a list item is created or modified.
  - When a task item is created,
    - Execute the actions from the **Before a task is assigned** phase.
    - Execute the actions from the **When a task is pending** phase.
  - When a task item is modified,
    - If the Outcome is not empty,
      - Set the Status to Completed.
    - If the Status is equal to Expired,
      - Execute the actions from the **When a task expires** phase.
    - If the Status is equal to Deleted,
      - Execute the actions from the **When a task is deleted** phase.
    - If the Status is equal to Completed,
      - Execute the actions from the **When a task completes** phase.
    - If the Status is equal to Expired, Deleted, or Completed,
      - Create the next task per the task process configuration **or** if there is no other task that needs to be created, change the Status of the task process item to Completed.
- **Task Expiration Flow**: Create a second Flow which will operate against the task list, but is set to execute on a schedule between once every hour and once every day, depending on how quickly tasks should be expired after their expiration dates.
  - Query the task list for all task items that have a status of Pending with an expiration date before the current date and time.
  - For all of those list items, update each one to have a Status of Expired.

## Scenario 2 - Medium complexity, single task, all stages have actions

For the second scenario, assume that the task process generates only a single task and that there are actions in each of the nine stages. This scenario is easier than the first scenario, as the task process and task tracking can be consolidated into a single list and Flow.

The main list should have one column added to it:
- Currently Running Task Item ID: A numeric column which will store the ID of the currently running task from the task list.

There will be no task process list as there is in Scenario 1.

The task list should contain the following columns:

- Main List Item ID: A numeric column which will store the ID of the item from the main list.
- Outcome: A choice column with the desired outcomes. (e.g. Approve, Reject, etc.)
- Status: A choice column with the values Pending, Expired, Canceled, Deleted, and Completed.
- Participants: A multiple user column.
- Serial or Parallel: A choice column with the values Serial and Parallel.
- CC: A multiple user column.
- Title: A single line of text column.
- Instructions: A plain text, multiple lines of text column.
- Expiration Date: A date and time column.

Once the main list and the task list are created, configure Flows against these lists as follows.

- **Main Flow**: This is the Flow that is executing against the desired list item on the main list. This is the Flow that is replacing the SharePoint Designer workflow which contains the task process. This Flow should include a single action which creates the list item in the task list as a replacement for the task process creation in SharePoint Designer. Once that item is created initially, the Flow should execute the actions from the **When the task process starts** and **When the task process is running** phases, then stop its execution. Upon the main list item being modified, the Flow should check to see if the task item has finished (Status is anything other than Pending) before continuing on with the remainder of the logic within the SharePoint Designer workflow. If the task item still has a Status of Pending, then exit the Flow without executing any actions. Once the task has finished, then the Flow should execute the actions from the **When the task process completes** phase. If the task is cancelled, then the Flow should execute the actions from the **When the task process is canceled** phase.
- **Task Flow**: Create a Flow against the task list that triggers whenever a list item is created or modified.
  - When a task item is created,
    - Execute the actions from the **Before a task is assigned** phase.
    - Execute the actions from the **When a task is pending** phase.
  - When a task item is modified,
    - If the Outcome is not empty,
      - Set the Status to Completed.
    - If the Status is equal to Expired,
      - Execute the actions from the **When a task expires** phase.
    - If the Status is equal to Deleted,
      - Execute the actions from the **When a task is deleted** phase.
    - If the Status is equal to Completed,
      - Execute the actions from the **When a task completes** phase.
    - If the Status is equal to Expired, Deleted, or Completed,
      - Update the main list item without changing any field values, so that the Main Flow will be triggered to run again.
- **Task Expiration Flow**: Create a second Flow which will operate against the task list, but is set to execute on a schedule between once every hour and once every day, depending on how quickly tasks should be expired after their expiration dates.
  - Query the task list for all task items that have a status of Pending with an expiration date before the current date and time.
  - For all of those list items, update each one to have a Status of Expired.

Note that Scenario 2 is a subset of Scenario 1 while still retaining all of the functionality of any actions that were in place on the phases in the SharePoint Designer task process. Scenario 3 will further decrease the complexity to show how the most generic framework can be tailored based on the actions present for a particular task process.

## Scenario 3 - Low complexity, single task, only some stages have actions

In this scenario, let's take a real-world example. Consider a task process that only assigns a task to a single individual. Furthermore, there are only actions in the following phases:

- When a task is pending
- When a task expires
- When a task completes

This is much closer to Scenario 2, so use that as a starting point. Also note that if there is an existing task list, it can likely be reused.

Based on the individual actions that needed to be implemented for this scenario, the main list and task list were configured as follows.

- The Main Flow had no need to access the task list item. Therefore, the suggested Currently Running Task Item ID column was not added to the main list.
- Since the task process is only assigning a single task to a single individual, the task list does not need to track Participants, Serial or Parallel, or CC. Those columns were omitted.
- Also, since the task process provided no instructions, that column was omitted as well.
- The existing task list was reused, adding three columns: Main List Item ID, Outcome, and Expiration Date. The existing Status column was adjusted to have the expected values of Pending, Expired, Canceled, Deleted, and Completed.
- The content type of the task list was adjusted to hide fields that should not be shown to the end user performing the approval or rejection. Only the following columns needed to be shown on the form: Title, Assigned To, Outcome, and Expiration Date.

Now that the main list and the task list are configured, three Flows were implemented.

- **Main Flow**: This flow contains all of the logic in the main workflow in SharePoint Designer. At the desired point in the Flow, a new list item was created in the task list representing the task to be executed. The Flow then stops executing. When it is executed again, it will check for a desired status in the main list item before continuing its actions.
- **Task Flow**: This flow contains the **When a task is pending** and **When a task completes** actions per Scenario 2 above. All other branches are excluded since they are not needed given that there are no actions in those phases.
- **Task Expiration Flow**: The **When a task expires** actions were moved into this Flow to simplify the implementation.

# References

- [Microsoft Support: Use the task process editor for approval workflows](https://support.microsoft.com/en-us/office/use-the-task-process-editor-for-approval-workflows-8680b4a4-36b1-441c-b070-e515976078aa)
- [The SharePoint 2010 Task Process Designer by Laura Rogers, a highly regarded Microsoft MVP with over 10 years experience with SharePoint](https://wonderlaura.com/2011/06/16/the-sharepoint-2010-task-process-designer/)
