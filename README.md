# Windows-Taskbar-Management-Script
Manage Windows 11 Taskbar

Considering App Removal

I'm now focusing on implementing the "remove apps" functionality. This involves managing the "Custom Apps" list, which is likely stored in JSON or memory. I'm distinguishing between removing custom entries and potentially hiding default apps, with removal mainly targeting the custom ones created by the user.

Implementing Removal Logic

I've added the "Remove Item" button and the initial logic. The system now checks for selection and differentiates between custom and default apps. For customs, the corresponding entry is removed from $Global:CustomApps, the JSON file is updated, and the item is removed from the listbox. I'm focusing solely on removing custom apps, as planned. Dealing with default app removal would add complexity I want to avoid for the current iteration.

Developing Removal Workflow

I'm now writing the script, integrating the removal logic into the existing features. I'm focusing on adding the Remove-App logic and button, and updating Save-Config. The refined logic I've developed gets the selected index and app name. It then checks if the app is custom, and if so, removes it from the array, saves the config, and removes it from the listbox. If not, it displays a "Cannot remove built-in apps" message.

Yes, we can add a "Remove App" button. This is useful if you add a custom app by mistake or no longer want it in your list.

What is new in this version?
Remove Button: A new button allows you to delete items you added via "Browse".

Note: It only deletes Custom Apps (the ones saved in the JSON file). It prevents you from deleting the core "Built-in" apps (like Edge or Word) so the script remains usable.

Smart Saving: When you remove an app, it immediately updates the saved_apps.json file, so the app is gone for good.
