# OutlookProfileRemoval
Prepare Outlook profile for autodiscover and log the migration to Office 365.

### Goal of this script
When a mailbox has been migrated to Exchange Online (or Office 365), Outlook must be reconfigured to point to the new Exchange servers. This is exactly what does this VBScript, by :

1. remove all existing MAPI profiles found in the user profile
2. create a brand new profile, and restart Outlook in "autodiscover mode"
3. log the date/time of the migration in the user's profile registry and in the event log to be able to follow the migration. Use your favorite management system to centralize the events.

The script take the CPU architecture into account so it will open the correct version of Outlook. In this case, we are focusing on Outlook 2010. Feel free to adapt the Outlook path to your environment.

### How to use the script ?
Deploy it with a GPO defined in your Active Directory. Be sure the users can read the *.PRF file !
