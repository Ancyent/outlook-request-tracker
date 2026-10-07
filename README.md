# Request Tracker - Outlook add-in (v0)

Track client requests (tasks) and the supplier quote requests sent for them (subtasks), with follow-up dates and reminder drafts.

- Data is stored in the user's mailbox (Outlook roaming settings, ~32 KB). Nothing is sent to any server.
- Sideload: Outlook > Get add-ins > My add-ins > Custom add-ins > Add from file > `manifest.xml`.
- Rebuild the manifest for another host: `./build.sh https://host/path`.
