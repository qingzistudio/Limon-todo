# Reminder Gen

Chrome extension prototype for turning quick todo text into Microsoft To Do tasks.

The point is not to clone Microsoft To Do. The point is to avoid its UI friction when capturing a batch of tasks.

## What It Does

- Paste a messy todo list into the popup.
- Pick a due shortcut: `Today`, `Tomorrow`, `Week`, `Month`, or `Year`.
- Optionally choose an exact due date with the calendar field.
- Choose the reminder time; default is `18:00`.
- Mark a batch as priority to make tasks appear in Microsoft To Do's Important smart list.
- Each non-empty line becomes one task title. The text is not parsed for dates or times.
- Sign in to Microsoft and push tasks to one target Microsoft To Do list.
- Create the target list if it does not exist.

## Microsoft Setup

1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Click `Load unpacked`.
4. Select this directory.
5. Open the extension Options page, or click `Setup` in the popup.
6. Copy the shown Redirect URI.
7. In Microsoft Entra app registrations, create an app and copy its Application (client) ID.
8. Choose an account type that includes the Microsoft account you use for To Do.
9. Add the Redirect URI as a Single-page application redirect URI. Do not add it as a Web redirect URI.
10. Add Microsoft Graph delegated permissions: `User.Read` and `Tasks.ReadWrite`.
11. Paste the Application client ID into the setup page and sign in.

For personal Outlook/Hotmail Microsoft To Do accounts, use the `consumers` tenant preset. For work or school accounts, use `organizations` or the tenant ID listed in Microsoft Entra Overview as `Directory (tenant) ID`.
Do not use `common` with a consumer-only app registration. Microsoft rejects that exact combination with a `userAudience` error.

Do not buy Microsoft Entra ID P1/P2 for this prototype. Those paid plans are enterprise identity features. The extension only needs an app registration with a public client ID and delegated Graph permissions.

Run Microsoft sign-in from the setup page, not from the small Chrome action popup. Chrome closes action popups as soon as focus moves to the Microsoft login window.

## Supported Input

```text
- pay rent
- review Graph app registration #admin
- 整理报销单
- 周五下午三点半 给妈妈打电话
- renew passport
```

## Limits

- No backend.
- No two-way sync.
- No real duplicate detection yet. The popup disables push after success, but editing and pushing the same text again can still create duplicates.
- No review step. Edit the text before pushing.
- No recurring tasks yet.
- Timezone uses the current system timezone.
- Tasks get a reminder on their due date at the selected reminder time.
- Due-today tasks appear in My Day automatically. Tasks with due dates or reminders appear in Planned. Priority tasks use Graph `importance: high`, which To Do shows in Important.
- OAuth tokens are stored in Chrome extension local storage for this local prototype.
