# Claude Mail Sorter for Outlook

Uses Claude to file emails into folders, catch spam, flag important mail, and apply categories. Sister project of the Thunderbird add-on, Claude Mail Sorter.

Two versions:

- **`outlook-vsto/`**: classic Outlook for Windows, all account types (including IMAP), with new-mail sorting, folder right-click and a spinning ribbon icon. See its README.
- **`outlook/`**: Office add-in for new Outlook, Outlook on the web and Mac (details below).

## Office add-in (Outlook.com / Microsoft 365)

Adds a **Sort with Claude** ribbon button that opens a task pane to sort the open email or a whole folder. It uses Microsoft Graph to move mail, flag it and set categories.

Differences from the Thunderbird add-on: no automatic sorting of new mail, no folder right-click item, no undo, and the spinner is in the task pane (Outlook ribbon icons can't animate).

Setup:

1. In the Azure portal, register an app (any account type, **Single-page application** redirect URI `brk-multihub://localhost:3000`) with delegated Graph permissions `Mail.ReadWrite` and `MailboxSettings.ReadWrite`.
2. Put its client ID in `outlook/taskpane.js` (`CLIENT_ID`) and `outlook/manifest.xml` (replace every `00000000-0000-0000-0000-000000000000`).
3. Run `npx office-addin-dev-certs install`, then `node outlook/serve.mjs`.
4. In Outlook, open **Get Add-ins → My add-ins → Add a custom add-in → Add from file** and pick `outlook/manifest.xml`.
5. Open the task pane and paste your Anthropic API key.

To host it elsewhere, replace `https://localhost:3000` in the manifest with your HTTPS URL.

## Privacy and cost

- Each sorted email (headers and up to 20,000 characters of body text) is sent to the Anthropic API.
- Each email is one API call. The default model is Claude Haiku 4.5.
- Claude can only choose folders that exist (or categories you allow it to create), so an email cannot instruct it to do anything else.

## Status

Neither version has been run yet; both were written without access to Outlook.
