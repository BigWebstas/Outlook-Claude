# Claude Mail Sorter for classic Outlook (VSTO)

C# add-in for Windows desktop Outlook (2016 and later, including classic Outlook in Microsoft 365). It works with any account type, including IMAP.

## What it does

- **Ribbon button** on the Mail tab: sorts the selected emails. The icon spins and the label reads "Claude: working…" while anything is being processed.
- **Right-click an email** → *Sort with Claude*. **Right-click a folder** → *Sort folder with Claude*.
- **Sort Inbox now** and **Claude settings** buttons in the same ribbon group.
- Optional: sort new mail automatically, and sort the Inbox when Outlook starts (up to 200 of the newest unsorted emails).
- Claude can mark spam (moved to Junk), flag important mail, apply categories, and move mail into an existing folder. It never files into Deleted Items, Junk, Sent, Drafts or Outbox.

Not included compared with the Thunderbird add-on: activity log, undo, and the OTP cleanup.

## Build

On Windows, from the `Outlook-Claude` folder in PowerShell:

```powershell
.\build.ps1            # build only
.\build.ps1 -Install   # build, trust the certificate and register with Outlook (close Outlook first)
```

Needs Visual Studio 2019/2022 (Community is fine) or Build Tools with the **Office/SharePoint development** workload, plus desktop Outlook. The script signs the add-in with a self-signed certificate it creates for your Windows user.

Prefer Visual Studio? Open `ClaudeMailSorter.csproj` and press F5.

## Use

Open **Claude settings** on the ribbon, paste your Anthropic API key and choose the options. The key is encrypted for your Windows user and stored in `%AppData%\ClaudeMailSorter`.

## Untested

The C# code compiles against the Outlook interop library, but the Visual Studio project file, `ThisAddIn.Designer.cs` and `build.ps1` were written without a Windows machine and have never been run. Expect to fix small build errors on first try. Two things worth checking: the folder right-click item should sort the folder you clicked (it falls back to the open folder if Outlook doesn't pass it), and the ribbon label and icon should update while sorting.
