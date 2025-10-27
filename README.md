# ITSS-PowerShell

Microsoft's goal for Windows PowerShell is to build 100% of a product's administrative functionality in PowerShell. Microsoft continues to build GUI consoles, but those consoles are executing PowerShell commands behind the scenes. This approach forces the company to make sure that every possible thing you can do with the product is accessible through PowerShell. If you need to automate a repetitive task or create a process that the GUI doesn't enable well, you can drop into PowerShell and take full control for yourself.

Several Microsoft products have already adopted this approach over the years, including Exchange, SharePoint, System Center products, Microsoft 365, Azure, and Windows Admin Center. Windows Server 2012 and higher are almost completely managed from PowerShell - or a GUI sitting on top of PowerShell. In other words, if you want to advance your career in IT, PowerShell is one of the best skills you can have in your toolbelt.

## Install

```Powershell
winget install --id Microsoft.Powershell --source winget
```

This command instructs winget to install the package with the ID `Microsoft.Powershell` from the default `winget` source, which corresponds to the latest stable version of PowerShell 7.

After the installation completes, you can launch PowerShell 7 from the Start Menu or by typing `pwsh` in a command prompt. Note that PowerShell 7 installs alongside Windows PowerShell (version 5.1) and does not replace it.

> For those of you who are familiar with the built-in Windows PowerShell ISE, you should be aware that *PowerShell 7 is not supported in ISE.*

### Recommended

Add your opid to the local admin group on your computer. This will allow you to run Powershell as an admin with your standard OPID, something normally blocked by group policy. This is not *bypassing* security as this is the method of allowing PowerShell to be used with a domain account recommended by EADS.

```PowerShell
Add-LocalGroupMember -Name 'Administrators' -Member ( Read-Host 'OPID' )
```

### VS Code

```Powershell
winget install -e --id Microsoft.VisualStudioCode
```

Microsoft Visual Studio Code is a code editor redefined and optimized for building and debugging modern web and cloud applications.

### Git

```PowerShell
winget install -e --id Git.Git
```
## Settng up Terminal

1. Open `Settings`.
2. Search for `Terminal`.
3. Select `Choose a terminal host app for command-line tools`.
4. Select `Windows Terminal` from the dropdown.

### Setting Powershell 7 as Default

1. Open `Settings (ctrl +,)`.
2. Set the default profile as `PowerShell 7`.
	1. *Note* PowerShell 7 must be installed before launching Terminal for it to appear as an option.

### Recommendations

- `Settings > Defaults > Appearance` Set font size to 16 (or higher).
- `Settings > Defaults > Appearance` Change to a Monospaced font.
	- [JetBrains Mono: A free and open source typeface for developers | JetBrains: Developer Tools for Professionals and Teams](https://www.jetbrains.com/lp/mono/)

## Setting up $PROFILE


```PowerShell
if ((-not (Test-Path -Path $PROFILE)) -and ($PSVersionTable.PSVersion.Major -GE 7)) {
	New-Item -Path $PROFILE -ItemType File -Force
} elseif ( $PSVersionTable.PSVersion.Major -LT 7 ) {
	Write-Host "Make sure to use PowerShell 7 when creating your profile." -ForegroundColor Yellow
} 
```

Make sure to run this in PowerShell, NOT WINDOWS POWERSHELL.

## SecretStore and SecretManagement

SecretStore and SecretManagement modules work together to provide a locally saved (meaning that it's saved on your computer, *not the network*) encrypted vault of secrets.

```Powershell
Install-Module Microsoft.PowerShell.SecretManagement
Install-Module Microsoft.PowerShell.SecretStore
```

### Create a Vault and Add a Secret

First you must register the vault. The `Name` parameter is a friendly name and can be any valid string.

```Powershell
Register-SecretVault -Name Credentials -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
```

If this is the first vault you've created, which should be the case for most of you, you will be prompted to create a password. Make this something memorable as if you forget this password you lose access to your vault forever and have to start over with a new one.

The `DefaultVault` parameter makes this the default vault. Now you can create a secret.

```PowerShell
Set-Secret -Name cred-a -Secret (Get-Credential)
```

Once again, the `Name` parameter is a friendly name and can be any valid string. I'm using `cred-a` in my vault for my -a credentials, but you can make it anything you want.

There are a couple of things to take note of here with how we're storing the secret. The first is that we're not simply saving our password as a string by typing it in directly; instead, we call `Get-Credential` which will prompt us for a username and password and save them as a credential object. The second is that when we call `Get-Credential`, we're doing so in parentheses, which tells the shell to resolve the `Get-Credential` call before resolving the entire command. There are a few reasons to do it this way:

1. By saving a credential object as the secret instead of a password, we can pass this object into other commands when we want to run something with our -a accounts. If we just saved the password, we would have to convert it to a credential object each time we want to use it.
2. Even though we'd be saving the password as an encrypted string, we would still have to type our password into the terminal as plain text to store it.

## Setting up VSCode

Go to the Plugins tab and search for: 

```
ms-vscode.powershell
```

## Recommended Settings
*`ctrl +,` to open settings*
1. Search `Format`.
2. Select `Powershell > Formatting` in the left pane.
3. Enable all options.
4. Choose `OTBS` for `Powershell > Code Formatting`.
5. Search `Format on save`.
6. Enable `Editor: Format On Save`.
7. *(Optional)* Enable `Editor: Format On Paste`.
8. Change to a Monospaced font for readability.
	- [JetBrains Mono: A free and open source typeface for developers | JetBrains: Developer Tools for Professionals and Teams](https://www.jetbrains.com/lp/mono/)

