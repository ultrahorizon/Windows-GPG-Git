# Windows GPG/Git Setup with YubiKey

Internally, Ultra Horizon and sister companies use YubiKeys to hold GPG keys
used for code signing, SSH and more. This repo collates some notes on getting a
GPG/Git environment setup in Windows, as well as a script that automates all
the steps.

## Script Usage

If running on a new machine, you may need to set the execution policy from an
elevated PowerShell:

```
Set-ExecutionLevel Unrestricted
```

Then run the script:

```
.\Setup-GPG-Git.ps1 --GitUserName "Bob Yards" --GitUserEmail bob.yards@ultra-horizon.com
```

Don't forget to reset the execution level (if appropriate).

---

## Notes

### Install dependencies with WinGet
```
winget install -e --id Git.Git
winget install -e --id GnuPG.Gpg4win
```
Check if SSH is installed, if not set up OpenSSH using [Windows optional features](https://learn.microsoft.com/en-us/windows-server/administration/openssh/openssh_install_firstuse).

### Ensure default Windows OpenSSH auth agent is not running and disabled

Requires elevated terminal
```
Stop-Service ssh-agent
Set-Service ssh-agent -StartupType Disabled
```

### Fetch key from keyserver, and get keyid
Plug in YubiKey then run
```
gpg --card-status
```
Fetch from URL in card status output
```
gpg --fetch https://keys.uh-net.com/XXXXXX.asc
```
Now run the following and take the ID of the **signing** key
```
gpg -K --keyid-format=long
```

### Write Config files:

`edit ~/AppData/Roaming/gnupg/gpg.conf` to contain:
```
use-agent
```

`edit ~/AppData/Roaming/gnupg/gpg-agent.conf` to contain:
```
enable-win32-openssh-support
enable-ssh-support
enable-putty-support
```

### Restart the agent
```
gpg-connect-agent.exe killagent /bye
gpg-connect-agent.exe /bye
```
or
```
gpgconf.exe --kill gpg-agent
gpgconf.exe --launch gpg-agent
```

### Set up git to use native OpenSSH and GPG install
```
git config --global core.sshCommand "C:/Windows/System32/OpenSSH/ssh.exe"
git config --global gpg.program "C:/Program Files/GnuPG/bin/gpg.exe"
```

### Set up git user to use code signing
```
git config --global user.name "Bob Yards"
git config --global user.email bob.yards@example.com
git config --global commit.gpgsign true
git config --global tag.gpgsign true
git config --global user.signingkey <KEYID FROM gpg -k --keyid-format=long>
```

### Setup automatic launch of GPG agent on login

Left as an exercise to the reader as there are multiple ways to do this.
Recommended method is to use a task scheduler task - reference implementation
in the accompanying script. 

Otherwise you can add to the startup folder, or add in the Run targets in the
Registry HKCU.. or any other method you fancy..

Note that the agent runs per user, so this startup method should be user
specific.
