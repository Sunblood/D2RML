# Diablo 2 Resurrected Multi-Launcher by Sunblood

Allows you to launch and connect with multiple instances of Diablo 2 Resurrected using pre-generated login tokens.

Requires handle64.exe from https://docs.microsoft.com/en-us/sysinternals/downloads/handle (included with the compiled exe)

### Discord: https://discord.gg/PwGR2rRafX

### BTC Donations: 3FLeQd9zt7H6zqG839apK5MxkVJvfNYxer

![2021-10-21 21_24_16-Window](https://user-images.githubusercontent.com/6067956/138388188-6e7b3dec-b07a-4036-99a5-b180348a4b75.png)

## Setup
1. Click the "Add Token" button and give your token a name
2. The D2R launcher will open. Log in with the Battle.net account you want to associate with this token.
3. Click "Play" to start D2R
4. Wait for a connection to D2R online servers. D2RML will hit Spacebar to skip intro videos for you.
5. The token is saved automatically to a .BIN file in the current working directory
6. Repeat as many times as required for the number of accounts you want to use (maximum is likely 4 concurrent connections)

## Usage
1. Check the box(es) for the token you want to use
2. Click "Launch Selected"
3. D2R will start and connect to your account automatically
4. If you've selected multiple tokens, they will each launch once the previous client connects to D2R servers

# Important Notes
* Tokens are **one-time use**. D2R generates a new token during each successful connection. D2RML scans for new tokens and saves them automatically *only if you're using D2RML to start the game.*
* This means that if you launch D2R manually via normal means, the D2RML saved token is invalidated and will no longer allow you to connect.
* If you try to connect using an invalid token, the server connection will fail and you'll get kicked to single-player.
* To fix this, check the box for the invalid token and click the "Refresh Token" button. You'll have to go through the login process again to save a new token.
* **ALWAYS USE D2RML TO LAUNCH THE GAME**

## Changelog
* 0.0.1
  -   Initial Commit
* 0.0.2
  -   Added field to specify cmdline arguments when launching D2R.exe  
  -   Increased verbosity of tooltip messages when creating a token  
  -   Spacebar is only pressed for the first 15 seconds after game launch (so it doesn't continue to spam while waiting in queue)  
  -   Attempt to close pre-existing Bnet windows before creating a token
* 0.0.3
  -   Added option to toggle Skip Intro
  -   Added option to rename D2R window to match token name
  -   Settings now save to D2RML.ini
  -   Handle64.exe is pre-run at launch
  -   Launch Bnet app directly instead of via D2R's launcher
  -   GUI now remains responsive while waiting on tokens
  -   Additional back-end work for upcoming features
   
## Virus warnings
"My antivirus flagged this as a virus! Are you trying to steal my account!?"  
No. Autoit executables are often flagged as malicious or trojans due to a long history of abuse. Feel free to [download Autoit](https://www.autoitscript.com/site/autoit/downloads/) and compile it yourself from source. Here's the Virustotal report from version 0.0.2, as you can see it gets flagged for multiple things.

![2021-10-22 11_08_16-VirusTotal - File - 00238f91343f8191c977980e5a17a55391cfe690f8cf8979948c4263423a](https://user-images.githubusercontent.com/6067956/138488135-dcb08250-b4ec-4163-8b09-27cccc2ff651.png)

[![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2FSunblood%2FD2RML&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false)](https://hits.seeyoufarm.com)
