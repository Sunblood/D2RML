# Diablo 2 Resurrected Multi-Launcher by Sunblood

Allows you to launch and connect with multiple instances of Diablo 2 Resurrected using pre-generated login tokens.

Requires handle64.exe from https://docs.microsoft.com/en-us/sysinternals/downloads/handle (included)

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
