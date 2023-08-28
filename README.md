# Master List Microsoft Teams Feed

This feed tells you when you have unread Microsoft Teams messages.

It uses puppeteer to read your Microsoft Teams DOM and notifies whenever it notices that you have unread messages.

## Installation

### NPM
- Install by running `npm install @zeyber/master-list-microsoft-teams-feed` in the terminal.
- Run from installed location with `PORT=XXXX node ./node_modules/@zeyber/master-list-microsoft-teams-feed/dist/main.js`. Define the port by replacing `PORT=XXXX` (eg. `PORT=3010 ...`). The default port is 3000.

### Clone from Github
- Clone with `git clone https://github.com/Zeyber/master-list-microsoft-teams-feed`.

#### Build
- Build with `npm run build`.
- Run with `PORT=3000 node dist/main.js`.

#### Run in Development mode
- Start with `PORT=3000 npm start`.

## Usage

### Authentication

When first starting the app, it will need to authenticate you in a browser session in order to initialize. In order to do this, chromium will be opened with the help of puppeteer and you will be taken to the service host's login page. After logging in, the browser will close and a headless version of chromium will be restarted to continue the session and complete initialization.

### Reading Feed Data

After the feed is initialized, you will be able to request JSON formatted data from the feed at its address. 
To see this in action:
- Open your browser and go url `http://localhost:3000` (or whichever address you have the app running).
- You will see a JSON-formatted response with relevant data.

This format is structured in a way interpretable by the [Master List](https://github.com/Zeyber/master-list) apps. But you could also use this feed for other purposes if you wanted.

## About Master List

An organizational list that leverages third-party APIs and displays information in a simple list.

Sometimes managing so many tasks can become overwhelming (eg. emails, agenda, tasks, social media, communications across multiple platforms). It is easy lose track of what needs to be done, when and how much you really need to do.

Master List has an [App version for Browser](https://github.com/Zeyber/master-list) and a [CLI version](https://github.com/Zeyber/master-list-cli). It features connecting to APIs or feeds that can be configured to read relevant important information that ordinary users may require.

[Find out more here](https://github.com/Zeyber/master-list)
