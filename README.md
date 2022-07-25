# spreadsheet-banking-telegram

Welcome to the open-source repository for the Spreadsheet Banking Telegram Bot Integration!

## Overview

This integration is built using an Add-In for Microsoft Office Excel, and the Telegram Bot API.


## Setup

- Step 1: Make a copy of the this banking spreadsheet [available here](https://1drv.ms/x/s!AtkUJn0N8CerboundGT57hRpQzg?e=naBg7S/) to your own OneDrive. You can also use your own spreadsheet, in which case you will need to copy the contents of the `💬 Telegram setup`, `💬 Telegram configurator` and `💬 Telegram log` sheets into your own spreadhseet. Only if your Microsoft user has full access to the OfferZen Sharepoint will you not need to make a copy to your OneDrive.


- Step 2: Clone this repository. Make sure you have Node and NPM installed by checking their versions using `node -v` and `npm -v` respectively. Install all the required dependencies by running `npm i` in the root directory of the cloned repository.


- Step 3: Obtain OpenAI API credentials [https://beta.openai.com/account/api-keys](on this website). Replace the hardcoded credential string named `GPT_API_SECRET` inside `src/functions/functions.js`.


- Step 4: Run the command `npm run build`.


- Step 5: Run the command `npm run dev-server`. This starts a server on `localhost:3000`, which might take a few moments to spin up. If prompted to install a self-signed certificate, agree. This is required for `localhost` to be treated as a secure origin (tested working in Edge).


- Step 6: Verify that the dev server is running on localhost by trying to access `https://localhost:3000/dist/functions.json`.


- Step 7: Almost there! Open your banking spreadsheet in Excel for Web (the one you copied to OneDrive earlier).


- Step 8: Go to __Insert__ > __Add-ins__ > __Upload My Add-in__. Browse to the root directory of the earlier-cloned repository, select `manifest.xml`, and click upload. If there is a pop-up message, ignore it.


- Step 9: The solution is now set up. To get started with your Telegram bot, go to the `💬 Telegram setup` sheet and follow the steps there. After that, take a look at `💬 Telegram configurator` to start setting up your own menu options. Hope you build something awesome!


## Troubleshooting

- If anything goes wrong, use Edge browser.
- In the `💬 Telegram setup` sheet, the cell under "Telegram Bot Status" should contain a timestamp that updates every 5 seconds. If this doesn't happen (or the timestamp is more than 5 seconds ago), refresh the browser window in which the worksheet is open.
- If any issues persist, clear the Add-in cache by clearing everything in [edge://settings/cookies/detail?site=excel.officeapps.live.com](edge://settings/cookies/detail?site=excel.officeapps.live.com).
- If any issues continue to persist, particularly during development, repeat Steps 4 to 8.
- If your bot is responding, but gets confused or stuck in a loop, select the "Delete and stop" option in Telegram, then "Start" the bot again.




## Problems with setup?

If you are unable to get the Add-in or Telegram bot working properly, please reach out to me on the [Programmable Banking Community Slack](https://app.slack.com/client/T8CRG18UC/).


## License


This project is released under the MIT license.