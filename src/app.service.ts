import { Injectable } from '@nestjs/common';
import { Browser, Page } from 'puppeteer';
import { getBrowser } from './puppeteer.utils';
import { of } from 'rxjs';
import * as fs from 'fs';

const ICON_PATH = '/assets/icon-teams.png';
const CLIENT_URL =
  'https://teams.microsoft.com/_#/conversations/48:notes?ctx=chat';

@Injectable()
export class AppService {
  browser: Browser;
  page: Page;
  initialized = false;

  async initialize() {
    this.browser = await getBrowser();
    this.page = await this.browser.newPage();

    // Disable timeout for slower devices
    this.page.setDefaultNavigationTimeout(0);
    this.page.setDefaultTimeout(0);

    this.page.goto(CLIENT_URL, {
      waitUntil: ['load', 'networkidle2'],
    });

    await this.page.waitForNavigation({
      waitUntil: 'networkidle0',
    });
    await this.waitForMainPage();
    console.log(this.page.url());

    const signedIn = this.page.url().includes(CLIENT_URL);

    if (!signedIn) {
      await this.page.waitForTimeout(1000);
      await this.browser.close();
      console.log('Deleting session data.');
      fs.rmSync('./puppeteer-teams-session', { recursive: true, force: true });
      console.log('Deleted session data.');
      await this.login();
      this.initialize();
    } else {
      this.initialized = true;
      console.log('Microsoft Teams initialized.');
    }
  }

  getData() {
    if (this.initialized) {
      return this.getThreads();
    } else {
      return of({
        data: [
          { message: 'Microsoft Teams feed not initialized', icon: ICON_PATH },
        ],
      });
    }
  }

  login() {
    return new Promise(async (resolve) => {
      this.browser = await getBrowser({
        headless: false,
        userDataDir: './puppeteer-teams-session',
      });
      this.page = await this.browser.newPage();
      // Disable timeout for slower devices
      this.page.setDefaultNavigationTimeout(0);
      this.page.setDefaultTimeout(0);

      console.log('Opening sign in page...');
      await this.page.goto(CLIENT_URL, {
        waitUntil: ['load', 'networkidle2'],
      });

      console.log('Please login to Microsoft Teams');
      await this.page.waitForFunction(
        `window.location.href.includes('https://teams.microsoft.com/_#/conversations/48:notes?ctx=chat')`,
      );
      await this.waitForMainPage();

      await this.browser.close();

      console.log('Microsoft Teams signed in!');
      resolve(true);
    });
  }

  async getThreads(): Promise<any> {
    return new Promise(async (resolve, reject) => {
      await (async () => {
        try {
          const threadList = await this.page.waitForSelector(
            '[aria-label="Chat list"]',
          );
          const threads = await threadList.$$('[role="treeitem"]');
          if (threads.length) {
            const items = [];
            for (const thread of threads) {
              const unread = await thread.evaluate((el) =>
                el.getElementsByClassName('ts-unread-channel'),
              );

              if (unread === undefined) {
                const threadText: string = await this.page.evaluate(
                  (el) => el.innerText,
                  thread,
                );
                const name = threadText.split('\n')[0];
                items.push({ message: name, icon: ICON_PATH });
              }
            }
            resolve({ data: items });
          }
        } catch (e) {
          reject(e);
        }
      })();
    });
  }

  protected async waitForMainPage() {
    const res = await this.page
      .waitForSelector('[id="chat-header-title"]', {
        timeout: 60000,
      })
      .catch(() => console.log('timed out waiting for gui'));

    if (res) {
      await this.page
        .waitForRequest(
          'https://teams.microsoft.com/registrar/prod/V2/registrations',
          { timeout: 60000 },
        )
        .catch(() => console.log('timed out waiting for request'));
      console.log('waiting for any last requests');
      await this.page.waitForTimeout(10000);
    } else {
      const tryAgainBtn = await this.page
        .waitForSelector('a[id="try-again-link"]', { timeout: 5000 })
        .catch((e) => console.log('Cannot try again'));
      if (tryAgainBtn) {
        console.log('Trying again');
        await tryAgainBtn.click();
        await this.waitForMainPage();
      }
    }
  }
}
