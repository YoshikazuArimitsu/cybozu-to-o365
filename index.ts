import "isomorphic-fetch";

var adal = require("adal-node");
var HttpsProxyAgent = require("https-proxy-agent");
import { Client } from "@microsoft/microsoft-graph-client/lib/src/Client";

const meow = require("meow");
const puppeteer = require("puppeteer");
const fs = require("fs");
const os = require("os");
const path = require("path");
const moment = require("moment");
const parseCsv = require("csv-parse/lib/sync");
const rc = require("rc");

const _ = require('underscore');

class FileCache {
  _entries: Array<any> = [];

  private save() {
    fs.writeFileSync('.token', JSON.stringify(this._entries, null, '    '));
  }

  private load() {
    try {
      this._entries = JSON.parse(fs.readFileSync('.token', 'utf8'));
    } catch (e) {

    }
  }

  public remove(entries, callback) {
    var updatedEntries = _.filter(this._entries, function (element) {
      if (_.findWhere(entries, element)) {
        return false;
      }
      return true;
    });

    this._entries = updatedEntries;
    this.save();
    callback();
  }

  public add(entries, callback) {
    // Remove any entries that are duplicates of the existing
    // cache elements.
    _.each(this._entries, function (element) {
      _.each(entries, function (addElement, index) {
        if (_.isEqual(element, addElement)) {
          entries[index] = null;
        }
      });
    });

    // Add the new entries to the end of the cache.
    entries = _.compact(entries);
    for (var i = 0; i < entries.length; i++) {
      this._entries.push(entries[i]);
    }

    this.save();
    callback(null, true);
  }

  public find(query, callback) {
    this.load();

    var results = _.where(this._entries, query);
    callback(null, results);
  }

  public first() {
    this.load();
    return this._entries.length > 0 ? this._entries[0] : null;
  }
}

class AdalAutenticator {
  constructor(public param: any) { }

  public async signIn(): Promise<string> {
    return new Promise(resolve => {
      const AuthenticationContext = adal.AuthenticationContext;

      var authorityUrl = this.param.authorityHostUrl + "/" + this.param.tenant;

      var cache = new FileCache();
      var resource = "https://graph.microsoft.com";

      var context = new AuthenticationContext(authorityUrl, null, cache);

      // DeviceCodeによるログインフロー
      const login = () => {
        context.acquireUserCode(
          resource,
          this.param.clientId,
          "es-mx",
          (err, response) => {
            if (err) {
              console.log("well that didn't work: " + err.stack);
            } else {
              console.log(response.message);
              context.acquireTokenWithDeviceCode(
                resource,
                this.param.clientId,
                response,
                function (err, tokenResponse) {
                  if (err) {
                    console.log(
                      "error happens when acquiring token with device code"
                    );
                    console.log(err);
                  } else {
                    console.log("Autehticate success");
                    resolve(tokenResponse.accessToken);
                  }
                }
              );
            }
          }
        );
      }

      if (cache.first()) {
        context.acquireTokenWithRefreshToken(cache.first().refreshToken, this.param.clientId, null, (err, tokenResponse) => {
          if (err) {
            login();
          } else {
            // トークンキャッシュ使用
            console.log("Autehticate success by tokenCache");
            resolve(tokenResponse.accessToken);
          }
        });
      } else {
        login();
      }
    });
  }
}

class O365Calendar {
  graphClient: Client;
  calendarId?: string;

  constructor(accessToken: string, proxy?: string, calendarId?: string) {
    const o = {
      authProvider: done => {
        done(null, accessToken);
      }
    };
    if (proxy) {
      o["fetchOptions"] = { agent: new HttpsProxyAgent(proxy) };
    }

    this.calendarId = calendarId;
    this.graphClient = Client.init(o);
  }

  private get eventsPath(): string {
    if (this.calendarId) {
      return `/me/calendars/${this.calendarId}/events`;
    } else {
      return '/me/events';
    }
  }

  async createCalendarEvent(e: any) {
    const event = {
      subject: e.summary,
      body: {
        contentType: "HTML",
        content: e.summary
      },
      start: {
        dateTime: e.start,
        timeZone: "Asia/Tokyo"
      },
      end: {
        dateTime: e.end,
        timeZone: "Asia/Tokyo"
      },
      location: {
        displayName: e.location
      }
    };

    try {
      await this.graphClient.api(this.eventsPath).post(event);
    } catch (err) {
      console.log(err);
    }
  }

  async deleteCalendarEvent(id: string) {
    await this.graphClient.api(`${this.eventsPath}/${id}`).delete();
  }

  private get calEventsPath(): string {
    if (this.calendarId) {
      return `/me/calendars/${this.calendarId}/events`;
    } else {
      return '/me/calendar/events';
    }
  }

  async getCalendarEvents() {
    const response = await this.graphClient
      .api(this.calEventsPath)
      .top(10000)
      .get();
    return response.value.map((v: any) => {
      return {
        id: v.id,
        summary: v.subject,
        start: v.start.dateTime,
        end: v.end.dateTime,
        location: v.location.displayName
      };
    });
  }
}




const cli = meow(
  `
  c2o - Move schedules from Cybozu to Google Calendar.

  Usage
    $ c2o

  Options
    --config, -c    Pass config file. By default, c2o will read "~/.config/c2o".
    --quiet,  -q    Hide debug messages
    --show, -s      Show browser window
    --version, -v   Show version
    --help, -h      Show this help

  Examples
    $ c2o -c ./config.json
      >>>> Fetching events from Cybozu Calendar...DONE
      >>>> Fetching events from Google Calendar...DONE
      >>>> Inserting new events...
        Inserted: [会議] B社MTG (2018/08/16)
        Inserted: [会議] 目標面談 (2018/10/31)
      >>>> Inserted 2 events.
      >>>> Deleting removed events...
      	Deleted: [外出] 幕張メッセ (2018/08/17)
      >>>> Deleted 1 events.
`,
  {
    flags: {
      config: {
        type: "string",
        alias: "c"
      },
      quiet: {
        type: "boolean",
        alias: "q"
      },
      show: {
        type: "boolean",
        alias: "s"
      },
      help: {
        type: "boolean",
        alias: "h"
      },
      version: {
        type: "boolean",
        alias: "v"
      }
    }
  }
);

// Detect CLI flags
if (cli.flags.help) {
  cli.showHelp();
}
if (cli.flags.version) {
  cli.showVersion();
}

const log = cli.flags.quiet ? () => { } : str => process.stdout.write(str);

let config = rc("c2o");
if (cli.flags.config) {
  try {
    config = JSON.parse(fs.readFileSync(cli.flags.config));
  } catch (e) {
    console.error(e);
    process.exit(-1);
  }
}

// Utils
const now = new Date();
const year = now.getFullYear();
const month = now.getMonth() + 1;
const day = now.getDate();

const csvDir = path.join(os.tmpdir(), Date.now() + "");
const csvPath = path.join(csvDir, "schedule.csv");

// Main
(async () => {
  // Office365 login
  const autehnicator = new AdalAutenticator(config.calendar);
  const accessToken = await autehnicator.signIn();
  const calendarClient = new O365Calendar(accessToken, config.proxy,
    config.calendar.calendarId);


  const browser = await puppeteer.launch({ headless: !cli.flags.show });
  const page = await browser.newPage();
  await page._client.send("Page.setDownloadBehavior", {
    behavior: "allow",
    downloadPath: csvDir
  });

  if (config.basicAuthUser && config.basicAuthPass) {
    await page.authenticate({
      username: config.basicAuthUser,
      password: config.basicAuthPass
    });
  }

  log(">>>> Fetching events from Cybozu Calendar...");

  // Login
  await page.goto(config.cybozuUrl);
  await page.type('input[name="username"]', config.username);
  await page.type('input[name="password"]', config.password);
  await page.waitFor(1000);
  await page.click("form input[type=submit]");
  await page.waitForNavigation({
    timeout: 30000,
    waitUntil: "domcontentloaded"
  });

  // Go to CSV exporter page
  await page.goto(`${config.cybozuUrl}o/ag.cgi?page=PersonalScheduleExport`, {
    waitUntil: "domcontentloaded"
  });

  // Input date
  await page.select('select[name="SetDate.Year"]', year + "");
  await page.select('select[name="SetDate.Month"]', month + "");
  await page.select('select[name="SetDate.Day"]', day + "");
  await page.select('select[name="EndDate.Year"]', year + 1 + "");
  await page.select('select[name="EndDate.Month"]', month + "");
  await page.select('select[name="EndDate.Day"]', day + "");
  await page.select('select[name="oencoding"]', "UTF-8");

  // Download CSV
  await page.click(".vr_hotButton");
  await page.waitFor(3000);
  await browser.close();

  // Parse CSV
  const newEvents: any[] = [];
  const csv = fs.readFileSync(csvPath, "utf8");
  parseCsv(csv).forEach((line, i) => {
    if (i === 0) {
      return;
    }

    const start = moment(new Date(line[0] + " " + line[1]));
    const end = moment(new Date(line[2] + " " + line[3]));
    const summary = `[${line[4]}] ${line[5]} (${start.format("YYYY/MM/DD")})`;
    const location = line[8];

    newEvents.push({
      start: start.toISOString(),
      end: end.toISOString(),
      location,
      summary
    });
  });

  log("DONE\n");


  log(">>>> Fetching events from Office365 Calendar...");

  const oldEvents = await calendarClient.getCalendarEvents();

  log("DONE\n");
  log(`>>>> Inserting new events...\n`);

  let insertedCount = 0;
  for (const event of newEvents) {
    if (oldEvents.findIndex(e => e.summary === event.summary) !== -1) {
      continue;
    }
    //await calendar.Events.insert(config.calendar.calendarId.primary, event);
    await calendarClient.createCalendarEvent(event);

    log(`\tInserted: ${event.summary}\n`);
    insertedCount++;
  }

  log(`>>>> Inserted ${insertedCount} events.\n`);
  log(`>>>> Deleting removed events...\n`);

  let deletedCount = 0;
  for (const event of oldEvents) {
    if (newEvents.findIndex(e => e.summary === event.summary) !== -1) {
      continue;
    }
    //await calendar.Events.delete(config.calendar.calendarId.primary, event.id, { sendNotifications: true });
    await calendarClient.deleteCalendarEvent(event.id);

    log(`\tDeleted: ${event.summary}\n`);
    deletedCount++;
  }
  log(`>>>> Deleted ${deletedCount} events.\n`);
})();
