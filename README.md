# c2o - Move schedules from Cybozu to Office365.

[cybozu-to-google](https://github.com/fand/cybozu-to-google) をフォークして Office365 のカレンダーに同期するよう作り変えたツールです。

# Usage

[ここらへん](https://docs.microsoft.com/ja-jp/azure/active-directory/develop/quickstart-register-app) で説明されている通りにお使いのADにアプリケーションを登録してください。

First, create '~/.config/c2o' like this:

```json
{
  "cybozuUrl": "https://XXXXXXX.cybozu.com/",
  "username": "YOUR_USERNAME_OF_CYBOZU",
  "password": "YOUR_PASSWORD_OF_CYBOZU",
  "proxy": "{PROXY_SERVER}",
  "basicAuthUser": "{BASIC_AUTH_USER}",
  "basicAuthPass": "{BASIC_AUTH_PASS}",
  "calendar": {
      "authorityHostUrl": "https://login.microsoftonline.com",
      "tenant": "{YOUR AZURE TENANT ID}",
      "clientId": "{YOUR AZURE CLIENT ID}"
  }
}
```

proxy, basicAuthUser, basicAuthPass は省略可能。

After that, just hit `c2o` in your terminal.

# Options

```
  --config, -c    Pass config file. By default, c2o will read "~/.config/c2o".
  --quiet,  -q    Hide debug messages
  --show, -s      Show browser window
  --version, -v   Show version
  --help, -h      Show this help
```

# Examples

```
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
```


# License

MIT
