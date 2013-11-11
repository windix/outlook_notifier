# Outlook Notification

New email notification (desktop / iOS push) for Outlook Web App

You need to have access to Microsoft EWS (Exchange Web Sercice)

## ALSO TRY THUNDERBIRD!

With the Thunderbird extension:
* [ExQuilla](https://exquilla.zendesk.com/home) - for email
* [Exchange EWS Provider Add-on for Lightning](http://www.1st-setup.nl/wordpress/?page_id=133) - for calendar (require lightning)

You can get email / calendar working (under Linux and other OS) in Thunderbird instead of webmail! It can replace desktop notification. 

If you are still interested in mobile notification (iOS push), keep reading.


## PREREQUISITES

* [Growl For Linux](http://mattn.github.com/growl-for-linux/) - desktop notification 
* [Prowl](http://www.prowlapp.com/) - iOS push notification


## INSTALL

### Growl

For Ubuntu:

```
# sudo add-apt-repository ppa:mattn/growl-for-linux
# sudo apt-get update
# sudo apt-get install growl-for-linux
```

To start, run `/usr/bin/gol`

Under Ubuntu, to set it start with X:

Click the gear icon on top-right corner: Startup Application...

Then add it to the list


### Prowl

Just download the app from app store and register your account, then create a new API key.


### Itself

```
bundle install
```

(all the necessary gems are packed for easy installation)


## SETUP

After running for the first time, edit config file


## RUNNING

```
./check.rb
```


## MISC

### Create desktop shortcut for Outlook Web App

Chrome > Tools > Create Application Shortcut... 

and it will create a desktop shortcut for it 

### Running Outlook Web App with Linux Chrome

Linux Chrome is not in the official support list, but we can spoof user agent as other browser (e.g. Firefox on Windows) using this extension:

https://chrome.google.com/webstore/detail/user-agent-switcher-for-c/djflhoibgkdhkhhcedjiklpkjnoahfmg

After installation, go to settings, and add Your webmail URL (e.g. "wmail.tabcorp.com.au") to "Permanent Spoof list", so it will only affect outlook web app

### Gems

* [ruby-prowl](https://github.com/augustl/ruby-prowl)
* [ruby-growl](https://github.com/drbrain/ruby-growl)
* [Viewpoint](https://github.com/windix/Viewpoint) (my version with a bug fix)

### License

© 2013 Wei Feng, Luxbet Pty Ltd. 

Released under [The BSD 3 clause License](http://www.opensource.org/licenses/BSD-3-Clause)


