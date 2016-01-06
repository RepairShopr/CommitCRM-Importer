# CommitCRM-Exporter
A simple Windows app to export data to our REST API

We are sharing the source code just to give people an idea how easy this type of work is, the first draft of this was probably just a half week of work for a single developer.

Feel free to fork and adapt to your REST API.

## Installation

Look for the link above to "Releases" and download the latest one.

Look for a file attached to the release called "install.exe"

## Operation

Once you install and run the tool, it wants you to login. Put in your RepairShopr login and password, and click login. 

Once logged in - it is ready to talk to our API.

Just check the box for customers (which includes contacts), and/or Tickets - and click "Export".

If you have a lot of data (thousands of tickets and customers) expect this to take quite a while.

It should be able to resume properly - and it also writes out an error log if you want to see why records fail, as this is usually normal to get some percentage (1%-10%) for one reason or another.

Any time you are asking for help with this tool, please share a dropbox/google drive link to your error log if possible.

## Developer Notes

Built with Visual Studio - community edition 2015 works

Installer built with http://www.sminstall.com/
