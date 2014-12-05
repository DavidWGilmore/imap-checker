# IMAP Checker

Imap checker is a simple command line utility for Windows that checks a given imap server mailbox for the latest unread message and returns the data and time stamp if there is one. It was originally developed by [David W. Gilmore](http://www.gilmoreipedia.org) and released in December 2014.

IMAP Checker utilizes the [MailSystem.Net ](https://mailsystem.codeplex.com/) libraries to make the connection and perform it's tasks

## Setup

The utility is a self contained application. Everything you need to run it is included in the zip file for the release. Simply extract the files to a directory, edit the imap_checker.config file and you're ready to go. You can find the config file refernece here: http://github.com/davidwgilmore/imap-checker/wiki/ConfigFile

## Usage

Open a command promot in the directory you extracted the files to and execute 'imap_checker.exe'. If an unread message exists, it will return the date and time.

## Support

If you find a bug (highly likely) or have problems making the application work, please use the GitHub Issues system to submit an issue. For general help or discussion, please use the imap-checker@googlegroups.com mailing list.

