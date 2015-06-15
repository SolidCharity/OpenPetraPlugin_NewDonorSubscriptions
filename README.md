# Plugin NewDonorSubscriptions

This is a plugin for [www.openpetra.org](http://www.openpetra.org).

## Functionality

You can send a letter to all new donors in a given time period.
They are recognized by donating for the first time in this time period, and being subscribed for a certain publication.
Only recipients that are part of a given extract will be included. The extract contains all the recipients for the next issue of the publication.

## Dependencies

This plugin depends on:

* https://github.com/SolidCharity/OpenPetraPlugin_PrintPreview

## Installation

Please copy this directory to your OpenPetra working directory, to csharp\ICT\Petra\Plugins, 
and then run
    nant generateSolution

Please check the config directory for changes to your config files.

## License

This plugin is licensed under the GPL v3.