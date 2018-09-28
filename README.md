# SSRS
String Search and Replace Script - AutoHotkey Script

This script was created to clean up the names of the TV recordings created by the DVB software. The source files are named according to the OTA EPG data and while the shows are most of the time reasonably named, there are some frequently occurring conventions. One, for example, is the "Movie:" tag in front of a movie. This ends up in files named as Movie_ Deadpool-TVChannel-YYYY-MM-DD-HH-MM.mkv .

While this is not particularly annoying, it causes the files to get sorted randomly and the ever-growing movie collection ends up a mess.

This script makes it easy to create rules to remove these unwanted tags from desired files in a specified location. It can also be applied to text files so that the tag is also removed inside the file. I use this for the .nfo (xml) files created from the EPG data.

The usage is explained with comments in the script. Of course, you need AutoHotkey installed to use the script.
