# EZ-Rename - TV Episode Renamer

EZ-Rename is a simple cross-platform python based TV episode renaming script that uses the TVMaze database to automatically identify episodes and rename video files in a consistent, human-readable format. It includes metadata writing, optional NFO generation, restore capability, and full preview before applying changes.

Features:

Automatic Episode Detection
Identifies season and episode numbers from filenames using patterns like S01E02 or 1x02. Extracts show names while filtering out noise tokens such as resolution flags, codec labels, and release group tags.

TVMaze Lookup
Matches shows using the TVMaze API and retrieves accurate episode titles. When multiple shows match a name, the user can select the correct one via a dialog.
If it can't find the show based on the title of the episode, it will prompt you to search for it.  

Clean Renaming
Renames files into a standardized format such as:

Subtitle Support
Subtitle files (.srt) with matching names are renamed alongside the video file.

Metadata Writing
Writes title metadata to MKV, MP4, M4V, MOV, and (with limitations) AVI files. On Windows, fallbacks to shell metadata writing are used when appropriate tools are unavailable.

Optional .nfo Sidecar Creation
Writes minimal Kodi-compatible episode NFO files containing title, season, episode number, and show name.

Backup and Restore
Can optionally create a restore log of original filenames prior to renaming. A restore command reverts all changes using this log.

Custom Noise Tokens
Users may define additional noise tokens to strip from filenames during show-name detection. These are saved in the persistent options file.


Dependencies: 
Required:
Python 3.9 or newer
https://www.python.org/downloads/

Optional:
mutagen (for MP4/MOV/AVI metadata)
MKVToolNix (for MKV metadata)
The application includes an installation helper for optional dependencies on Windows, macOS, and Linux.

Options File:
Settings are stored in:
~/.tv_renamer_options.json
This file contains folder paths, theme preferences, metadata settings, custom noise tokens, and other configuration values.

Simple Usage:

1. Run the script and either select a folder containing TV episode files, or it can be run directly from the folder you'd like to scan.
2. Click Scan to detect episodes and retrieve titles.
3. Review the planned renames and skipped items.
4. Click Apply Renames to perform the operations.

License:
CC0 (Creative Commons Zero)  
ie: It doesn't matter.  Do whatever the hell you want with this; there are no restrictions. 
