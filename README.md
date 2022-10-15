# File-Explorer-Video-Transcoder

## About
This is a helper script for HandbrakeCLI to run through the right click function in file explorer.

The script will transcode a selected video file with a bitrate based on the video length as to have the output be below a certain size. This is helpful if you are uploading videos to sites with size limits, by default the script has a target of 6mb to make sure the result is below Discords 8mb file limit

### How it Works
There are only 2 scripts, the first is to edit the Windows registry to add an option to the right click function in file explorer and to associate it with the other script which runs the video transcoder on the selected video

## Getting Started
### Dependencies
- HandbrakeCLI 
- Python
- [Optional] pywin32

### Installation
1. Clone this repo with `git clone https://github.com/goofbrush/File-Explorer-Video-Transcoder.git`
2. Download python from https://www.python.org/downloads/
3. Download the HandbrakeCLI executatable from https://handbrake.fr/downloads2.php
4. Place `HandBrakeCLI.exe` in the same folder as the python scripts and move the folder to someplace permanent
5. [Optional] run `pip install pywin32` (if not installed all videos are encoded with the same bitrate)
6. Enter admin and run `make_key.py`

## Usage
If everything is setup correctly, you will be able to right click a video file and be able to compress the video from the menu. The transcode settings can also be changed by selecting config from the right click menu

### Config
The current with an extension from this list ["mp4", "mkv", "mov", "wmv", "avi"] and select "Compress Video"  If you would like to extend this to other video formats that handbrake supports, add them to the "fileType" list in  [make_key.py](make_key.py)

If you opt to install pywin32, the script will calculate what bitrate to use based on the length of your video and force the output to be around the target filesize. You can change the target size in the config file, and if set to 0 the transcoder will always use the bitrate defined in the config

You can also change the transcodeing parameters by editing the config file, just seperate each input by a comma
