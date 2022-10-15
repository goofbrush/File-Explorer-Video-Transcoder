import subprocess
import sys

bitrate = 874 #default value


def getMeta(path):  # gets metadata of a file
    import win32com.client  # install with 'pip install pywin32'
    # to remove this dependancy comment out section where getMeta is called

    split = path.rfind('\\')

    sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)
    ns = sh.NameSpace(path[:split])
    item = ns.ParseName(path[split+1:]) 

    colnum = 0
    metadata = {}   # creates dictionary to return

    while True:
        colname=ns.GetDetailsOf(None, colnum)   # create key
        if not colname:
            break
        colval=ns.GetDetailsOf(item, colnum)    # create value

        metadata[colname] = colval  # add association to dictionary
        colnum += 1

    return metadata # returns dictionary of metadata components


def convertTime(time):  # converts video length in metadata to seconds as an int
    seconds = 0
    toSecond = 3600

    hms = time.split(':')   # format is "hh:mm:ss"
    for x in hms:
        seconds += int(x) * toSecond
        toSecond /= 60  # 3600 -> 60 -> 1

    return seconds



handbrake_path = " ".join(sys.argv[:1])     # get path to script
handbrake_path = handbrake_path[:handbrake_path.rfind('\\')] + "\\HandBrakeCLI.exe" # change to point to handbrake exe
# ^ only works when the exe and the script are in the same folder

file_path = " ".join(sys.argv[1:])  # get path to video file, passed in by right click function
new_path = file_path[:file_path.rfind('.')] + "2.mp4"   # rename to prevent overwrite


# -- if win32com is not installed bitrate defaults to 874
try:
    metadata = getMeta(file_path)

    sizeTarget = 6 #in mb
    bitrate = (sizeTarget*1024 * 8) / convertTime(metadata["Length"])   # bitrate(kbps) = fileSize(kb) *8 / videoLength(s)
    bitrate = 2000 if bitrate > 2000 else bitrate   # cap bitrate at 2000(kbps)
except:
    print("win32 not installed")
# -- 


# # command line arguments to pass through, documentation can be found here: https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
handbrake_command = [handbrake_path, "-i", file_path, "-o", new_path,
    "-O", "-e", "x264", "-b", str(int(bitrate)), "-2", "-T", "-r", "30",
    "--encoder-preset", "VerySlow", "--encoder-profile", "High",
    "--non-anamorphic", "-B", "96"]
subprocess.run(handbrake_command, shell=True)
