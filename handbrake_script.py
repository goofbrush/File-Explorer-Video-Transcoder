import subprocess
import sys

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

def isComment(line):    # For reading config file, sees if line should be ignored
    for char in line:
        if(char == "#" or char == "\n"):
            return True
        elif(char != " "):
            return False
    return True

def getNum(string, default=0): # gets first number occuring in a string
    out = ""
    for char in string:
        if(char.isnumeric()):
            out += char
        if(len(out) > 0 and not char.isnumeric()):
            break
    out = default if(not out) else int(out)
    return out



folder_path = " ".join(sys.argv[:1])     # get path to script
folder_path = folder_path[:folder_path.rfind('\\')] # remove extra path
handbrake_path = folder_path + "\\HandBrakeCLI.exe" # change to pointer to handbrake exe
cfg_path = folder_path + "\\config.cfg" # change to pointer to config file
# ^ only works when the exe and the script are in the same folder

file_path = " ".join(sys.argv[1:])  # get path to video file, passed in by right click function
new_path = file_path[:file_path.rfind('.')] + "2.mp4"   # rename to prevent overwrite


# default values
sizeTarget = 6 #in mb
bitrate = 874 #in kbps
handbrakeConfig = []


# -- Reading config file
f = open(cfg_path,"r")
lines = f.readlines()
for line in lines: 
    if(not isComment(line)): # discards comments and empty lines
        if(line.lower().find("sizetarget") >= 0):
            sizeTarget = getNum(line[line.find("="):], sizeTarget)
        elif(line.lower().find("bitrate") >= 0):
            bitrate = getNum(line[line.find("="):], bitrate)
        elif(line.lower().find("handbrake") >= 0):
            handbrakeConfig = line[line.find("=")+1:].split(",")
# --end



# -- calculates bitrate on video length
if(sizeTarget > 0):
    try:
        metadata = getMeta(file_path)   # obtains metadata from video

        bitrate = (sizeTarget*1024 * 8) / convertTime(metadata["Length"])   # bitrate(kbps) = fileSize(kb) *8 / videoLength(s)
        bitrate = 2000 if bitrate > 2000 else bitrate   # cap bitrate at 2000(kbps)
    except:
        print("pywin32 not installed")
# --end


# command line arguments to pass through, documentation can be found here: https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
handbrake_command = [handbrake_path, "-i", file_path, "-o", new_path, 
    "-b", str(int(bitrate))] + handbrakeConfig  # <- adds extra arguments from config

subprocess.run(handbrake_command, shell=True) # runs handbrake exe
