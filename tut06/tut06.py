import os
import re
import shutil

# list of all web series
web_series = ["Breaking Bad", "Game of Thrones", "Lucifer"]
# master pattern for matching srt and mp4 files of all series
pattern = re.compile(
    r"(Breaking Bad|Game of Thrones|Lucifer) -? ?s?(\d+)[ex]?(\d+) -? ?([\w ]+).*\.(mp4|srt)"
)


def get_input():
    """
    Receive series number and season and episode padding
    as input from user and return them as a tuple
    """
    # Print menu items
    print("*" * 5, "Web Series", "=" * 5)
    for i, series in enumerate(web_series):
        print(f"{i+1}. {series}")

    # initalize reqd variables
    webseries_num = 0
    season_padding = 0
    episode_padding = 0
    # receive input from user
    try:
        webseries_num = int(
            input(
                "Enter the number of the web series that you wish to rename. 1/2/3: "
            )
        )
        season_padding = int(input("Enter the Season Number Padding: "))
        episode_padding = int(input("Enter the Episode Number Padding: "))
    # if user enters string or invalid input, print error message and exit
    except ValueError:
        print("Invalid input!")
        exit(1)
    # return tuple of reqd variables
    return webseries_num, season_padding, episode_padding


def setup_dir(series_name):
    """
    Set up the correct output folder for srt and mp4 files of the
    given seriesby deleting them if they exist and creating them anew
    Return the destination directory for the given series
    """
    # source directory
    source_dir = os.path.join("wrong_srt", series_name)
    # destination directory
    dest_dir = os.path.join("corrected_srt", series_name)
    # create output folder if it does not exist
    if not os.path.exists("corrected_srt"):
        os.makedirs("corrected_srt")
    # if previous files exist, delete them
    if os.path.exists(dest_dir):
        shutil.rmtree(dest_dir)
    # copy incorrect files to be renamed later
    shutil.copytree(source_dir, dest_dir)
    # return string of destination directory
    return dest_dir


def rename_files(dest_dir, season_padding, episode_padding):
    """
    Use regex to search for
    """
    # prepare list of files in destination directory
    files = [
        f
        for f in os.listdir(dest_dir)
        if os.path.isfile(os.path.join(dest_dir, f))
    ]
    # iterate through list of files and rename them
    for file in files:
        # search for the pattern in filename
        match = pattern.search(file)
        # if no match
        if not match:
            print("File name did not match regex pattern:", file)
            continue
        # extract reqd variables from match groups
        (
            series_name,
            season_num,
            episode_num,
            episode_name,
            file_extension,
        ) = match.groups()
        # declare and initalize string to store corrected filename
        corrected_filename = ""
        # add series name
        corrected_filename += series_name
        # add season number with correct padding
        corrected_filename += " - Season " + str(int(season_num)).zfill(
            season_padding
        )
        # add episode number with correct padding
        corrected_filename += " Episode " + str(int(episode_num)).zfill(
            episode_padding
        )
        # edge case: Breaking Bad files do not contain episode names and only resolution
        # add episode name otherwise
        if episode_name != "720p":
            corrected_filename += " - " + episode_name
        # add file extension
        corrected_filename += "." + file_extension
        # full path of corrected file
        corrected_filename = os.path.join(dest_dir, corrected_filename)
        # full path of incorrectly-named file
        file = os.path.join(dest_dir, file)
        # rename files
        os.rename(file, corrected_filename)


def regex_renamer():
    """
    Driver code for renaming srt and mp4 files using regex
    """
    # get input from user
    webseries_num, season_padding, episode_padding = get_input()
    # set up and get the name of destination directory
    dest_dir = setup_dir(web_series[webseries_num - 1])
    # rename files using regex
    rename_files(dest_dir, season_padding, episode_padding)


# renaming function call
regex_renamer()
