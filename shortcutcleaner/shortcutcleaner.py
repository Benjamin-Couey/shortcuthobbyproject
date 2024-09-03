import argparse
import os
from pathlib import Path
from pywintypes import com_error
import time
import tkinter as tk
from tkinter import filedialog
from win32com import client
from win32com.client import CDispatch
from urllib.parse import urlparse

# Some things to consider about optimizing the performance of this project:
# Practically speaking, it probably isn't going to be run very frequently. After one
# run that cleans up all broken shortcuts, it will take a significant amount of time
# (I assume?) before enough shortcuts build up to justify running this again.
#
# Probably fine to just make the project and clarify in the README all the assumptions
# I made and the reason I didn't make certian optimizations
# "the bottleneck is IO not CPU cycles, so parallelization won't have any benefit"
# "predicting the distribution of broken shortcuts on a system would require more
# research than makes sense for this project"
#
# Will be a good oportunity to show the test harness for a project like this that
# relies heavily on the file system.
#
# Can't easily determine whether a URL shortcut is broken or not # without
# following the URL which seems like something a shortcut cleaner shouldn't
# automatically do. Current thought on how to handle this is to have the program just
# report url shortcuts it finds and let the user sort them out? Make it configurable?
# Option to just scrub all url shortcuts? Urls that match a certain pattern?
#
# Should auto scrub url shortcuts that don't have valid urls.
# Possible this doesn't matter, or at least doesn't practiclaly matter as Windows
# won't let you create a url shortcut with an invalid url (it will always prepend
# http://).
#
# Ideally, script would have some way to avoid installed packages as those probably
# don't have shortcuts, much less broken ones.
# TOOD: Update Windows python install to 3.9

FILE_SHORTCUT_EXT = '.lnk'
NET_SHORTCUT_EXT = '.url'

def parse_clean_drives( clean_drives ):
    parsed_drives = []

    for drive in clean_drives:
        if len(drive) > 0:
            alpha_drive = list( filter( str.isalpha, drive ) )
            first_letter = alpha_drive.pop(0)
            if len(alpha_drive) > 0:
                print("There are multiple drive letter in the input " + drive + ". The whole input will be ignored.")
            else:
                parsed_drives.append( first_letter.upper() + ":" )
        else:
            print("There is an empty input which will be ignored.")

    return parsed_drives

def is_file_shortcut( shortcut ):
    if isinstance( shortcut, str ):
        _, extension = os.path.splitext( shortcut )
        return extension == FILE_SHORTCUT_EXT
    elif issubclass( type(shortcut), Path ):
        return shortcut.suffix == FILE_SHORTCUT_EXT
    elif isinstance( shortcut, CDispatch ):
        try:
            _, extension = os.path.splitext( shortcut.FullName )
            # Since URL shortcuts do not have the RelativePath attribute, an
            # alternative way to check if a COMObject shortcut is a file short
            # cut would be to check for the presence of that attribute.
            # return hasattr( shortcut, 'RelativePath' )
            return extension == FILE_SHORTCUT_EXT
        except AttributeError:
            return False
    else:
        raise ValueError("Not a string, Path, or CDispatch shortcut.")

def is_net_shortcut( shortcut ):
    if isinstance( shortcut, str ):
        _, extension = os.path.splitext( shortcut )
        return extension == NET_SHORTCUT_EXT
    elif issubclass( type(shortcut), Path ):
        return shortcut.suffix == NET_SHORTCUT_EXT
    elif isinstance( shortcut, CDispatch ):
        try:
            _, extension = os.path.splitext( shortcut.FullName )
            return extension == NET_SHORTCUT_EXT
        except AttributeError:
            return False
    else:
        raise ValueError("Not a string, Path, or CDispatch shortcut.")

def is_valid_url( url ):
    if isinstance( url, str ):
        result = urlparse( url )
        # It is possible for a URL shortcut to point to a file on the local
        # filesystem, in which case it should be handled like a file shortcut.
        if result.scheme and result.scheme == "file" and not result.netloc and result.path:
            path = result.path.strip("/")
            return os.path.isfile( path ) or os.path.isdir( path )
        else:
            return bool( result.scheme and result.netloc )
            # return all( [ result.scheme, result.netloc ] )
    else:
        raise ValueError("Not a string.")

def is_broken_shortcut( shortcut ):
    # Convert to a shortcut object if necessary.
    try:
        if isinstance( shortcut, str ):
            shell = client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut( shortcut )
        elif issubclass( type(shortcut), Path ):
            shell = client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut( str( shortcut ) )
        elif isinstance( shortcut, CDispatch ):
            pass
        else:
            raise ValueError("Not a string, Path, or CDispatch shortcut.")
    except com_error as e:
        # Will be raised by CreateShortCut() if shortcut is a str or Path that
        # does not point to a shortcut file, in which case it can't be a broken
        # shortcut.
        print(e)
        return False

    try:
        if is_file_shortcut( shortcut ):
            if not shortcut.TargetPath:
                # Address corner cases here.
                # It is possible for .lnk files to have a TargetPath that points to
                # a URL or a system program like Control Panel. In these cases, the
                # TargetPath of the shortcut object returned by shell.CreateShortCut
                # will be empty (even if the shortcut itself works). Since the script
                # currently can't differentiate between a working or broken shortcut
                # in this case, it should just return False so the mainloop doesn't
                # try to clean it.
                # TODO: Figure out a way to actually get the target path of .lnk files
                # in these corner cases and more accurately determine if they are
                # valid.
                print("Encountered a CDispatch object with an empty TargetPath at " + shortcut.FullName)
                return False
            else:
                return not ( os.path.isfile( shortcut.TargetPath ) or os.path.isdir( shortcut.TargetPath ) )
        elif is_net_shortcut( shortcut ):
            return not is_valid_url( shortcut.TargetPath )
        else:
            print("Encountered a CDispatch object that is not a recognized shortcut type.")
            return False
    except AttributeError as e:
        print(e)
        return False


def is_target_drive_missing( shortcut ):
    # Convert to a shortcut object if necessary.
    try:
        if isinstance( shortcut, str ):
            shell = client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut( shortcut )
        elif issubclass( type(shortcut), Path ):
            shell = client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut( str( shortcut ) )
        elif isinstance( shortcut, CDispatch ):
            pass
        else:
            raise ValueError("Not a string, Path, or CDispatch shortcut.")
    except com_error as e:
        # Will be raised by CreateShortCut() if shortcut is a str or Path that
        # does not point to a shortcut file, in which case it can't be a broken
        # shortcut.
        print(e)
        return False

    try:
        if is_file_shortcut( shortcut ):
            drive, _ = os.path.splitdrive( shortcut.TargetPath )
            return not os.path.exists( drive )
        elif is_net_shortcut( shortcut ):
            return False
        else:
            print("Encountered a CDispatch object that is not a recognized shortcut type.")
            return False
    except AttributeError as e:
        print(e)
        return False

def main():
    parser = argparse.ArgumentParser(
        prog="shortcutcleaner",
        description="Search for and clean broken shortcuts."
    )
    parser.add_argument(
        '--clean',
        help='Delete broken shortcuts that are found (default: report broken shortcuts).',
        action='store_true',
    )
    parser.add_argument(
        '--clean_drives',
        help='A list of drive letters. If shortcuts target these missing drives, they will be treated as broken shortcuts. Strings with multiple letters will be ignored.',
        action='store',
        nargs='+',
        default=[]
    )
    args = parser.parse_args()

    args.clean_drives = parse_clean_drives( args.clean_drives )

    root = tk.Tk()
    # Hide the Tkinter root so we only get the file dialog.
    root.withdraw()

    start_dir = filedialog.askdirectory()
    print( "Starting search at: " + start_dir )
    print( "Cleaning shortcuts to drives: " + str(args.clean_drives) )

    start_time = time.time()

    total_count = 0
    total_size = 0
    dirs_to_search = [ start_dir ]
    while len( dirs_to_search ) > 0:
        dir_to_search = dirs_to_search.pop(0)

        try:
            for filename in os.listdir( dir_to_search ):
                path = os.path.join( dir_to_search, filename )
                if os.path.isfile( path ):
                    shortcut = None
                    try:
                        shell = client.Dispatch("WScript.Shell")
                        shortcut = shell.CreateShortCut( path )
                    except com_error as e:
                        # Not a shortcut file.
                        pass
                    if shortcut:
                        broken = False
                        if is_target_drive_missing( shortcut ):
                            # Possible the drive is just disconnected, so leave the
                            # shortcut be.
                            drive, _ = os.path.splitdrive( shortcut.TargetPath )
                            print("Found shortcut to missing drive " + drive + " at: " + path)
                            if drive in args.clean_drives:
                                print("Treating as broken because " + drive + " is in clean drives list.")
                                broken = True
                        else:
                            broken = is_broken_shortcut( shortcut )

                        if broken:
                            total_size += os.path.getsize( path )
                            total_count += 1
                            if args.clean:
                                os.remove( path )
                            else:
                                print("Found broken shortcut at: " + path)
                elif os.path.isdir( path ):
                    dirs_to_search.append( path )
        except PermissionError as e:
            print(e)

    print("Took %s seconds to run." % (time.time() - start_time))
    print("Found %s broken shortcuts using %s total bytes." % (total_count, total_size))

if __name__=="__main__":
    main()
