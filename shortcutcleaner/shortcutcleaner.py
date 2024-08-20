import argparse
import os
from pathlib import Path
from pywintypes import com_error
import time
import tkinter as tk
from tkinter import filedialog
from win32com import client
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

def is_file_shortcut( shortcut ):
    if isinstance( shortcut, str ):
        _, extension = os.path.splitext( shortcut )
        return extension == FILE_SHORTCUT_EXT
    elif issubclass( type(shortcut), Path ):
        return shortcut.suffix == FILE_SHORTCUT_EXT
    # TODO: Find a better way to check if the shortcut is a COMObject shortcut
    else:
        try:
            _, extension = os.path.splitext( shortcut.FullName )
            # Since URL shortcuts do not have the RelativePath attribute, an
            # alternative way to check if a COMObject shortcut is a file short
            # cut would be to check for the presence of that attribute.
            # return hasattr( shortcut, 'RelativePath' )
            return extension == FILE_SHORTCUT_EXT
        except AttributeError:
            return False

def is_net_shortcut( shortcut ):
    if isinstance( shortcut, str ):
        _, extension = os.path.splitext( shortcut )
        return extension == NET_SHORTCUT_EXT
    elif issubclass( type(shortcut), Path ):
        return shortcut.suffix == NET_SHORTCUT_EXT
    else:
        try:
            _, extension = os.path.splitext( shortcut.FullName )
            return extension == NET_SHORTCUT_EXT
        except AttributeError:
            return False

def is_valid_url( url ):
    try:
        result = urlparse( url )
        return bool( result.scheme and result.netloc )
        # return all( [ result.scheme, result.netloc ] )
    except AttributeError:
        return False

def is_broken_shortcut( shortcut ):
    try:
        if isinstance( shortcut, str ):
            shell = client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut( shortcut )
        elif issubclass( type(shortcut), Path ):
            shell = client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut( str( shortcut ) )
    except com_error as e:
        # Will be raised if shortcut is not a path to a shortcut, in which case
        # it can't be a broken shortcut.
        # print( e )
        return False
    # TODO: Find a better way to check if the shortcut is a COMObject shortcut
    # instead of just assuming.
    try:
        if is_file_shortcut( shortcut ):
            return not ( os.path.isfile( shortcut.TargetPath ) or os.path.isdir( shortcut.TargetPath ) )
        elif is_net_shortcut( shortcut ):
            return not is_valid_url( shortcut.TargetPath )
        else:
            # TODO: Report unknown type of shortcut encountered.
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
        action='store',
        default=False
    )
    args = parser.parse_args()

    root = tk.Tk()
    # Hide the Tkinter root so we only get the file dialog.
    root.withdraw()

    start_dir = filedialog.askdirectory()
    print( "Starting search at: " + start_dir )

    start_time = time.time()

    total_count = 0
    total_size = 0
    dirs_to_search = [ start_dir ]
    while len( dirs_to_search ) > 0:
        dir_to_search = dirs_to_search.pop(0)

        try:
            for filename in os.listdir( dir_to_search ):
                path = os.path.join( dir_to_search, filename )
                if os.path.isfile( path ) and is_broken_shortcut( path ):
                    if args.clean:
                        os.remove( path )
                    else:
                        print("Found broken shortcut at: " + path)
                    total_size += os.path.getsize( path )
                    total_count += 1
                elif os.path.isdir( path ):
                    dirs_to_search.append( path )
        except PermissionError as e:
            print(e)

    print("Took %s seconds to run." % (time.time() - start_time))
    print("Found %s broken shortcuts using %s total bytes." % (total_count, total_size))

if __name__=="__main__":
    main()
