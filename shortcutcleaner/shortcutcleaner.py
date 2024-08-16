import argparse
import os
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

def is_file_shortcut( filepath ):
    _, extension = os.path.splitext( filepath )
    return extension == FILE_SHORTCUT_EXT

def is_net_shortcut( filepath ):
    _, extension = os.path.splitext( filepath )
    return extension == NET_SHORTCUT_EXT

def alt_is_file_shortcut( filepath ):
    try:
        shell = client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut( filepath )
        return hasattr( shortcut, 'RelativePath' )
    except com_error as e:
        return False

def alt_is_net_shortcut( filepath ):
    try:
        shell = client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut( filepath )
        return not hasattr( shortcut, 'RelativePath' )
    except com_error as e:
        return False

def is_valid_url( url ):
    try:
        result = urlparse( url )
        return bool( result.scheme and result.netloc )
        # return all( [ result.scheme, result.netloc ] )
    except AttributeError:
        return False

def is_broken_shortcut( filepath ):
    try:
        shell = client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut( filepath )
        if is_file_shortcut( filepath ):
            return not os.path.isfile( shortcut.Targetpath )
        elif is_net_shortcut( filepath ):
            return not is_valid_url( shortcut.Targetpath )
        else:
            # TODO: Report unknown type of shortcut encountered.
            return False
    except com_error as e:
        # Will be raised if filepath is not a shortcut, in which case it can't
        # be a broken shortcut.
        # print( e )
        return False

def search_dir( dir ):

    sub_dirs = []

    for filename in os.listdir( dir ):
        path = os.path.join( dir, filename )
        if os.path.isfile( path ) and is_broken_shortcut( path ):
            print("Found broken shortcut at: " + path)
        elif os.path.isdir( path ):
            sub_dirs.append( path )

    return sub_dirs

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

dirs_to_search = [ start_dir ]
while len( dirs_to_search ) > 0:
    dir_to_search = dirs_to_search.pop(0)
    dirs_to_search += search_dir( dir_to_search )

print("Took %s seconds to run." % (time.time() - start_time))
