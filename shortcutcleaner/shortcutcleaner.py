"""
Script to search for and clean broken Windows shortcut files.

The script will ask you to select a starting directory, and then search that
directory and its subdirectories for broken shortcuts. A shortcut is considered
broken if it:
    Targets a file or directory that is not there.
    Targets an invalid URL.
    Targets a file or directory on a detached drive specified in the --clean_drives
    option.

By default, the script will only report broken shortcuts. Run it with the --clean
option to delete broken shortcuts.
"""
import argparse
import os
from pathlib import Path
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from urllib.parse import urlparse

from pywintypes import com_error
from win32com import client
from win32com.client import CDispatch

type Shortcut = str | Path | CDispatch
FILE_SHORTCUT_EXT = '.lnk'
NET_SHORTCUT_EXT = '.url'

def parse_clean_drives( clean_drives: list[str] ) -> bool:
    """
    Given list of characters, parse them to be drive letters and return
    resulting list.
    """
    parsed_drives = []

    for drive in clean_drives:
        if len(drive) > 0:
            alpha_drive = list( filter( str.isalpha, drive ) )
            first_letter = alpha_drive.pop(0)
            if len(alpha_drive) > 0:
                print( f"There are multiple drive letter in the input {drive}. The whole input will be ignored." )
            else:
                parsed_drives.append( first_letter.upper() + ":" )
        else:
            print("There is an empty input which will be ignored.")

    return parsed_drives

def is_file_shortcut( shortcut: Shortcut ) -> bool:
    """
    Given a string, Path, or CDispatch object, return whether shortcut is a .lnk
    file shortcut.

    Exceptions:
        Raises ValueError if shortcut is not a string, Path, or CDispatch object.
    """
    if isinstance( shortcut, str ):
        _, extension = os.path.splitext( shortcut )
        return extension == FILE_SHORTCUT_EXT
    if issubclass( type(shortcut), Path ):
        return shortcut.suffix == FILE_SHORTCUT_EXT
    if isinstance( shortcut, CDispatch ):
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

def is_net_shortcut( shortcut: Shortcut ) -> bool:
    """
    Given a string, Path, or CDispatch object, return whether shortcut is a .url
    net shortcut.

    Exceptions:
        Raises ValueError if shortcut is not a string, Path, or CDispatch object.
    """
    if isinstance( shortcut, str ):
        _, extension = os.path.splitext( shortcut )
        return extension == NET_SHORTCUT_EXT
    if issubclass( type(shortcut), Path ):
        return shortcut.suffix == NET_SHORTCUT_EXT
    if isinstance( shortcut, CDispatch ):
        try:
            _, extension = os.path.splitext( shortcut.FullName )
            return extension == NET_SHORTCUT_EXT
        except AttributeError:
            return False
    else:
        raise ValueError("Not a string, Path, or CDispatch shortcut.")

def is_valid_url( url: str ) -> bool:
    """
    Given a string, return whether it is a valid URL or points to a file on the
    local filesystem.

    Exceptions:
        Raises ValueError if url is not a string.
    """
    if isinstance( url, str ):
        result = urlparse( url )
        # It is possible for a URL shortcut to point to a file on the local
        # filesystem, in which case it should be handled like a file shortcut.
        if result.scheme and result.scheme == "file" and not result.netloc and result.path:
            path = result.path.strip("/")
            return os.path.isfile( path ) or os.path.isdir( path )
        return bool( result.scheme and result.netloc )
    raise ValueError("Not a string.")

def is_broken_shortcut( shortcut: Shortcut ) -> bool:
    """
    Given a string, Path, or CDispatch object, return whether shortcut is a
    broken shortcut. That is, it targets a file that doesn't exist or an invalid
    URL.

    Exceptions:
        Raises ValueError if shortcut is not a string, Path, or CDispatch object.
    """
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
                print(f"Encountered a CDispatch object with an empty TargetPath at {shortcut.FullName}.")
                return False
            return not ( os.path.isfile( shortcut.TargetPath ) or os.path.isdir( shortcut.TargetPath ) )
        if is_net_shortcut( shortcut ):
            return not is_valid_url( shortcut.TargetPath )
        print("Encountered a CDispatch object that is not a recognized shortcut type.")
        return False
    except AttributeError as e:
        print(e)
        return False


def is_target_drive_missing( shortcut: Shortcut ) -> bool:
    """
    Given a string, Path, or CDispatch object, return whether shortcut targets
    a file on a drive that is not currently connected.

    Exceptions:
        Raises ValueError if shortcut is not a string, Path, or CDispatch object.
    """
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
        if is_net_shortcut( shortcut ):
            return False
        print("Encountered a CDispatch object that is not a recognized shortcut type.")
        return False
    except AttributeError as e:
        print(e)
        return False

def main():
    """
    Parse user input, get starting directory, and enter main loop. In main loop,
    search for broken shortcuts in starting directory or subdirectories, and
    either report or delete them based on user input.
    """
    parser = argparse.ArgumentParser(
        prog="shortcutcleaner",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter # Preserve docstring formatting
    )
    parser.add_argument(
        '--clean',
        help='Delete broken shortcuts that are found (default: report broken shortcuts).',
        action='store_true',
    )
    parser.add_argument(
        '--clean_drives',
        help='''A list of drive letters. If shortcuts target these missing drives,
        they will be treated as broken shortcuts. Strings with multiple letters
        will be ignored.''',
        action='store',
        nargs='+',
        default=[]
    )
    args = parser.parse_args()

    args.clean_drives = parse_clean_drives( args.clean_drives )

    # Start building Tkinter window
    root = tk.Tk()

    start_dir_var = tk.StringVar( root, "" )
    clean_var = tk.BooleanVar( root, args.clean )
    clean_drives = args.clean_drives
    add_drive_var = tk.StringVar( root, "" )

    frame = ttk.Frame( root, padding=10 )
    frame.grid()

    ttk.Label( frame, text="Starting directory" ).grid( column=0, row=0 )
    start_dir_entry = ttk.Entry( frame, textvariable=start_dir_var )
    start_dir_entry.grid( column=0, row=1 )

    def browse_start_dir():
        start_dir = filedialog.askdirectory()
        start_dir_entry.insert(tk.END, start_dir)

    start_dir_button = ttk.Button( frame, text="Select", command=browse_start_dir )
    start_dir_button.grid( column=0, row=2 )

    ttk.Checkbutton( frame, text="Clean broken shortcuts", variable=clean_var ).grid( column=0, row=3 )

    def validate_add_drive(input):
        if not input:
            return True
        if len(input) > 1 or not input.isalpha():
            return False
        return True
    validate_add_drive_wrapper = (root.register(validate_add_drive), '%P')

    ttk.Label( frame, text="Clean drives" ).grid( column=0, row=4 )
    add_drive_entry = ttk.Entry(
        frame,
        textvariable=add_drive_var,
        validate="key",
        validatecommand=validate_add_drive_wrapper
    )
    add_drive_entry.grid( column=0, row=5 )

    def add_clean_drive( clean_drives ):
        drive_to_add = add_drive_var.get()
        if drive_to_add:
            clean_drives.append( drive_to_add )
            drive_frame = RemovableDrive( clean_drive_frame, drive_to_add )
            drive_frame.pack()
            print( clean_drives )
        add_drive_var.set("")

    add_drive_button = ttk.Button( frame, text="Add drive", command=lambda: add_clean_drive(clean_drives) )
    add_drive_button.grid( column=0, row=6 )

    class RemovableDrive(ttk.Frame):
        def __init__( self, parent, drive, **options ):
            self.drive = drive
            ttk.Frame.__init__( self, parent, **options )
            ttk.Label( self, text=drive ).grid(row=0, column=0)
            ttk.Button( self, text="X", command=self.remove ).grid(row=0, column=1)

        def remove( self ):
            clean_drives.remove( self.drive )
            self.destroy()

    clean_drive_frame = ttk.Frame( frame, padding=10 )
    clean_drive_frame.grid( column=0, row=7 )

    for drive in clean_drives:
        drive_frame = RemovableDrive( clean_drive_frame, drive )
        drive_frame.pack()

    ttk.Button( frame, text="Run", command=root.destroy ).grid( column=0, row=8 )

    root.mainloop()

    start_dir = start_dir_var.get()
    clean = clean_var.get()
    clean_drives = parse_clean_drives( clean_drives )
    print( f"Starting search at {start_dir}." )
    if clean:
        print( "Cleaning broken drives." )
    print( f"Treating shortcuts to drives as broken: {clean_drives}." )

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
                    except com_error:
                        # Not a shortcut file.
                        pass
                    if shortcut:
                        broken = False
                        if is_target_drive_missing( shortcut ):
                            # Possible the drive is just disconnected, so leave the
                            # shortcut be.
                            drive, _ = os.path.splitdrive( shortcut.TargetPath )
                            print(f"Found shortcut to missing drive {drive} at {path}.")
                            if drive in clean_drives:
                                print(f"Treating as broken because {drive} is in clean drives list.")
                                broken = True
                        else:
                            broken = is_broken_shortcut( shortcut )

                        if broken:
                            total_size += os.path.getsize( path )
                            total_count += 1
                            if clean:
                                os.remove( path )
                            else:
                                print(f"Found broken shortcut at {path}.")
                elif os.path.isdir( path ):
                    dirs_to_search.append( path )
        except PermissionError as e:
            print(e)

    print(f"Took {(time.time() - start_time)} seconds to run.")
    print(f"Found {total_count} broken shortcuts using {total_size} total bytes.")

if __name__=="__main__":
    main()
