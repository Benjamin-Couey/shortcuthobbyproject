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

class NoTargetPathException(Exception):
    """
    Exception for when a .lnk shortcut has an empty TargetPath attribute.

    Attributes:
        message: The error message.
        path: Path to the shortcut file which caused the exception.
    """

    def __init__(self, message, path):
        super().__init__(message)
        self.path = path

def parse_drive_str( drive: str ) -> str:
    """
    Given string, parse it to be a drive letter and return. Empty strings, or
    strings containing multiple characters will return None.
    """
    alpha_drive = list( filter( str.isalpha, drive ) )
    if not len(alpha_drive) == 1:
        return None
    return alpha_drive[0].upper() + ":"

def parse_clean_drives( clean_drives: list[str] ) -> list[str]:
    """
    Given list of characters, parse them to be drive letters and return
    resulting list. Intended to parse user input and so reports cases where user
    input is malformed and ignored.
    """
    parsed_drives = []

    for drive in clean_drives:
        if len(drive) > 0:
            alpha_drive = list( filter( str.isalpha, drive ) )
            first_letter = alpha_drive.pop(0)
            if len(alpha_drive) > 0:
                print( f"There are multiple drive letter in the input {drive}. The whole input will be ignored." )
            else:
                parsed_drives.append( parse_drive_str(first_letter) )
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
        Raises NoTargetPathException if shortcut is a file shortcut without a
        target path.
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
                # It is possible for .lnk files to have a TargetPath that points
                # to a URL or a program like Control Panel not in the file system
                # (I do not fully understand this second case). In these cases,
                # the TargetPath of the shortcut object returned by shell.CreateShortCut
                # will be empty (even if the shortcut itself works) and the sciprt
                # won't be able to differentiate between a working or broken
                # shortcut.
                # TODO: Figure out a way to actually get the target path of .lnk files
                # in these corner cases and more accurately determine if they are
                # valid.
                raise NoTargetPathException(
                    f"Encountered a CDispatch object with an empty TargetPath at {shortcut.FullName}.",
                    shortcut.FullName
                )
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
        Raises NoTargetPathException if shortcut is a file shortcut without a
        target path.
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
                # It is possible for .lnk files to have a TargetPath that points
                # to a URL or a program like Control Panel not in the file system
                # (I do not fully understand this second case). In these cases,
                # the TargetPath of the shortcut object returned by shell.CreateShortCut
                # will be empty (even if the shortcut itself works). While in these
                # two cases, the shortcut is not targeting a missing drive,
                # because I do not fully understand this behavior, I will treat
                # it as exceptional so it does not become harder to track later.
                raise NoTargetPathException(
                    f"Encountered a CDispatch object with an empty TargetPath at {shortcut.FullName}.",
                    shortcut.FullName
                )
            drive, _ = os.path.splitdrive( shortcut.TargetPath )
            return not os.path.exists( drive )
        if is_net_shortcut( shortcut ):
            return False
        print("Encountered a CDispatch object that is not a recognized shortcut type.")
        return False
    except AttributeError as e:
        print(e)
        return False

def search_loop( start_dir: str, clean: bool, clean_drives: list[str] ):
    """
    Search for broken shortcuts in starting directory or subdirectories, and
    either report or delete them based on user input.

    Arguments:
        start_dir: The directory to start the search at.
        clean: Whether or not to delete broken shortcuts that are found.
        clean_dirves: A list of drive letters. If shortcuts target these missing
        drives, they will be treated as broken shortcuts.
    """
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
                        try:
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
                        except NoTargetPathException as e:
                            print( str(e) )
                            # Since the script can't currently handle this weird
                            # case, it should report it to the user and not try
                            # to clean the shortcut.
                            broken = False
                            pass

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


class TkinterGUI(ttk.Frame):
    """
    Frame which holds the TKinter GUI and manages the data it represents. Stores
    all child widgets as attributes.

    Attributes:
        parent: The parent TKinter widget of this TkinterGUI.
        start_dir_var: The StringVar which holds the path to the starting directory.
        clean_var: The BooleanVar which holds whether --clean is enabled.
        clean_drives: The list of string drive letters which holds --clean_drives.
        add_drive_var: The StringVar which holds the drive letter the user has
            entered before it is submitted to clean_drives.

    Functions:
        browse_start_dir
        add_clean_drive
        validate_add_drive
        remove_clean_drive
        run_search_loop
    """

    def __init__(self, parent, clean, clean_drives, **options):
        """
        Initialize the TkinterGUI's attributes, and then initializes all the
        widgets that make up the GUI.
        """
        ttk.Frame.__init__( self, parent, **options )
        self.parent = parent

        self.start_dir_var = tk.StringVar( self, "" )
        self.clean_var = tk.BooleanVar( self, clean )
        self.clean_drives = clean_drives
        self.add_drive_var = tk.StringVar( self, "" )

        self.grid()

        # GUI for selecting starting directory.
        self.start_dir_label = ttk.Label( self, text="Starting directory" )
        self.start_dir_label.grid( column=0, row=0 )
        self.start_dir_entry = ttk.Entry( self, textvariable=self.start_dir_var )
        self.start_dir_entry.grid( column=0, row=1 )
        self.start_dir_button = ttk.Button( self, text="Select", command=self.browse_start_dir )
        self.start_dir_button.grid( column=0, row=2 )

        # GUI for toggling clean.
        self.clean_check = ttk.Checkbutton( self, text="Clean broken shortcuts", variable=self.clean_var )
        self.clean_check.grid( column=0, row=3 )

        # GUI for selecting drives to treat as broken.
        self.clean_drives_label = ttk.Label( self, text="Clean drives" )
        self.clean_drives_label.grid( column=0, row=4 )
        validate_add_drive_wrapper = (parent.register(self.validate_add_drive), '%P')
        self.add_drive_entry = ttk.Entry(
            self,
            textvariable=self.add_drive_var,
            validate="key",
            validatecommand=validate_add_drive_wrapper
        )
        self.add_drive_entry.grid( column=0, row=5 )

        self.add_drive_button = ttk.Button( self, text="Add drive", command=self.add_clean_drive )
        self.add_drive_button.grid( column=0, row=6 )

        self.clean_drive_frame = ttk.Frame( self, padding=10 )
        self.clean_drive_frame.grid( column=0, row=7 )

        for drive in clean_drives:
            drive_frame = RemovableDrive( self.clean_drive_frame, drive )
            drive_frame.bind( "<Destroy>", self.remove_clean_drive )
            drive_frame.pack()

        # Button for starting search loop.
        self.run_button = ttk.Button( self, text="Run", command=self.run_search_loop )
        self.run_button.grid( column=0, row=8 )


    def browse_start_dir(self):
        """
        Opens a filedialog and inserts the result into the start_dir_entry.
        This will change the value of the start_dir_var.
        """
        start_dir = filedialog.askdirectory()
        self.start_dir_entry.insert(tk.END, start_dir)

    def validate_add_drive( self, entry_input ):
        """
        Given an input, returns whether input is a valid drive letter to add.
        Returns true if input is empty, or a single alphabetic character. Returns
        false if input is more than one character, non-alphabetic, or corresponds
        to a drive letter already in clean_drives.
        Intended to be used as the validatecommand of a TKinter Entry widget.
        """
        if not entry_input:
            return True
        if len(entry_input) > 1 or not entry_input.isalpha():
            return False
        if parse_drive_str( entry_input ) in self.clean_drives:
            return False
        return True

    def add_clean_drive(self):
        """
        Gets the value from the add_drive_var. If it isn't empty, parses the entry,
        adds it to clean_drives, and creates a RemovableDrive frame for the new
        drive. Always empties add_drive_var.
        """
        drive_to_add = self.add_drive_var.get()
        if drive_to_add:
            parsed_drive = parse_drive_str( drive_to_add )
            self.clean_drives.append( parsed_drive )
            drive_frame = RemovableDrive( self.clean_drive_frame, parsed_drive )
            drive_frame.bind( "<Destroy>", self.remove_clean_drive )
            drive_frame.pack()
        self.add_drive_var.set("")

    def remove_clean_drive( self, event ):
        """
        Given a Destroy event, remove from clean_drives the drive of the destroyed
        widget.
        Intended to be used as a callback function for when a RemovableDrive frame
        is destroyed.
        """
        self.clean_drives.remove( event.widget.drive )

    def run_search_loop(self):
        """
        Run the search_loop then destroys the TkinterGUI's parent to close the GUI.
        """
        search_loop( self.start_dir_var.get(), self.clean_var.get(), self.clean_drives )
        self.parent.destroy()

class RemovableDrive(ttk.Frame):
    """
    Frame which represents a drive letter that was added to clean_drives. Includes
    a button to destroy the frame.

    Attributes:
        drive: The string drive letter this RemovableDrive represents.
        parent: The parent TKinter widget of this RemovableDrive.
        label: The TKinter label which displays the drive letter.
        button: The TKinter button which destroys teh RemovableDrive frame.
    """
    def __init__( self, parent, drive, **options ):
        self.drive = drive
        ttk.Frame.__init__( self, parent, **options )
        self.parent = parent
        self.label = ttk.Label( self, text=drive )
        self.label.grid(row=0, column=0)
        self.button = ttk.Button( self, text="X", command=self.destroy )
        self.button.grid(row=0, column=1)

def main():
    """
    Parse user input, get starting directory, and open GUI.
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
    parser.add_argument(
        '--no_gui',
        help='''Don't open the Tkinter GUI and immediately run the search.''',
        action='store_true',
    )
    args = parser.parse_args()

    args.clean_drives = parse_clean_drives( args.clean_drives )

    # Start building Tkinter window
    root = tk.Tk()

    if args.no_gui:
        # Hide the Tkinter root so we only get the file dialog.
        root.withdraw()
        start_dir = filedialog.askdirectory()
        search_loop( start_dir, args.clean, args.clean_drives )
        return

    gui = TkinterGUI( root, args.clean, args.clean_drives, padding=10 )
    gui.parent.mainloop()

if __name__=="__main__":
    main()
