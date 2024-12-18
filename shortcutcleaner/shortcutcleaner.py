"""
Script to search for and optionally delete broken Windows shortcut files.

The script will ask you to select a starting directory, and then search that
directory and its subdirectories for broken shortcuts. A shortcut is considered
broken if it:
    Targets a file or directory that is not there.
    Targets an invalid URL.
    Targets a file or directory on a detached drive not specified in the
    --removable_drives option.

By default, the script will only report broken shortcuts. Run it with the --delete
option to delete broken shortcuts.
"""
import argparse
import os
from pathlib import Path
import sys
from threading import Event, Thread
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

class UnfamiliarShortcutExtException(Exception):
    """
    Exception for when a CDispatch object returned by shell.CreateShortCut does
    not have a file extension of .lnk or .url.

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

def parse_removable_drives( removable_drives: list[str] ) -> list[str]:
    """
    Given list of characters, parse them to be drive letters and return
    resulting list. Intended to parse user input and so reports cases where user
    input is malformed and ignored.
    """
    parsed_drives = []

    for drive in removable_drives:
        alpha_drive = list( filter( str.isalpha, drive ) )
        if alpha_drive:
            first_letter = alpha_drive.pop(0)
            if len(alpha_drive) > 0:
                print( f"There are multiple drive letter in the input {drive}. The whole input will be ignored." )
            else:
                parsed_drives.append( parse_drive_str(first_letter) )
        else:
            print("There is an empty input which will be ignored.")

    return parsed_drives

def shortcut_has_ext( shortcut: Shortcut, ext: str ) -> bool:
    """
    Given a string, Path, or CDispatch object, return whether shortcut has the
    extension ext.

    Exceptions:
        Raises AttributeError is shortcut is a CDispatch object without a FullName.
        Raises ValueError if shortcut is not a string, Path, or CDispatch object.
    """
    if isinstance( shortcut, str ):
        _, extension = os.path.splitext( shortcut )
        return extension.lower() == ext
    if issubclass( type(shortcut), Path ):
        return shortcut.suffix.lower() == ext
    if isinstance( shortcut, CDispatch ):
        _, extension = os.path.splitext( shortcut.FullName )
        return extension.lower() == ext
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


def get_shortcut_object( shortcut: Shortcut ) -> CDispatch:
    """
    Given a string, Path, or CDispatch object, return a CDispatch object.

    Exceptions:
        Raises com_error if shortcut is not a str or Path that points to a
        shortcut file.
        Raises ValueError if shortcut is not a string, Path, or CDispatch object.
    """
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

    return shortcut

def is_broken_shortcut( shortcut: Shortcut ) -> bool:
    """
    Given a string, Path, or CDispatch object, return whether shortcut is a
    broken shortcut. That is, it targets a file that doesn't exist or an invalid
    URL.

    Exceptions:
        Raises AttributeError if shortcut is a CDispatch object missing a full
        name or target path.
        Raises com_error if shortcut is not a str or Path that points to a
        shortcut file.
        Raises NoTargetPathException if shortcut is a file shortcut with an empty
        target path.
        Raises UnfamiliarShortcutExtException if shortcut does not have recognized
        extension.
    """
    shortcut = get_shortcut_object( shortcut )
    if shortcut_has_ext( shortcut, FILE_SHORTCUT_EXT ):
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
    if shortcut_has_ext( shortcut, NET_SHORTCUT_EXT ):
        return not is_valid_url( shortcut.TargetPath )
    raise UnfamiliarShortcutExtException(
        f"Encountered a CDispatch object that is not a recognized shortcut extension at {shortcut.FullName}.",
        shortcut.FullName
    )

def is_target_drive_missing( shortcut: Shortcut ) -> bool:
    """
    Given a string, Path, or CDispatch object, return whether shortcut targets
    a file on a drive that is not currently connected.

    Exceptions:
        Raises AttributeError if shortcut is a CDispatch object missing a full
        name or target path.
        Raises com_error if shortcut is not a str or Path that points to a
        shortcut file.
        Raises NoTargetPathException if shortcut is a file shortcut without a
        target path.
        Raises UnfamiliarShortcutExtException if shortcut does not have recognized
        extension.
    """
    shortcut = get_shortcut_object( shortcut )
    if shortcut_has_ext( shortcut, FILE_SHORTCUT_EXT ):
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
    if shortcut_has_ext( shortcut, NET_SHORTCUT_EXT ):
        return False
    raise UnfamiliarShortcutExtException(
        f"Encountered a CDispatch object that is not a recognized shortcut extension at {shortcut.FullName}.",
        shortcut.FullName
    )

def search_loop( start_dir: str, delete: bool, removable_drives: list[str], stop_event: Event=None ):
    """
    Search for broken shortcuts in starting directory or subdirectories, and
    either report or delete them based on user input.

    Arguments:
        start_dir: The directory to start the search at.
        delete: Whether or not to delete broken shortcuts that are found.
        removable_drives: A list of drive letters. Shortcuts will not be treated
        as broken if they target these missing drives.
        stop_event: Optionally, a threading.Event object which, if set, will cause
        the loop to exit early. Intended for use when running search_loop in a
        thread.
    """
    print( f"Starting search at {start_dir}." )
    if delete:
        print( "Deleting broken drives." )
    print( f"Ignoring broken shortcuts to these missing drives: {removable_drives}." )

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
                    broken = False
                    try:
                        if is_target_drive_missing( path ):
                            shortcut = get_shortcut_object( path )
                            drive, _ = os.path.splitdrive( shortcut.TargetPath )
                            print(f"Found broken shortcut to missing drive {drive} at {path}.")
                            if drive in removable_drives:
                                print(f"Ignoring this shortcut because {drive} is in removable drives list.")
                            else:
                                broken = True
                        else:
                            broken = is_broken_shortcut( path )
                    # Common case of path not being a shortcut file, so don't need
                    # to report anything.
                    except com_error:
                        pass
                    # AttributeError: the shortcut was missing some important data.
                    # NoTargetPathException: a weird case the script can't handle.
                    # UnfamiliarShortcutExtException: unfamiliar type of shortcut.
                    # In all these cases, the script should report to the user and
                    # not delete the shortcut.
                    except AttributeError as e:
                        print( f"Encountered {e} while processing file at {path}" )
                    except ( NoTargetPathException, UnfamiliarShortcutExtException ) as e:
                        print( str(e) )
                        broken = False
                    if broken:
                        total_size += os.path.getsize( path )
                        total_count += 1
                        if delete:
                            os.remove( path )
                        else:
                            print(f"Found broken shortcut at {path}.")
                elif os.path.isdir( path ):
                    dirs_to_search.append( path )

                if stop_event and stop_event.is_set():
                    break

        except PermissionError as e:
            print(e)

        if stop_event and stop_event.is_set():
            break

    print(f"Took {(time.time() - start_time)} seconds to run.")
    print(f"Found {total_count} broken shortcuts using {total_size} total bytes.")

class TextRedirector():
    """
    An object which stores a TKinter text area and writes to that area.
    Intended to be assigned to sys.stdout.

    Attributes:
        text_area: The TKinter text widget to write to.
    """
    def __init__( self, text_area ):
        self.text_area = text_area

    def write( self, string ):
        """
        Given a string, writes it to the text area.
        """
        self.text_area.config( state=tk.NORMAL ) # Enable writing.
        self.text_area.insert( tk.END, string )
        self.text_area.config( state=tk.DISABLED ) # Disable writing again.
        self.text_area.update_idletasks() # Call for the text area to be updated.

    def flush(self):
        pass

class TkinterGUI(ttk.Frame):
    """
    Frame which holds the TKinter GUI and manages the data it represents. Stores
    all child widgets as attributes.

    Attributes:
        parent: The parent TKinter widget of this TkinterGUI.
        start_dir_var: The StringVar which holds the path to the starting directory.
        delete_var: The BooleanVar which holds whether --delete is enabled.
        removable_drives: The list of string drive letters which holds --removable_drives.
        add_drive_var: The StringVar which holds the drive letter the user has
            entered before it is submitted to removable_drives.

    Functions:
        browse_start_dir
        add_removable_drive
        validate_add_drive
        remove_removable_drive
        start_search_thread
        run_search_loop
    """

    def __init__(self, parent, delete, removable_drives, **options):
        """
        Initialize the TkinterGUI's attributes, and then initializes all the
        widgets that make up the GUI.
        """
        ttk.Frame.__init__( self, parent, **options )
        self.parent = parent

        self.start_dir_var = tk.StringVar( self, "" )
        self.delete_var = tk.BooleanVar( self, delete )
        self.removable_drives = removable_drives
        self.add_drive_var = tk.StringVar( self, "" )

        self.stop_event = Event()

        self.grid( sticky="NESW")

        self.control_frame = ttk.Frame( self )
        self.control_frame.grid( column=0, row=0, columnspan=2, sticky="NESW" )

        # GUI for selecting starting directory.
        self.start_dir_label = ttk.Label( self.control_frame, text="Starting directory" )
        self.start_dir_label.grid( column=0, row=0 )
        self.start_dir_button = ttk.Button( self.control_frame, text="Select", command=self.browse_start_dir )
        self.start_dir_button.grid( column=0, row=1 )
        self.start_dir_entry = ttk.Entry( self.control_frame, textvariable=self.start_dir_var )
        self.start_dir_entry.grid( column=1, row=1, columnspan=2, sticky="NESW" )

        # GUI for toggling delete.
        self.delete_check = ttk.Checkbutton( self.control_frame, text="Delete broken shortcuts", variable=self.delete_var )
        self.delete_check.grid( column=0, row=2 )

        # GUI for selecting drives to treat as broken.
        self.removable_drives_label = ttk.Label( self.control_frame, text="Removable drives" )
        self.removable_drives_label.grid( column=0, row=3 )

        self.add_drive_button = ttk.Button( self.control_frame, text="Add drive", command=self.add_removable_drive )
        self.add_drive_button.grid( column=0, row=4 )

        validate_add_drive_wrapper = (parent.register(self.validate_add_drive), '%P')
        self.add_drive_entry = ttk.Entry(
            self.control_frame,
            textvariable=self.add_drive_var,
            validate="key",
            validatecommand=validate_add_drive_wrapper
        )
        self.add_drive_entry.grid( column=1, row=4 )

        self.removable_drive_frame = ttk.Frame( self.control_frame, padding=10 )
        self.removable_drive_frame.grid( column=2, row=3, rowspan=4 )

        for drive in removable_drives:
            drive_frame = RemovableDrive( self.removable_drive_frame, drive )
            drive_frame.bind( "<Destroy>", self.remove_removable_drive )
            drive_frame.pack()

        # Button for starting search loop.
        self.run_button = ttk.Button( self.control_frame, text="Run", command=self.start_search_thread )
        self.run_button.grid( column=0, row=5 )

        # Text box to display result of search loop.
        self.text_area = tk.Text( self, state=tk.DISABLED )
        self.text_area.grid( column=0, row=1, sticky="NESW" )

        self.text_scroller = ttk.Scrollbar( self )
        self.text_scroller.grid( column=1, row=1, sticky="NESW" )

        self.text_scroller.config( command=self.text_area.yview )
        self.text_area.config( yscrollcommand=self.text_scroller.set )

        # Store original stdout object so it can be restored later.
        self.old_stdout = sys.stdout
        sys.stdout = TextRedirector( self.text_area )

        # Configure grid to expand to fill the full window.
        self.parent.rowconfigure(0, weight=1)
        self.parent.columnconfigure (0, weight=1)

        self.rowconfigure(0, weight=1)
        # Preference expanding the text area vertically.
        self.rowconfigure(1, weight=3)
        self.columnconfigure (0, weight=1)

        # Let the starting dir entry expand horizontally.
        self.control_frame.columnconfigure(2, weight=1)
        # Let the 6th row of the control frame, which only contains the
        # removable_drive_frame, expand vertically.
        self.control_frame.rowconfigure(6, weight=1)

    def destroy(self):
        """
        Restores the original stdout object and sets an event signalling the
        potential search_loop_thread to stop before destroying the TkinterGUI.
        """
        sys.stdout = self.old_stdout
        self.stop_event.set()
        super().destroy()

    def browse_start_dir(self):
        """
        Opens a filedialog and inserts the result into the start_dir_entry.
        This will change the value of the start_dir_var.
        """
        start_dir = filedialog.askdirectory()
        self.start_dir_entry.delete(0, tk.END)
        self.start_dir_entry.insert(tk.END, start_dir)

    def validate_add_drive( self, entry_input ):
        """
        Given an input, returns whether input is a valid drive letter to add.
        Returns true if input is empty, or a single alphabetic character. Returns
        false if input is more than one character, non-alphabetic, or corresponds
        to a drive letter already in removable_drives.
        Intended to be used as the validatecommand of a TKinter Entry widget.
        """
        if not entry_input:
            return True
        if len(entry_input) > 1 or not entry_input.isalpha():
            return False
        if parse_drive_str( entry_input ) in self.removable_drives:
            return False
        return True

    def add_removable_drive(self):
        """
        Gets the value from the add_drive_var. If it is valid and isn't empty,
        parses the entry, adds it to removable_drives, and creates a RemovableDrive
        frame for the new drive. Always empties add_drive_var.
        """
        drive_to_add = self.add_drive_var.get()
        if drive_to_add and self.validate_add_drive( drive_to_add ):
            parsed_drive = parse_drive_str( drive_to_add )
            self.removable_drives.append( parsed_drive )
            drive_frame = RemovableDrive( self.removable_drive_frame, parsed_drive )
            drive_frame.bind( "<Destroy>", self.remove_removable_drive )
            drive_frame.pack()
        self.add_drive_var.set("")

    def remove_removable_drive( self, event ):
        """
        Given a Destroy event, remove from removable_drives the drive of the destroyed
        widget.
        Intended to be used as a callback function for when a RemovableDrive frame
        is destroyed.
        """
        self.removable_drives.remove( event.widget.drive )

    def start_search_thread(self):
        """
        A wrapper function to disable the run_button and start the thread which
        will run the search loop.
        """
        self.run_button.config( state=tk.DISABLED )
        search_loop_thread = Thread(target=self.run_search_loop)
        search_loop_thread.start()

    def run_search_loop(self):
        """
        A wrapper function to run the search_loop and then reactivate the run_button.
        Intended to be ran in a seperate thread.
        """
        search_loop( self.start_dir_var.get(), self.delete_var.get(), self.removable_drives, stop_event=self.stop_event )
        # Don't need to reactivate the run_button if the TKinterGUI has been closed
        # stop_event is set.
        if not self.stop_event.is_set():
            self.run_button.config( state=tk.NORMAL )

class RemovableDrive(ttk.Frame):
    """
    Frame which represents a drive letter that was added to removable_drives. Includes
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
        '--delete',
        help='Delete broken shortcuts that are found (default: report broken shortcuts).',
        action='store_true',
    )
    parser.add_argument(
        '--removable_drives',
        help='''A list of drive letters. If shortcuts target these missing drives,
        and thus are broken, they will be ignored. Strings with multiple letters
        will be ignored. Intended for shortcuts to drives which are not always
        connected.''',
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

    args.removable_drives = parse_removable_drives( args.removable_drives )

    # Start building Tkinter window
    root = tk.Tk()

    if args.no_gui:
        # Hide the Tkinter root so we only get the file dialog.
        root.withdraw()
        start_dir = filedialog.askdirectory()
        search_loop( start_dir, args.delete, args.removable_drives )
        return

    gui = TkinterGUI( root, args.delete, args.removable_drives, padding=10 )
    gui.parent.mainloop()

if __name__=="__main__":
    main()
