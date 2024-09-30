from pathlib import WindowsPath
import pytest
from shortcutcleaner.shortcutcleaner import *
from shutil import rmtree
from unittest.mock import call, patch
import win32com.client

def test_parse_drive_str():
    assert parse_drive_str( "" ) == None
    assert parse_drive_str( "abc" ) == None
    assert parse_drive_str( ";:" ) == None
    assert parse_drive_str( "c:" ) == "C:"
    assert parse_drive_str( "A" ) == "A:"

def test_parse_clean_drives():
    assert parse_clean_drives( [ "a", "B", "c:", "D:" ] ) == [ "A:", "B:", "C:", "D:" ]
    assert parse_clean_drives( [ "abc", ";:,", "c:", "" ] ) == [ "C:" ]
    assert parse_clean_drives( [] ) == []

def test_shortcut_has_ext_file_shortcut_ext():
    file_path = WindowsPath() / "testfile.lnk"
    url_path = WindowsPath() / "testfile.url"
    upper_case_path = WindowsPath() / "testfile.LNK"
    weird_case_path = WindowsPath() / "testfile.LnK"
    shell = win32com.client.Dispatch("WScript.Shell")
    file_shortcut = shell.CreateShortCut( str( file_path ) )
    url_shortcut = shell.CreateShortCut( str( url_path ) )
    upper_case_shortcut = shell.CreateShortCut( str( upper_case_path ) )
    weird_case_shortcut = shell.CreateShortCut( str( weird_case_path ) )
    assert shortcut_has_ext(str(file_path), FILE_SHORTCUT_EXT) == True
    assert shortcut_has_ext(str(url_path), FILE_SHORTCUT_EXT) == False
    assert shortcut_has_ext(file_path, FILE_SHORTCUT_EXT) == True
    assert shortcut_has_ext(url_path, FILE_SHORTCUT_EXT) == False
    assert shortcut_has_ext(file_shortcut, FILE_SHORTCUT_EXT) == True
    assert shortcut_has_ext(url_shortcut, FILE_SHORTCUT_EXT) == False
    assert shortcut_has_ext("testfile.txt", FILE_SHORTCUT_EXT) == False
    assert shortcut_has_ext(upper_case_shortcut, FILE_SHORTCUT_EXT) == True
    assert shortcut_has_ext(weird_case_shortcut, FILE_SHORTCUT_EXT) == True
    with pytest.raises(ValueError):
        shortcut_has_ext(1, FILE_SHORTCUT_EXT)
        shortcut_has_ext(b'1', FILE_SHORTCUT_EXT)
        shortcut_has_ext(True, FILE_SHORTCUT_EXT)

def test_shortcut_has_ext_net_shortcut_ext():
    file_path = WindowsPath() / "testfile.lnk"
    url_path = WindowsPath() / "testfile.url"
    shell = win32com.client.Dispatch("WScript.Shell")
    file_shortcut = shell.CreateShortCut( str( file_path ) )
    url_shortcut = shell.CreateShortCut( str( url_path ) )
    assert shortcut_has_ext(str(file_path), NET_SHORTCUT_EXT) == False
    assert shortcut_has_ext(str(url_path), NET_SHORTCUT_EXT) == True
    assert shortcut_has_ext(file_path, NET_SHORTCUT_EXT) == False
    assert shortcut_has_ext(url_path, NET_SHORTCUT_EXT) == True
    assert shortcut_has_ext(file_shortcut, NET_SHORTCUT_EXT) == False
    assert shortcut_has_ext(url_shortcut, NET_SHORTCUT_EXT) == True
    assert shortcut_has_ext("testfile.txt", NET_SHORTCUT_EXT) == False
    with pytest.raises(ValueError):
        shortcut_has_ext(1, NET_SHORTCUT_EXT)
        shortcut_has_ext(b'1', NET_SHORTCUT_EXT)
        shortcut_has_ext(True, NET_SHORTCUT_EXT)

def test_is_valid_url( tmp_path ):
    assert is_valid_url('http://www.cwi.nl:80/%7Eguido/Python.html') == True
    assert is_valid_url('https://stackoverflow.com') == True
    target_path = tmp_path / "target_dir" / "target_file"
    target_path.parent.mkdir() # Create temporary directory
    target_path.touch() # Create temporary file
    target_path_url = "file:///" + str( target_path )
    target_dir_url = "file:///" + str( tmp_path / "target_dir" )
    not_a_file_url = "file:///" + str( tmp_path / "not_a_file" )
    assert is_valid_url(str(target_path)) == False
    assert is_valid_url(not_a_file_url) == False
    assert is_valid_url(target_path_url) == True
    assert is_valid_url(target_dir_url) == True
    assert is_valid_url("https://stackoverflow.com/" + str( target_path )) == True
    assert is_valid_url("https://stackoverflow.com/" + str( tmp_path / "not_a_file" )) == True
    with pytest.raises(ValueError):
        is_valid_url(1)
        is_valid_url(b'1')
        is_valid_url(True)

@pytest.fixture(scope="session")
def dir_of_shortcuts( tmp_path_factory ):
    # Make basic directory structure.
    start_dir = tmp_path_factory.mktemp("start")
    sub_dir = start_dir / "sub"
    sub_dir.mkdir()
    # Make target dir and file for .lnk shortcuts.
    target_dir = start_dir / "target_dir"
    target_path = target_dir / "target_file"
    target_dir.mkdir()
    target_path.touch()
    # TODO: Come up with a better way to choose a drive that doesn't exist.
    different_drive_path = "A:" / WindowsPath( *target_path.parts[1:] )

    # Make paths for the various shortcuts to test.
    working_file_lnk_path = start_dir / "working_file_shortcut.lnk"
    working_dir_lnk_path = start_dir / "working_dir_shortcut.lnk"

    working_url_path = start_dir / "working_url_shortcut.url"
    working_file_url_path = start_dir / "working_file_url_shortcut.url"
    working_dir_url_path = start_dir / "working_dir_url_shortcut.url"

    not_a_file_lnk_path = start_dir / "not_a_file_lnk_shortcut.lnk"
    wrong_dir_lnk_path = start_dir / "wrong_dir_lnk_shortcut.lnk"
    not_a_dir_lnk_path = start_dir / "not_a_dir_lnk_shortcut.lnk"
    different_drive_lnk_path = start_dir / "different_drive_lnk_shortcut.lnk"

    # TODO: Find an appropriate way to test with a broken URL shortcut.
    # An error will be raised if you try to assign an invalid URL to a URL
    # shortcut.
    # TODO: Find an appropriate way to test with a shortcut that doesn't have
    # a .lnk or .url extension.
    shortcuts = {}
    not_a_file_url_path = start_dir / "not_a_file_url_shortcut.url"
    wrong_dir_url_path = start_dir / "wrong_dir_url_shortcut.url"
    not_a_dir_url_path = start_dir / "not_a_dir_url_shortcut.url"

    empty_lnk_path = start_dir / "empty_shortcut.lnk"
    empty_url_path = start_dir / "empty_shortcut.url"
    lnk_to_net_path = start_dir / "lnk_to_net.lnk"

    # Create the actual shortcut files.
    shell = win32com.client.Dispatch("WScript.Shell")
    working_file_lnk_shortcut = shell.CreateShortCut( str( working_file_lnk_path ) )
    working_file_lnk_shortcut.TargetPath = str( target_path )
    working_file_lnk_shortcut.save()
    shortcuts["working_file_lnk_shortcut"] = working_file_lnk_shortcut

    working_dir_lnk_shortcut = shell.CreateShortCut( str( working_dir_lnk_path ) )
    working_dir_lnk_shortcut.TargetPath = str( target_dir )
    working_dir_lnk_shortcut.save()
    shortcuts["working_dir_lnk_shortcut"] = working_dir_lnk_shortcut

    working_url_shortcut = shell.CreateShortCut( str( working_url_path ) )
    working_url_shortcut.TargetPath = "https://stackoverflow.com"
    working_url_shortcut.save()
    shortcuts["working_url_shortcut"] = working_url_shortcut

    working_file_url_shortcut = shell.CreateShortCut( str( working_file_url_path ) )
    working_file_url_shortcut.TargetPath = f"file:///{target_path}"
    working_file_url_shortcut.save()
    shortcuts["working_file_url_shortcut"] = working_file_url_shortcut

    working_dir_url_shortcut = shell.CreateShortCut( str( working_dir_url_path ) )
    working_dir_url_shortcut.TargetPath = f"file:///{target_dir}"
    working_dir_url_shortcut.save()
    shortcuts["working_dir_url_shortcut"] = working_dir_url_shortcut

    not_a_file_lnk_shortcut = shell.CreateShortCut( str( not_a_file_lnk_path ) )
    not_a_file_lnk_shortcut.TargetPath = str( start_dir / "not_a_file" )
    not_a_file_lnk_shortcut.save()
    shortcuts["not_a_file_lnk_shortcut"] = not_a_file_lnk_shortcut

    wrong_dir_lnk_shortcut = shell.CreateShortCut( str( wrong_dir_lnk_path ) )
    wrong_dir_lnk_shortcut.TargetPath = str( start_dir / "wrong_dir" / "target_file" )
    wrong_dir_lnk_shortcut.save()
    shortcuts["wrong_dir_lnk_shortcut"] = wrong_dir_lnk_shortcut

    not_a_dir_lnk_shortcut = shell.CreateShortCut( str( not_a_dir_lnk_path ) )
    not_a_dir_lnk_shortcut.TargetPath = str( start_dir / "not_a_dir" )
    not_a_dir_lnk_shortcut.save()
    shortcuts["not_a_dir_lnk_shortcut"] = not_a_dir_lnk_shortcut

    different_drive_lnk_shortcut = shell.CreateShortCut( str( different_drive_lnk_path ) )
    different_drive_lnk_shortcut.TargetPath = str( different_drive_path )
    different_drive_lnk_shortcut.save()
    shortcuts["different_drive_lnk_shortcut"] = different_drive_lnk_shortcut

    not_a_file_url_shortcut = shell.CreateShortCut( str( not_a_file_url_path ) )
    not_a_file_url_shortcut.TargetPath = str( start_dir / "not_a_file" )
    not_a_file_url_shortcut.save()
    shortcuts["not_a_file_url_shortcut"] = not_a_file_url_shortcut

    wrong_dir_url_shortcut = shell.CreateShortCut( str( wrong_dir_url_path ) )
    wrong_dir_url_shortcut.TargetPath = str( start_dir / "wrong_dir" / "target_file" )
    wrong_dir_url_shortcut.save()
    shortcuts["wrong_dir_url_shortcut"] = wrong_dir_url_shortcut

    not_a_dir_url_shortcut = shell.CreateShortCut( str( not_a_dir_url_path ) )
    not_a_dir_url_shortcut.TargetPath = str( start_dir / "not_a_dir" )
    not_a_dir_url_shortcut.save()
    shortcuts["not_a_dir_url_shortcut"] = not_a_dir_url_shortcut

    empty_lnk_shortcut = shell.CreateShortCut( str( empty_lnk_path ) )
    empty_lnk_shortcut.save()
    shortcuts["empty_lnk_shortcut"] = empty_lnk_shortcut

    empty_url_shortcut = shell.CreateShortCut( str( empty_url_path ) )
    empty_url_shortcut.save()
    shortcuts["empty_url_shortcut"] = empty_url_shortcut

    lnk_to_net_shortcut = shell.CreateShortCut( str( lnk_to_net_path ) )
    lnk_to_net_shortcut.TargetPath = "https://stackoverflow.com"
    lnk_to_net_shortcut.save()
    shortcuts["lnk_to_net_shortcut"] = lnk_to_net_shortcut

    yield start_dir, shortcuts

    rmtree( start_dir )

def test_get_shortcut_object( dir_of_shortcuts ):
    _, shortcuts = dir_of_shortcuts
    assert get_shortcut_object( shortcuts["working_file_lnk_shortcut"] ) == shortcuts["working_file_lnk_shortcut"]
    # TODO: Write function to compare CDispatch objects.
    shortcut_from_strpath = get_shortcut_object( shortcuts["working_file_lnk_shortcut"].FullName )
    assert isinstance( shortcut_from_strpath, CDispatch )
    assert shortcut_from_strpath.TargetPath == shortcuts["working_file_lnk_shortcut"].TargetPath

    shortcut_from_path = get_shortcut_object( Path( shortcuts["working_file_lnk_shortcut"].FullName ) )
    assert isinstance( shortcut_from_path, CDispatch )
    assert shortcut_from_path.TargetPath == shortcuts["working_file_lnk_shortcut"].TargetPath

    with pytest.raises(com_error):
        get_shortcut_object( shortcuts["working_file_lnk_shortcut"].TargetPath )

    with pytest.raises(ValueError):
        get_shortcut_object(1)
        get_shortcut_object(b'1')
        get_shortcut_object(True)

def test_is_target_drive_missing( dir_of_shortcuts ):
    _, shortcuts = dir_of_shortcuts

    assert is_target_drive_missing( shortcuts["working_file_lnk_shortcut"] ) == False
    assert is_target_drive_missing( shortcuts["working_dir_lnk_shortcut"] ) == False
    # Broken paths, but not on a missing drive.
    assert is_target_drive_missing( shortcuts["not_a_file_lnk_shortcut"] ) == False
    assert is_target_drive_missing( shortcuts["not_a_file_url_shortcut"] ) == False
    # Regardless of their target, URL shortcuts aren't targeting a missing drive.
    assert is_target_drive_missing( shortcuts["working_url_shortcut"] ) == False
    assert is_target_drive_missing( shortcuts["empty_url_shortcut"] ) == False

    assert is_target_drive_missing( shortcuts["different_drive_lnk_shortcut"] ) == True

    # String and Path path to shortcut should also work.
    assert is_target_drive_missing( shortcuts["working_file_lnk_shortcut"].FullName ) == False
    assert is_target_drive_missing( Path( shortcuts["working_file_lnk_shortcut"].FullName ) ) == False

    with pytest.raises(ValueError):
        is_target_drive_missing(1)
        is_target_drive_missing(b'1')
        is_target_drive_missing(True)

    with pytest.raises(NoTargetPathException):
        is_target_drive_missing( shortcuts["empty_lnk_shortcut"] )
        is_target_drive_missing( shortcuts["lnk_to_net_shortcut"] )

def test_is_broken_shortcut( dir_of_shortcuts ):
    _, shortcuts = dir_of_shortcuts

    assert is_broken_shortcut( shortcuts["working_file_lnk_shortcut"] ) == False
    assert is_broken_shortcut( shortcuts["working_dir_lnk_shortcut"] ) == False
    assert is_broken_shortcut( shortcuts["working_url_shortcut"] ) == False

    assert is_broken_shortcut( shortcuts["working_file_url_shortcut"] ) == False
    assert is_broken_shortcut( shortcuts["working_dir_url_shortcut"] ) == False

    assert is_broken_shortcut( shortcuts["not_a_file_lnk_shortcut"] ) == True
    assert is_broken_shortcut( shortcuts["wrong_dir_lnk_shortcut"] ) == True
    assert is_broken_shortcut( shortcuts["not_a_dir_lnk_shortcut"] ) == True
    assert is_broken_shortcut( shortcuts["different_drive_lnk_shortcut"] ) == True

    assert is_broken_shortcut( shortcuts["not_a_file_url_shortcut"] ) == True
    assert is_broken_shortcut( shortcuts["wrong_dir_url_shortcut"] ) == True
    assert is_broken_shortcut( shortcuts["not_a_dir_url_shortcut"] ) == True

    assert is_broken_shortcut( shortcuts["empty_url_shortcut"] ) == True

    # String and Path path to shortcut should also work.
    assert is_target_drive_missing( shortcuts["working_file_lnk_shortcut"].FullName ) == False
    assert is_target_drive_missing( Path( shortcuts["working_file_lnk_shortcut"].FullName ) ) == False

    with pytest.raises(ValueError):
        is_broken_shortcut(1)
        is_broken_shortcut(b'1')
        is_broken_shortcut(True)

    with pytest.raises(NoTargetPathException):
        is_broken_shortcut( shortcuts["empty_lnk_shortcut"] )
        is_broken_shortcut( shortcuts["lnk_to_net_shortcut"] )

@patch('os.path.getsize')
@patch('time.time')
@patch('os.remove')
@patch('builtins.print')
def test_search_loop( mock_print, mock_remove, mock_time, mock_getsize, dir_of_shortcuts ):
    starting_dir, shortcuts = dir_of_shortcuts

    # Mock time() and getsize() for consistent output.
    mock_time.return_value = 1
    mock_getsize.return_value = 1

    # Assert that print is being called with specific output messages. Since this
    # is what the user will see, and it changing unexpectedly could be confusing,
    # it should be verified.
    start_calls = [
        call( f"Starting search at {starting_dir}." ),
        call( f"Treating shortcuts to drives as broken: {[]}." )
    ]

    broken_shortcut_calls = [
        call(f"Found broken shortcut at {shortcuts["not_a_file_lnk_shortcut"].FullName}."),
        call(f"Found broken shortcut at {shortcuts["wrong_dir_lnk_shortcut"].FullName}."),
        call(f"Found broken shortcut at {shortcuts["not_a_dir_lnk_shortcut"].FullName}."),
        call(f"Found broken shortcut at {shortcuts["not_a_file_url_shortcut"].FullName}."),
        call(f"Found broken shortcut at {shortcuts["wrong_dir_url_shortcut"].FullName}."),
        call(f"Found broken shortcut at {shortcuts["not_a_dir_url_shortcut"].FullName}."),
    ]

    exception_calls = [
        call( f"Encountered a CDispatch object with an empty TargetPath at {shortcuts["empty_lnk_shortcut"].FullName}." ),
        call( f"Encountered a CDispatch object with an empty TargetPath at {shortcuts["lnk_to_net_shortcut"].FullName}." )
    ]

    end_calls = [
        call( f"Took {time.time() - time.time()} seconds to run." ),
        call( f"Found {len(broken_shortcut_calls)} broken shortcuts using {len(broken_shortcut_calls)} total bytes." )
    ]

    shortcut_path = shortcuts["different_drive_lnk_shortcut"].FullName
    shortcut_drive, _ = os.path.splitdrive( shortcuts["different_drive_lnk_shortcut"].TargetPath )

    search_loop( starting_dir, delete=False, clean_drives=[] )

    mock_print.assert_has_calls( start_calls, any_order=False )
    mock_print.assert_has_calls( broken_shortcut_calls, any_order=True )
    mock_print.assert_has_calls( exception_calls, any_order=True )
    mock_print.assert_any_call( f"Found shortcut to missing drive {shortcut_drive} at {shortcut_path}." )
    mock_print.assert_has_calls( end_calls, any_order=False )

    mock_print.reset_mock()
    search_loop( starting_dir, delete=True, clean_drives=[] )

    clean_start_calls = [
        call( f"Starting search at {starting_dir}." ),
        call( "Deleting broken drives." ),
        call( f"Treating shortcuts to drives as broken: {[]}." )
    ]

    # Mock os.remove to verify that search_loop calls it when run with the clean
    # option without actually deleting any of the test files.
    remove_calls = [
        call(shortcuts["not_a_file_lnk_shortcut"].FullName),
        call(shortcuts["wrong_dir_lnk_shortcut"].FullName),
        call(shortcuts["not_a_dir_lnk_shortcut"].FullName),
        call(shortcuts["not_a_file_url_shortcut"].FullName),
        call(shortcuts["wrong_dir_url_shortcut"].FullName),
        call(shortcuts["not_a_dir_url_shortcut"].FullName),
    ]

    mock_print.assert_has_calls( clean_start_calls, any_order=False )
    mock_print.assert_has_calls( exception_calls, any_order=True )
    mock_remove.assert_has_calls( remove_calls, any_order=True )
    mock_print.assert_any_call( f"Found shortcut to missing drive {shortcut_drive} at {shortcut_path}." )
    mock_print.assert_has_calls( end_calls, any_order=False )

    # Verify that the shortcut that targets a different drive is treated as broken
    # when the drive is included in clean_drives.
    mock_print.reset_mock()
    search_loop( starting_dir, delete=False, clean_drives=[ shortcut_drive ] )

    clean_drives_start_call = [
        call( f"Starting search at {starting_dir}." ),
        call( f"Treating shortcuts to drives as broken: {[ shortcut_drive ]}." )
    ]

    clean_drives_calls = [
        call( f"Found shortcut to missing drive {shortcut_drive} at {shortcut_path}." ),
        call( f"Treating as broken because {shortcut_drive} is in clean drives list." ),
    ]

    clean_drives_broken_calls = broken_shortcut_calls + [ call( f"Found broken shortcut at {shortcut_path}." ) ]

    clean_drives_end_call = [
        call( f"Took {time.time() - time.time()} seconds to run." ),
        call( f"Found {len(clean_drives_broken_calls)} broken shortcuts using {len(clean_drives_broken_calls)} total bytes." )
    ]

    mock_print.assert_has_calls( clean_drives_start_call, any_order=False )
    mock_print.assert_has_calls( clean_drives_broken_calls, any_order=True )
    mock_print.assert_has_calls( exception_calls, any_order=True )
    mock_print.assert_has_calls( clean_drives_calls, any_order=True )
    mock_print.assert_has_calls( clean_drives_end_call, any_order=False )


@pytest.fixture
def tkinter_gui( tmp_path_factory ):
    root = tk.Tk()
    gui = TkinterGUI( root, False, [], padding=10 )
    yield gui
    root.destroy()

def test_validate_add_drive( tkinter_gui ):
    assert tkinter_gui.validate_add_drive( None ) == True
    assert tkinter_gui.validate_add_drive( "a" ) == True
    assert tkinter_gui.validate_add_drive( "A" ) == True
    assert tkinter_gui.validate_add_drive( "A:" ) == False
    assert tkinter_gui.validate_add_drive( "aa" ) == False
    assert tkinter_gui.validate_add_drive( "1" ) == False
    assert tkinter_gui.validate_add_drive( ":" ) == False
    # A new drive isn't valid if it's already in clean_drives.
    tkinter_gui.clean_drives.append( parse_drive_str( "A:" ) )
    assert tkinter_gui.validate_add_drive( "A" ) == False

@patch('tkinter.ttk.Frame.bind')
def test_add_clean_drive( mock_bind, tkinter_gui ):
    drive_letter = "A"
    assert not tkinter_gui.clean_drive_frame.winfo_children()
    tkinter_gui.add_drive_var.set( drive_letter )
    tkinter_gui.add_clean_drive()
    assert not tkinter_gui.add_drive_var.get()
    assert len( tkinter_gui.clean_drive_frame.winfo_children() ) == 1
    new_removable_drive = tkinter_gui.clean_drive_frame.winfo_children()[0]
    assert new_removable_drive.drive == parse_drive_str( drive_letter )
    # Based on this: https://stackoverflow.com/questions/138029/get-bound-event-handler-in-tkinter
    # there isn't a clean way to get the name of a function bound to a TK widget.
    # For now, easiest to just mock bind to verify that add_clean_drive is hooking
    # up the RemovableDrive to remove_clean_drive correctly.
    mock_bind.assert_called_with( "<Destroy>", tkinter_gui.remove_clean_drive )

    tkinter_gui.add_clean_drive()
    assert not tkinter_gui.add_drive_var.get()
    assert len( tkinter_gui.clean_drive_frame.winfo_children() ) == 1

    tkinter_gui.add_drive_var.set( "A:" )
    tkinter_gui.add_clean_drive()
    assert not tkinter_gui.add_drive_var.get()
    assert len( tkinter_gui.clean_drive_frame.winfo_children() ) == 1
