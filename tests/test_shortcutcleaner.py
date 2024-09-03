from pathlib import WindowsPath
import pytest
from shortcutcleaner.shortcutcleaner import *
import win32com.client

def parse_clean_drives():
    assert parse_clean_drives( [ "a", "B", "c:", "D:" ] ) == [ "A:", "B:", "C:", "D:" ]
    assert parse_clean_drives( [ "abc", ";:,", "c:", "" ] ) == [ "C:" ]
    assert parse_clean_drives( [] ) == []

def test_is_file_shortcut():
    file_path = WindowsPath() / "testfile.lnk"
    url_path = WindowsPath() / "testfile.url"
    shell = win32com.client.Dispatch("WScript.Shell")
    file_shortcut = shell.CreateShortCut( str( file_path ) )
    url_shortcut = shell.CreateShortCut( str( url_path ) )
    assert is_file_shortcut(str(file_path)) == True
    assert is_file_shortcut(str(url_path)) == False
    assert is_file_shortcut(file_path) == True
    assert is_file_shortcut(url_path) == False
    assert is_file_shortcut(file_shortcut) == True
    assert is_file_shortcut(url_shortcut) == False
    assert is_file_shortcut("testfile.txt") == False
    with pytest.raises(ValueError):
        is_file_shortcut(1)
        is_file_shortcut(b'1')
        is_file_shortcut(True)

def test_is_net_shortcut():
    file_path = WindowsPath() / "testfile.lnk"
    url_path = WindowsPath() / "testfile.url"
    shell = win32com.client.Dispatch("WScript.Shell")
    file_shortcut = shell.CreateShortCut( str( file_path ) )
    url_shortcut = shell.CreateShortCut( str( url_path ) )
    assert is_net_shortcut(str(file_path)) == False
    assert is_net_shortcut(str(url_path)) == True
    assert is_net_shortcut(file_path) == False
    assert is_net_shortcut(url_path) == True
    assert is_net_shortcut(file_shortcut) == False
    assert is_net_shortcut(url_shortcut) == True
    assert is_net_shortcut("testfile.txt") == False
    with pytest.raises(ValueError):
        is_net_shortcut(1)
        is_net_shortcut(b'1')
        is_net_shortcut(True)

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

def test_is_target_drive_missing( tmp_path ):
    working_path = tmp_path / "working_shortcut.lnk"
    broken_path = tmp_path / "broken_shortcut.lnk"
    # TODO: Come up with a better way to choose a drive that doesn't exist.
    different_drive_path = "A:" / tmp_path.relative_to( tmp_path.drive )

    shell = win32com.client.Dispatch("WScript.Shell")
    working_shortcut = shell.CreateShortCut( str( working_path ) )
    working_shortcut.TargetPath = str( tmp_path )
    working_shortcut.save()

    broken_shortcut = shell.CreateShortCut( str( broken_path ) )
    broken_shortcut.TargetPath = str( different_drive_path )
    broken_shortcut.save()

    assert is_target_drive_missing( str( working_path ) ) == False
    assert is_target_drive_missing( working_path ) == False
    assert is_target_drive_missing( working_shortcut ) == False

    assert is_target_drive_missing( str( broken_path ) ) == True
    assert is_target_drive_missing( broken_path ) == True
    assert is_target_drive_missing( broken_shortcut ) == True

    with pytest.raises(ValueError):
        is_target_drive_missing(1)
        is_target_drive_missing(b'1')
        is_target_drive_missing(True)

def test_is_broken_shortcut( tmp_path ):
    target_path = tmp_path / "target_dir" / "target_file"
    target_path.parent.mkdir() # Create temporary directory
    target_path.touch() # Create temporary file
    working_path1 = tmp_path / "working_shortcut1.lnk"
    working_path2 = tmp_path / "working_shortcut2.lnk"
    working_path3 = tmp_path / "working_shortcut3.url"
    empty_path = tmp_path / "empty_shortcut.lnk"
    broken_path2 = tmp_path / "broken_shortcut2.lnk"
    broken_path3 = tmp_path / "broken_shortcut3.url"
    # broken_path4 = tmp_path / "broken_shortcut4.url"
    broken_path5 = tmp_path / "broken_shortcut5.lnk"
    broken_path6 = tmp_path / "broken_shortcut6.lnk"

    shell = win32com.client.Dispatch("WScript.Shell")
    working_shortcut1 = shell.CreateShortCut( str( working_path1 ) )
    working_shortcut1.TargetPath = str( target_path )
    working_shortcut1.save()

    working_shortcut2 = shell.CreateShortCut( str( working_path2 ) )
    working_shortcut2.TargetPath = str( tmp_path / "target_dir" )
    working_shortcut2.save()

    working_shortcut3 = shell.CreateShortCut( str( working_path3 ) )
    working_shortcut3.TargetPath = "https://a_valid_url"
    working_shortcut3.save()

    empty_shortcut = shell.CreateShortCut( str( empty_path ) )
    empty_shortcut.save()

    broken_shortcut2 = shell.CreateShortCut( str( broken_path2 ) )
    broken_shortcut2.TargetPath = str( tmp_path / "not_a_file" )
    broken_shortcut2.save()

    broken_shortcut3 = shell.CreateShortCut( str( broken_path3 ) )
    broken_shortcut3.save()

    # TODO: Find an appropriate way to test with a broken URL shortcut.
    # An error will be raised if you try to assign an invalid URL to a URL
    # shortcut.
    # broken_shortcut4 = shell.CreateShortCut( str( broken_path4 ) )
    # broken_shortcut4.TargetPath = "not_a_valid_url"
    # broken_shortcut4.save()

    broken_shortcut5 = shell.CreateShortCut( str( broken_path5 ) )
    broken_shortcut5.TargetPath = str( tmp_path / "wrong_dir" / "target_file" )
    broken_shortcut5.save()

    broken_shortcut6 = shell.CreateShortCut( str( broken_path6 ) )
    broken_shortcut6.TargetPath = str( tmp_path / "not_a_dir" )
    broken_shortcut6.save()

    assert is_broken_shortcut( str( working_path1 ) ) == False
    assert is_broken_shortcut( working_path1 ) == False
    assert is_broken_shortcut( working_shortcut1 ) == False

    assert is_broken_shortcut( str( working_path2 ) ) == False
    assert is_broken_shortcut( working_path2 ) == False
    assert is_broken_shortcut( working_shortcut2 ) == False

    assert is_broken_shortcut( str( working_path3 ) ) == False
    assert is_broken_shortcut( working_path3 ) == False
    assert is_broken_shortcut( working_shortcut3 ) == False

    assert is_broken_shortcut( str( empty_path ) ) == False
    assert is_broken_shortcut( empty_path ) == False
    assert is_broken_shortcut( empty_shortcut ) == False

    assert is_broken_shortcut( str( broken_path2 ) ) == True
    assert is_broken_shortcut( broken_path2 ) == True
    assert is_broken_shortcut( broken_shortcut2 ) == True

    assert is_broken_shortcut( str( broken_path3 ) ) == True
    assert is_broken_shortcut( broken_path3 ) == True
    assert is_broken_shortcut( broken_shortcut3 ) == True

    # assert is_broken_shortcut( str( broken_path4 ) ) == True

    assert is_broken_shortcut( str( broken_path5 ) ) == True
    assert is_broken_shortcut( broken_path5 ) == True
    assert is_broken_shortcut( broken_shortcut5 ) == True

    assert is_broken_shortcut( str( broken_path6 ) ) == True
    assert is_broken_shortcut( broken_path6 ) == True
    assert is_broken_shortcut( broken_shortcut6 ) == True

    with pytest.raises(ValueError):
        is_broken_shortcut(1)
        is_broken_shortcut(b'1')
        is_broken_shortcut(True)
