import pytest
from shortcutcleaner.shortcutcleaner import *
import win32com.client

def test_is_file_shortcut():
    assert is_file_shortcut("testfile.lnk") == True
    assert is_file_shortcut("testfile.url") == False
    assert is_file_shortcut("testfile.txt") == False

def test_is_net_shortcut():
    assert is_net_shortcut("testfile.lnk") == False
    assert is_net_shortcut("testfile.url") == True
    assert is_net_shortcut("testfile.txt") == False

def test_alt_is_file_shortcut():
    assert alt_is_file_shortcut("testfile.lnk") == True
    assert alt_is_file_shortcut("testfile.url") == False
    assert alt_is_file_shortcut("testfile.txt") == False

def test_alt_is_net_shortcut():
    assert alt_is_net_shortcut("testfile.lnk") == False
    assert alt_is_net_shortcut("testfile.url") == True
    assert alt_is_net_shortcut("testfile.txt") == False

def test_is_valid_url():
    assert is_valid_url('http://www.cwi.nl:80/%7Eguido/Python.html') == True
    assert is_valid_url('https://stackoverflow.com') == True
    assert is_valid_url('/data/Python.html') == False
    assert is_valid_url(532) == False
    assert is_valid_url(u'dkakasdkjdjakdjadjfalskdjfalk') == False

def test_is_broken_shortcut( tmp_path ):
    target_path = tmp_path / "target_dir" / "target_file"
    target_path.parent.mkdir() # Create temporary directory
    target_path.touch() # Create temporary file
    working_path1 = tmp_path / "working_shortcut1.lnk"
    working_path2 = tmp_path / "working_shortcut2.lnk"
    broken_path1 = tmp_path / "broken_shortcut1.lnk"
    broken_path2 = tmp_path / "broken_shortcut2.lnk"
    # broken_path3 = tmp_path / "broken_shortcut3.url"
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

    broken_shortcut1 = shell.CreateShortCut( str( broken_path1 ) )
    broken_shortcut1.save()

    broken_shortcut2 = shell.CreateShortCut( str( broken_path2 ) )
    broken_shortcut2.TargetPath = str( tmp_path / "not_a_file" )
    broken_shortcut2.save()

    # broken_shortcut3 = shell.CreateShortCut( str( broken_path3 ) )
    # broken_shortcut3.save()
    #
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
    assert is_broken_shortcut( str( working_path2 ) ) == False
    assert is_broken_shortcut( str( broken_path1 ) ) == True
    assert is_broken_shortcut( str( broken_path2 ) ) == True
    # assert is_broken_shortcut( str( broken_path3 ) ) == True
    # assert is_broken_shortcut( str( broken_path4 ) ) == True
    assert is_broken_shortcut( str( broken_path5 ) ) == True
    assert is_broken_shortcut( str( broken_path6 ) ) == True
