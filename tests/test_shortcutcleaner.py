from shortcutcleaner.shortcutcleaner import *

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
