import styleframe as sf
import os

tests_commandline_dir = os.path.join(os.path.dirname(sf.__file__),
                                     'command_line',
                                     'tests')
TEST_JSON_FILE = os.path.join(tests_commandline_dir, 'test_json.json')
TEST_JSON_STRING_FILE = os.path.join(tests_commandline_dir, 'test_json_string.json')