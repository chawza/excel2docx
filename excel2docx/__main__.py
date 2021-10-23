#By: Nabeel Kahlil Maulana

import sys
import os

from .gui import app as GUIApp
from .app import windows_to_unix_path, command_line_app
from .story import SCENARIO_MODE

def print_help():
    print('--help to print help')
    print('usage:\t<source> [<destination> <mode>]')
    print('')
    print('<source>\texcel file full path')
    print('<destination>\t(optional) output .docx full path location')
    print('<mode>\t\t(optional) read mode for reading scenario title')
    print('\t--uac-comment (default)')
    print('\t--uac-sheet')
    print('\t--uac-tc')

if __name__ == '__main__':
    argv = sys.argv

    if '--help' in argv or 'help' in argv or len(argv) == 1:
        print_help()
        exit()
    
    if 'gui' in argv:
        GUIApp.mainloop()
        exit()

    scenario_read_mode = '--uac-comment'
    if '--uac-sheet' in argv:
        scenario_read_mode = SCENARIO_MODE.UAC_COMMENT
        argv.remove('--uac-sheet')
    elif '--uac-comment' in argv:
        scenario_read_mode = SCENARIO_MODE.UAC_COMMENT
        argv.remove('--uac-comment')
    elif '--uac-tc' in argv:
        scenario_read_mode = SCENARIO_MODE.UAC_TC
        argv.remove('--uac-tc')

    try:
        source_location = argv[1]
        target_directory = argv[2]
    except IndexError:
        if len(argv) < 2:
            raise Exception("argument required: target source is not defined")
        if len(argv) < 3:
            target_directory = os.path.join(os.getcwd(), 'results')
    
    source_location = windows_to_unix_path(source_location)
    target_directory = windows_to_unix_path(target_directory)

    command_line_app(source_location, target_directory, scenario_read_mode=scenario_read_mode)
