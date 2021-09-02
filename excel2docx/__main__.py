#By: Nabeel Kahlil Maulana

import sys
import os

from .gui import app as GUIApp
from .app import windows_to_unix_path, command_line_app

if __name__ == '__main__':
    argv = sys.argv
    
    if 'gui' in argv:
        GUIApp.mainloop()
        exit()

    is_uac_sheet = True
    if '--uac-tc' in argv:
        is_uac_sheet = False
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

    command_line_app(source_location, target_directory, uac_sheet=is_uac_sheet)
