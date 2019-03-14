#!C:\Users\EcBerry\PycharmProjects\MetRetailFrontEndAutomation\venv\Scripts\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'tap-db2==0.3.3','console_scripts','tap-db2'
__requires__ = 'tap-db2==0.3.3'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('tap-db2==0.3.3', 'console_scripts', 'tap-db2')()
    )
