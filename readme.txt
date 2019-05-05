pyinstaller --clean --win-private-assemblies --noupx --win-no-prefer-redirects -F -n "Chips" -i "C:\Users\kerne\PycharmProjects\chips\fish.ico" start.py
pyinstaller --clean -F -n "Chips" -i "C:\Users\kerne\PycharmProjects\chips\fish.ico" start.py
pyinstaller --clean -y -i "C:/Users/kerne/PycharmProjects/chips/fish.ico" -n "Chips" start.py


#update all packages
import pkg_resources
from subprocess import call

packages = [dist.project_name for dist in pkg_resources.working_set]
call("pip install --upgrade " + ' '.join(packages), shell=True)

pyinstaller --clean -F -n "Chips" -i "C:\Users\kerne\PycharmProjects\chips\fish.ico" "C:\Users\kerne\PycharmProjects\chips\start.py"