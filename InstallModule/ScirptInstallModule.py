import sys
import subprocess
import pkg_resources

print("=============================Scitpt Install And Update All Module python=======================================")
print("")
"""
print("=============================Updating All Module python================================")
print("")
packages = [dist.project_name for dist in pkg_resources.working_set]
subprocess.call("pip install --upgrade " + ' '.join(packages), shell=True)
print("")
print("==============================Finished Updating All Module=======================================")
print("")
"""
#moduleName = input("install Module: ")
print("")
print("=============================Installing Module python================================")
print("")
# implement pip as a subprocess:
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas','--upgrade'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl','--upgrade'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'click','--upgrade'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'seleniumbase','--upgrade'])
print("")
print("==============================Finished Install=======================================")
input('Press any key to Exit program.')
    
