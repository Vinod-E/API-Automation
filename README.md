# Automation :pizza:
Delete the data from automation tenant then run the duplication_rule script afterwards upload candidates


"----------------------"
"DB connection Scripts "
"----------------------"

Eligibility_Criteria.py

Password_Policy.py

** ----------------- details ------------**

New Check Out
-------------
vinodkumar@vinod:~$ mkdir hirepro_automation
vinodkumar@vinod:~$ cd hirepro_automation/
vinodkumar@vinod:~/hirepro_automation$ git clone git@github.com:Vinod-E/API-Automation.git
vinodkumar@vinod:~/hirepro_automation$ cd API-Automation/
vinodkumar@vinod:~/hirepro_automation/API-Automation$ ls

Check python versions and pip
-----------------------------
vinodkumar@vinod:~/hirepro_automation$ python
vinodkumar@vinod:~/hirepro_automation$ pip

Download the get-pip.py from google
-----------------------------------
vinodkumar@vinod:~/hirepro_automation$ sudo python3.7 ~/Desktop/get-pip.py 
vinodkumar@vinod:~/hirepro_automation$ sudo python3.7 ~/Desktop/get-pip.py 
vinodkumar@vinod:~/hirepro_automation$ sudo pip3.7 install virtualenv

xyz = any name which you like
-----------------------------
vinodkumar@vinod:~/hirepro_automation$ virtualenv --python=python3.7 xyz

activate your virtual environment
---------------------------------
vinodkumar@vinod:~/hirepro_automation$ source api-automation/bin/activate

vinodkumar@vinod:~/hirepro_automation/API-Automation$ source  ../api-automation/bin/activate
(api-automation) vinodkumar@vinod:~/hirepro_automation/API-Automation$ pip

Requirement.txt is packages which we required in framework
----------------------------------------------------------
(api-automation) vinodkumar@vinod:~/hirepro_automation/API-Automation$ pip3.7 install requirements.txt 

Remove folder
-------------
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git rm -rf venv/
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git status 
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git reset 
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git status
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git reset HEAD
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git add .
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git status
vinodkumar@vinod:~/hirepro_automation/API-Automation$ git commit -m "removed virtual env"