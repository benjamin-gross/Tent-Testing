#!/bin/bash
cd ~
cd Desktop/

filename="Tent-Testing"
if [ -d "$filename" ]; then
echo "Starting Testing"
cd Tent-Testing
git pull
python3 testing.py
else 
echo "begining download"
pip3 install openpyxl
pip3 install requests 

cd ~
cd Desktop/
git clone https://github.com/benjamin-gross/Tent-Testing.git
cd Tent-Testing 
cp start.sh ~/

fi
