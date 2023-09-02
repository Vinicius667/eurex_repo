#!/bin/bash

sudo apt update
sudo apt upgrade -y
mkdir ~/Downloads
cd ./Downloads
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb

sudo apt -y install ./google-chrome-stable_current_amd64.deb --fix-broken 

cd ~

git clone https://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@github.com/Vinicius667/eurex_repo.git


cd eurex_repo

pip install -r requirements.txt
#chmod +x install.sh

#git fetch --all
# git reset --hard origin
# https://crontab.guru/#0_5_*_*_1-5 => 0 5 * * 1-5
#sudo visudo
# add:  %sudo ALL=NOPASSWD: /usr/sbin/service cron start
python3 main.py