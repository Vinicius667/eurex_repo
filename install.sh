#!/bin/bash

sudo apt update
sudo apt upgrade -y
mkdir ~/Downloads
cd /Downloads
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo apt -y install ./google-chrome-stable_current_amd64.deb --fix-broken 
