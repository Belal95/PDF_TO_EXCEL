@echo off;
cd /d "%~dp0";
cd "./src";
npm run dev --o;
# git status;
# read -p "Press Enter to continue..."