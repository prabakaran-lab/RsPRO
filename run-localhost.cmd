@echo off
set "PORT=8001"
set "HOST_URL=http://localhost:8001"
cd /d "%~dp0"
call "C:\Program Files\nodejs\npm.cmd" run dev >> "codex-persistent.out.log" 2>> "codex-persistent.err.log"
