@echo off
title Naeil Path AI MVP Server
cd /d "C:\Users\USER\Documents\Codex\2026-05-08\https-www-moel-go-kr-info"
echo Starting Naeil Path AI MVP server...
echo.
echo Keep this window open while using http://localhost:4173
echo If this window closes, the site will stop.
echo.
if "%OPENAI_API_KEY%"=="" (
  set /p OPENAI_API_KEY=Paste OpenAI API key for this session, or press Enter to run fallback mode: 
)
"C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" server.mjs
echo.
echo Server stopped or failed to start.
pause
