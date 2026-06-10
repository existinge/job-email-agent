@echo off
cd /d C:\Users\exist\Downloads\job_email_agent_openrouter\job_email_agent_openrouter
py job_email_agent.py processed-status
py job_email_agent.py review --days 14 --max 30
py job_email_agent.py processed-status
pause
