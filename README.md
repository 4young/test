# test

domain eventlog:

wevtutil qe security /rd:true /f:text /q:"*[system/eventid=4624 and 4623 and 4672]" /r:dc1 /u:administrator /p:password 
