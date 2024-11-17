# Dependency
python3.8+, tested on python3.8.5 on centos 7.8

# ENV
python3.8 -m pip install -r requirements.txt

# Start lingling
python3.8 /root/lingling/app.py

# Add watchdog
## 1. install watchdog

```
yum install supervisor
```

## 2. Add below config to end of `/etc/supervisord.conf`

```
[program:lingling]
command=/usr/local/bin/python3.8 /root/lingling/app.py
autostart=true
autorestart=true
stderr_logfile=/var/log/lingling.err.log
stdout_logfile=/var/log/lingling.out.log
```

## 3. start watchdog

```
systemctl daemon-reload
systemctl enable supervisord
systemctl start supervisord
```

## 4. start lingling

```
supervisorctl start lingling
```
