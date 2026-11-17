import os, sys, time, subprocess

proj = r"C:\Users\smd32\Downloads\PM Project"
log = os.path.join(proj, "sandbox", "sync_watch.log")
with open(log, "w") as f:
    f.write("starting\n")
    proc = subprocess.Popen([sys.executable, "scripts\\sync_template.py", "--skip-package"],
                            cwd=proj, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    time.sleep(40)
    if proc.poll() is None:
        f.write("WATCHDOG: sync still running after 40s, killing\n")
        proc.kill()
        try:
            out, _ = proc.communicate(timeout=5)
        except Exception:
            out = ""
    else:
        out, _ = proc.communicate(timeout=5)
    f.write("rc: %s\n" % proc.returncode)
    f.write("output:\n")
    f.write(out[-3000:])
f.write("done\n")