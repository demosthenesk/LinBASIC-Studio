#!/usr/bin/env python3
"""The SysInfo control on Linux, without VB6 (its events need a computer that changes): a battery made
of files (VBRT_POWER_SUPPLY_DIR) that starts charging -> PowerStatusChanged; a systemd-logind of our own
on a private D-Bus (fake_logind.py, needs dbus-daemon and `pip install jeepney`) that puts the computer
to sleep -> PowerQuerySuspend, PowerSuspend (the sleep waits for them: the inhibitor lock), PowerResume.

    python3 tests/dev/sysinfo/run_check.py
"""
import os, shutil, subprocess, sys, time
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
sys.path.insert(0, os.path.join(ROOT, 'compiler'))
from lbvc.driver import compile_project, build_exe, find_runtime   # noqa: E402

WANT = ['load ac= 0  st= 1  pct= 80  life= 14400  full= 18000',
        'event PowerQuerySuspend cancel=False', 'event PowerSuspend (0.3s of work)', 'event PowerResume',
        'event PowerStatusChanged ac= 1  st= 9  pct= 80  life=-1  full=-1']


def main():
    try:
        import jeepney  # noqa: F401
    except ImportError:
        print('---- sysinfo: needs `pip install jeepney` (the fake logind)')
        return 0
    if not shutil.which('dbus-daemon'):
        print('---- sysinfo: needs dbus-daemon')
        return 0
    work = os.path.join(HERE, 'work')
    shutil.rmtree(work, ignore_errors=True)
    shutil.copytree(HERE, work, ignore=shutil.ignore_patterns('work', '*.py'))
    pw = os.path.join(work, 'pw')
    for d, files in (('AC', {'type': 'Mains', 'online': '0'}),
                     ('BAT0', {'type': 'Battery', 'present': '1', 'capacity': '80', 'status': 'Discharging',
                               'energy_now': '40000000', 'energy_full': '50000000', 'power_now': '10000000'})):
        os.makedirs(os.path.join(pw, d))
        for k, v in files.items():
            open(os.path.join(pw, d, k), 'w').write(v + '\n')
    csrc, _ = compile_project([os.path.join(work, 'Si.vbp')], ide=True)   # (its output: Debug.Print, as in the IDE)
    exe = build_exe(csrc, os.path.join(work, 'prog'), find_runtime()) or os.path.join(work, 'prog')
    bus = subprocess.Popen(['dbus-daemon', '--session', '--nofork', '--print-address=1'], stdout=subprocess.PIPE, text=True)
    addr = bus.stdout.readline().strip()
    env = dict(os.environ, DBUS_SESSION_BUS_ADDRESS=addr)
    logind = subprocess.Popen([sys.executable, os.path.join(HERE, 'fake_logind.py'), '2'], env=env,
                              stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True)
    logind.stdout.readline()                    # ready
    env = dict(os.environ, DBUS_SYSTEM_BUS_ADDRESS=addr, VBRT_POWER_SUPPLY_DIR=pw, QT_QPA_PLATFORM='offscreen')
    prog = subprocess.Popen([exe], cwd=work, env=env, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True)
    time.sleep(4.5)                             # after the sleep: the charger is plugged in
    open(os.path.join(pw, 'AC', 'online'), 'w').write('1\n')
    open(os.path.join(pw, 'BAT0', 'status'), 'w').write('Charging\n')
    out, _ = prog.communicate(timeout=30)
    logind.kill()
    bus.kill()
    got = [l.rstrip() for l in out.splitlines() if l.startswith(('load', 'event'))]
    ok = got == WANT
    print('ok   sysinfo' if ok else 'FAIL sysinfo\n  want: %r\n  got:  %r' % (WANT, got))
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
