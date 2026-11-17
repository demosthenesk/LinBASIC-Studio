"""a fake systemd-logind on a private bus: Inhibit gives a pipe; after `delay` seconds it sends
PrepareForSleep(true), waits until the inhibitor is closed (the program ran its events), sleeps a
moment, sends PrepareForSleep(false). Prints what happened."""
import os, sys, time, select, threading
from jeepney import new_method_return, new_signal, DBusAddress, MessageType, HeaderFields
from jeepney.io.blocking import open_dbus_connection
from jeepney.bus_messages import message_bus
conn = open_dbus_connection(bus='SESSION', enable_fds=True)
conn.send_and_get_reply(message_bus.RequestName('org.freedesktop.login1'))
print('fake logind ready', flush=True)
mgr = DBusAddress('/org/freedesktop/login1', interface='org.freedesktop.login1.Manager')
pipes = []
def cycle():
    time.sleep(float(sys.argv[1]))
    print('-> PrepareForSleep(true), inhibitors:', len(pipes), flush=True)
    t0 = time.time()
    conn.send(new_signal(mgr, 'PrepareForSleep', 'b', (True,)))
    # the inhibitor: its write end closed by the program -> EOF on the read end
    for r in list(pipes):
        select.select([r], [], [], 5)
        print('inhibitor released after %.2fs' % (time.time() - t0), os.read(r, 1), flush=True)
        pipes.remove(r)
    time.sleep(0.5)
    print('-> PrepareForSleep(false)', flush=True)
    conn.send(new_signal(mgr, 'PrepareForSleep', 'b', (False,)))
threading.Thread(target=cycle, daemon=True).start()
while True:
    msg = conn.receive(timeout=30)
    h = msg.header
    if h.message_type == MessageType.method_call and h.fields.get(HeaderFields.member) == 'Inhibit':
        r, w = os.pipe()
        pipes.append(r)
        print('Inhibit', msg.body, flush=True)
        conn.send(new_method_return(msg, 'h', (w,)))
        os.close(w)
