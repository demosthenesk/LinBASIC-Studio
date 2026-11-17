#!/usr/bin/env python3
# LinBASIC Studio - the servers of the Inet tests (tests/dev/f_inet, f_inetftp): HTTP and FTP with fixed
# answers (fixed Date and Server headers, fixed dates in the FTP listings), so that the output of a program
# is the same in VB6 (the VM) and in lbvc. What the clients send is written to the log: the headers of
# HTTP requests, the commands of FTP.
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
# With --tls-dir also HTTPS, with self-signed certificates made with openssl (for localhost, 127.0.0.1 and the
# addresses of the VM, 10.0.2.2 and 192.168.56.1):
#   47122 trusted.pem     the tests trust it (tls-dir/ca.pem, given to the programs as VBRT_CA_FILE)
#   47123 untrusted.pem   not trusted: WinINet 12045
#   47124 expired.pem     trusted, but it expired in 2021: 12037
#   47125 wrongname.pem   trusted, for wrong.invalid only: 12038
# ca.pem = trusted + expired + wrongname: VBRT_CA_FILE for libcurl; WinINet trusts only the store of Windows
# (run_tests.py puts them in the one of Wine while the tests run; the VM: Trusted Root Certification Authorities).
# --certs-only makes them and stops (ce.sh inetserver: for the VM).
#   python3 tests/dev/inet_server.py [--bind 127.0.0.1] [--http 47120] [--ftp 47121] [--pasv-ip IP]
#                                    [--root folder] [--log file] [--tls-dir folder] [--certs-only]
import argparse, base64, os, shutil, socket, socketserver, ssl, subprocess, sys, threading, time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

LOG = None
LOCK = threading.Lock()


def log(*a):
    line = time.strftime('%H:%M:%S ') + ' '.join(str(x) for x in a)
    with LOCK:
        print(line, flush=True)
        if LOG:
            with open(LOG, 'a', encoding='utf-8') as f:
                f.write(line + '\n')


# ---------------------------------------------------------------- HTTP
HELLO = b'Hello, Inet!\r\nline 2\r\n'
BIN = bytes(range(256))
BIG = bytes((i * 7 + i // 256) % 251 for i in range(100000))


class Http(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'
    server_version = 'LinBASIC-Test/1.0'
    sys_version = ''

    def version_string(self):
        return 'LinBASIC-Test/1.0'

    def date_time_string(self, timestamp=None):
        return 'Mon, 01 Jan 2024 00:00:00 GMT'

    def log_message(self, fmt, *a):
        pass

    def reply(self, code, body=b'', ctype='text/plain', extra=()):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('X-Test', 'one')
        for k, v in extra:
            self.send_header(k, v)
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)

    def body_in(self):
        n = int(self.headers.get('Content-Length') or 0)
        return self.rfile.read(n) if n else b''

    def handle_any(self):
        data = self.body_in()
        hdrs = [(k, v) for k, v in self.headers.items()]
        log('HTTP', self.command, self.path, '|', ' | '.join(f'{k}: {v}' for k, v in hdrs), '| body', len(data))
        p = self.path.split('?')[0]
        if p == '/hello.txt':
            return self.reply(200, HELLO)
        if p == '/bin':
            return self.reply(200, BIN, 'application/octet-stream')
        if p == '/big':
            return self.reply(200, BIG, 'application/octet-stream')
        if p == '/empty':
            return self.reply(200, b'')
        if p == '/greek':
            return self.reply(200, 'Καλημέρα\r\n'.encode('cp1253'), 'text/plain; charset=windows-1253')
        if p == '/echo':
            host = self.headers.get('Host', '')
            lines = [f'{self.command} {self.path} {self.request_version}']
            for k, v in hdrs:
                if k.lower() == 'host':
                    v = '<server>' + (v[v.rfind(':'):] if ':' in v else '')
                lines.append(f'{k}: {v}')
            lines.append(f'<body {len(data)}> ' + data.decode('latin-1'))
            del host
            return self.reply(200, ('\r\n'.join(lines) + '\r\n').encode('latin-1'))
        if p == '/missing':
            return self.reply(404, b'not here')
        if p == '/redirect':
            return self.reply(302, b'moved', extra=[('Location', '/hello.txt')])
        if p == '/auth':
            want = 'Basic ' + base64.b64encode(b'vb:six').decode()
            if self.headers.get('Authorization') != want:
                return self.reply(401, b'who are you?', extra=[('WWW-Authenticate', 'Basic realm="test"')])
            return self.reply(200, b'welcome')
        if p == '/slow':
            time.sleep(3)
            return self.reply(200, b'late')
        if p == '/error':
            return self.reply(500, b'broken')
        return self.reply(404, b'not here')

    do_GET = do_HEAD = do_POST = do_PUT = do_DELETE = handle_any

    def handle_one_request(self):
        try:
            super().handle_one_request()
        except (ConnectionError, OSError):
            self.close_connection = True


# ---------------------------------------------------------------- FTP
FILES = {'hello.txt': HELLO, 'bin.dat': BIN, 'sub/inner.txt': b'inside\r\n', 'big.dat': BIG}


def make_root(root):
    shutil.rmtree(root, ignore_errors=True)
    for name, data in FILES.items():
        p = os.path.join(root, *name.split('/'))
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, 'wb') as f:
            f.write(data)


class Ftp(socketserver.StreamRequestHandler):
    def send(self, line):
        log('FTP <', line)
        self.wfile.write((line + '\r\n').encode('latin-1'))

    def setup(self):
        super().setup()
        self.cwd = '/'
        self.user = None
        self.ok = False
        self.pasv = None
        self.port = None
        self.rnfr = None

    def real(self, path):
        """the path on the disk of an FTP path (relative to cwd); None outside the root"""
        if not path.startswith('/'):
            path = self.cwd.rstrip('/') + '/' + path
        parts = []
        for x in path.split('/'):
            if x in ('', '.'):
                continue
            if x == '..':
                if parts:
                    parts.pop()
                continue
            parts.append(x)
        return os.path.join(self.server.root, *parts), '/' + '/'.join(parts)

    def data_conn(self):
        if self.pasv:
            s, _ = self.pasv.accept()
            self.pasv.close()
            self.pasv = None
            return s
        if self.port:
            s = socket.create_connection(self.port, timeout=5)
            self.port = None
            return s
        return None

    def listing(self, real, names_only):
        if os.path.isfile(real):
            items = [os.path.basename(real)]
            base = os.path.dirname(real)
        else:
            items = sorted(os.listdir(real))
            base = real
        out = []
        for n in items:
            p = os.path.join(base, n)
            if names_only:
                out.append(n)
            elif os.path.isdir(p):
                out.append(f'drwxr-xr-x 1 vb vb {0:>8} Jan 01  2024 {n}')
            else:
                out.append(f'-rw-r--r-- 1 vb vb {os.path.getsize(p):>8} Jan 01  2024 {n}')
        return ('\r\n'.join(out) + ('\r\n' if out else '')).encode('latin-1')

    def transfer(self, data=None, store=None):
        self.send('150 Opening data connection.')
        try:
            s = self.data_conn()
        except OSError as e:                # PORT to an address this computer cannot reach (VirtualBox NAT)
            log('FTP data connection failed:', e)
            return self.send("425 Can't open data connection.")
        if s is None:
            return self.send('425 Use PORT or PASV first.')
        try:
            if store is not None:
                buf = b''
                while True:
                    x = s.recv(65536)
                    if not x:
                        break
                    buf += x
                with open(store, 'wb') as f:
                    f.write(buf)
            else:
                s.sendall(data)
        finally:
            s.close()
        self.send('226 Transfer complete.')

    def handle(self):
        log('FTP connection from', self.client_address[0])
        self.send('220 LinBASIC test FTP server ready.')
        while True:
            line = self.rfile.readline()
            if not line:
                break
            line = line.decode('latin-1').rstrip('\r\n')
            cmd, _, arg = line.partition(' ')
            cmd = cmd.upper()
            log('FTP >', 'PASS ***' if cmd == 'PASS' else line)
            if cmd == 'USER':
                self.user = arg
                self.send('331 Password required.')
            elif cmd == 'PASS':
                if (self.user or '').lower() == 'anonymous' or (self.user == 'vb' and arg == 'six'):
                    self.ok = True
                    self.send('230 Logged in.')
                else:
                    self.send('530 Login incorrect.')
            elif cmd == 'QUIT':
                self.send('221 Goodbye.')
                break
            elif not self.ok:
                self.send('530 Please login with USER and PASS.')
            elif cmd == 'SYST':
                self.send('215 UNIX Type: L8')
            elif cmd in ('FEAT', 'OPTS', 'EPSV', 'EPRT', 'MLSD', 'MLST', 'REST', 'ALLO', 'SITE', 'STAT', 'HELP', 'MDTM'):
                self.send('502 Command not implemented.')
            elif cmd == 'NOOP':
                self.send('200 NOOP ok.')
            elif cmd == 'TYPE':
                self.send(f'200 Type set to {arg}.')
            elif cmd in ('MODE', 'STRU'):
                self.send('200 OK.')
            elif cmd == 'PWD' or cmd == 'XPWD':
                self.send(f'257 "{self.cwd}" is current directory.')
            elif cmd in ('CWD', 'XCWD', 'CDUP', 'XCUP'):
                real, path = self.real('..' if cmd in ('CDUP', 'XCUP') else arg)
                if os.path.isdir(real):
                    self.cwd = path
                    self.send(f'250 CWD command successful.')
                else:
                    self.send(f'550 {arg}: No such file or directory.')
            elif cmd == 'PASV':
                if self.pasv:
                    self.pasv.close()
                self.pasv = socket.socket()
                me = self.connection.getsockname()[0]
                self.pasv.bind((me, 0))
                self.pasv.listen(1)
                self.pasv.settimeout(15)
                # through VirtualBox NAT the VM comes from 127.0.0.1 and reaches this computer as --pasv-ip
                ip = self.server.pasv_ip if (self.server.pasv_ip and me == '127.0.0.1') else me
                p = self.pasv.getsockname()[1]
                self.send('227 Entering Passive Mode (%s,%d,%d).' % (ip.replace('.', ','), p // 256, p % 256))
            elif cmd == 'PORT':
                x = arg.split(',')
                self.port = ('.'.join(x[:4]), int(x[4]) * 256 + int(x[5]))
                self.send('200 PORT command successful.')
            elif cmd in ('LIST', 'NLST'):
                a = ' '.join(w for w in arg.split() if not w.startswith('-'))
                real, _ = self.real(a or '.')
                if not os.path.exists(real):
                    self.send(f'550 {a}: No such file or directory.')
                else:
                    self.transfer(self.listing(real, cmd == 'NLST'))
            elif cmd == 'RETR':
                real, _ = self.real(arg)
                if not os.path.isfile(real):
                    self.send(f'550 {arg}: No such file or directory.')
                else:
                    with open(real, 'rb') as f:
                        self.transfer(f.read())
            elif cmd == 'STOR':
                real, _ = self.real(arg)
                self.transfer(store=real)
            elif cmd == 'SIZE':
                real, _ = self.real(arg)
                if os.path.isfile(real):
                    self.send(f'213 {os.path.getsize(real)}')
                else:
                    self.send(f'550 {arg}: No such file or directory.')
            elif cmd == 'DELE':
                real, _ = self.real(arg)
                if os.path.isfile(real):
                    os.remove(real)
                    self.send('250 DELE command successful.')
                else:
                    self.send(f'550 {arg}: No such file or directory.')
            elif cmd in ('MKD', 'XMKD'):
                real, path = self.real(arg)
                if os.path.exists(real):
                    self.send(f'550 {arg}: File exists.')
                else:
                    os.makedirs(real)
                    self.send(f'257 "{path}" directory created.')
            elif cmd in ('RMD', 'XRMD'):
                real, _ = self.real(arg)
                try:
                    os.rmdir(real)
                    self.send('250 RMD command successful.')
                except OSError:
                    self.send(f'550 {arg}: Directory not empty or not found.')
            elif cmd == 'RNFR':
                real, _ = self.real(arg)
                if os.path.exists(real):
                    self.rnfr = real
                    self.send('350 File exists, ready for destination name.')
                else:
                    self.send(f'550 {arg}: No such file or directory.')
            elif cmd == 'RNTO':
                real, _ = self.real(arg)
                if self.rnfr:
                    os.rename(self.rnfr, real)
                    self.rnfr = None
                    self.send('250 Rename successful.')
                else:
                    self.send('503 Bad sequence of commands.')
            else:
                self.send(f'500 {cmd} not understood.')
        log('FTP connection closed')


# the addresses also as DNS names: Wine's crypt32 compares the host with the DNS names only (Windows and OpenSSL: IP)
SAN = ('DNS:localhost,IP:127.0.0.1,DNS:127.0.0.1,IP:10.0.2.2,DNS:10.0.2.2,IP:192.168.56.1,DNS:192.168.56.1')
CERTS = (('trusted', SAN, None), ('untrusted', SAN, None), ('expired', SAN, ('20200101000000Z', '20210101000000Z')),
         ('wrongname', 'DNS:wrong.invalid', None))


def make_cert(folder, name, san, dates):
    """a self-signed certificate (openssl): (certificate, key); dates: (start, end) in the past, for 'expired'"""
    os.makedirs(folder, exist_ok=True)
    crt, key = os.path.join(folder, name + '.pem'), os.path.join(folder, name + '.key')
    if os.path.exists(crt) and os.path.exists(key):
        return crt, key
    q = dict(check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    cn = '/CN=' + san.split(',')[0].split(':', 1)[1]
    if not dates:
        subprocess.run(['openssl', 'req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', key, '-out', crt,
                        '-days', '36500', '-subj', cn, '-addext', 'subjectAltName=' + san,
                        '-addext', 'basicConstraints=critical,CA:FALSE',      # trusted, it signs nothing else
                        '-addext', 'keyUsage=digitalSignature,keyEncipherment', '-addext', 'extendedKeyUsage=serverAuth'], **q)
        return crt, key
    import tempfile                         # openssl ca: the only way of openssl 3.0 to dates in the past
    with tempfile.TemporaryDirectory() as t:
        open(os.path.join(t, 'index.txt'), 'w').close()
        with open(os.path.join(t, 'serial'), 'w') as f:
            f.write('1000\n')
        cnf = os.path.join(t, 'ca.cnf')
        with open(cnf, 'w') as f:
            f.write(f'[ca]\ndefault_ca = d\n[d]\ndatabase = {t}/index.txt\nnew_certs_dir = {t}\nserial = {t}/serial\n'
                    f'default_md = sha256\npolicy = p\nunique_subject = no\n[p]\ncommonName = supplied\n'
                    f'[x]\nsubjectAltName = {san}\nbasicConstraints = critical,CA:FALSE\n'
                    f'keyUsage = digitalSignature,keyEncipherment\nextendedKeyUsage = serverAuth\n')
        csr = os.path.join(t, 'r.csr')
        subprocess.run(['openssl', 'req', '-new', '-newkey', 'rsa:2048', '-nodes', '-keyout', key, '-out', csr,
                        '-subj', cn], **q)
        subprocess.run(['openssl', 'ca', '-batch', '-selfsign', '-config', cnf, '-keyfile', key, '-in', csr, '-out', crt,
                        '-startdate', dates[0], '-enddate', dates[1], '-extensions', 'x', '-notext'], **q)
    return crt, key


def make_certs(folder):
    """all of them, and ca.pem (the ones the tests trust): {name: (certificate, key)}"""
    stamp = os.path.join(folder, 'v3')     # older ones (other names, CA:TRUE): made again
    if os.path.isdir(folder) and not os.path.exists(stamp):
        for f in os.listdir(folder):
            os.remove(os.path.join(folder, f))
    made = {n: make_cert(folder, n, san, d) for n, san, d in CERTS}
    open(stamp, 'w').close()
    with open(os.path.join(folder, 'ca.pem'), 'w') as f:
        for n in ('trusted', 'expired', 'wrongname'):
            f.write(open(made[n][0]).read())
    return made


class TlsServer(ThreadingHTTPServer):
    def __init__(self, addr, handler, crt, key):
        super().__init__(addr, handler)
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(crt, key)
        try:                                # Windows 7 (VB6 in the VM): TLS 1.0 and its ciphers too
            import warnings
            warnings.simplefilter('ignore', DeprecationWarning)
            ctx.minimum_version = ssl.TLSVersion.TLSv1
            ctx.set_ciphers('DEFAULT:@SECLEVEL=0')
        except (ValueError, ssl.SSLError):
            pass
        self.ctx = ctx

    def get_request(self):
        s, a = self.socket.accept()
        s.settimeout(10)
        return self.ctx.wrap_socket(s, server_side=True, do_handshake_on_connect=False), a

    def process_request_thread(self, request, client_address):
        try:                                          # the handshake in the thread of the connection
            request.do_handshake()
        except (ssl.SSLError, OSError) as e:          # a client that does not trust the certificate
            if not isinstance(e, ssl.SSLEOFError) and 'EOF' not in str(e):   # not the port check of run_tests.py
                log('HTTPS handshake failed:', e)
            self.shutdown_request(request)
            return
        super().process_request_thread(request, client_address)


class FtpServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def main():
    global LOG
    ap = argparse.ArgumentParser()
    ap.add_argument('--bind', default='127.0.0.1', help='addresses, separated by commas')
    ap.add_argument('--http', type=int, default=47120)
    ap.add_argument('--ftp', type=int, default=47121)
    ap.add_argument('--pasv-ip', default='')
    ap.add_argument('--root', default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'work', 'ftproot'))
    ap.add_argument('--log', default='')
    ap.add_argument('--tls-dir', default='')
    ap.add_argument('--certs-only', action='store_true')
    a = ap.parse_args()
    certs = make_certs(a.tls_dir) if a.tls_dir and shutil.which('openssl') else {}
    if a.certs_only:
        print('\n'.join(sorted(c for c, k in certs.values())) if certs else 'no openssl')
        return 0 if certs else 1
    LOG = a.log or None
    if LOG:
        open(LOG, 'w').close()
    make_root(a.root)
    ThreadingHTTPServer.allow_reuse_address = True
    ThreadingHTTPServer.daemon_threads = True
    up = 0
    for ip in [x.strip() for x in a.bind.split(',') if x.strip()]:
        try:
            h = ThreadingHTTPServer((ip, a.http), Http)
            f = FtpServer((ip, a.ftp), Ftp)
        except OSError as e:                # 192.168.56.1: the host-only network of VirtualBox is not there
            log(f'{ip}: {e} (skipped; VirtualBox host-only network: ip addr show vboxnet0)')
            continue
        up += 1
        if certs:
            for i, (name, san, d) in enumerate(CERTS):
                t = TlsServer((ip, a.http + 2 + i), Http, *certs[name])
                t.daemon_threads = True
                threading.Thread(target=t.serve_forever, daemon=True).start()
        f.root, f.pasv_ip = a.root, a.pasv_ip
        threading.Thread(target=h.serve_forever, daemon=True).start()
        threading.Thread(target=f.serve_forever, daemon=True).start()
    if not up:
        return 1
    log(f'HTTP on {a.bind}:{a.http}, FTP on {a.bind}:{a.ftp} (root {a.root}); Ctrl+C stops')
    try:
        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    sys.exit(main())
