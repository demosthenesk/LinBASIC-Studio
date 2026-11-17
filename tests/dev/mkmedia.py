# mkmedia.py - the sound and video files of the device tests (our own: no Microsoft files)
#   python3 tests/dev/mkmedia.py <folder>   -> tone.wav, beep8.wav, tune.mid, clip.avi
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
import struct, math, sys, os
out = sys.argv[1]
def wav(name, rate, bits, ch, secs, freq):
    n = int(rate*secs); frames=[]
    for i in range(n):
        v = math.sin(2*math.pi*freq*i/rate)*0.5
        for c in range(ch):
            if bits == 16: frames.append(struct.pack('<h', int(v*32767)))
            else: frames.append(struct.pack('B', int(128+v*127)))
    data = b''.join(frames); ba = ch*bits//8
    fmt = struct.pack('<HHIIHH', 1, ch, rate, rate*ba, ba, bits)
    body = b'WAVE' + b'fmt ' + struct.pack('<I', 16) + fmt + b'data' + struct.pack('<I', len(data)) + data
    open(os.path.join(out, name), 'wb').write(b'RIFF' + struct.pack('<I', len(body)) + body)
wav('tone.wav', 22050, 16, 1, 1.5, 440)
wav('beep8.wav', 11025, 8, 2, 0.5, 880)
# MIDI format 1, 2 tracks, 480 ppq: tempo 120 bpm then 90 bpm after 4 beats
def vlq(n):
    b=[n&0x7f]; n>>=7
    while n: b.append(0x80|(n&0x7f)); n>>=7
    return bytes(reversed(b))
def trk(ev): 
    d=b''.join(vlq(t)+e for t,e in ev)+vlq(0)+b'\xff\x2f\x00'
    return b'MTrk'+struct.pack('>I',len(d))+d
t0=[(0,b'\xff\x51\x03'+(500000).to_bytes(3,'big')),(0,b'\xff\x58\x04\x04\x02\x18\x08'),(1920,b'\xff\x51\x03'+(666667).to_bytes(3,'big'))]
notes=[]; 
for i,n in enumerate([60,62,64,65,67,69,71,72]):
    notes += [(0 if i==0 else 0, bytes([0x90,n,100])),(480,bytes([0x80,n,0]))]
notes=[(0,b'\xc0\x00')]+notes
open(os.path.join(out,'tune.mid'),'wb').write(b'MThd'+struct.pack('>IHHH',6,1,2,480)+trk(t0)+trk(notes))
# AVI: 32x24 RGB24 uncompressed, 10 frames, 10 fps
W,H,N=32,24,10; stride=W*3; fsz=stride*H
frames=[bytes([(x*8+f*20)&255 for y in range(H) for x in range(W) for _ in (0,1,2)]) for f in range(N)]
def chunk(id,d): return id+struct.pack('<I',len(d))+d+(b'\0' if len(d)%2 else b'')
def lst(t,d): return chunk(b'LIST',t+d)
avih=struct.pack('<IIIIIIIIIIIIII',100000,fsz*10,0,0x10,N,0,1,fsz,W,H,0,0,0,0)
strh=b'vids'+b'DIB '+struct.pack('<IHHIIIIIIIIhhhh',0,0,0,0,1,10,0,N,fsz,0xffffffff,0,0,0,W,H)
strf=struct.pack('<IiiHHIIiiII',40,W,H,1,24,0,fsz,0,0,0,0)
hdrl=lst(b'hdrl',chunk(b'avih',avih)+lst(b'strl',chunk(b'strh',strh)+chunk(b'strf',strf)))
movi_items=b''; idx=b''; off=4
for f in frames:
    c=chunk(b'00db',f); idx+=b'00db'+struct.pack('<III',0x10,off,len(f)); movi_items+=c; off+=len(c)
movi=lst(b'movi',movi_items)
body=b'AVI '+hdrl+movi+chunk(b'idx1',idx)
open(os.path.join(out,'clip.avi'),'wb').write(b'RIFF'+struct.pack('<I',len(body))+body)
