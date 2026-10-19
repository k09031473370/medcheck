import struct
from oletools.olevba import decompress_stream
def parse_dir(d):
    """Return list of (rid, raw_bytes_of_whole_record) preserving layout; VERSION (0x09) special-cased."""
    i=0; recs=[]
    while i < len(d):
        rid=struct.unpack_from('<H', d, i)[0]
        if rid==0x0009:
            recs.append((rid, d[i:i+12])); i+=12; continue
        size=struct.unpack_from('<I', d, i+2)[0]
        recs.append((rid, d[i:i+6+size])); i+=6+size
    return recs
def modules(recs):
    mods=[]; cur=None
    for rid,raw in recs:
        data=raw[6:]
        if rid==0x0019: cur={'name':data.decode('cp932')}
        elif rid==0x001A: cur['stream']=data.decode('cp932')
        elif rid==0x0031: cur['offset']=struct.unpack('<I',data)[0]; mods.append(cur)
    return mods
