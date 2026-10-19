"""MS-OVBA (2.4.1) compression for VBA module/dir streams."""
def _compress_chunk(data: bytes, start: int, end: int) -> bytes:
    out = bytearray(); pos = start
    while pos < end:
        flag = 0; tokens = bytearray()
        for bit in range(8):
            if pos >= end: break
            diff = pos - start
            bit_count = max(4, (diff - 1).bit_length()) if diff > 0 else 4
            max_len = (0xFFFF >> bit_count) + 3
            max_off = 1 << bit_count
            best_len = 0; best_off = 0
            if diff >= 1:
                limit = min(max_len, end - pos)
                for off in range(1, min(diff, max_off) + 1):
                    l = 0
                    while l < limit and data[pos - off + l] == data[pos + l]: l += 1
                    if l > best_len: best_len, best_off = l, off
                    if best_len == limit: break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                tokens += token.to_bytes(2, 'little'); flag |= (1 << bit); pos += best_len
            else:
                tokens.append(data[pos]); pos += 1
        out.append(flag); out += tokens
    return bytes(out)

def compress(data: bytes) -> bytes:
    out = bytearray(b'\x01'); start = 0
    while start < len(data):
        end = min(start + 4096, len(data))
        body = _compress_chunk(data, start, end)
        if len(body) + 2 > 4098:   # incompressible: raw chunk (must be exactly 4096 bytes)
            raw = data[start:end].ljust(4096, b'\x00')
            header = 0x3000 | (len(raw) + 2 - 3)
            out += header.to_bytes(2, 'little') + raw
        else:
            header = 0xB000 | (len(body) + 2 - 3)
            out += header.to_bytes(2, 'little') + body
        start = end
    return bytes(out)
