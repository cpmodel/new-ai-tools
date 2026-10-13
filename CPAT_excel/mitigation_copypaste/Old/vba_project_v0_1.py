"""Write an Excel VBA project (xl/vbaProject.bin) from VBA source text, without Excel, and embed it in a workbook.

Implements the parts of [MS-OVBA] and [MS-CFB] that a source-only project needs:
  - MS-OVBA compression (2.4.1) of the dir stream and the module streams (source at offset 0, no p-code);
  - the dir stream (2.3.4.2): project information, a reference to stdole, one record per module;
  - the PROJECT stream (2.3.1) with the encrypted protection state, password and visibility (2.4.3: not protected,
    no password, visible) and PROJECTwm (2.3.3);
  - _VBA_PROJECT (2.3.4.1) with version 0xFFFF and no performance cache, so the VBA host compiles from source;
  - a version 3 compound file (512-byte sectors, 64-byte mini sectors for streams below 4096 bytes).
The format was checked against the legacy CPAT project (same records; its CMG/DPB/GC decrypt with this module's
algorithm). Document modules (ThisWorkbook, one per sheet) carry only their attributes and must match the code
names in workbook.xml and each sheet's sheetPr.

    from vba_project_v0_1 import embed_vba
    embed_vba('book.xlsx', 'book.xlsm', {'Module1': source}, sheet_codenames)
"""
import os
import re
import shutil
import struct
import zipfile

ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF
CP = 'cp1252'
PROJECT_ID = '{6F3B2A10-5C1D-4E8A-9B7F-2D4C6E8A0B13}'   # fixed: the file is the same on every build
DOC_ATTRS = ('Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n'
             'Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
             'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n')
BASE_WORKBOOK = '0{00020819-0000-0000-C000-000000000046}'
BASE_SHEET = '0{00020820-0000-0000-C000-000000000046}'


# ------------------------------------------------------------------ MS-OVBA 2.4.1 compression
def _copy_token_bits(difference):
    bits = 4
    while (1 << bits) < difference:
        bits += 1
    return bits


def compress(data):
    out = bytearray(b'\x01')
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        body = bytearray()
        seen = {}                                               # 3-byte prefix -> earlier positions in the chunk
        i = 0

        def remember(k):
            if k + 3 <= len(chunk):
                seen.setdefault(chunk[k:k + 3], []).append(k)
        while i < len(chunk):
            flag_pos = len(body)
            body.append(0)
            flag = 0
            for bit in range(8):
                if i >= len(chunk):
                    break
                bits = _copy_token_bits(i) if i else 4
                max_len = (0xFFFF >> bits) + 3
                best_len, best_off = 0, 0
                for j in reversed(seen.get(chunk[i:i + 3], [])):    # nearest first
                    if i - j > (1 << bits):
                        break
                    n = 0
                    while n < max_len and i + n < len(chunk) and chunk[j + n] == chunk[i + n]:
                        n += 1
                    if n > best_len:
                        best_len, best_off = n, i - j
                        if n == max_len:
                            break
                if best_len >= 3:
                    body += struct.pack('<H', ((best_off - 1) << (16 - bits)) | (best_len - 3))
                    flag |= 1 << bit
                    for k in range(i, i + best_len):
                        remember(k)
                    i += best_len
                else:
                    body.append(chunk[i])
                    remember(i)
                    i += 1
            body[flag_pos] = flag
        if len(body) > 4096 and len(chunk) == 4096:            # raw chunk
            out += struct.pack('<H', 0x3000 | (4096 + 2 - 3)) + chunk
        else:
            out += struct.pack('<H', 0xB000 | (len(body) + 2 - 3)) + body
    return bytes(out)


# ------------------------------------------------------------------ MS-OVBA 2.4.3 data encryption
def encrypt(data, project_id, seed=0x5A):
    key = sum(project_id.encode(CP)) & 0xFF
    venc, kenc = seed ^ 2, seed ^ key
    out = [seed, venc, kenc]
    ub1, eb1, eb2 = key, kenc, venc
    plain = [0] * ((seed & 6) // 2) + list(struct.pack('<I', len(data))) + list(data)
    for b in plain:
        be = b ^ ((eb2 + ub1) & 0xFF)
        out.append(be)
        eb2, eb1, ub1 = eb1, be, b
    return bytes(out).hex().upper()


# ------------------------------------------------------------------ streams
def _rec(rid, payload):
    return struct.pack('<HI', rid, len(payload)) + payload


def dir_stream(project_name, modules):
    """modules: list of (name, is_document)."""
    u = lambda s: s.encode('utf-16-le')
    b = lambda s: s.encode(CP)
    d = bytearray()
    d += _rec(0x0001, struct.pack('<I', 3))                    # SYSKIND: 64-bit Windows (as legacy)
    d += _rec(0x004A, struct.pack('<I', 6))                    # COMPATVERSION
    d += _rec(0x0002, struct.pack('<I', 0x409))                # LCID
    d += _rec(0x0014, struct.pack('<I', 0x409))                # LCIDINVOKE
    d += _rec(0x0003, struct.pack('<H', 1252))                 # CODEPAGE
    d += _rec(0x0004, b(project_name))                         # NAME
    d += _rec(0x0005, b'') + _rec(0x0040, b'')                 # DOCSTRING (+ unicode)
    d += _rec(0x0006, b'') + _rec(0x003D, b'')                 # HELPFILEPATH 1, 2
    d += _rec(0x0007, struct.pack('<I', 0))                    # HELPCONTEXT
    d += _rec(0x0008, struct.pack('<I', 0))                    # LIBFLAGS
    d += struct.pack('<HIIH', 0x0009, 4, 1, 0)                 # VERSION (reserved size 4, major, minor)
    d += _rec(0x000C, b'') + _rec(0x003C, b'')                 # CONSTANTS (+ unicode)
    libid = b('*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation')
    d += _rec(0x0016, b('stdole')) + _rec(0x003E, u('stdole'))
    d += _rec(0x000D, struct.pack('<I', len(libid)) + libid + struct.pack('<IH', 0, 0))
    d += struct.pack('<HIH', 0x000F, 2, len(modules))          # PROJECTMODULES
    d += struct.pack('<HIH', 0x0013, 2, 0xFFFF)                # PROJECTCOOKIE
    for name, is_doc in modules:
        d += _rec(0x0019, b(name)) + _rec(0x0047, u(name))     # MODULENAME, MODULENAMEUNICODE
        d += _rec(0x001A, b(name)) + _rec(0x0032, u(name))     # MODULESTREAMNAME (+ unicode)
        d += _rec(0x001C, b'') + _rec(0x0048, b'')             # MODULEDOCSTRING (+ unicode)
        d += _rec(0x0031, struct.pack('<I', 0))                # MODULEOFFSET: source at 0 (no p-code)
        d += _rec(0x001E, struct.pack('<I', 0))                # MODULEHELPCONTEXT
        d += _rec(0x002C, struct.pack('<H', 0xFFFF))           # MODULECOOKIE
        d += struct.pack('<HI', 0x0022 if is_doc else 0x0021, 0)   # MODULETYPE: document / procedural
        d += struct.pack('<HI', 0x002B, 0)                     # module terminator
    d += struct.pack('<HI', 0x0010, 0)                         # dir terminator
    return bytes(d)


def project_stream(project_id, modules):
    lines = [f'ID="{project_id}"']
    lines += [f'Document={n}/&H00000000' if doc else f'Module={n}' for n, doc in modules]
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              f'CMG="{encrypt(bytes(4), project_id)}"', f'DPB="{encrypt(bytes(1), project_id)}"',
              f'GC="{encrypt(bytes([0xFF]), project_id)}"', '',
              '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
              '[Workspace]'] + [f'{n}=0, 0, 0, 0, C' for n, _d in modules]
    return ('\r\n'.join(lines) + '\r\n').encode(CP)


def projectwm_stream(modules):
    return b''.join(n.encode(CP) + b'\x00' + n.encode('utf-16-le') + b'\x00\x00' for n, _d in modules) + b'\x00\x00'


def module_source(name, body, base=None):
    head = f'Attribute VB_Name = "{name}"\r\n'
    if base:
        head += f'Attribute VB_Base = "{base}"\r\n' + DOC_ATTRS
    body = re.sub(r'\r?\n', '\r\n', body)
    return (head + body).encode(CP)


# ------------------------------------------------------------------ MS-CFB writer (version 3)
def _name_key(name):
    return (len(name), name.upper())


def write_cfb(streams):
    """streams: {'PROJECT': bytes, 'VBA/dir': bytes, ...} (one storage level). Returns the compound file bytes."""
    storages = sorted({p.split('/')[0] for p in streams if '/' in p})
    entries = [{'name': 'Root Entry', 'type': 5, 'children': []}]
    index = {}
    for s in storages:
        index[s] = len(entries)
        entries.append({'name': s, 'type': 1, 'children': []})
        entries[0]['children'].append(index[s])
    for path, data in streams.items():
        parent = index[path.split('/')[0]] if '/' in path else 0
        k = len(entries)
        entries.append({'name': path.split('/')[-1], 'type': 2, 'data': data, 'children': []})
        entries[parent]['children'].append(k)

    def tree(ids):                                              # balanced binary search tree of siblings
        if not ids:
            return NOSTREAM
        ids = sorted(ids, key=lambda i: _name_key(entries[i]['name']))
        m = len(ids) // 2
        entries[ids[m]]['left'] = tree(ids[:m])
        entries[ids[m]]['right'] = tree(ids[m + 1:])
        return ids[m]
    for e in entries:
        e['child'] = tree(e['children'])
    # mini stream (streams < 4096 bytes) and regular streams
    mini, minifat = bytearray(), []
    big = []
    for e in entries:
        if e['type'] != 2:
            continue
        data = e['data']
        if len(data) < 4096:
            n = max(1, -(-len(data) // 64)) if data else 0
            e['start'] = len(mini) // 64 if n else ENDOFCHAIN
            for k in range(n):
                minifat.append(len(mini) // 64 + 1 if k < n - 1 else ENDOFCHAIN)
                mini += data[k * 64:(k + 1) * 64].ljust(64, b'\x00')
        else:
            big.append(e)
    sectors, fat = [], []

    def add_chain(data, start_fat=None):
        n = -(-len(data) // 512)
        first = len(sectors)
        for k in range(n):
            sectors.append(bytes(data[k * 512:(k + 1) * 512]).ljust(512, b'\x00'))
            fat.append(first + k + 1 if k < n - 1 else ENDOFCHAIN)
        return first if n else ENDOFCHAIN
    entries[0]['start'] = add_chain(mini)
    entries[0]['size'] = len(mini)
    for e in big:
        e['start'] = add_chain(e['data'])
    mf = b''.join(struct.pack('<I', x) for x in minifat)
    mf += struct.pack('<I', FREESECT) * ((-len(minifat)) % 128)
    minifat_start = add_chain(mf) if minifat else ENDOFCHAIN
    n_minifat = -(-len(mf) // 512) if minifat else 0
    dirs = bytearray()
    for e in entries:
        nm = e['name'].encode('utf-16-le') + b'\x00\x00'
        size = e.get('size', len(e.get('data', b'')))
        dirs += nm.ljust(64, b'\x00') + struct.pack('<HBB', len(nm), e['type'], 1)
        dirs += struct.pack('<III', e.get('left', NOSTREAM), e.get('right', NOSTREAM), e['child'])
        dirs += bytes(16) + struct.pack('<I', 0) + bytes(16)
        dirs += struct.pack('<IQ', e.get('start', ENDOFCHAIN) if e['type'] != 1 else 0, size)
    while len(dirs) % 512:
        dirs += (b'\x00' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM)
                 + bytes(16 + 4 + 16 + 4 + 8))
    dir_start = add_chain(dirs)
    n_fat = 1
    while (len(sectors) + n_fat) > n_fat * 128:
        n_fat += 1
    fat_start = len(sectors)
    fat += [FATSECT] * n_fat
    fat += [FREESECT] * (n_fat * 128 - len(fat))
    for k in range(n_fat):
        sectors.append(b''.join(struct.pack('<I', x) for x in fat[k * 128:(k + 1) * 128]))
    assert n_fat <= 109
    hdr = bytearray(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + bytes(16))
    hdr += struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + bytes(6)
    hdr += struct.pack('<IIIIIIIII', 0, n_fat, dir_start, 0, 4096, minifat_start, n_minifat, ENDOFCHAIN, 0)
    difat = [fat_start + k for k in range(n_fat)] + [FREESECT] * (109 - n_fat)
    hdr += b''.join(struct.pack('<I', x) for x in difat)
    assert len(hdr) == 512
    return bytes(hdr) + b''.join(sectors)


def vba_project(std_modules, sheet_codenames, project_id=None):
    """std_modules: {name: source}; sheet_codenames: code names of the sheets (document modules)."""
    project_id = project_id or PROJECT_ID
    mods = [('ThisWorkbook', True)] + [(n, True) for n in sheet_codenames] + [(n, False) for n in std_modules]
    streams = {'PROJECT': project_stream(project_id, mods), 'PROJECTwm': projectwm_stream(mods),
               'VBA/_VBA_PROJECT': b'\xcc\x61\xff\xff\x00\x00\x00',
               'VBA/dir': compress(dir_stream('VBAProject', mods))}
    streams['VBA/ThisWorkbook'] = compress(module_source('ThisWorkbook', '', BASE_WORKBOOK))
    for n in sheet_codenames:
        streams[f'VBA/{n}'] = compress(module_source(n, '', BASE_SHEET))
    for n, src in std_modules.items():
        streams[f'VBA/{n}'] = compress(module_source(n, src))
    return write_cfb(streams)


# ------------------------------------------------------------------ embedding in an OOXML workbook
def embed_vba(src, dst, std_modules, project_id=None):
    """Copy workbook src (.xlsx, saved with workbookPr codeName="ThisWorkbook" and a sheetPr codeName on every
    sheet) to dst (.xlsm) with a VBA project holding std_modules."""
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}
    wbxml = files['xl/workbook.xml'].decode('utf-8')
    rels = files['xl/_rels/workbook.xml.rels'].decode('utf-8')
    order = re.findall(r'<sheet [^>]*r:id="([^"]+)"', wbxml)
    targets = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    targets.update({a: b for b, a in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    codenames = []
    for rid in order:
        t = targets[rid].lstrip('/')
        path = t if t.startswith('xl/') else 'xl/' + t
        m = re.search(r'<sheetPr[^>]*codeName="([^"]+)"', files[path].decode('utf-8'))
        assert m, f'no codeName in {path}'
        codenames.append(m.group(1))
    assert 'codeName="ThisWorkbook"' in wbxml, 'workbook code name missing'
    files['xl/vbaProject.bin'] = vba_project(std_modules, codenames, project_id)
    ct = files['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    if 'Extension="bin"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                                     '<Default ', 1)
    files['[Content_Types].xml'] = ct.encode('utf-8')
    files['xl/_rels/workbook.xml.rels'] = rels.replace(
        '</Relationships>', '<Relationship Id="rIdVBA1" Type="http://schemas.microsoft.com/office/2006/relationships/'
                            'vbaProject" Target="vbaProject.bin"/></Relationships>').encode('utf-8')
    tmp = dst + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names + ['xl/vbaProject.bin']:
            z.writestr(n, files[n])
    shutil.move(tmp, dst)
    return codenames
