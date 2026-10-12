"""Replace the embedded MS-OVBA project in legacy Visio 11 .vss copies.

Requires olefile and oletools. Never writes the input directory. Unchanged Visio
payloads and form streams are verified byte-for-byte, then VBA is re-extracted.
Native compilation and execution in Microsoft Visio are separate checks.
"""
import argparse, collections, hashlib, io, json, re, struct
from pathlib import Path
import olefile
from oletools.olevba import VBA_Project, decompress_stream
from compound_writer import read_compound, write_compound

PTR = struct.Struct('<IIIIH')


def visio_unpack(data):
    ring = bytearray(4096)
    out = bytearray()
    at = 0
    while at < len(data):
        flag = data[at]; at += 1
        for bit in range(8):
            if at == len(data):
                break
            if flag & (1 << bit):
                value = data[at]; at += 1
                ring[len(out) % 4096] = value
                out.append(value)
            else:
                assert at+2 <= len(data)
                a, b = data[at:at+2]; at += 2
                ref = (((b & 240) << 4) | a) + 18
                for j in range((b & 15)+3):
                    value = ring[(ref+j) % 4096]
                    ring[len(out) % 4096] = value
                    out.append(value)
    return bytes(out)


def visio_pack(data):
    # Valid literal-only LZSS: fewer moving parts than a new match encoder.
    return b''.join(bytes([(1 << len(data[i:i+8]))-1]) + data[i:i+8]
                    for i in range(0, len(data), 8))


def ovba_pack(data):
    output = bytearray(b'\x01')
    for base in range(0, len(data), 4096):
        chunk = data[base:base+4096]
        packed = bytearray()
        seen = collections.defaultdict(list)
        pos = 0
        while pos < len(chunk):
            flags_at = len(packed); packed.append(0)
            for bit in range(8):
                if pos == len(chunk):
                    break
                bits = max(4, (pos-1).bit_length())
                length_bits = 16-bits
                longest, previous = 0, 0
                limit = min((1 << length_bits)+2, len(chunk)-pos)
                for candidate in reversed(seen[chunk[pos:pos+3]][-128:]):
                    if pos-candidate > (1 << bits):
                        continue
                    n = 3
                    while n < limit and chunk[candidate+n] == chunk[pos+n]:
                        n += 1
                    if n > longest:
                        longest, previous = n, candidate
                if longest >= 3:
                    packed[flags_at] |= 1 << bit
                    packed.extend(struct.pack('<H', ((pos-previous-1) << length_bits) | (longest-3)))
                    taken = longest
                else:
                    packed.append(chunk[pos]); taken = 1
                for offset in range(pos, pos+taken):
                    seen[chunk[offset:offset+3]].append(offset)
                pos += taken
        if len(packed) >= 4096:
            assert len(chunk) == 4096, 'Incompressible partial OVBA chunk'
            output.extend(b'\xff\x3f' + chunk)
        else:
            output.extend(struct.pack('<H', 0xb000 | (len(packed)-1)) + packed)
    result = bytes(output)
    assert decompress_stream(bytearray(result)) == data
    return result


def parse_visio(stream):
    assert stream.startswith(b'Visio (TM) Drawing')
    assert struct.unpack_from('<H', stream, 26)[0] == 11, 'Only reviewed Visio 11 files are supported'
    visiting = set()

    def visit(pointer):
        typ, ignored, offset, length, fmt = pointer
        assert offset+length <= len(stream)
        key = offset, length
        node = {'pointer': pointer, 'children': []}
        if not typ:
            node['data'] = b''
            return node
        assert key not in visiting
        visiting.add(key)
        raw = stream[offset:offset+length]
        node['data'] = visio_unpack(raw) if fmt & 2 else raw
        shift = 4 if fmt & 2 else 0
        if fmt >> 4 == 5:
            loc = struct.unpack_from('<I', node['data'], shift)[0]+shift-4
            _, count, _ = struct.unpack_from('<III', node['data'], loc)
            assert 0 <= count <= 10000
            assert loc+12+count*18 <= len(node['data'])
            node['loc'] = loc
            for i in range(count):
                node['children'].append(visit(PTR.unpack_from(node['data'], loc+12+i*18)))
        visiting.remove(key)
        return node
    return visit(PTR.unpack_from(stream, 36))


def vba_payload(root):
    found = []
    def walk(node):
        if node['pointer'][0] == 13:
            assert all(c['pointer'][0] == 31 for c in node['children'])
            found.append(node)
        else:
            for child in node['children']:
                walk(child)
    walk(root)
    assert len(found) == 1, 'Expected one embedded VBA project'
    node = found[0]
    chunks = []
    for child in node['children']:
        data = child['data']
        size = struct.unpack_from('<I', data)[0]
        assert size == len(data)-4
        chunks.append(data[4:])
    size = struct.unpack_from('<I', node['data'], 8)[0]
    payload = b''.join(chunks)
    assert size <= len(payload)
    return node, payload[:size]


def record(identifier, payload):
    return struct.pack('<HI', identifier, len(payload)) + payload


def source_from_file(path):
    raw = path.read_bytes()
    try:
        source = raw.decode('utf8')
    except UnicodeDecodeError:
        source = raw.decode('cp1251')
    lines = source.splitlines()
    start = next(i for i, line in enumerate(lines) if line.startswith('Attribute VB_Name'))
    return '\r\n'.join(lines[start:]) + '\r\n'


def replace_project(payload, modules, repo):
    meta, streams = read_compound(payload)
    ole = olefile.OleFileIO(io.BytesIO(payload))
    project = VBA_Project(ole, '', 'PROJECT', 'VBA/dir', relaxed=False)
    project.parse_project_stream()
    modules_at = project.dir_stream.tell()-2
    list(project.parse_modules())
    original = {module.name: module for module in project.modules}
    updates = {module['name']: source_from_file(repo/module['path']) for module in modules}
    extra = [m for m in modules if m['name'] not in original]
    assert all(m['is_new'] for m in extra)
    directory = bytearray(project.dir_stream.getvalue())
    cursor = modules_at+16
    for module in project.modules:
        while True:
            identifier, length = struct.unpack_from('<HI', directory, cursor)
            if identifier == 0x31:
                assert length == 4
                struct.pack_into('<I', directory, cursor+6, 0)
            cursor += 6+length
            if identifier == 0x2b:
                break
        source = updates.get(module.name)
        if source is not None:
            # Preserve host-bound module identities/designer attributes.
            attributes = []
            for line in module.code.splitlines():
                if line.startswith('Attribute '):
                    attributes.append(line)
                elif line.strip():
                    break
            if attributes:
                lines = source.splitlines()
                while lines and lines[0].startswith('Attribute '):
                    lines.pop(0)
                source = '\r\n'.join(attributes+lines)+'\r\n'
            raw = source.encode(project.codec)
        else:
            raw = module.code_raw
        streams[('VBA', module.streamname)] = ovba_pack(raw)
    suffix = directory[cursor:]
    assert suffix == bytes.fromhex('100000000000')
    directory = directory[:cursor]
    for module in extra:
        name = module['name']
        ansi, unicode_name = name.encode(project.codec), name.encode('utf-16le')
        for identifier, value in ((0x19, ansi), (0x47, unicode_name),
                (0x1a, ansi), (0x32, unicode_name), (0x1c, b''), (0x48, b''),
                (0x31, bytes(4)), (0x1e, bytes(4)), (0x2c, b'\xff\xff'),
                (0x22 if module['component_type'] == 2 else 0x21, b''), (0x2b, b'')):
            directory.extend(record(identifier, value))
        streams[('VBA', name)] = ovba_pack(updates[name].encode(project.codec))
    directory.extend(suffix)
    struct.pack_into('<H', directory, modules_at+6, len(original)+len(extra))
    streams[('VBA', 'dir')] = ovba_pack(bytes(directory))
    streams[('VBA', '_VBA_PROJECT')] = bytes.fromhex('cc61ffff000100')
    for path in list(streams):
        if path[0] == 'VBA' and path[-1].startswith('__SRP_'):
            del streams[path]; meta.pop(path, None)
    project_text = streams[('PROJECT',)].decode(project.codec)
    new_declarations = ''.join(('Class=' if m['component_type'] == 2 else 'Module=')+
                              m['name']+'\r\n' for m in extra)
    position = project_text.index('Name=')
    streams[('PROJECT',)] = (project_text[:position]+new_declarations+project_text[position:]).encode(project.codec)
    wm = streams.get(('PROJECTwm',), b'\0\0')
    assert wm.endswith(b'\0\0')
    streams[('PROJECTwm',)] = wm[:-2] + b''.join(m['name'].encode(project.codec)+b'\0'+
        m['name'].encode('utf-16le')+b'\0\0' for m in extra) + b'\0\0'
    rebuilt = write_compound(meta, streams)
    check = olefile.OleFileIO(io.BytesIO(rebuilt))
    checked_project = VBA_Project(check, '', 'PROJECT', 'VBA/dir', relaxed=False)
    checked_project.parse_project_stream()
    list(checked_project.parse_modules())
    expected = {module.name: updates.get(module.name, module.code) for module in project.modules}
    expected.update({module['name']: updates[module['name']] for module in extra})
    def body(text):
        return '\n'.join(line.rstrip() for line in text.splitlines()
                         if not line.startswith('Attribute ')).strip()
    assert {m.name: body(m.code) for m in checked_project.modules} == {name: body(code) for name, code in expected.items()}
    before_meta, before_streams = read_compound(payload)
    after_meta, after_streams = read_compound(rebuilt)
    for path, data in before_streams.items():
        if path[0] != 'VBA' and path not in (('PROJECT',), ('PROJECTwm',)):
            assert after_streams[path] == data, ('Designer stream altered', path)
    for path, entry in before_meta.items():
        if path in after_meta:
            assert entry.clsid == after_meta[path].clsid
    return rebuilt, len(expected)


def replace_visio(stream, root, vba_node, payload):
    output = bytearray(stream)
    blocks = [payload[i:i+8192].ljust(8192, b'\0') for i in range(0, len(payload), 8192)]
    children = []
    template = vba_node['children'][0]['pointer']
    for block in blocks:
        encoded = visio_pack(struct.pack('<I', len(block))+block)
        pointer = (31, template[1], len(output), len(encoded), template[4])
        output.extend(encoded)
        children.append({'pointer': pointer, 'children': [], 'data': struct.pack('<I',len(block))+block})
    old_count = len(vba_node['children'])
    loc = vba_node['loc']
    data = bytearray(vba_node['data'][:loc+12])
    struct.pack_into('<III', data, 4, len(children)*22, len(payload), 0x20000000 | len(children))
    struct.pack_into('<I', data, loc+4, len(children))
    data.extend(b''.join(PTR.pack(*child['pointer']) for child in children))
    data.extend(vba_node['data'][loc+12+old_count*18:])
    vba_node['data'], vba_node['children'] = bytes(data), children

    def update(node):
        modified = node is vba_node
        for index, child in enumerate(node['children']):
            if node is vba_node:
                break
            if update(child):
                raw = bytearray(node['data'])
                PTR.pack_into(raw, node['loc']+12+index*18, *child['pointer'])
                node['data'] = bytes(raw)
                modified = True
        if modified:
            typ, ignored, offset, size, fmt = node['pointer']
            raw = node['data']
            if fmt & 2:
                raw = struct.pack('<I', len(raw)-4)+raw[4:]
                raw = visio_pack(raw)
            node['pointer'] = typ, ignored, len(output), len(raw), fmt
            output.extend(raw)
        return modified
    update(root)
    PTR.pack_into(output, 36, *root['pointer'])
    struct.pack_into('<I', output, 28, len(output))
    return bytes(output)


def payload_inventory(root):
    results = []
    def walk(node, path=()):
        if node['pointer'][0] == 13:
            return
        if not node['children']:
            results.append((path, node['pointer'][0], hashlib.sha256(node['data']).hexdigest()))
        for index, child in enumerate(node['children']):
            walk(child, path+(index,))
    walk(root)
    return results


def rewrite(source, target, modules, repo):
    metadata, streams = read_compound(source.read_bytes())
    original = streams[('VisioDocument',)]
    root = parse_visio(original)
    unchanged = payload_inventory(root)
    node, payload = vba_payload(root)
    rebuilt, module_count = replace_project(payload, modules, repo)
    streams[('VisioDocument',)] = replace_visio(original, root, node, rebuilt)
    result = write_compound(metadata, streams)
    checked_meta, checked_streams = read_compound(result)
    checked_root = parse_visio(checked_streams[('VisioDocument',)])
    assert payload_inventory(checked_root) == unchanged, 'Visio figure/style payload changed'
    assert vba_payload(checked_root)[1] == rebuilt
    for path, data in streams.items():
        assert checked_streams[path] == data
    target.write_bytes(result)
    return {'file': source.name, 'modules': module_count, 'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
            'output_sha256': hashlib.sha256(result).hexdigest(), 'native_verified': False}


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-directory', type=Path, required=True)
    parser.add_argument('--source-directory', type=Path)
    parser.add_argument('--only')
    args = parser.parse_args()
    repo = Path(__file__).resolve().parents[1]
    source_directory = (args.source_directory or repo).resolve()
    destination = args.output_directory.resolve()
    assert destination != source_directory and source_directory not in destination.parents
    destination.mkdir(parents=True, exist_ok=True)
    manifest = json.loads((repo/'manifest.json').read_text())
    report = []
    for project in manifest['projects']:
        if args.only and project['stencil'] != args.only:
            continue
        target = destination/project['stencil']
        assert not target.exists(), 'Refusing to overwrite an existing result'
        report.append(rewrite(source_directory/project['stencil'], target, project['modules'], repo))
        print(project['stencil'], 'PASS: embedded source, non-VBA payloads, designer streams', flush=True)
    (destination/'binary-rewrite-report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
