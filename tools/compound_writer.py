"""Write MS-CFB v3 while preserving storage metadata and stream bytes.

Small streams use the mini FAT. Storage children use validated red-black trees.
The reader is deliberately separate (olefile) and verifies every emitted stream.
"""
import io, struct, uuid
import olefile

FREE = 0xffffffff
END = 0xfffffffe
FAT = 0xfffffffd
DIF = 0xfffffffc


def read_compound(data):
    ole = olefile.OleFileIO(io.BytesIO(data))
    metadata = {(): ole.direntries[0]}
    streams = {}
    for path in ole.listdir(streams=True, storages=True):
        key = tuple(path)
        entry = ole.direntries[ole._find(path)]
        metadata[key] = entry
        if entry.entry_type == 2:
            streams[key] = ole.openstream(path).read()
    return metadata, streams


def write_compound(metadata, streams):
    paths = set(metadata) | set(streams)
    for path in list(paths):
        for n in range(len(path)):
            paths.add(path[:n])
    paths = [()] + sorted(paths - {()}, key=lambda p: (len(p), p))
    ids = {p: i for i, p in enumerate(paths)}
    count = len(paths)
    left, right, parent = [[FREE] * count for _ in range(3)]
    color = [1] * count
    child = [FREE] * count

    def key(i):
        name = paths[i][-1].upper()
        return len(name.encode('utf-16le')), tuple(ord(c) for c in name)

    for storage in paths:
        if storage in streams:
            continue
        members = [ids[p] for p in paths if p and p[:-1] == storage]
        root = FREE

        def rotate(x, direction):
            nonlocal root
            a, b = (right, left) if direction == 'left' else (left, right)
            y = a[x]
            a[x] = b[y]
            if b[y] != FREE:
                parent[b[y]] = x
            parent[y] = parent[x]
            if parent[x] == FREE:
                root = y
            elif x == left[parent[x]]:
                left[parent[x]] = y
            else:
                right[parent[x]] = y
            b[y] = x
            parent[x] = y

        for z in members:
            y, x = FREE, root
            while x != FREE:
                y = x
                x = left[x] if key(z) < key(x) else right[x]
            parent[z] = y
            if y == FREE:
                root = z
            elif key(z) < key(y):
                left[y] = z
            else:
                right[y] = z
            color[z] = 0
            while parent[z] != FREE and color[parent[z]] == 0:
                p, g = parent[z], parent[parent[z]]
                if p == left[g]:
                    uncle = right[g]
                    if uncle != FREE and color[uncle] == 0:
                        color[p] = color[uncle] = 1
                        color[g] = 0
                        z = g
                    else:
                        if z == right[p]:
                            z = p
                            rotate(z, 'left')
                        color[parent[z]] = 1
                        color[parent[parent[z]]] = 0
                        rotate(parent[parent[z]], 'right')
                else:
                    uncle = left[g]
                    if uncle != FREE and color[uncle] == 0:
                        color[p] = color[uncle] = 1
                        color[g] = 0
                        z = g
                    else:
                        if z == left[p]:
                            z = p
                            rotate(z, 'right')
                        color[parent[z]] = 1
                        color[parent[parent[z]]] = 0
                        rotate(parent[parent[z]], 'left')
            color[root] = 1
        child[ids[storage]] = root

        def validate(node):
            if node == FREE:
                return 1
            if color[node] == 0:
                assert all(c == FREE or color[c] == 1 for c in (left[node], right[node]))
            a, b = validate(left[node]), validate(right[node])
            assert a == b, 'Unequal CFB tree black heights'
            return a + color[node]
        validate(root)

    sectors, fat, mini, minifat = [], [], bytearray(), []
    start, size = [END] * count, [0] * count

    def allocate(data):
        if not data:
            return END
        first = len(sectors)
        for pos in range(0, len(data), 512):
            sectors.append(data[pos:pos+512].ljust(512, b'\0'))
            fat.append(len(sectors))
        fat[-1] = END
        return first

    for path, data in streams.items():
        i = ids[path]
        size[i] = len(data)
        if not data:
            continue
        if len(data) < 4096:
            start[i] = len(minifat)
            for pos in range(0, len(data), 64):
                mini.extend(data[pos:pos+64].ljust(64, b'\0'))
                minifat.append(len(minifat)+1)
            minifat[-1] = END
        else:
            start[i] = allocate(data)
    start[0], size[0] = allocate(mini), len(mini)
    directory = bytearray()
    for i, path in enumerate(paths):
        name = path[-1] if path else 'Root Entry'
        encoded = (name + '\0').encode('utf-16le')
        assert len(encoded) <= 64
        old = metadata.get(path)
        clsid = uuid.UUID(old.clsid).bytes_le if old and old.clsid else bytes(16)
        typ = 5 if not path else 2 if path in streams else 1
        directory.extend(struct.pack('<64sHBBIII16sIQQIQ', encoded, len(encoded), typ,
            color[i], left[i], right[i], child[i], clsid,
            old.dwUserFlags if old else 0, old.createTime if old else 0,
            old.modifyTime if old else 0, start[i], size[i]))
    first_dir = allocate(directory)
    first_mini = allocate(struct.pack('<%dI' % len(minifat), *minifat) +
        b'\xff' * ((-len(minifat)*4) % 512)) if minifat else END
    nmini = (len(minifat)*4 + 511)//512
    nfat, ndif = 1, 0
    while True:
        newfat = (len(sectors)+nfat+ndif+127)//128
        newdif = max(0, (newfat-109+126)//127)
        if (newfat, newdif) == (nfat, ndif):
            break
        nfat, ndif = newfat, newdif
    dif_ids = list(range(len(sectors), len(sectors)+ndif))
    fat_ids = list(range(len(sectors)+ndif, len(sectors)+ndif+nfat))
    for k, sid in enumerate(dif_ids):
        entries = fat_ids[109+k*127:109+(k+1)*127]
        entries += [FREE] * (127-len(entries))
        entries += [dif_ids[k+1] if k+1 < ndif else END]
        sectors.append(struct.pack('<128I', *entries))
        fat.append(DIF)
    fat += [FAT] * nfat
    fat += [FREE] * (nfat*128-len(fat))
    for k in range(nfat):
        sectors.append(struct.pack('<128I', *fat[k*128:(k+1)*128]))
    header = bytearray(512)
    header[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<HHHHH', header, 24, 0x3e, 3, 0xfffe, 9, 6)
    struct.pack_into('<IIIIIIIII', header, 40, 0, nfat, first_dir, 0, 4096,
        first_mini, nmini, dif_ids[0] if ndif else END, ndif)
    struct.pack_into('<109I', header, 76, *(fat_ids[:109] + [FREE]*max(0,109-nfat)))
    result = bytes(header) + b''.join(sectors)
    new_meta, new_streams = read_compound(result)
    assert new_streams == streams, 'CFB stream read-back mismatch'
    assert set(new_meta) == set(paths)
    return result
