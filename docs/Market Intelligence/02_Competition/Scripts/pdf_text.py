#!/usr/bin/env python3
"""Izvlacenje teksta iz konkurentskih PDF uputstava, bez spoljnih zavisnosti.

Zasto postoji: u web/Linux sesiji nema `pdftotext` (poppler), a `pip install pypdf`
pada jer je `cryptography` u okruzenju polomljen. Ovaj skript koristi samo stdlib
(`re`, `zlib`, `struct`) i resava tri stvari zbog kojih naivni ekstraktori pucaju
bas na ovim dokumentima:

1. **ObjStm** — PDF 1.5+ pakuje page/font recnike u komprimovane object streams,
   pa ih regex nad sirovim fajlom ne vidi. Skript ih raspakuje.
2. **Type0 / Identity-H fontovi** — tekst je u 2-bajtnim CID-ovima. Mapa se cita iz
   `/ToUnicode`, a gde je ona nepotpuna (Word ume da izostavi slova, npr. 'A' u
   SOFTEK uputstvu) dopunjuje se iz `cmap` tabele ugradjenog TrueType fajla.
   Kod SOFTEK-a je `/ToUnicode` i pogresna za 'c' sa kvacicom, pa TTF ima prednost.
3. **Rekonstrukcija redova** — tekst se cesto postavlja po glifu (`Td`/`Tm`), pa se
   prelom reda izvodi iz promene Y koordinate, a ne iz operatora.

Koriscenje:

    python3 pdf_text.py <ulaz.pdf> <izlaz.txt>

Izlaz je obican tekst sa markerima `===== STRANA n =====`.

Provereno na: AgroSoft-Korisnicko-Uputsvo.pdf (161 str, Type0+ToUnicode),
SOFTEK_uputstvp_otkup_poljoproizvoda.pdf (34 str, ObjStm + nepotpun ToUnicode),
Softek-otkup.pdf (16 str, PDFCreator/Ghostscript).
"""

import re
import struct
import sys
import zlib

# --- TrueType cmap -> {gid: unicode} ---------------------------------------


def gid_to_unicode(ttf):
    """Invertuje `cmap` tabelu ugradjenog TTF-a. Za CIDFontType2 sa Identity
    mapiranjem vazi cid == gid, pa je ovo tacna dopuna za /ToUnicode."""
    if len(ttf) < 12:
        return {}
    num_tables = struct.unpack(">H", ttf[4:6])[0]
    cmap_off = None
    for i in range(num_tables):
        rec = 12 + 16 * i
        if ttf[rec:rec + 4] == b"cmap":
            cmap_off = struct.unpack(">I", ttf[rec + 8:rec + 12])[0]
            break
    if cmap_off is None or cmap_off + 4 > len(ttf):
        return {}

    best = None
    for i in range(struct.unpack(">H", ttf[cmap_off + 2:cmap_off + 4])[0]):
        rec = cmap_off + 4 + 8 * i
        pid, eid, off = struct.unpack(">HHI", ttf[rec:rec + 8])
        sub = cmap_off + off
        if sub + 2 > len(ttf):
            continue
        fmt = struct.unpack(">H", ttf[sub:sub + 2])[0]
        score = {(3, 10): 5, (3, 1): 4, (0, 4): 3, (0, 3): 3,
                 (0, 6): 3, (0, 1): 2, (3, 0): 1}.get((pid, eid), 0)
        if best is None or score > best[0]:
            best = (score, sub, fmt)
    if not best:
        return {}

    _, sub, fmt = best
    out = {}
    if fmt == 4:
        seg_x2 = struct.unpack(">H", ttf[sub + 6:sub + 8])[0]
        seg = seg_x2 // 2
        end_o = sub + 14
        start_o = end_o + seg_x2 + 2
        delta_o = start_o + seg_x2
        range_o = delta_o + seg_x2
        for s in range(seg):
            end = struct.unpack(">H", ttf[end_o + 2 * s:end_o + 2 * s + 2])[0]
            start = struct.unpack(">H", ttf[start_o + 2 * s:start_o + 2 * s + 2])[0]
            delta = struct.unpack(">h", ttf[delta_o + 2 * s:delta_o + 2 * s + 2])[0]
            ro = struct.unpack(">H", ttf[range_o + 2 * s:range_o + 2 * s + 2])[0]
            if start == 0xFFFF:
                continue
            for c in range(start, min(end, 0xFFFE) + 1):
                if ro == 0:
                    gid = (c + delta) & 0xFFFF
                else:
                    gi = range_o + 2 * s + ro + 2 * (c - start)
                    if gi + 2 > len(ttf):
                        continue
                    gid = struct.unpack(">H", ttf[gi:gi + 2])[0]
                    if gid:
                        gid = (gid + delta) & 0xFFFF
                if gid and gid not in out:
                    out[gid] = chr(c)
    elif fmt == 12:
        ngroups = struct.unpack(">I", ttf[sub + 12:sub + 16])[0]
        for i in range(ngroups):
            o = sub + 16 + 12 * i
            sc, ec, sg = struct.unpack(">III", ttf[o:o + 12])
            for c in range(sc, min(ec, sc + 65535) + 1):
                out.setdefault(sg + (c - sc), chr(c))
    return out


# --- PDF ---------------------------------------------------------------------

WIN = {0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„',
       0x85: '…', 0x86: '†', 0x87: '‡', 0x88: 'ˆ',
       0x89: '‰', 0x8a: 'Š', 0x8b: '‹', 0x8c: 'Œ',
       0x8e: 'Ž', 0x91: '‘', 0x92: '’', 0x93: '“',
       0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—',
       0x98: '˜', 0x99: '™', 0x9a: 'š', 0x9b: '›',
       0x9c: 'œ', 0x9e: 'ž', 0x9f: 'Ÿ'}

TOK = re.compile(rb"/([A-Za-z0-9#+,.-]+)\s+[-\d.]+\s+Tf"
                 rb"|(-?[\d.]+)\s+(-?[\d.]+)\s+(?:TD|Td)"
                 rb"|(?:-?[\d.]+\s+){4}(-?[\d.]+)\s+(-?[\d.]+)\s+Tm"
                 rb"|\((?:\\.|[^\\()])*\)"
                 rb"|<[0-9A-Fa-f\s]+>"
                 rb"|\[(?:[^\]\\]|\\.)*\]\s*TJ"
                 rb"|T\*", re.S)


class Pdf:
    def __init__(self, data):
        self.d = data
        self.objs = {}
        for m in re.finditer(rb"(?<![0-9])(\d+)\s+(\d+)\s+obj\b(.*?)\bendobj",
                             data, re.S):
            self.objs[int(m.group(1))] = m.group(3)
        self._expand_objstm()
        self._fcache = {}

    # -- streams --
    def stream(self, body):
        m = re.search(rb"stream\r?\n", body)
        if not m:
            return None
        end = body.rfind(b"endstream")
        if end < 0:
            return None
        hdr, raw = body[:m.start()], body[m.end():end]
        if b"/FlateDecode" not in hdr:
            return raw
        for trim in (0, 1, 2):
            try:
                chunk = raw[:len(raw) - trim] if trim else raw
                return zlib.decompressobj().decompress(chunk)
            except Exception:
                continue
        return None

    def _expand_objstm(self):
        for num in list(self.objs):
            body = self.objs[num]
            if b"/ObjStm" not in body[:400]:
                continue
            data = self.stream(body)
            if not data:
                continue
            n = int(re.search(rb"/N\s+(\d+)", body).group(1))
            first = int(re.search(rb"/First\s+(\d+)", body).group(1))
            head = data[:first].split()
            for i in range(n):
                onum = int(head[2 * i])
                off = int(head[2 * i + 1])
                end = int(head[2 * i + 3]) + first if i + 1 < n else len(data)
                self.objs.setdefault(onum, data[first + off:end])

    def ref(self, body, key):
        m = re.search(key + rb"\s+(\d+)\s+\d+\s+R", body)
        return self.objs.get(int(m.group(1)), b"") if m else None

    # -- fonts --
    @staticmethod
    def _tounicode(data):
        cmap = {}
        for m in re.finditer(rb"beginbfchar(.*?)endbfchar", data, re.S):
            for src, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>",
                                       m.group(1)):
                try:
                    cmap[int(src, 16)] = bytes.fromhex(
                        dst.decode()).decode("utf-16-be", "replace")
                except Exception:
                    pass
        for m in re.finditer(rb"beginbfrange(.*?)endbfrange", data, re.S):
            for lo, hi, dst in re.findall(
                    rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>",
                    m.group(1)):
                a, b, base = int(lo, 16), int(hi, 16), int(dst, 16)
                for i in range(a, min(b, a + 65535) + 1):
                    try:
                        cmap[i] = chr(base + (i - a))
                    except Exception:
                        pass
        return cmap

    def font(self, num):
        """-> (cmap, dvobajtni)"""
        if num in self._fcache:
            return self._fcache[num]
        body = self.objs.get(num, b"")
        two = b"/Type0" in body
        cmap = {}
        s = self.ref(body, rb"/ToUnicode")
        if s is not None:
            data = self.stream(s)
            if data:
                cmap = self._tounicode(data)
        if two:
            desc = self.ref(body, rb"/DescendantFonts")
            if desc is not None and b"/FontDescriptor" not in desc:
                inner = re.search(rb"(\d+)\s+\d+\s+R", desc)
                desc = self.objs.get(int(inner.group(1)), b"") if inner else desc
            if desc:
                fd = self.ref(desc, rb"/FontDescriptor")
                ff = self.ref(fd, rb"/FontFile2") if fd is not None else None
                ttf = self.stream(ff) if ff is not None else None
                if ttf:
                    try:
                        # TTF ima prednost: /ToUnicode ume da bude i nepotpuna i
                        # pogresna (SOFTEK: 'c' sa kvacicom mapirana na 'd').
                        merged = dict(cmap)
                        merged.update(gid_to_unicode(ttf))
                        cmap = merged
                    except Exception:
                        pass
        if not cmap and not two:
            cmap = dict(WIN)
            m = re.search(rb"/Differences\s*\[(.*?)\]", body, re.S)
            if m:
                code = 0
                for tok in re.findall(rb"(\d+)|/([A-Za-z0-9.]+)", m.group(1)):
                    if tok[0]:
                        code = int(tok[0])
                    else:
                        name = tok[1].decode()
                        if name.startswith("uni"):
                            try:
                                cmap[code] = chr(int(name[3:], 16))
                            except Exception:
                                pass
                        code += 1
        self._fcache[num] = (cmap, two)
        return self._fcache[num]

    # -- pages --
    def pages(self):
        out, seen = [], set()

        def walk(num):
            if num in seen:
                return
            seen.add(num)
            b = self.objs.get(num, b"")
            if re.search(rb"/Type\s*/Pages", b):
                m = re.search(rb"/Kids\s*\[(.*?)\]", b, re.S)
                if m:
                    for k in re.findall(rb"(\d+)\s+\d+\s+R", m.group(1)):
                        walk(int(k))
            elif re.search(rb"/Type\s*/Page[^s]", b):
                out.append(num)

        m = re.search(rb"/Root\s+(\d+)\s+\d+\s+R", self.d)
        if m:
            pm = re.search(rb"/Pages\s+(\d+)\s+\d+\s+R",
                           self.objs.get(int(m.group(1)), b""))
            if pm:
                walk(int(pm.group(1)))
        if not out:
            out = sorted(n for n, b in self.objs.items()
                         if re.search(rb"/Type\s*/Page[^s]", b))
        return out

    def text(self, page_num):
        body = self.objs[page_num]
        res = self.ref(body, rb"/Resources")
        res = body if res is None else res
        fonts = {}
        fm = re.search(rb"/Font\s*<<(.*?)>>", res, re.S)
        src = fm.group(1) if fm else (self.ref(res, rb"/Font") or b"")
        for name, num in re.findall(rb"/([A-Za-z0-9#+,.-]+)\s+(\d+)\s+\d+\s+R", src):
            fonts[name.decode()] = int(num)

        cont = re.findall(rb"/Contents\s+(\d+)\s+\d+\s+R", body)
        if not cont:
            m = re.search(rb"/Contents\s*\[(.*?)\]", body, re.S)
            cont = re.findall(rb"(\d+)\s+\d+\s+R", m.group(1)) if m else []
        data = b""
        for c in cont:
            s = self.stream(self.objs.get(int(c), b""))
            if s:
                data += s + b"\n"

        out, cmap, two, lasty = [], {}, False, None
        for m in TOK.finditer(data):
            t = m.group(0)
            if m.group(1):
                n = fonts.get(m.group(1).decode())
                cmap, two = self.font(n) if n else ({}, False)
            elif m.group(2) is not None:
                out.append("\n" if abs(float(m.group(3))) > 0.5 else " ")
            elif m.group(4) is not None:
                y = float(m.group(5))
                out.append("\n" if (lasty is None or abs(y - lasty) > 0.5) else " ")
                lasty = y
            elif t.startswith(b"("):
                out.append(_decode(_unescape(t[1:-1]), cmap, two))
            elif t.startswith(b"<"):
                out.append(_hex(t, cmap, two))
            elif t.startswith(b"["):
                inner = t[1:t.rfind(b"]")]
                for mm in re.finditer(
                        rb"\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]+>|(-?[\d.]+)",
                        inner, re.S):
                    s = mm.group(0)
                    if s.startswith(b"("):
                        out.append(_decode(_unescape(s[1:-1]), cmap, two))
                    elif s.startswith(b"<"):
                        out.append(_hex(s, cmap, two))
                    else:
                        try:
                            if float(s) < -150:
                                out.append(" ")
                        except ValueError:
                            pass
            elif t == b"T*":
                out.append("\n")

        txt = "".join(out)
        txt = re.sub(r"[ \t]+", " ", txt)
        txt = re.sub(r"\n[ ]+", "\n", txt)
        return re.sub(r"\n{3,}", "\n\n", txt)


def _unescape(b):
    out, i = bytearray(), 0
    mp = {ord('n'): 10, ord('r'): 13, ord('t'): 9, ord('b'): 8,
          ord('f'): 12, ord('('): 40, ord(')'): 41, ord('\\'): 92}
    while i < len(b):
        c = b[i]
        if c == 0x5C and i + 1 < len(b):
            n = b[i + 1]
            if n in mp:
                out.append(mp[n])
                i += 2
                continue
            if 0x30 <= n <= 0x37:
                j, o = i + 1, b""
                while j < len(b) and len(o) < 3 and 0x30 <= b[j] <= 0x37:
                    o += bytes([b[j]])
                    j += 1
                out.append(int(o, 8) & 0xFF)
                i = j
                continue
            i += 2
            continue
        out.append(c)
        i += 1
    return bytes(out)


def _decode(raw, cmap, two):
    if two:
        return "".join(cmap.get((raw[i] << 8) | raw[i + 1], "")
                       for i in range(0, len(raw) - 1, 2))
    if not cmap:
        return raw.decode("latin-1", "replace")
    return "".join(cmap.get(ch, chr(ch)) for ch in raw)


def _hex(tok, cmap, two):
    h = re.sub(rb"\s", b"", tok[1:-1])
    if len(h) % 2:
        h += b"0"
    try:
        return _decode(bytes.fromhex(h.decode()), cmap, two)
    except ValueError:
        return ""


def main(argv):
    if len(argv) != 3:
        print(__doc__.strip().splitlines()[0])
        print("\nupotreba: python3 pdf_text.py <ulaz.pdf> <izlaz.txt>")
        return 2
    pdf = Pdf(open(argv[1], "rb").read())
    pages = pdf.pages()
    with open(argv[2], "w", encoding="utf-8") as f:
        for i, p in enumerate(pages, 1):
            f.write("\n===== STRANA %d =====\n" % i)
            f.write(pdf.text(p))
    print("strana: %d -> %s" % (len(pages), argv[2]))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
