#!/usr/bin/env python3
"""Pretvori brend logotip (img/*.png) u VBA modul src-vba/modLogo.bas.

ZASTO OVO POSTOJI
-----------------
MSForms sliku moze da dobije samo kroz `LoadPicture`, a `LoadPicture` cita
FAJL -- ne bajtove. Ranije je logotip ziveo u `.frx`-u forme (frmSplash), sto
znaci dve stvari koje ne zelimo: `.frx` se ne pravi iz koda (CLAUDE.md par.3) i
`.frx` NE putuje kroz self-update, pa bi svaka promena logotipa trazila
REINSTALL na svakoj masini.

Resenje: logotip se pece u GIF, upisuje u `.bas` modul kao Base64 (cist ASCII,
pa prolazi `vba_check`), a u runtime-u se dekodira u privremeni fajl i ucitava
`LoadPicture`-om. Time logotip putuje kao KOD -- obicnim self-update-om.

ZASTO GIF, A NE PNG
-------------------
`LoadPicture` podrzava BMP, RLE, ICO, WMF, EMF, GIF i JPEG -- PNG NE. BMP je
nekomprimovan (pola megabajta za ovaj logotip), JPEG bi na ostrim ivicama
zlatnog teksta dao prsten. GIF je bez gubitka na ovoj slici: znak ima svega
nekoliko boja, pa 64-clana paleta pokriva i antialiasing.

ZASTO SE POZADINA PECE U SLIKU
------------------------------
MSForms ne zna per-pixel alfu -- providan PNG bi se svejedno spljostio, samo
na boju koju bi MSForms izabrao umesto nas. Zato kompozit radimo OVDE, na tacno
onu boju na kojoj slika stoji, a tu istu boju modul izvozi kao konstantu
(`LOGO_BG_*`), da se crtanje i slika ne mogu razici.

Splash ima gradijent, pa se boja uzima na sredini logotipa (t = 0.35). Razlika
boje gradijenta preko visine logotipa je najvise 1 jedinica po kanalu -- ravna
ploca iza slike se ne vidi.

UPOTREBA
--------
    python tools/logo_to_vba.py            # regenerisi src-vba/modLogo.bas
    python tools/logo_to_vba.py --proveri  # samo self-test enkodera, bez upisa

Self-test je dokaz u OBA smera (CLAUDE.md par.5): GIF se dekodira nazad i
poredi piksel po piksel, pa se namerno pokvari jedan bajt -- provera mora da
prijavi razliku. Zeleni enkoder koji nikad nije pokazan crven ne dokazuje nista.
"""
import base64
import os
import struct
import sys
import zlib

ROOT = os.path.dirname(os.path.abspath(os.path.dirname(__file__)))
SRC_PNG = os.path.join(ROOT, "img", "AgriX-Otkup-Logo-Final.png")
OUT_BAS = os.path.join(ROOT, "src-vba", "modLogo.bas")

# Boje ljuske -- iste vrednosti kao Public Const u modOtkupUI.bas.
C_FOREST = (0x1E, 0x2D, 0x14)
C_FOREST_DK = (0x14, 0x20, 0x0D)
C_CREAM = (0xF7, 0xF4, 0xEE)
# Gde na gradijentu splash-a stoji sredina logotipa. Isti broj stoji kao
# LOGO_T_SPLASH u modUiFaze -- ploca iza slike se crta bas tom bojom.
T_SPLASH = 0.35

# kljuc -> (pozadina, sirina, visina, broj boja)
VARIJANTE = [
    ("SPLASH", "forest", 480, 157, 64),
    ("KARTICA", "cream", 300, 98, 64),
    ("MINI", "cream", 160, 52, 48),
]


# ------------------------------------------------------------------ PNG citanje
def png_read(path):
    d = open(path, "rb").read()
    assert d[:8] == b"\x89PNG\r\n\x1a\n", "nije PNG"
    pos, idat, pal, trns = 8, [], None, None
    w = h = ct = None
    while pos < len(d):
        ln, typ = struct.unpack_from(">I4s", d, pos)
        body = d[pos + 8:pos + 8 + ln]
        if typ == b"IHDR":
            w, h, bd, ct, _, _, il = struct.unpack(">IIBBBBB", body)
            assert bd == 8 and il == 0, "podrzan je samo 8-bit PNG bez interlace-a"
        elif typ == b"PLTE":
            pal = body
        elif typ == b"tRNS":
            trns = body
        elif typ == b"IDAT":
            idat.append(body)
        elif typ == b"IEND":
            break
        pos += 12 + ln
    raw = zlib.decompress(b"".join(idat))
    ch = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ct]
    stride = w * ch
    out = bytearray(stride * h)
    prev = bytearray(stride)
    p = 0
    for y in range(h):
        f = raw[p]
        p += 1
        line = bytearray(raw[p:p + stride])
        p += stride
        if f == 1:
            for i in range(ch, stride):
                line[i] = (line[i] + line[i - ch]) & 255
        elif f == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 255
        elif f == 3:
            for i in range(stride):
                a = line[i - ch] if i >= ch else 0
                line[i] = (line[i] + ((a + prev[i]) >> 1)) & 255
        elif f == 4:
            for i in range(stride):
                a = line[i - ch] if i >= ch else 0
                b = prev[i]
                c = prev[i - ch] if i >= ch else 0
                pp = a + b - c
                pa, pb, pc = abs(pp - a), abs(pp - b), abs(pp - c)
                pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                line[i] = (line[i] + pr) & 255
        out[y * stride:(y + 1) * stride] = line
        prev = line
    rgba = bytearray(w * h * 4)
    for i in range(w * h):
        if ct == 6:
            r, g, b, a = out[i * 4:i * 4 + 4]
        elif ct == 2:
            r, g, b = out[i * 3:i * 3 + 3]
            a = 255
        elif ct == 0:
            r = g = b = out[i]
            a = 255
        elif ct == 4:
            r = g = b = out[i * 2]
            a = out[i * 2 + 1]
        else:
            idx = out[i]
            r, g, b = pal[idx * 3:idx * 3 + 3]
            a = trns[idx] if trns and idx < len(trns) else 255
        rgba[i * 4:i * 4 + 4] = bytes((r, g, b, a))
    return w, h, rgba


# --------------------------------------------------- kompozit i skaliranje
def lerp(c1, c2, t):
    """Isti racun koji modUiKit.Lerp radi u VBA -- gradijent mora da se poklopi."""
    return tuple(int(round(c1[i] + (c2[i] - c1[i]) * t)) for i in range(3))


def composite(w, h, rgba, bg):
    br, bgc, bb = bg
    out = bytearray(w * h * 3)
    for i in range(w * h):
        r, g, b, a = rgba[i * 4:i * 4 + 4]
        out[i * 3 + 0] = (r * a + br * (255 - a) + 127) // 255
        out[i * 3 + 1] = (g * a + bgc * (255 - a) + 127) // 255
        out[i * 3 + 2] = (b * a + bb * (255 - a) + 127) // 255
    return out


def resize_box(w, h, rgb, nw, nh):
    """Prosek povrsine -- pri smanjivanju cuva ivice bolje od najblizeg suseda."""
    out = bytearray(nw * nh * 3)
    for ny in range(nh):
        y0, y1 = ny * h // nh, max(ny * h // nh + 1, (ny + 1) * h // nh)
        for nx in range(nw):
            x0, x1 = nx * w // nw, max(nx * w // nw + 1, (nx + 1) * w // nw)
            sr = sg = sb = n = 0
            for y in range(y0, y1):
                base = y * w * 3
                for x in range(x0, x1):
                    o = base + x * 3
                    sr += rgb[o]
                    sg += rgb[o + 1]
                    sb += rgb[o + 2]
                    n += 1
            o = (ny * nw + nx) * 3
            out[o], out[o + 1], out[o + 2] = sr // n, sg // n, sb // n
    return out


def quantize(w, h, rgb, ncol):
    """Median cut na najvise ncol boja. Vraca (paleta, indeksi)."""
    px = [tuple(rgb[i * 3:i * 3 + 3]) for i in range(w * h)]
    uniq = {}
    for p in px:
        uniq[p] = uniq.get(p, 0) + 1
    if len(uniq) <= ncol:
        pal = sorted(uniq)
    else:
        boxes = [list(uniq.keys())]
        while len(boxes) < ncol:
            bi, best = -1, -1
            for i, b in enumerate(boxes):
                if len(b) < 2:
                    continue
                rng = max(max(p[c] for p in b) - min(p[c] for p in b) for c in range(3))
                if rng > best:
                    best, bi = rng, i
            if bi < 0:
                break
            b = boxes.pop(bi)
            c = max(range(3), key=lambda c: max(p[c] for p in b) - min(p[c] for p in b))
            b.sort(key=lambda p: p[c])
            m = len(b) // 2
            boxes += [b[:m], b[m:]]
        pal = []
        for b in boxes:
            tw = sum(uniq[p] for p in b)
            pal.append(tuple(sum(p[c] * uniq[p] for p in b) // tw for c in range(3)))
    cache, idx = {}, bytearray(w * h)
    for i, p in enumerate(px):
        j = cache.get(p)
        if j is None:
            j = min(range(len(pal)), key=lambda k: (pal[k][0] - p[0]) ** 2
                    + (pal[k][1] - p[1]) ** 2 + (pal[k][2] - p[2]) ** 2)
            cache[p] = j
        idx[i] = j
    return pal, idx


# --------------------------------------------------------------- GIF upis
def gif_encode(w, h, pal, idx):
    bits = max(2, (len(pal) - 1).bit_length())
    size = 1 << bits
    gct = bytearray()
    for i in range(size):
        gct += bytes(pal[i]) if i < len(pal) else b"\x00\x00\x00"
    out = bytearray(b"GIF89a")
    out += struct.pack("<HH", w, h)
    out += bytes([0xF0 | (bits - 1), 0, 0]) + gct
    out += b"\x2C" + struct.pack("<HHHH", 0, 0, w, h) + b"\x00"
    out += bytes([bits]) + _lzw(idx, bits)
    out += b"\x3B"
    return bytes(out)


def _lzw(data, bits):
    clear, end = 1 << bits, (1 << bits) + 1
    state = {"dic": {}, "nxt": end + 1, "cw": bits + 1, "buf": 0, "cnt": 0}
    chunk = bytearray()

    def emit(code):
        state["buf"] |= code << state["cnt"]
        state["cnt"] += state["cw"]
        while state["cnt"] >= 8:
            chunk.append(state["buf"] & 255)
            state["buf"] >>= 8
            state["cnt"] -= 8

    def reset():
        state["dic"] = {bytes([i]): i for i in range(clear)}
        state["nxt"], state["cw"] = end + 1, bits + 1

    reset()
    emit(clear)
    cur = b""
    for b in data:
        nc = cur + bytes([b])
        if nc in state["dic"]:
            cur = nc
        else:
            emit(state["dic"][cur])
            state["dic"][nc] = state["nxt"]
            state["nxt"] += 1
            if state["nxt"] > (1 << state["cw"]):
                if state["cw"] < 12:
                    state["cw"] += 1
                else:
                    emit(clear)
                    reset()
            cur = bytes([b])
    if cur:
        emit(state["dic"][cur])
    emit(end)
    if state["cnt"]:
        chunk.append(state["buf"] & 255)
    out = bytearray()
    for i in range(0, len(chunk), 255):
        blk = chunk[i:i + 255]
        out.append(len(blk))
        out += blk
    out.append(0)
    return bytes(out)


# ------------------------------------------------------- GIF citanje (samo test)
def gif_decode(d):
    assert d[:6] in (b"GIF89a", b"GIF87a"), "nije GIF"
    flags = d[10]
    assert flags & 0x80, "nema globalne palete"
    bits = (flags & 7) + 1
    n = 1 << bits
    p = 13
    pal = [tuple(d[p + i * 3:p + i * 3 + 3]) for i in range(n)]
    p += n * 3
    while d[p] != 0x2C:
        assert d[p] == 0x21, "neocekivan blok"
        p += 2
        while d[p]:
            p += d[p] + 1
        p += 1
    _, _, iw, ih = struct.unpack_from("<HHHH", d, p + 1)
    assert d[p + 9] & 0xC0 == 0, "lokalna paleta ili interlace"
    p += 10
    mincode = d[p]
    p += 1
    data = bytearray()
    while d[p]:
        ln = d[p]
        data += d[p + 1:p + 1 + ln]
        p += 1 + ln
    clear, end = 1 << mincode, (1 << mincode) + 1
    cw = mincode + 1
    dic = {i: bytes([i]) for i in range(clear)}
    nxt, prev, bitpos = end + 1, None, 0
    out = bytearray()
    total = len(data) * 8
    while bitpos + cw <= total:
        byte, off = bitpos >> 3, bitpos & 7
        code = (int.from_bytes(data[byte:byte + 3].ljust(3, b"\x00"), "little") >> off) \
            & ((1 << cw) - 1)
        bitpos += cw
        if code == clear:
            dic = {i: bytes([i]) for i in range(clear)}
            nxt, cw, prev = end + 1, mincode + 1, None
            continue
        if code == end:
            break
        if code in dic:
            entry = dic[code]
        elif code == nxt and prev is not None:
            entry = prev + prev[:1]
        else:
            raise AssertionError("nevazeci LZW kod %d" % code)
        out += entry
        if prev is not None:
            dic[nxt] = prev + entry[:1]
            nxt += 1
            if nxt == (1 << cw) and cw < 12:
                cw += 1
        prev = entry
    return iw, ih, pal, out


# ------------------------------------------------------------------ VBA upis
def vba_color(rgb):
    """RGB -> VBA Long (&HBBGGRR), kako ga RGB() gradi."""
    r, g, b = rgb
    return "&H%02X%02X%02X" % (b, g, r)


def b64_lines(b64, per=200):
    return [b64[i:i + per] for i in range(0, len(b64), per)]


def gradi(png_w, png_h, rgba):
    """Vrati listu (kljuc, bg_rgb, w, h, gif_bytes)."""
    rez = []
    for kljuc, poz, nw, nh, ncol in VARIJANTE:
        bg = lerp(C_FOREST, C_FOREST_DK, T_SPLASH) if poz == "forest" else C_CREAM
        rgb = composite(png_w, png_h, rgba, bg)
        small = resize_box(png_w, png_h, rgb, nw, nh)
        pal, idx = quantize(nw, nh, small, ncol)
        rez.append((kljuc, bg, nw, nh, gif_encode(nw, nh, pal, idx), pal, idx))
    return rez


def self_test(rez):
    """Dokaz u oba smera: povratno dekodiranje mora da vrati iste piksele, a
    pokvaren bajt mora da obori proveru."""
    ok = True
    for kljuc, _bg, nw, nh, gif, pal, idx in rez:
        dw, dh, dpal, didx = gif_decode(gif)
        isto = (dw, dh) == (nw, nh) and len(didx) == len(idx) \
            and all(dpal[didx[i]] == pal[idx[i]] for i in range(len(idx)))
        print("  %-8s %dx%d povratno dekodiranje: %s" % (kljuc, nw, nh, "ISTO" if isto else "RAZLIKA"))
        ok = ok and isto
    kljuc, _bg, nw, nh, gif, pal, idx = rez[-1]
    pokvaren = bytearray(gif)
    pokvaren[-40] ^= 0xFF
    try:
        _, _, dpal, didx = gif_decode(bytes(pokvaren))
        razlika = len(didx) != len(idx) \
            or any(dpal[didx[i]] != pal[idx[i]] for i in range(min(len(didx), len(idx))))
    except AssertionError:
        razlika = True
    print("  pokvaren bajt -> %s" % ("RAZLIKA (provera meri)" if razlika
                                     else "ISTO -- PROVERA NE MERI NISTA"))
    return ok and razlika


def ispisi_bas(rez):
    L = []
    a = L.append
    crta = "'" + "=" * 69
    a(dict(t='Attribute VB_Name = "modLogo"')["t"])
    a("Option Explicit")
    a("")
    a(crta)
    a("' modLogo - brend logotip kao KOD, ne kao .frx")
    a("'")
    a("' GENERISAN FAJL. Ne menjaj rukom: pokreni  python tools/logo_to_vba.py")
    a("' (izvor: img/AgriX-Otkup-Logo-Final.png).")
    a("'")
    a("' ZASTO OVAKO: MSForms sliku uzima samo kroz LoadPicture, a LoadPicture cita")
    a("' FAJL. Ranije je logotip ziveo u .frx-u forme, sto znaci dve stvari koje ne")
    a("' zelimo: .frx se ne pravi iz koda (CLAUDE.md par.3), i .frx NE putuje kroz")
    a("' self-update -- svaka promena logotipa bi trazila REINSTALL na svakoj masini.")
    a("' Ovako logotip putuje kao kod, obicnim self-update-om.")
    a("'")
    a("' Format je GIF jer LoadPicture ne cita PNG (zna BMP, RLE, ICO, WMF, EMF, GIF,")
    a("' JPEG). Na ovom znaku je GIF bez gubitka -- ima svega nekoliko boja.")
    a("'")
    a("' POZADINA JE PECENA U SLIKU. MSForms ne zna per-pixel alfu, pa bi providan")
    a("' PNG svejedno bio spljosten -- samo na boju koju bi MSForms izabrao umesto")
    a("' nas. Zato je kompozit uradjen unapred, na tacno onu boju na kojoj slika")
    a("' stoji, a ista ta boja izlazi kao LOGO_BG_* -- crtanje i slika se ne mogu")
    a("' razici.")
    a("'")
    a("' Fajl je 100% ASCII (Base64 i jeste ASCII).")
    a(crta)
    a("")
    for kljuc, bg, nw, nh, gif, _pal, _idx in rez:
        a("' %s: %dx%d, %d bajtova GIF-a" % (kljuc, nw, nh, len(gif)))
        a('Public Const LOGO_%s As String = "%s"' % (kljuc, kljuc))
    a("")
    a("' Boja na koju je slika pecena. Ploca iza slike se crta BAS ovim, pa se")
    a("' pravougaonik oko znaka ne vidi.")
    for kljuc, bg, _nw, _nh, _gif, _pal, _idx in rez:
        a("Public Const LOGO_BG_%s As Long = %s   ' RGB(%d, %d, %d)"
          % (kljuc, vba_color(bg), bg[0], bg[1], bg[2]))
    a("")
    a("' Odnos stranica (sirina / visina). Okvir slike se racuna po njemu, pa Zoom")
    a("' nema sta da doda sa strane -- inace bi se oko znaka video pojas pozadine.")
    for kljuc, _bg, nw, nh, _gif, _pal, _idx in rez:
        a("Public Const LOGO_ODNOS_%s As Single = %.4f" % (kljuc, nw / float(nh)))
    a("")
    a("' Ucitane slike po kljucu -- dekodiranje i upis na disk idu jednom po sesiji.")
    a("Private mKes As Object")
    a("")
    a(crta)
    a("' Slika za dati kljuc, ili Nothing ako je ucitavanje palo.")
    a("'")
    a("' NOTHING JE OCEKIVAN ISHOD, ne greska: MSXML ili ADODB mogu da nedostaju, a")
    a("' TEMP ume da bude nedostupan. Pozivalac tada crta tekstualni znak -- zato")
    a("' modUiFaze i dalje nosi natpise AX / OtkupApp.")
    a(crta)
    a("Public Function LogoSlika(ByVal kljuc As String) As Object")
    a('    Dim p As String')
    a("    On Error GoTo EH")
    a("    If mKes Is Nothing Then Set mKes = CreateObject(\"Scripting.Dictionary\")")
    a("    If mKes.Exists(kljuc) Then")
    a("        Set LogoSlika = mKes(kljuc)")
    a("        Exit Function")
    a("    End If")
    a("    p = UpisiPrivremeni(kljuc)")
    a('    If Len(p) = 0 Then Exit Function')
    a("    Set mKes(kljuc) = LoadPicture(p)")
    a("    Set LogoSlika = mKes(kljuc)")
    a("    Exit Function")
    a("EH:")
    a('    LogErr "modLogo.LogoSlika"')
    a("End Function")
    a("")
    a("' Otpusti ucitane slike (self-update rusi runtime, pa i ovaj kes).")
    a("Public Sub LogoOtpusti()")
    a("    Set mKes = Nothing")
    a("End Sub")
    a("")
    a("' Base64 -> bajtovi -> privremeni GIF. Vraca putanju ili \"\".")
    a("'")
    a("' Binarni upis ide istim obrascem kao modDrive.DriveDownloadToFile")
    a("' (ADODB.Stream, Type = 1, SaveToFile ... 2) -- nema drugog nacina u ovom")
    a("' projektu i ne uvodi se treci.")
    a("Private Function UpisiPrivremeni(ByVal kljuc As String) As String")
    a("    Dim dom As Object, cvor As Object, stm As Object, p As String, b64 As String")
    a("    On Error GoTo EH")
    a("    b64 = Base64Za(kljuc)")
    a("    If Len(b64) = 0 Then Exit Function")
    a('    p = Environ$("TEMP") & "\\AgriX_logo_" & kljuc & ".gif"')
    a("")
    a("    ' Fallback ide kroz Resume Next, ne kroz 'If dom Is Nothing': CreateObject")
    a("    ' nad nepostojecim ProgID-em DIZE gresku, ne vraca Nothing -- provera na")
    a("    ' Nothing se nikad ne bi izvrsila, a stara masina bi ostala bez logotipa.")
    a("    On Error Resume Next")
    a("    Set dom = CreateObject(\"MSXML2.DOMDocument.6.0\")")
    a("    If dom Is Nothing Then")
    a("        Err.Clear")
    a("        Set dom = CreateObject(\"MSXML2.DOMDocument\")")
    a("    End If")
    a("    On Error GoTo EH")
    a("    If dom Is Nothing Then Exit Function")
    a("    Set cvor = dom.createElement(\"b\")")
    a("    cvor.DataType = \"bin.base64\"")
    a("    cvor.text = b64")
    a("")
    a("    Set stm = CreateObject(\"ADODB.Stream\")")
    a("    stm.Type = 1                 ' adTypeBinary")
    a("    stm.Open")
    a("    stm.Write cvor.nodeTypedValue")
    a("    stm.SaveToFile p, 2          ' adSaveCreateOverWrite")
    a("    stm.Close")
    a("")
    a("    UpisiPrivremeni = p")
    a("    Exit Function")
    a("EH:")
    a('    LogErr "modLogo.UpisiPrivremeni"')
    a("End Function")
    a("")
    a("Private Function Base64Za(ByVal kljuc As String) As String")
    a("    Select Case kljuc")
    for kljuc, _bg, _nw, _nh, _gif, _pal, _idx in rez:
        a("        Case LOGO_%s: Base64Za = B64_%s()" % (kljuc, kljuc))
    a("    End Select")
    a("End Function")
    a("")
    a(crta)
    a("' Slike. Svaka je svoja procedura -- VBA ima granicu velicine procedure, a")
    a("' jedan zajednicki blok bi je s vremenom probio.")
    a(crta)
    for kljuc, _bg, nw, nh, gif, _pal, _idx in rez:
        b64 = base64.b64encode(gif).decode("ascii")
        a("")
        a("' %s -- %dx%d, %d B GIF-a, %d znakova Base64" % (kljuc, nw, nh, len(gif), len(b64)))
        a("Private Function B64_%s() As String" % kljuc)
        a("    Dim s As String")
        for ln in b64_lines(b64):
            a('    s = s & "%s"' % ln)
        a("    B64_%s = s" % kljuc)
        a("End Function")
    return "\r\n".join(L) + "\r\n"


def main():
    samo_test = "--proveri" in sys.argv
    w, h, rgba = png_read(SRC_PNG)
    print("izvor: %s  %dx%d" % (os.path.relpath(SRC_PNG, ROOT), w, h))
    rez = gradi(w, h, rgba)
    print("self-test enkodera:")
    if not self_test(rez):
        print("PAD: enkoder nije dokazan -- modLogo.bas NIJE upisan.")
        return 1
    if samo_test:
        print("--proveri: nista nije upisano.")
        return 0
    tekst = ispisi_bas(rez)
    assert all(ord(c) < 128 for c in tekst), "izlaz nije ASCII"
    open(OUT_BAS, "wb").write(tekst.encode("ascii"))
    print("upisano: %s (%d B)" % (os.path.relpath(OUT_BAS, ROOT), len(tekst)))
    for kljuc, _bg, nw, nh, gif, _pal, _idx in rez:
        print("  %-8s %3dx%-3d  gif %5d B  base64 %5d znakova"
              % (kljuc, nw, nh, len(gif), len(base64.b64encode(gif))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
