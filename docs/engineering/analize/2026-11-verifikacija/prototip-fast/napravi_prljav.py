"""Prljava kopija src-vba: nasumicne (ali ponovljive) povrede, da stara i nova
verzija checkera imaju STA da prijave -- identican izlaz nad cistim repoom je
slab dokaz jer je tamo nalaza nula."""
import os, random, re, shutil, sys
src, dst, seed = sys.argv[1], sys.argv[2], int(sys.argv[3])
BEZ_NEASCII = set(open(sys.argv[4]).read().split()) if len(sys.argv) > 4 else set()
rnd = random.Random(seed)
os.makedirs(dst, exist_ok=True)
EXT = (".bas", ".cls", ".frm", ".doccls")
n_izmena = 0
for ime in sorted(os.listdir(src)):
    p = os.path.join(src, ime)
    if not ime.endswith(EXT):
        shutil.copy(p, os.path.join(dst, ime)); continue
    raw = open(p, "rb").read()
    nl = b"\r\n" if b"\r\n" in raw else b"\n"
    redovi = raw.replace(b"\r\n", b"\n").split(b"\n")
    out = []
    for r in redovi:
        x = rnd.random()
        s = r.decode("latin-1")
        if x < 0.0006: n_izmena += 1; continue                       # obrisan red
        if x < 0.0012: out += [r, r]; n_izmena += 1; continue         # dupliran red
        if x < 0.0015 and ime not in BEZ_NEASCII: out.append(r + " \xc5\xa1".encode("latin-1")); n_izmena += 1; continue   # ne-ASCII
        if x < 0.0019 and re.match(r"\s+Dim \w+ As ", s):
            out += [r, b"    Dim eNum As Long", b"    eNum = poruka(1)"]; n_izmena += 1; continue
        if x < 0.0023 and re.match(r"\s*(Public|Private) (Sub|Function) \w+\(.*,", s):
            out.append(re.sub(r",[^,]*\)", ")", s, count=1).encode("latin-1")); n_izmena += 1; continue  # arnost
        if x < 0.0027 and re.match(r"\s+On Error Resume Next", s):
            out += [r, b'    LogErr "X"']; n_izmena += 1; continue
        if x < 0.0031 and s.startswith("End Sub"):
            out += [r, b'Private Const KASNA_%d As String = "x"' % n_izmena]; n_izmena += 1; continue
        if x < 0.0035 and re.match(r"\s+\w+ = ", s):
            out += [r, b'    MsgBox Poruka("NEMA_KLJUCA_%d"): NemaTeProcedure_%d 1, 2' % (n_izmena, n_izmena),
                    b"    frmOtkupUI.NemaTogClana_%d = 1" % n_izmena]; n_izmena += 1; continue
        if x < 0.0038 and re.match(r"\s+\w+ = ", s):
            out += [r, b"    nedeklarisanoIme_%d = 1 ' kom \"x" % n_izmena,
                    b'    d = ExcludeStornirano(d, TBL_NEPOZNATA_%d)' % n_izmena]; n_izmena += 1; continue
        if x < 0.0040 and re.match(r"\s+Dim (\w+) As String", s):
            ime_v = re.match(r"\s+Dim (\w+) As String", s).group(1)
            out += [r, ("    Dim %s As Long" % ime_v.upper()).encode(), ("    x = %s(1) & \"a'b\" ' k" % ime_v).encode()]; n_izmena += 1; continue
        out.append(r)
    open(os.path.join(dst, ime), "wb").write(nl.join(out))
print("izmena:", n_izmena)
