"""Dvosmerni dokaz za nove provere: pokvari -> self-test pada PO IMENU -> vrati."""
import subprocess, sys, os
ROOT = "/home/user/otkupapp-pwa"
os.chdir(ROOT)
D, R, G = "tools/dokaz.py", "tools/run_vba.py", "tools/vba_gate.py"
MT, BFP = "src-vba/modTest.bas", "src-vba/modBusinessFlowProTests.bas"
ST = {D: [sys.executable, D, "--self-test"], R: [sys.executable, R, "--self-test"],
      G: [sys.executable, G, "--self-test"]}
SLUCAJEVI = [
 (D, D, "            baze[kljuc] = _baza_zelena(suite, testovi)\n",
        "            baze[kljuc] = (True, 'x')\n", "baza izbora se ne meri PRE"),
 (D, D, "        run = _pusti(*_run_suite(suite, testovi))\n",
        "        run = _pusti(*_run_suite(suite))\n", "ne pustaju TACNO testove clanova"),
 (D, D, "        if kljuc not in baze:\n", "        if True:\n", "placa ponovo"),
 (D, D, "        if not ok:\n            for s in clanovi:\n",
        "        if False:\n            for s in clanovi:\n", "crvena baza grupe ipak"),
 (D, D, "                if grupno and len(clanovi) > 1:\n", "                if False:\n",
        "ne rastura na pojedinacna"),
 (D, D, "    if testovi and int(m.group(1)) == 0:\n", "    if False:\n", "bez ijedne provere"),
 (D, D, "    if suite in izbor:\n", "    if izbor:\n", "suite bez ulazne tacke izbora"),
 (D, D, "        if s[\"suite\"] not in izbor or kljuc in vidjeni:\n",
        "        if s[\"suite\"] not in izbor:\n", "IZOLACIJA: test se pusta vise puta"),
 (D, D, "        if not ok:\n            crveni.append", "        if ok:\n            crveni.append",
        "IZOLACIJA: ne vraca bas test"),
 (R, R, "    if args.all or len(args.suite) != 1:\n", "    if args.all:\n", "problem_izbora puca"),
 (R, R, "    if args.all or len(args.suite) != 1:\n", "    if args.all or not args.suite:\n",
        "--test uz DVE suite se prihvata"),
 (R, R, "    if not izabrani_testovi(args):\n", "    if False:\n", "--test bez ijednog imena se prihvata"),
 (R, R, "    if not meta.get(\"izbor\"):\n", "    if False:\n", "BEZ ulazne tacke izbora se prihvata"),
 (R, R, "    if args.compile_only:\n        return \"--test", "    if False:\n        return \"--test",
        "--compile-only se prihvata"),
 (R, R, "    out.sort(key=lambda x: -x[0])\n", "    pass\n", "vreme se ne cita kako ga VBA pise"),
 (R, R, "                        entry[\"izbor\"] = testovi\n", "                        pass\n",
        "izvestaj ne oznacava izbor"),
 (R, R, "    razlog = problem_izbora(args)\n", "    razlog = ''\n", "main() ne proverava kombinaciju"),
 (R, MT, "    m_Izbor = PreuzmiIzbor(m_IzborZahtev)\r\n", "    m_Izbor = m_IzborZahtev\r\n",
        "RunAllTests ne preuzima"),
 (R, MT, "    If Not TestIzabran(nm) Then Exit Sub\r\n", "    ' bez izbora\r\n", "RunOne ne pita TestIzabran"),
 (R, BFP, "    m_BfpIzbor = BfpPreuzmiIzbor(m_BfpIzborZahtev)\r\n", "    m_BfpIzbor = m_BfpIzborZahtev\r\n",
        "ne preuzima izbor kao PRVU"),
 (R, BFP, "    If BfpIzabran(\"Test_MalinaVozacMirror\") Then Test_MalinaVozacMirror\r\n",
        "    Test_MalinaVozacMirror\r\n", "zove test mimo BfpIzabran"),
 (G, G, "and d[\"vrsta\"] == \"sub\" and d[\"obaveznih\"] == 1 for d in svi):",
        "and d[\"vrsta\"] == \"sub\" for d in svi):", "izbora 'BEZ argumenta' prolazi"),
 (G, G, "    if izbori:\n        return (", "    if False:\n        return (",
        "izbor testova (--test) se upisuje"),
]
ok = True
for alat, fajl, staro, novo, igla in SLUCAJEVI:
    raw = open(fajl, "rb").read()
    s = raw.decode("ascii")
    assert s.count(staro) == 1, (fajl, staro)
    open(fajl, "wb").write(s.replace(staro, novo).encode("ascii"))
    try:
        r = subprocess.run(ST[alat], capture_output=True, text=True)
    finally:
        open(fajl, "wb").write(raw)
    izl = r.stdout + r.stderr
    pogodak = r.returncode != 0 and igla in izl and "Traceback" not in izl
    ok &= pogodak
    print("%-5s %-14s %-48s rc=%d" % ("CRVEN" if pogodak else "PROMASIO",
          os.path.basename(alat), igla[:48], r.returncode))
    if not pogodak:
        print(izl[-600:])
# posle vracanja: sve zeleno
for alat in (D, R, G):
    r = subprocess.run(ST[alat], capture_output=True, text=True)
    print("ZELEN " if r.returncode == 0 else "PAO   ", alat, (r.stdout + r.stderr).strip()[:90])
    ok &= r.returncode == 0
print("DVOSMERNO:", "OK" if ok else "NIJE")
