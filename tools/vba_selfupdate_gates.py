"""Strukturne kapije nad modSelfUpdate: nijedan save bez dokaza.

Atomarnost self-update-a ne pociva na jednoj proveri nego na tome da se
`ThisWorkbook.Save` NIKAD ne dosegne bez zavrsne provere release-a. To je
invarijanta RASPOREDA koda, a takva invarijanta ne moze da se testira suite-om:
put do nje se otvara tek kad neko doda TRECI uspesan izlaz i zaboravi kapiju.
Tada nema crvenog testa -- ima samo klijenta koji je snimio polu-nov projekat.

    python tools/vba_selfupdate_gates.py              # provera nad src-vba/
    python tools/vba_selfupdate_gates.py --self-test  # dokaz da provera hvata kvar

Izlazni kod: 0 = cisto, 2 = ima nalaza.

Provere:
  KAPIJA_SAVE   -- svaki poziv `SaveWorkbookVerified` mora imati `VerifyReleaseProject`
                   RANIJE u ISTOJ proceduri. Bez toga postoji uspesan put do save-a
                   koji nije dokazao da projekat odgovara release-u.
  KAPIJA_ABORT  -- svaka procedura koja zove `VerifyReleaseProject` mora na neuspeh
                   zvati `AbortSelfUpdate`. Provera koja se ignorise nije kapija.
  KAPIJA_POSTOJI-- obe procedure moraju postojati (da preimenovanje ne ugasi proveru
                   tiho, tako sto vise nema sta da se nadje).
"""

import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")

MODUL = "modSelfUpdate.bas"
SAVE_PROC = "SaveWorkbookVerified"
GATE_PROC = "VerifyReleaseProject"
ABORT_PROC = "AbortSelfUpdate"

_PROC_HEAD = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE,
)
_PROC_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\s*$", re.IGNORECASE)


def strip_comment(line):
    """Odseci trailing komentar. Apostrof UNUTAR stringa nije komentar.

    Bitno: komentari OBILNO pominju i SaveWorkbookVerified i VerifyReleaseProject.
    Provera koja bi ih brojala citala bi objasnjenje kao kod -- i ostala zelena
    nad modulom u kome kapije vise nema, a komentar o njoj je ostao.
    """
    out = []
    inq = False
    for ch in line:
        if ch == '"':
            inq = not inq
        if ch == "'" and not inq:
            break
        out.append(ch)
    return "".join(out)


def read_procs(path):
    """[(ime, [kodni redovi bez komentara])] redom kojim se javljaju."""
    txt = io.open(path, encoding="ascii", newline="").read().replace("\r\n", "\n")
    res = []
    cur = None
    buf = []
    for ln in txt.split("\n"):
        if cur is None:
            m = _PROC_HEAD.match(ln)
            if m:
                cur = m.group(1)
                buf = []
            continue
        if _PROC_END.match(ln):
            res.append((cur, buf))
            cur = None
            continue
        buf.append(strip_comment(ln))
    return res


def _prvi_indeks(telo, ime):
    """Indeks prvog reda koji ZOVE `ime`, ili -1."""
    rx = re.compile(r"(?<![A-Za-z0-9_])" + re.escape(ime) + r"(?![A-Za-z0-9_])")
    for i, ln in enumerate(telo):
        if rx.search(ln):
            return i
    return -1


def check_gates(path=None):
    path = path or os.path.join(SRC_VBA, MODUL)
    if not os.path.isfile(path):
        return ["KAPIJA_POSTOJI  {}: fajl ne postoji".format(path)]

    procs = read_procs(path)
    imena = set(n for n, _ in procs)
    nalazi = []

    for ime in (SAVE_PROC, GATE_PROC, ABORT_PROC):
        if ime not in imena:
            nalazi.append("KAPIJA_POSTOJI  nema procedure {} -- provera bi ostala "
                          "zelena jer vise nema sta da nadje".format(ime))
    if nalazi:
        return nalazi

    pozivalaca = 0
    for ime, telo in procs:
        if ime in (SAVE_PROC, GATE_PROC):
            continue                      # definicija same procedure nije poziv

        i_save = _prvi_indeks(telo, SAVE_PROC)
        i_gate = _prvi_indeks(telo, GATE_PROC)

        if i_save >= 0:
            pozivalaca += 1
            if i_gate < 0:
                nalazi.append(
                    "KAPIJA_SAVE  {}: zove {} BEZ {} -- uspesan put do save-a koji "
                    "nije dokazao da projekat odgovara release-u"
                    .format(ime, SAVE_PROC, GATE_PROC))
            elif i_gate > i_save:
                nalazi.append(
                    "KAPIJA_SAVE  {}: {} je POSLE {} -- kapija posle save-a nije "
                    "kapija".format(ime, GATE_PROC, SAVE_PROC))

        # Abort se trazi BAS U PROZORU izmedju kapije i save-a. Traziti ga bilo
        # gde u proceduri ne bi merilo nista: grana "save nije uspeo" ionako zove
        # AbortSelfUpdate, pa bi provera bila zelena i kad se ishod kapije potpuno
        # ignorise. (Ovo je uhvatio sopstveni self-test, slucaj kapija-bez-aborta.)
        if i_gate >= 0:
            kraj = i_save if i_save > i_gate else len(telo)
            if _prvi_indeks(telo[i_gate + 1:kraj], ABORT_PROC) < 0:
                nalazi.append(
                    "KAPIJA_ABORT  {}: izmedju {} i save-a nema {} -- provera cija "
                    "se neuspesnost ignorise nije kapija"
                    .format(ime, GATE_PROC, ABORT_PROC))

    # Nula pozivalaca nije "cisto" nego provera koja nista ne meri: ako se oba
    # uspesna puta preimenuju ili obrisu, gornja petlja bi tiho prosla.
    if pozivalaca < 2:
        nalazi.append(
            "KAPIJA_SAVE  ocekivana su BAR DVA puta do save-a (soft-only i faza 2), "
            "nadjeno {} -- provera vise ne meri ono zbog cega postoji"
            .format(pozivalaca))
    return nalazi


# --- self-test: dokaz u oba smera -------------------------------------------

def _napisi(put, telo):
    with io.open(put, "w", encoding="ascii", newline="\r\n") as fh:
        fh.write(telo)


CIST = (
    "Private Function VerifyReleaseProject(ByVal folder As String) As String\n"
    "End Function\n"
    "Private Function SaveWorkbookVerified() As Boolean\n"
    "End Function\n"
    "Private Sub AbortSelfUpdate(ByVal msg As String)\n"
    "End Sub\n"
    "Private Sub PutA()\n"
    "    p = VerifyReleaseProject(d)\n"
    "    If Len(p) > 0 Then AbortSelfUpdate p\n"
    "    If Not SaveWorkbookVerified() Then AbortSelfUpdate \"ne\"\n"
    "End Sub\n"
    "Public Sub PutB()\n"
    "    p = VerifyReleaseProject(d)\n"
    "    If Len(p) > 0 Then AbortSelfUpdate p\n"
    "    If Not SaveWorkbookVerified() Then AbortSelfUpdate \"ne\"\n"
    "End Sub\n"
)


def self_test():
    import shutil
    import tempfile

    palo = []
    slucajevi = []

    def slucaj(naziv, izvor, ocekuj_kod):
        """ocekuj_kod=None -> mora biti CISTO; inace nalaz mora poceti tim kodom."""
        slucajevi.append(naziv)
        tmp = tempfile.mkdtemp(prefix="sugates_")
        try:
            put = os.path.join(tmp, MODUL)
            _napisi(put, izvor)
            n = check_gates(put)
            if ocekuj_kod is None:
                if n:
                    palo.append("  {}: ocekivano CISTO, dobijeno {}".format(naziv, n))
            elif not any(x.startswith(ocekuj_kod) for x in n):
                palo.append("  {}: ocekivan nalaz {}, dobijeno {}"
                            .format(naziv, ocekuj_kod, n))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # Sidro: cist modul mora biti ZELEN. Provera koja uvek vristi ne razlikuje
    # ispravno od pokvarenog, pa bi svi ostali slucajevi bili prazan hod.
    slucaj("cist-modul", CIST, None)

    # Kapija uklonjena sa jednog puta.
    slucaj("save-bez-kapije",
           CIST.replace("    p = VerifyReleaseProject(d)\n"
                        "    If Len(p) > 0 Then AbortSelfUpdate p\n"
                        "    If Not SaveWorkbookVerified() Then AbortSelfUpdate \"ne\"\n"
                        "End Sub\n"
                        "Public Sub PutB()\n",
                        "    If Not SaveWorkbookVerified() Then AbortSelfUpdate \"ne\"\n"
                        "End Sub\n"
                        "Public Sub PutB()\n", 1),
           "KAPIJA_SAVE")

    # Kapija POSLE save-a nije kapija.
    slucaj("kapija-posle-save",
           CIST.replace("Private Sub PutA()\n"
                        "    p = VerifyReleaseProject(d)\n"
                        "    If Len(p) > 0 Then AbortSelfUpdate p\n"
                        "    If Not SaveWorkbookVerified() Then AbortSelfUpdate \"ne\"\n",
                        "Private Sub PutA()\n"
                        "    If Not SaveWorkbookVerified() Then AbortSelfUpdate \"ne\"\n"
                        "    p = VerifyReleaseProject(d)\n"
                        "    If Len(p) > 0 Then AbortSelfUpdate p\n", 1),
           "KAPIJA_SAVE")

    # Kapija cija se neuspesnost ignorise.
    slucaj("kapija-bez-aborta",
           CIST.replace("Private Sub PutA()\n"
                        "    p = VerifyReleaseProject(d)\n"
                        "    If Len(p) > 0 Then AbortSelfUpdate p\n",
                        "Private Sub PutA()\n"
                        "    p = VerifyReleaseProject(d)\n", 1),
           "KAPIJA_ABORT")

    # Preimenovana kapija ne sme da ostavi proveru zelenom.
    slucaj("kapija-preimenovana",
           CIST.replace("VerifyReleaseProject", "VerifyReleaseProjectV2"),
           "KAPIJA_POSTOJI")

    # Komentar NIJE poziv: modul u kome je kapija samo pomenuta u komentaru mora
    # da padne. Bez strip_comment-a provera bi ga citala kao ispravan.
    slucaj("kapija-samo-u-komentaru",
           CIST.replace("    p = VerifyReleaseProject(d)\n"
                        "    If Len(p) > 0 Then AbortSelfUpdate p\n",
                        "    ' ovde je nekad stajao VerifyReleaseProject poziv\n", 2),
           "KAPIJA_SAVE")

    # Manje od dva puta do save-a = provera vise nista ne meri.
    slucaj("jedan-put-do-savea",
           CIST.replace("Public Sub PutB()\n"
                        "    p = VerifyReleaseProject(d)\n"
                        "    If Len(p) > 0 Then AbortSelfUpdate p\n"
                        "    If Not SaveWorkbookVerified() Then AbortSelfUpdate \"ne\"\n"
                        "End Sub\n", "", 1),
           "KAPIJA_SAVE")

    # Stvarni src-vba mora biti zelen.
    slucajevi.append("stvarni-modul")
    n = check_gates()
    if n:
        palo.append("  stvarni-modul: {} nije cist -- {}".format(MODUL, n))

    if palo:
        print("\n".join(palo), file=sys.stderr)
        print("\nself-test: {} od {} slucajeva palo."
              .format(len(palo), len(slucajevi)), file=sys.stderr)
        return 2
    print("self-test: cisto ({} slucajeva: {})."
          .format(len(slucajevi), ", ".join(slucajevi)))
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()
    nalazi = check_gates()
    if nalazi:
        for x in nalazi:
            print(x, file=sys.stderr)
        print("\nvba_selfupdate_gates: {} nalaza.".format(len(nalazi)), file=sys.stderr)
        return 2
    print("vba_selfupdate_gates: cisto (svaki put do save-a prolazi kroz {})."
          .format(GATE_PROC))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
