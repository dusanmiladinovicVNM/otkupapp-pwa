"""Wall-clock po fazi vba_check.main, bez cProfile rezije.
Upotreba: python3 stage_timer.py <tools-dir> [argumenti vba_check-a]
Ne pise nista u stablo (PYTHONDONTWRITEBYTECODE spolja)."""
import importlib.util, sys, time, os, io, contextlib
from collections import defaultdict

tools = sys.argv[1]
argv = sys.argv[2:]
sys.path.insert(0, tools)
spec = importlib.util.spec_from_file_location("vba_check_t", os.path.join(tools, "vba_check.py"))
t_imp = time.perf_counter()
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
t_imp = time.perf_counter() - t_imp

T = defaultdict(float); N = defaultdict(int)
def wrap(mod, name, label=None):
    f = getattr(mod, name)
    label = label or name
    def w(*a, **k):
        t = time.perf_counter()
        try:
            return f(*a, **k)
        finally:
            T[label] += time.perf_counter() - t; N[label] += 1
    setattr(mod, name, w)

TOP = ["check_katalog_sabotaza", "check_dokaz_grupe", "check_popis_suita",
       "collect_definitions", "collect_module_members", "collect_arities",
       "check_file", "check_poruke", "check_rollback_tvrdnja", "check_rollback_red",
       "check_storno_registar", "check_sema_registar", "check_amb_bind_vlasnik",
       "check_clan_forme"]
SUB = ["check_ascii", "check_eol", "check_decl_after_proc", "check_reserved",
       "check_undefined", "check_local_dupes", "check_dupli_lokal", "check_scalar_call",
       "check_dead_log", "check_truncated", "check_array_copy", "check_test_registry",
       "check_storno_progutan", "check_nedeklarisan", "check_proc_size", "collect_public"]
for n in TOP: wrap(m, n)
for n in SUB: wrap(m, n, "  file:" + n)

t0 = time.perf_counter()
out, err = io.StringIO(), io.StringIO()
with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
    rc = m.main(argv)
total = time.perf_counter() - t0
print(f"rc={rc} main={total:.2f}s import_vba_check={t_imp:.2f}s  argv={argv}")
top_sum = 0
for n in TOP:
    print(f"{T[n]:7.2f}s {100*T[n]/total:5.1f}%  x{N[n]:<4} {n}")
    top_sum += T[n]
print(f"{total-top_sum:7.2f}s {100*(total-top_sum)/total:5.1f}%        (ostalo u main)")
for n in SUB:
    l = "  file:" + n
    print(f"{T[l]:7.2f}s {100*T[l]/total:5.1f}%  x{N[l]:<4} {l}")
