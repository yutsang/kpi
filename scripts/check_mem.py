"""Definitive memory diagnosis — why does a *tiny* (1-8 MiB) allocation fail on a box with RAM?

The usual culprit is a 32-bit Python: a 32-bit process can address at most ~2 GB of virtual
memory REGARDLESS of how much physical RAM the machine has. Once the frame + a transient copy
cross that 2 GB ceiling, even a 1 MiB allocation fails — exactly the Galaxy step5 symptom.

This prints Python's bitness + its virtual-address ceiling + physical RAM + commit headroom,
then says plainly what the bottleneck is.

Run on Windows:
  python scripts/check_mem.py
"""
from __future__ import annotations
import ctypes, struct, sys


def _gb(n):
    return f"{n/1024**3:,.2f} GB"


def main():
    bits = struct.calcsize("P") * 8
    is64 = sys.maxsize > 2**32
    print("=== Python ===")
    print(f"  bitness        : {bits}-bit   (sys.maxsize>2^32 = {is64})")
    print(f"  version        : {sys.version.splitlines()[0]}")
    print(f"  executable     : {sys.executable}")

    try:
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                        ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                        ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
        m = MEMORYSTATUSEX(); m.dwLength = ctypes.sizeof(m)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
        print("\n=== Machine (physical / commit) ===")
        print(f"  memory load    : {m.dwMemoryLoad}%")
        print(f"  physical RAM   : total {_gb(m.ullTotalPhys)}   avail {_gb(m.ullAvailPhys)}")
        print(f"  page file      : total {_gb(m.ullTotalPageFile)}   avail {_gb(m.ullAvailPageFile)}")
        print("\n=== THIS PROCESS (virtual address space) ===")
        print(f"  total virtual  : {_gb(m.ullTotalVirtual)}   ← the hard ceiling for THIS process")
        print(f"  avail virtual  : {_gb(m.ullAvailVirtual)}")

        print("\n=== 結論 ===")
        if m.ullTotalVirtual < 3 * 1024**3:
            print("  ⚠⚠ 32-bit Python —— 個 process 最多用到 ~2GB virtual,唔理你部機幾多 RAM。")
            print("     呢個就係 Galaxy 1MB 都 alloc 唔到嘅主因(撞 2GB 天花板,唔關 free RAM 事)。")
            print("     徹底解決:裝 64-bit Python 3.13(同版本),re-pip install,重跑。")
        elif m.ullAvailPhys < 1 * 1024**3 or m.ullAvailPageFile < 1 * 1024**3:
            print("  64-bit Python,但 free physical/pagefile 偏低 → 真係 RAM/commit 緊。")
            print("     關閉其他 app,或加大 Windows pagefile(虛擬記憶體)。")
        else:
            print("  64-bit Python + RAM/commit 充足 → 1MB alloc 唔應該失敗。")
            print("     若仍爆,可能 pagefile 設定 / 公司 policy 限制咗 process working set。截圖畀我睇。")
    except Exception as e:
        print(f"\n(GlobalMemoryStatusEx 失敗: {e} — 應該唔係 Windows?)")
        print("  但 bitness 已經夠判斷:32-bit = 主因,裝 64-bit Python。")


if __name__ == "__main__":
    main()
