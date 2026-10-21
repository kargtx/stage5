# -*- coding: utf-8 -*-
"""
Генератор файлов зон DNS для ЛР3 (Вариант 17)
Зона: gr16.ugatu.ru
IP: 10.10.16.1 - 10.10.16.33
Имена: stud161 - stud1633
NS: stud161.gr16.ugatu.ru (10.10.16.1), stud162.gr16.ugatu.ru (10.10.16.2)
Кодировка: UTF-8 БЕЗ BOM, Unix-переводы строк (LF)
"""
import os
import dns.zone

def generate_forward():
    lines = [
        "$TTL 3600",
        "$ORIGIN gr16.ugatu.ru.",
        "@       IN      SOA     stud161.gr16.ugatu.ru. root.gr16.ugatu.ru. (",
        "                        2026102001      ; Serial",
        "                        10800           ; Refresh (3 hours)",
        "                        1800            ; Retry (30 mins)",
        "                        3600000         ; Expire (1000 hours)",
        "                        259200 )        ; Minimum TTL (3 days)",
        "",
        "@       IN      NS      stud161.gr16.ugatu.ru.",
        "@       IN      NS      stud162.gr16.ugatu.ru.",
        ""
    ]

    for i in range(1, 34):  # stud161 ... stud1633
        name = f"stud16{i}"
        ip = f"10.10.16.{i}"
        alias = f"alias16{i}"

        lines.append(f"{name:<15} IN      A       {ip}")
        lines.append(f"{name:<15} IN      HINFO   \"PC-INTEL\" \"LINUX\"")
        lines.append(f"{name:<15} IN      MX      10 stud16{i}.gr16.ugatu.ru.")
        lines.append(f"{name:<15} IN      TXT     \"Workstation 16{i}\"")
        lines.append(f"{alias:<15} IN      CNAME   {name}.gr16.ugatu.ru.")
        lines.append("")

    return "\n".join(lines) + "\n"

def generate_reverse():
    lines = [
        "$TTL 3600",
        "$ORIGIN 16.10.10.in-addr.arpa.",
        "@       IN      SOA     stud161.gr16.ugatu.ru. root.gr16.ugatu.ru. (",
        "                        2026102001      ; Serial",
        "                        10800           ; Refresh (3 hours)",
        "                        1800            ; Retry (30 mins)",
        "                        3600000         ; Expire (1000 hours)",
        "                        259200 )        ; Minimum TTL (3 days)",
        "",
        "@       IN      NS      stud161.gr16.ugatu.ru.",
        "@       IN      NS      stud162.gr16.ugatu.ru.",
        ""
    ]

    for i in range(1, 34):
        lines.append(f"{i:<7} IN      PTR     stud16{i}.gr16.ugatu.ru.")

    return "\n".join(lines) + "\n"

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    fwd_path = os.path.join(base_dir, "gr_17.txt")
    rev_path = os.path.join(base_dir, "gr_b_17.txt")

    fwd_content = generate_forward()
    rev_content = generate_reverse()

    # Запись строго в binary / UTF-8 без BOM
    with open(fwd_path, "wb") as f:
        f.write(fwd_content.encode("utf-8"))
    print(f"Updated {fwd_path} (size={os.path.getsize(fwd_path)} bytes, NO BOM)")

    with open(rev_path, "wb") as f:
        f.write(rev_content.encode("utf-8"))
    print(f"Updated {rev_path} (size={os.path.getsize(rev_path)} bytes, NO BOM)")

    # Валидация через dnspython
    try:
        z_fwd = dns.zone.from_file(fwd_path, origin="gr16.ugatu.ru.")
        print(f"Validation gr_17.txt: OK! Loaded {len(z_fwd.nodes)} nodes.")
    except Exception as e:
        print(f"Validation gr_17.txt FAILED: {e}")

    try:
        z_rev = dns.zone.from_file(rev_path, origin="16.10.10.in-addr.arpa.")
        print(f"Validation gr_b_17.txt: OK! Loaded {len(z_rev.nodes)} nodes.")
    except Exception as e:
        print(f"Validation gr_b_17.txt FAILED: {e}")
