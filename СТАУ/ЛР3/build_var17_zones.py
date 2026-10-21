# -*- coding: utf-8 -*-
import os
import dns.zone

# Generate forward zone gr_17.txt exactly following gr_24.txt
fwd_lines = [
    '$TTL 3600',
    '@\tIN\tSOA\tns1.gr16.ugatu.ru. admin.gr16.ugatu.ru. (',
    '\t\t\t2026102001 ; Serial',
    '\t\t\t10800      ; Refresh',
    '\t\t\t1800       ; Retry',
    '\t\t\t3600000    ; Expire',
    '\t\t\t259200 )   ; Minimum',
    '',
    '; Name Servers',
    '@\tIN\tNS\tns1.gr16.ugatu.ru.',
    '@\tIN\tNS\tns2.gr16.ugatu.ru.',
    '',
    'ns1\tIN\tA\t10.10.16.1',
    'ns2\tIN\tA\t10.10.16.2',
    '',
    '; Hosts (10.10.16.1 - 10.10.16.32)'
]

for i in range(1, 33):
    name = f'stud16{i}'
    ip = f'10.10.16.{i}'
    cname = f'www-stud16{i}'
    fwd_lines.append(f'{name}\tIN\tA\t{ip}')
    fwd_lines.append(f'{name}\tIN\tHINFO\t"INTEL-CORE" "UBUNTU-LINUX"')
    fwd_lines.append(f'{name}\tIN\tMX 10\t{name}.gr16.ugatu.ru.')
    fwd_lines.append(f'{name}\tIN\tTXT\t"Student 16{i} Workstation"')
    fwd_lines.append(f'{cname}\tIN\tCNAME\t{name}')
    fwd_lines.append('')

fwd_content = '\n'.join(fwd_lines).rstrip('\n') + '\n'

# Generate reverse zone gr_b_17.txt exactly following gr_b_24.txt
rev_lines = [
    '$TTL 3600',
    '@\tIN\tSOA\tns1.gr16.ugatu.ru. admin.gr16.ugatu.ru. (',
    '\t\t\t2026102001 ; Serial',
    '\t\t\t10800      ; Refresh',
    '\t\t\t1800       ; Retry',
    '\t\t\t3600000    ; Expire',
    '\t\t\t259200 )   ; Minimum',
    '',
    '; Name Servers',
    '@\tIN\tNS\tns1.gr16.ugatu.ru.',
    '@\tIN\tNS\tns2.gr16.ugatu.ru.',
    '',
    '; Name Server PTRs',
    '1\tIN\tPTR\tns1.gr16.ugatu.ru.',
    '2\tIN\tPTR\tns2.gr16.ugatu.ru.',
    '',
    '; Host PTRs (10.10.16.1 - 10.10.16.32)'
]

for i in range(1, 33):
    name = f'stud16{i}'
    rev_lines.append(f'{i}\tIN\tPTR\t{name}.gr16.ugatu.ru.')

rev_content = '\n'.join(rev_lines).rstrip('\n') + '\n'

# Destinations
base_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5"
destinations = [
    (os.path.join(base_dir, "СТАУ", "gr_17.txt"), os.path.join(base_dir, "СТАУ", "gr_b_17.txt")),
    (os.path.join(base_dir, "СТАУ", "ЛР3", "gr_17.txt"), os.path.join(base_dir, "СТАУ", "ЛР3", "gr_b_17.txt"))
]

for f_fwd, f_rev in destinations:
    with open(f_fwd, 'wb') as f:
        f.write(fwd_content.encode('utf-8'))
    with open(f_rev, 'wb') as f:
        f.write(rev_content.encode('utf-8'))
    print(f'Wrote {f_fwd} ({os.path.getsize(f_fwd)} bytes)')
    print(f'Wrote {f_rev} ({os.path.getsize(f_rev)} bytes)')

# Validation
z_fwd = dns.zone.from_text(fwd_content, 'gr16.ugatu.ru.')
print('Validation gr_17.txt: OK! Nodes count =', len(z_fwd.nodes))

z_rev1 = dns.zone.from_text(rev_content, '16.10.10.in-addr.arpa.')
print('Validation gr_b_17.txt (reverse in-addr.arpa): OK! Nodes count =', len(z_rev1.nodes))

z_rev2 = dns.zone.from_text(rev_content, 'gr16.ugatu.ru.')
print('Validation gr_b_17.txt (gr16.ugatu.ru): OK! Nodes count =', len(z_rev2.nodes))
