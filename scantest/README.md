## scantest.sh

A Bash script to automate running a variety of Nmap scan types against a target host or IP address, with debug output and flexible options for port ranges, scan types, and timing templates.

### Features
- Supports multiple TCP and UDP scan techniques (e.g., SYN, FIN, NULL, XMAS, UDP, SCTP).
- Optional "idle" (zombie) scan if you have a host with predictable IP-ID sequence (set via `ZOMBIE` environment variable).
- Configurable port range (defaults to 1–1000) or full range (1–65535) when using `all`.
- Flexible scan-type filtering to run a single scan or all supported scans.
- Customizable timing templates (`-T0` through `-T5`), defaulting to `T2`.
- Captures verbose output (`-v`) and logs both stdout and stderr for troubleshooting.

### Prerequisites
- **nmap** must be installed and in your `PATH`.
- Root (or `sudo`) access is recommended for raw socket scans (SYN, UDP, SCTP, idle).

### Usage
```bash
sudo ./scantest.sh <target> [ports] [scan_type] [timing]
```

- `<target>` — IP address or hostname to scan (required).
- `[ports]` — Port range to scan (e.g., `1-1000`, `80`, or `all` for `1-65535`). Default: `1-1000`.
- `[scan_type]` — (Optional) One of the supported scan names (e.g., `TCP_SYN`, `UDP`). If omitted, runs _all_ scan types.
- `[timing]` — (Optional) Nmap timing template (`T0` through `T5`). Default: `T2`.

### Scan Types
| Name             | Flag  |
|------------------|-------|
| TCP_Connect      | `-sT` |
| TCP_SYN          | `-sS` |
| TCP_ACK          | `-sA` |
| TCP_Window       | `-sW` |
| TCP_Maimon       | `-sM` |
| TCP_FIN          | `-sF` |
| TCP_NULL         | `-sN` |
| TCP_Xmas         | `-sX` |
| UDP              | `-sU` |
| SCTP_INIT        | `-sY` |
| SCTP_COOKIE_ECHO | `-sZ` |
| IP_Protocol      | `-sO` |

### Examples
```bash
# Scan default ports (1–1000) with all scan types at T2
sudo ./scantest.sh 10.0.0.5

# Scan all ports (1–65535) with all scan types at T2
sudo ./scantest.sh 10.0.0.5 all

# Scan ports 80–443 with only SYN scan at T4
sudo ./scantest.sh example.com 80-443 TCP_SYN T4

# Scan default ports with only UDP scan at the slowest rate (T0)
sudo ./scantest.sh 192.168.1.1 1-1000 UDP T0
```

### Logs
All output is written into the `scantest/` directory, with filenames:
```
<target>_<scan_name>_T<timing>.log
```

E.g., `scantest/10_0_0_5_TCP_SYN_T2.log`
