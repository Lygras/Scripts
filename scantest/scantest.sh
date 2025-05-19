#!/usr/bin/env bash
# scantest.sh — run nmap scans with debug output,
# optional scan-type filtering, and timing templates
# Usage: sudo ./scantest.sh <target> [ports (default: 1-1000 or 'all')] [scan_type] [timing (default: T2 or specify T0–T5)]
# Example: sudo ./scantest.sh 10.0.0.5        # scans 1–1000 at T2
#          sudo ./scantest.sh 10.0.0.5 all    # scans 1–65535 at T2
#          sudo ./scantest.sh 10.0.0.5 80-443 TCP_SYN T4

set -uo pipefail

#–– Usage check ––
if [[ $# -lt 1 ]] || [[ $# -gt 4 ]]; then
  echo "Usage: $0 <target> [ports (default: 1-1000 or 'all')] [scan_type] [timing (default: T2 or specify T0–T5)]"
  exit 1
fi

TARGET="$1"

#–– Port range logic ––
# default to 1-1000; if user passes "all", expand to 1-65535
PORTS="${2:-1-1000}"
if [[ "${2:-}" == "all" ]]; then
  PORTS="1-65535"
fi

LOGDIR="scantest"
mkdir -p "$LOGDIR"

#–– Define all scan types ––
declare -A SCANS=(
  ["TCP_Connect"]="-sT"
  ["TCP_SYN"]="-sS"
  ["TCP_ACK"]="-sA"
  ["TCP_Window"]="-sW"
  ["TCP_Maimon"]="-sM"
  ["TCP_FIN"]="-sF"
  ["TCP_NULL"]="-sN"
  ["TCP_Xmas"]="-sX"
  ["UDP"]="-sU"
  ["SCTP_INIT"]="-sY"
  ["SCTP_COOKIE_ECHO"]="-sZ"
  ["IP_Protocol"]="-sO"
)

#–– Optional scan-type filter ––
if [[ $# -ge 3 ]]; then
  if [[ -z "${SCANS["$3"]+x}" ]]; then
    echo "Invalid scan type: $3"
    exit 1
  fi
  SCANS=( ["$3"]="${SCANS["$3"]}" )
fi

#–– Timing templates ––
# Default to only T2 if not overridden
TIMINGS=(2)
if [[ $# -eq 4 ]]; then
  t="${4#T}"
  if [[ "$t" =~ ^[0-5]$ ]]; then
    TIMINGS=("$t")
  else
    echo "Invalid timing: $4  (must be T0–T5)"
    exit 1
  fi
fi

#–– Prerequisites ––
if ! command -v nmap &>/dev/null; then
  echo "ERROR: nmap not found in PATH."
  exit 1
fi
if [[ $EUID -ne 0 ]]; then
  echo "WARNING: not running as root; some scans (SYN/UDP/Idle) may fail."
fi

#–– Run scans ––
set +e  # don’t exit on first failure

for timing in "${TIMINGS[@]}"; do
  for name in "${!SCANS[@]}"; do
    flags=${SCANS[$name]}
    logfile="${LOGDIR}/${TARGET//./_}_${name}_T${timing}.log"
    echo "[+] $name at -T${timing} on ports $PORTS → $logfile"
    nmap $flags -T${timing} -v -p "$PORTS" "$TARGET" \
      >"$logfile" 2>&1
    echo "    → exit code: $?"
  done
done

#–– Idle (“zombie”) scans ––
if [[ -n "${ZOMBIE-}" ]]; then
  for timing in "${TIMINGS[@]}"; do
    logfile="${LOGDIR}/${TARGET//./_}_Idle_T${timing}.log"
    echo "[+] Idle scan (zombie: $ZOMBIE) at -T${timing} on ports $PORTS → $logfile"
    nmap -sI "$ZOMBIE" -T${timing} -v -p "$PORTS" "$TARGET" \
      >"$logfile" 2>&1
    echo "    → exit code: $?"
  done
else
  echo "[!] ZOMBIE not set; skipping Idle scans."
  echo "    To include Idle: export ZOMBIE=<zombie-ip> and re-run."
fi

echo "✔ All requested scans complete. Logs are in ./$LOGDIR/"
