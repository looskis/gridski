#!/usr/bin/env python3
"""End-to-end smoke test: drives the gridski binary over stdio against a running Excel.

Usage: scripts/smoke.py [path/to/gridski]   (Excel must be open with a writable active sheet)
Writes to Z1000:AA1001 on the active sheet, then restores the previous contents.
"""
import json, subprocess, sys

binary = sys.argv[1] if len(sys.argv) > 1 else "target/debug/gridski"
proc = subprocess.Popen([binary], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
next_id = 0

def send(method, params=None, notify=False):
    global next_id
    msg = {"jsonrpc": "2.0", "method": method}
    if params is not None:
        msg["params"] = params
    if not notify:
        next_id += 1
        msg["id"] = next_id
    proc.stdin.write(json.dumps(msg) + "\n")
    proc.stdin.flush()
    if notify:
        return None
    resp = json.loads(proc.stdout.readline())
    if "error" in resp:
        raise SystemExit(f"{method} failed: {resp['error']}")
    return resp["result"]

def call(name, **args):
    result = send("tools/call", {"name": name, "arguments": args})
    text = result["content"][0]["text"]
    if result.get("isError"):
        return {"__error__": text}
    return json.loads(text)

init = send("initialize", {"protocolVersion": "2025-11-25", "capabilities": {},
                           "clientInfo": {"name": "smoke", "version": "0"}})
print("server:", init["serverInfo"], "protocol:", init["protocolVersion"])
send("notifications/initialized", notify=True)
print("tools:", [t["name"] for t in send("tools/list")["tools"]])

print("\nlist_workbooks:", json.dumps(call("list_workbooks"), indent=1))
print("\nget_selection:", json.dumps(call("get_selection"), indent=1))
print("\nread_range (used range):", json.dumps(call("read_range"), indent=1))

w = call("write_range", start="Z1000", values=[["x", 2], ["=AA1000*3", None]])
print("\nwrite_range:", w)
r = call("read_range", range="Z1000:AA1001")
print("read back:", r)
assert r["values"] == [["x", 2], [6, None]], r
assert r["formulas"] == {"Z1001": "=AA1000*3"}, r
restored = call("write_range", start="Z1000", values=w["previous_formulas"])
print("restored:", call("read_range", range="Z1000:AA1001")["values"])

print("\ntruncation:", json.dumps(call("read_range", range="A1:C3", max_cells=4)))
print("bad sheet:", call("read_range", sheet="Nope"))
print("bad range:", call("read_range", range="not a range"))
print("ragged write:", call("write_range", start="A1", values=[[1], [1, 2]]))
proc.stdin.close()
proc.wait(timeout=5)
print("\nOK")
