# -*- coding: utf-8 -*-
"""Depuracao do marcador do patch."""
import ast

h = open("porsche_intelligence_dashboard.html", encoding="utf-8").read()
tree = ast.parse(open("patch_dashboard.py", encoding="utf-8").read())

for node in ast.walk(tree):
    if isinstance(node, ast.Call) and getattr(node.func, "id", "") == "replace_once":
        label = node.args[2].value
        old = node.args[0].value
        cnt = h.count(old)
        print(f"--- {label!r}: len(old)={len(old)} count={cnt}")
        if cnt == 0 and "tr.innerHTML" in old:
            print(repr(old))
