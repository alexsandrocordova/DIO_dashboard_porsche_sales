# -*- coding: utf-8 -*-
"""Extrai os scripts inline do dashboard para checagem de sintaxe."""

import re

html = open("porsche_intelligence_dashboard.html", encoding="utf-8").read()

# Remove <script src="..."></script> externos; pega apenas inline
blocks = re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", html, re.S)
print("blocos inline:", len(blocks))

# Bloco principal = o maior (logica do dashboard)
main = max(blocks, key=len)
open("_inline_check.js", "w", encoding="utf-8").write(main)
print("main chars:", len(main))
