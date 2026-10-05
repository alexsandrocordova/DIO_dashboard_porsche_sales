# -*- coding: utf-8 -*-
"""Testes unitarios rapidos das regras de sanitizacao."""
import sys

sys.path.insert(0, ".")
from sanitize_data import (sanitize_city, sanitize_date, sanitize_delivery,
                           sanitize_mileage, sanitize_model, sanitize_payment,
                           sanitize_price, sanitize_state, sanitize_year)

fails = 0


def chk(label, got, exp):
    global fails
    ok = got == exp
    if not ok:
        fails += 1
    print(("OK  " if ok else "FAIL"), label, "->", repr(got),
          "" if ok else ("(esperado " + repr(exp) + ")"))


for i, e in [
    ("2024-13-05", "INVALID"), ("2024-02-30", "INVALID"),
    ("July 15th, 2024", "2024-07-15"), ("02/05/24", "2024-02-05"),
    ("2024.07.11", "2024-07-11"), ("05-27-27", "2027-05-27"),
    ("06/31/2025", "INVALID"), ("2024-04-18 00:00:00", "2024-04-18"),
    ("2025/18/11", "INVALID"), ("2025/04/31", "INVALID"),
]:
    chk("date " + i, sanitize_date(i), e)

for i, e in [
    ("20-24", 2024), ("20 23", 2023), ("twenty twenty four", 2024),
    ("two thousand twenty one", 2021), (2022, 2022), ("1990", 1990),
    ("1985", "INVALID"),
]:
    chk("year " + str(i), sanitize_year(i), e)

for i, e in [
    ("$985,000.00", 985000.0), ("188k USD", 188000.0), ("$645k", 645000.0),
    ("eighty two thousand USD", 82000.0), ("two hundred thousand USD", 200000.0),
    ("$103.750,00", 103750.0), ("104,500 USD", 104500.0),
    ("104.600 USD", 104600.0), ("USD 112.750", 112750.0),
    ("73500", 73500.0), ("USD $146,800", 146800.0),
    ("$153,200.50", 153200.5), ("USD 99.950", 99950.0),
]:
    chk("price " + i, sanitize_price(i), e)

for i, e in [
    ("12,500 miles", 12500), ("KM 15.200", 9445), ("new", 0),
    ("zero miles", 0), ("twelve thousand miles", 12000),
    ("1.100 miles", 1100), ("9.5", 9500), ("42", 42),
    ("Miles: 24,100", 24100), ("19,250 mi.", 19250), ("0 mi", 0),
    ("KM 8,900", 5530), ("fifteen thousand miles", 15000),
]:
    chk("mile " + i, sanitize_mileage(i), e)

for i, e in [
    ("CreditCard", "Credit Card"), ("bank-transfer", "Bank Transfer"),
    ("crypto payment", "Crypto Payment"), ("Bank wire", "Wire Transfer"),
    ("Leasing", "Lease"), ("ACH payment", "ACH Payment"),
    ("CASH payment", "Cash"), ("finance", "Financing"),
    ("debit card", "Debit Card"), ("wire-transfer", "Wire Transfer"),
]:
    chk("pay " + i, sanitize_payment(i), e)

for i, e in [
    ("California", "CA"), ("california", "CA"), ("ca", "CA"),
    ("New York", "NY"), ("la", "LA"), ("minnesota", "MN"), ("XX", "INVALID"),
]:
    chk("state " + i, sanitize_state(i), e)

for i, e in [
    ("delivered!!!", "Delivered"), ("DELIVERD", "Delivered"),
    ("in-transit", "In Transit"), ("IN TRANSIT", "In Transit"),
    ("pending!!", "Pending"), ("awaiting review", "Awaiting Review"),
    ("CANCELLED", "Cancelled"), ("shipped", "Shipped"),
    ("awaiting pickup", "Awaiting Pickup"),
]:
    chk("deliv " + i, sanitize_delivery(i), e)

for i, e in [
    ("911 Carrera Cabriolet", "911 Carrera Cabriolet"),
    ("macan electric", "Macan Electric"), ("911 targa 4s", "911 Targa 4S"),
]:
    chk("model " + i, sanitize_model(i), e)

chk("city", sanitize_city("st. louis"), "St. Louis")

print("TOTAL FAILS:", fails)
sys.exit(1 if fails else 0)
