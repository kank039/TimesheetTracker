from core_engine import setup_database, get_all_leaves, add_leave, remove_leave
setup_database()
leaves = get_all_leaves(2026)
print(f"{len(leaves)} leaves in DB for 2026")
for l in leaves[:5]:
    print(f"  {l['date']}  {l['type']:20s}  {l['name']}")
# Test personal leave add/remove
add_leave("2026-07-15", "Personal Leave", "Doctor appointment")
leaves2 = get_all_leaves(2026)
personal = [l for l in leaves2 if l["type"] == "Personal Leave"]
print(f"\nPersonal leaves: {len(personal)}")
for l in personal:
    print(f"  {l['date']}  {l['name']}")
remove_leave("2026-07-15")
print("Remove OK")
# Try removing a company holiday (should fail)
try:
    remove_leave("2026-01-26")
    print("ERROR: should have raised!")
except ValueError as e:
    print(f"Correctly blocked: {e}")