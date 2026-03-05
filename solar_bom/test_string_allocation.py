from src.utils.string_allocation import compute_allocation_cycle, allocate_strings, format_allocation_summary

# Your example: 3-string trackers, 10 strings per inverter
print("=== Cycle Pattern ===")
cycle = compute_allocation_cycle(3, 10)
for i, pattern in enumerate(cycle):
    print(f"  Inverter {i+1}: {pattern}")

print("\n=== Full Allocation: 30 trackers ===")
result = allocate_strings(3, 10, 30)
print(format_allocation_summary(result, 3))

print("\n=== Edge case: 7 trackers (not evenly divisible) ===")
result2 = allocate_strings(3, 10, 7)
print(format_allocation_summary(result2, 3))