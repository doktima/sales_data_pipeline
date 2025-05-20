from etl.parser import get_apply_months_and_days
# Test the function with your date range
test_start = "20250502"
test_end = "20250603"
result = get_apply_months_and_days(test_start, test_end)
print(f"Result for {test_start} to {test_end}: {result}")
print(f"Number of months: {len(result)}")# Tests for parser functions
