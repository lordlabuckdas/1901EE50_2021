def meraki_helper(n: int) -> bool:
    """This will detect meraki number"""
    prev_dig = n % 10
    n //= 10
    while n:
        curr_dig = n % 10
        if abs(curr_dig - prev_dig) != 1:
            return False
        n //= 10
    return True


input = [12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345, 987654321]

meraki_count = 0
for num in input:
    if meraki_helper(num):
        print("Yes,", num, "is a meraki number")
        meraki_count += 1
    else:
        print("No,", num, "is not a meraki number")

print("The input list contains", meraki_count, "meraki numbers and", len(input) - meraki_count, "non-meraki numbers")
