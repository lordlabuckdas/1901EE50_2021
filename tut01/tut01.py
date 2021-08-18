def meraki_helper(n: int) -> bool:
    """This will detect meraki number"""
    # fetch last digit
    prev_dig = n % 10
    n //= 10
    # loop while number is non-zero
    # start from the right end towards left
    while n:
        # fetch last dig
        curr_dig = n % 10
        # if absolute diff is not 1, non-meraki number
        if abs(curr_dig - prev_dig) != 1:
            return False
        # divide by 10 to shift right
        n //= 10
        # set current dig to be previous
        prev_dig = curr_dig
    # code reaches here if all the digits satisfy meraki condition
    return True


input = [12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345, 987654321]

# variable to track count of meraki numbers in input list
meraki_count = 0
# loop through input list checking each number
for num in input:
    assert num >= 0, str(num) + " - Negative integer provided"
    if meraki_helper(num):
        print("Yes -", num, "is a meraki number")
        meraki_count += 1
    else:
        print("No -", num, "is not a meraki number")

print("The input list contains", meraki_count, "meraki numbers and", len(input) - meraki_count, "non-meraki numbers")
