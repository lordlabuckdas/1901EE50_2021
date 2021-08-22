def get_memory_score(nums: list) -> int:
    # initialize empty list to store invalid values
    invalid_values = []
    for value in nums:
        # check if each value in list is of float or string type
        # if int, check if it lies in [0, 10)
        if (
            isinstance(value, float)
            or isinstance(value, str)
            or (isinstance(value, int) and (value < 0 or value >= 10))
        ):
            invalid_values.append(value)
    # print message and exit if list is not empty
    if len(invalid_values):
        print("Please enter a valid input list.")
        print("Invalid inputs detected:", invalid_values)
        exit(1)
    # initialize score to 0
    score = 0
    # initialize user memory as an empty list
    mem = []
    for num in nums:
        # if number is already present, increment score by 1
        if num in mem:
            score += 1
        else:
            # if memory already has 5 nums, delete the first number
            if len(mem) == 5:
                del mem[0]
            # add new number to memory
            mem.append(num)
    # return accumulated score
    return score


input_nums = [3, 4, 5, 3, 2, 1]

print(get_memory_score(input_nums))

# print("1. Score:", get_memory_score([3, 4, 1, 6, 3, 3, 9, 0, 0, 0]))
# print("2. Score:", get_memory_score([1, 2, 2, 2, 2, 3, 1, 1, 8, 2]))
# print("3. Score:", get_memory_score([2, 2, 2, 2, 2, 2, 2, 2, 2]))
# print("4. Score:", get_memory_score([1, 2, 3, 4, 5, 6, 7, 8, 9]))
# input_nums = [7, 5, 8, 6, 3, 5, 9, 7, 9, 7, 5, 6, 4, 1, 7, 4, 6, 5, 8, 9, 4, 8, 3, 0, 3]
# print("5. Score:", get_memory_score(input_nums))
