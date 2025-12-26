


def unic_tree(nums: list[int]) -> list[list[int]]:
    """
    Находит все уникальные тройки чисел [a, b, c], такие что a + b + c == 0.
    """
    if len(nums) < 3:
        return []
    nums.sort()
    result = []
    len_size = len(nums)

    if nums[0] + nums[1] + nums[2] > 0:
        return []
    if nums[-1] + nums[-2] + nums[-3] < 0:
        return []

    for first_idx in range(len_size - 2):
        first_value = nums[first_idx]

        if first_idx > 0 and first_value == nums[first_idx - 1]:
            continue

        if first_value > 0:
            break

        left_idx = first_idx + 1
        right_idx = len_size - 1

        while left_idx < right_idx:
            left_value = nums[left_idx]
            right_value = nums[right_idx]
            sum_tree = first_value + left_value + right_value

            if sum_tree == 0:
                result.append([first_value, left_value, right_value])

                while left_idx < right_idx and nums[left_idx] == left_value:
                    left_idx += 1

                while left_idx < right_idx and nums[right_idx] == right_value:
                    right_idx -= 1

            elif sum_tree < 0:
                left_idx += 1
            else:
                right_idx -= 1

    return result

if __name__ == "__main__":
    print(unic_tree([]))
    print(unic_tree([-1, 0, 1]))
    print(unic_tree([0, 0, 0, 0]))
    print(unic_tree([-2, 0, 1, 1, 2]))