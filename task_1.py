
def arab(num: int) -> str:
    """
    Конвертирует арабское число в римскую запись.

    Args:
        num: Целое число для конвертации. Должно быть в диапазоне 1-3999.

    Returns:
        Строку с римской записью числа.
    """
    if num <= 0:
        return "Введите число больше нуля"
    if num > 3999:
        return "Римские цифры обычно не больше 3999"

    thousands_list = ("","M", "MM", "MMM")
    hundreds_list = ("","C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM")
    tens_list = ("","X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC")
    ones_list = ("","I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX")

    thousands = num // 1000
    hundreds = (num % 1000) // 100
    tens = (num % 100) // 10
    ones = num % 10

    return (thousands_list[thousands] + hundreds_list[hundreds]
            + tens_list[tens] + ones_list[ones])



if __name__ == "__main__":
    print(arab(3749))
    print(arab(58))
    print(arab(1994))
