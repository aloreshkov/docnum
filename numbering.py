# В модуле создаются классы, описывающие способ получения значения элемента списка по числовому значению.
#
# Экземпляры классов не создаются. От родительского класса Numbering создаются дочерние в которых происходит
# преобразование числового значения в требуемое строковое. Соответствие между типом элементов списка и именем
# класса задается в свойстве класса numbering_class.
#
# Для получения значений элементов списка в нижнем регистре, применяется вызов метода num_value
# класса, возвращающего требуемое значение в верхнем регистре, с последующим вызовом для результата метода lower().
#
# Класс AnyLetterNumbering реализует получение значения для элемента списка по любому перечню символов.
# От класса AnyLetterNumbering дочерние классы не создаются.

# This module defines classes that describe how to retrieve a list item value from a numeric value.
#
# Class instances are not created directly. Instead, child classes are derived from the base `Numbering` class,
# where the conversion from a numeric value to the required string representation is implemented.
# The mapping between the list item type and the corresponding class name is defined in the class property `numbering_class`.
#
# To obtain list item values in lowercase, call the `num_value` method of the class which returns the value in uppercase
# and then apply the `lower()` method to the result.
#
# The `AnyLetterNumbering` class implements value retrieval for a list item based on any set of characters.
# No child classes are created from `AnyLetterNumbering`.


class Numbering:
    """Родительский класс для создания дочерних классов,
    в которых описывается порядок преобразования числа в значение элемента списка.
    Метод num_value должен возвращать строку, в которую преобразуется значение value.

    Base class for creating child classes that define the logic for converting a numeric value
    into the corresponding list item value.

    The `num_value` method must return a string representing the converted value for the given `value`.
    """
    numbering_class = 'default'

    def num_value(value:int)->str:
        """Возвращает результат преобразования значения value (int) в строку"""
        pass


class DecimalNumbering(Numbering):
    """Дочерний класс для класса `Numbering`, реализует преобразование в арабские цифры.

    Child class of `Numbering` that implements conversion to Arabic numerals.
    """
    numbering_class = 'decimal'

    def num_value(value:int)->str:
        return str(value)


class RomanNumbering(Numbering):
    """Дочерний класс для класса Numbering, реализует преобразование в римские цифры (латиница верхний регистр).

    Child class of `Numbering` that implements conversion to Roman numerals (Latin uppercase letters).
    """
    numbering_class = 'upperRoman'

    def num_value(value:int)->str:
        roman_numbers = {'M': 1000, 'CM': 900, 'D': 500, 'CD': 400,
                         'C': 100, 'XC': 90, 'L': 50, 'XL': 40,
                         'X': 10, 'IX': 9, 'V': 5, 'IV': 4, 'I': 1}

        roman = ''
        remains = value
        for letter, val in roman_numbers.items():
            while remains >= val:
                roman += letter
                remains -= val
        return roman


class RomanNumberingLower(Numbering):
    """Дочерний класс для класса Numbering, реализует преобразование в римские цифры (латиница нижний регистр).

    Child class of `Numbering` that implements conversion to Roman numerals (Latin lowercase letters).
    """
    numbering_class = 'lowerRoman'

    def num_value(value:int)->str:
        return RomanNumbering.num_value(value).lower()

class AnyLetterNumbering:
    """Класс реализует формирование номера по любому списку символов, передаваемому в аргументе letters_lst.

    Class that generates a number based on any list of characters passed in the ``letters_lst`` argument.
    """
    def num_value(value:int,letters_lst:list)->str:
        # Получаем частное от целочисленного деления на длину списка символов и остаток
        quotient, remainder = divmod(value, len(letters_lst))
        # Определяем основной символ: элемент списка в позиции остаток-1,
        # для нулевого остатка будет -1, т.е. последний элемент списка
        main_chr = letters_lst[remainder-1]
        # Если остаток от деления равен нулю, то количество букв равно частному,
        # в ином случае, частому + 1
        result = main_chr * (quotient + (1 if remainder > 0 else 0))
        return result


class LetterNumbering(Numbering):
    """Дочерний класс для класса `Numbering`, реализует значение в виде букв латинского алфавита (латиница верхний регистр)

    Child class of `Numbering` that implements value generation as Latin uppercase letters.
    """
    numbering_class = 'upperLetter'

    def num_value(value:int)->str:
        letters_lst = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
        return AnyLetterNumbering.num_value(value, letters_lst)


class LetterNumberingLower(Numbering):
    """Дочерний класс для класса `Numbering`, реализует значение в виде букв латинского алфавита (латиница нижний регистр).

    Child class of `Numbering` that implements value generation as Latin lowercase letters.
    """
    numbering_class = 'lowerLetter'

    def num_value(value:int)->str:
        return LetterNumbering.num_value(value).lower()


class RussianNumbering(Numbering):
    """Дочерний класс для класса `Numbering`, реализует значение в виде букв русского алфавита (кириллица верхний регистр).

    Child class of `Numbering` that implements value generation as Cyrillic uppercase letters.
    """
    numbering_class = 'russianUpper'

    def num_value(value:int)->str:
        letters_lst = list('АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЫЭЮЯ')
        return AnyLetterNumbering.num_value(value, letters_lst)


class RussianNumberingLower(Numbering):
    """Дочерний класс для класса `Numbering`, реализует значение в виде букв русского алфавита (кириллица нижний регистр).

    Child class of `Numbering` that implements value generation as Cyrillic lowercase letters.
    """
    numbering_class = 'russianLower'

    def num_value(value:int)->str:
        return RussianNumbering.num_value(value).lower()
        # Если остаток от деления равен нулю, то количество букв равно частному,
        # в ином случае, частому + 1
        result = main_chr * (quotient + (1 if remainder > 0 else 0))
        return result


class LetterNumbering(Numbering):
    """Дочерний класс для класса Numbering, реализует значение в виде букв латинского алфавита (латиница верхний регистр)"""
    numbering_class = 'upperLetter'

    def num_value(value:int)->str:
        letters_lst = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
        return AnyLetterNumbering.num_value(value, letters_lst)


class LetterNumberingLower(Numbering):
    """Дочерний класс для класса Numbering, реализует значение в виде букв латинского алфавита (латиница нижний регистр)"""
    numbering_class = 'lowerLetter'

    def num_value(value:int)->str:
        return LetterNumbering.num_value(value).lower()


class RussianNumbering(Numbering):
    """Дочерний класс для класса Numbering, реализует значение в виде букв русского алфавита (кириллица верхний регистр)"""
    numbering_class = 'russianUpper'

    def num_value(value:int)->str:
        letters_lst = list('АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЫЭЮЯ')
        return AnyLetterNumbering.num_value(value, letters_lst)


class RussianNumberingLower(Numbering):
    """Дочерний класс для класса Numbering, реализует значение в виде букв русского алфавита (кириллица нижний регистр)"""
    numbering_class = 'russianLower'

    def num_value(value:int)->str:
        return RussianNumbering.num_value(value).lower()
