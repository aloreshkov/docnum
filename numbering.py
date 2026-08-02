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

class Numbering:
    """Родительский класс для создания дочерних классов,
     в которых описывается порядок преобразования числа в значение элемента списка.
      Метод num_value должен возвращать строку, в которую преобразуется значение value"""
    numbering_class = 'default'

    def num_value(value:int)->str:
        """Возвращает результат преобразования значения value (int) в строку"""
        pass


class DecimalNumbering(Numbering):
    """Дочерний класс для класса Numbering, реализует преобразование в арабские цифры"""
    numbering_class = 'decimal'

    def num_value(value:int)->str:
        return str(value)


class RomanNumbering(Numbering):
    """Дочерний класс для класса Numbering, реализует преобразование в римские цифры (латиница верхний регистр)"""
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
    """Дочерний класс для класса Numbering, реализует преобразование в римские цифры (латиница нижний регистр)"""
    numbering_class = 'lowerRoman'

    def num_value(value:int)->str:
        return RomanNumbering.num_value(value).lower()

class AnyLetterNumbering:
    """Класс реализует формирование номера по любому списку символов, передаваемому в аргументе letters_lst"""
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
