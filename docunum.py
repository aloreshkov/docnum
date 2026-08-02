# Модуль реализует класс DocuNum, который вносит в свойство paragraph._element.pPr.numPr.text документа word
# значения элемента списка. Документ word передается экземпляру класса DocuNum в виде объекта skelmis.docx.Document.
# Пример использования
# from skelmis.docx import Document
# from docnum import docunum
#
# document = Document('имя_файла.docx')
# docunum = docunum.DocuNum(document)
#
# for paragraph in document.paragraphs:
#     if paragraph._element.pPr is not None and paragraph._element.pPr.numPr is not None:
#         print(paragraph._element.pPr.numPr.text)

# The module implements the DocuNum class, which sets the value of the list item element
# into the paragraph._element.pPr.numPr.text property of a Word document.
# The Word document is passed to an instance of the DocuNum class as a skelmis.docx.Document object.
# Example of usage:
# from skelmis.docx import Document
# from docnum import docunum
#
# document = Document('имя_файла.docx')
# docunum = docunum.DocuNum(document)
#
# for paragraph in document.paragraphs:
#     if paragraph._element.pPr is not None and paragraph._element.pPr.numPr is not None:
#         print(paragraph._element.pPr.numPr.text)

import docnum.numbering as numbering
from skelmis.docx import Document
from skelmis.docx.oxml import parse_xml
from skelmis.docx.oxml.xmlchemy import BaseOxmlElement
import re
from lxml import etree as ET


def get_attribute_uri_mame(nsmap:dict, value:str)->str:
    """Функция преобразует имя аттрибута из сокращенного описания в имя URI вида {namespace}attribute

    nsmap - словарь с префиксами (ключ) областей имен (значение)

    value - преобразуемое имя атрибута

    Convert an attribute name from a short description to a URI-style name of the form {namespace}attribute.

    Parameters
    ----------
    nsmap : dict
        Dictionary mapping prefixes (keys) to namespace URIs (values).
    value : str
        The attribute name to convert.

    Returns
    -------
    str
        The converted attribute name in the format {namespace}attribute.
    """
    pref,tag = value.split(':')
    return "{"+nsmap[pref]+"}"+tag


def get_numbering_classes()->dict:
    """Функция создает словарь описывающий соответствие между типами нумерации в документе и классами,
    которые их реализуют. в модуле ``docnum.numbering``

    Create a dictionary that maps numbering types in a Word document to their
    implementing classes in ``docnum.numbering``.

    Returns:
        dict: Mapping from numbering type IDs (e.g., abstractNumId, style name)
              to corresponding handler classes."""
    numbering_classes_dict = {}
    # Получаем список классов дочерних от класса Numbering
    subclasses = numbering.Numbering.__subclasses__()
    for subclass in subclasses:
        # Проверяем наличие у класса атрибута numbering_class
        if hasattr(subclass, 'numbering_class'):
            numbering_class = subclass.numbering_class
            # Добавляем соответствие в словарь
            numbering_classes_dict[numbering_class] = subclass
    return numbering_classes_dict


numbering_classes_dict = get_numbering_classes()


class Lvl:
    """Класс содержит описание уровня нумерации, полученное из файла numbering.xml.

    Represents a single numbering level definition extracted from ``numbering.xml``
    in a Word document (Office Open XML format)."""

    def __init__(self,src,ilvl:int,start:int=1,numFmt:str='',lvlText:str=''):
        """src - исходный объект, описывающий уровень
    ilvl - номер уровня
    start - стартовое значение для уровня
    numFmt - строка, указывающая тип нумерации (ключ для словаря numbering_classes_dict)
    lvlText - строка, описывающая формат значения номера для текущего уровня
    current - текущее значение номера для уровня
    lastCallLvl - номер уровня нумерации, с которого в последний раз был вызов за значением данного уровня нумерации.
    Если обращений не было, то None

    Initialize a numbering level handler for Word document processing.

    Parameters
    ----------
    src : object
        Source object describing the numbering level.
    ilvl : int
        Level index (0-based) of the numbering hierarchy.
    start : int, optional
        Starting value for the level counter (default is 1).
    numFmt : str, optional
        Numbering format string (default is '').
    lvlText : str, optional
        Format template for the displayed number at this level (e.g. '%1.', '(%1)'),
        used when rendering the number text (default is '').

    Attributes
    ----------
    current : int
        Current counter value for this level, initialized to `start`.
    lastCallLvl : int | None
        The level index from which the current level was last requested.
        `None` if no such request has occurred yet.
"""
        self.src = src
        self.ilvl = ilvl
        self.start = start
        self.numFmt = numFmt
        self.lvlText = lvlText
        self.current = self.start
        self.lastCallLvl = None

    def __str__(self):
        """Метод реализован только для отладки

        This method is implemented solely for debugging and should not be used
        in production code."""
        s = (f"ilvl={self.ilvl}\n"
             f"start={self.start}\n"
             f"numFmt={self.numFmt}\n"
             f"lvlText={self.lvlText}\n")
        return s

    def get_value(self, call_lvl:int)->str:
        """Метод возвращает текущее значение для уровня.
         call_lvl - номер уровня, с которого происходит вызов.

        Return the current numbering value for this level, optionally incrementing
        the counter based on the caller's level.

        This method handles two main cases:

        - If the numbering type is a bullet (``self.numFmt == 'bullet'``), it returns the static bullet text from ``self.lvlText`` without modifying state.
        - For numeric formats, it increments the internal counter (``self.current``)
        only when called from the same level as this one, and only if this is a
        new logical step in the list traversal (determined via ``self.lastCallLvl``).

        After computing the value, the method records ``call_lvl`` as the last caller
        in ``self.lastCallLvl`` to support correct increment behavior on subsequent calls.

        Parameters
        ----------
        call_lvl : int
          The level index from which this method is invoked. Used to determine
          whether the counter should be incremented and to track traversal state.

        Returns
        -------
        str
            The rendered numbering string for the current state of this level
            (e.g., "1.", "1.2.3", or a bullet symbol).
"""
        # Если тип нумерации символ, то сразу возвращаем его значение
        if self.numFmt == 'bullet':
            return self.lvlText
        else:
            # Если тип нумерации не символ, то проверяем, с какого уровня был последний вызов
            if self.lastCallLvl is not None:
            # Если сейчас вызов идет с этого же уровня
            # и предыдущий вызов был с другого уровня или предыдущий вызов был с того уже уровня, что и текущий вызов
                if call_lvl == self.ilvl and (self.lastCallLvl != self.ilvl or self.lastCallLvl == call_lvl):
                    # увеличиваем текущий номер
                    self.increment_value()
        # Получаем строку со значением номера
        value = self.get_value_by_type(self.current, self.numFmt)
        # Запоминаем номер уровня с которого идет текущий вызов
        self.lastCallLvl = call_lvl
        return value

    def increment_value(self)->None:
        """Метод увеличивает текущее значение для уровня нумерации,
        исключая нумерацию с помощью символов.

        Increment the current counter value for this numbering level, except bullet."""
        if self.numFmt != 'bullet':
            self.current += 1

    def get_value_by_type(self, value:int, nfmt:str)->str:
        """Метод возвращает строку со значением нумерации по типу.
        Если в словаре ``numbering_classes_dict`` нет соответствующего формату номера,
        используется десятичная нумерация

        value - значение
        nfmt - тип нумерации

        Render a numbering string for a given counter value and format type.

        The method looks in the global `numbering_classes_dict`` dictionary.
        If no matching key is found, it falls back to the 'decimal'.

        Parameters
        ----------
        value : int
            The numeric counter value to format (e.g., 1, 2, 3).
        nfmt : str
            The numbering format identifier, typically taken from the ``<w:numFmt>``
            element in ``numbering.xml`` (e.g. 'lowerLetter', 'upperRoman').

        Returns
        -------
        str
            The formatted numbering string (e.g. "1.", "a)", "(i)").
        """
        if nfmt not in numbering_classes_dict:
            nfmt = 'decimal'
        return numbering_classes_dict[nfmt].num_value(value)


class AbstructNum:
    """Класс содержит описание набора уровней списка, определенного в документе,
    полученного из одного тэга ``w:abstructNumbering`` файла ``numbering.xml``.

     Represents a single abstract numbering style defined in a Word document’s
     ``numbering.xml``, corresponding to one ``<w:abstractNum>`` element."""
    def __init__(self,src,abstructNumId:int,restartNumberingAfterBreak:int,multiLevelType:str=''):
        """src - исходный объект, описывающий набор уровней,
        abstructNumId - номер набора,
        restartNumberingAfterBreak - свойство набора, из файла,
        multiLevelType - свойство набора, из файла,
        levels - список объектов ``Lvl``

        Initialize a numbering set definition.

        Parameters
        ----------
        src : object
            The source XML element describing the levels configuration.
        abstract_num_id : int
            The ID of the abstract numbering style (abstractNumId).
        restart_numbering_after_break : bool
            Whether numbering should restart after a break (value from the file).
        multi_level_type : str
            The multi-level type property of the set (value from the file).
        levels : list of ``Lvl`` objects
            List of level definitions (Lvl objects). Defaults to an empty list.
        """
        self.abstructNumId = abstructNumId
        self.restartNumberingAfterBreak = restartNumberingAfterBreak
        self.multiLevelType = multiLevelType
        self.levels = []
        self.src = src

    def add_level(self, level:Lvl):
        """Метод добавляет в свойство levels объект ``Lvl`` c описанием уровня нумерации.
        level - объект ``Lvl``
        Add or update a numbering level definition

        Parameters
        ----------
        level : ``Lvl`` The level object """
        # Если объект с таким уровнем нумерации уже есть, то замещаем имеющийся
        if level.ilvl < len(self.levels):
            self.levels[level.ilvl] = level
        else:
            # Если объекта с таким уровнем нет, то создаем новый элемент списка в позиции, с соответствующим индексом,
            # в котором располагаем объект с описанием уровня нумерации
            self.levels.insert(level.ilvl,level)

    def __str__(self):
        """Метод реализован только для отладки

        This method is implemented solely for debugging and should not be used
        in production code."""
        s = (f"abstructNumId={self.abstructNumId}\n"
             f"restartNumberingAfterBreak={self.restartNumberingAfterBreak}\n"
             f"multiLevelType={self.multiLevelType}\n")
        for level in self.levels:
            s += f"\t{str(level)}\n"
        return s


class Num:
    """Класс содержит описание набора уровней списка, используемого в документе,
    полученного из одного тэга ``w:num`` файла ``numbering.xml``.

     Represents a single numbered list instance in a Word document,
     corresponding to one ``<w:num>`` element in ``numbering.xml``."""
    def __init__(self, src, numId:int, abstructNum:AbstructNum):
        """src - исходный объект, описывающий набор уровней,
        numId - номер используемого уровня,
        abstructNum - объект типа ``AbstructNum``, на который ссылается текущий набор,
        levels - список объектов ``Lvl``, описывающих уровни набора.

        Initialize a numbered list instance ``w:num`` from numbering.xml.

        Parameters
        ----------
        src : object
            The source XML element describing the levels configuration.
        num_id : int
            The numeric ID of this numbering instance (used to match paragraphs).
        abstract_num : AbstractNum
            The base abstract numbering style (<w:abstractNum>) that this
            numbering instance references.
        levels : list[Lvl]
            List of level definitions
        """
        self.numId = numId
        self.abstructNum = abstructNum
        self.levels = []
        self.src = src

    def add_level(self, level:Lvl):
        """Метод добавляет в свойство levels объект Lvl c описанием уровня нумерации по индексу, равному его номеру - ``level.ilvl``\n
       level - объект Lvl

       Add a numbering level in the ``levels`` list at index ``level.ilvl``.

       Parameters
        ----------
        level : Lvl The level object
       """
        # Если объект с таким уровнем нумерации уже есть, то замещаем имеющийся
        if level.ilvl < len(self.levels):
            self.levels[level.ilvl] = level
        else:
            # Если объекта с таким уровнем нет, то создаем новый элемент списка в позиции, с соответствующим индексом,
            # в котором располагаем объект с описанием уровня нумерации
            self.levels.insert(level.ilvl,level)

    def __str__(self):
        """Метод реализован только для отладки

        This method is implemented solely for debugging and should not be used
        in production code."""
        s = (f"numId={self.numId}\n"
             f"abstructNumId={self.abstructNum}\n")
        for level in self.levels:
            s += f"\t{str(level)}\n"
        return s


class Numberings:
    """Класс содержит список описанных в документе тэгов ``w:abstractNum``.

    The class contains a list of the ``w:abstractNum`` tags described in the document."""
    def __init__(self):
        self.abstructNums = []

    def add_num(self, an:AbstructNum):
        """Метод добавляет объект ``AbstructNum`` в список abstructNums объекта
        по индексу, равному его номеру - abstructNumId

        Add an AbstractNum object to the abstructNums list at the index equal to its abstractNumId.

        Parameters
        ----------
        abstract_num : AbstractNum The AbstractNum object to add.
        """
        if an.abstructNumId < len(self.abstructNums):
            self.abstructNums[an.abstructNumId] = an
        else:
            self.abstructNums.insert(an.abstructNumId,an)

    # def get_levels(self, abstractNumId:int)->list:
    #     """Метод возвращает список объектов Lvl, содержащих описание уровней списка"""
    #     return self.abstructNums[abstractNumId].levels

    def get_abstract_num(self, abstractNumId:int)->AbstructNum:
        """Метод возвращает объект AbstructNum по его номеру.

        Return the AbstractNum object corresponding to the given ID."""
        return self.abstructNums[abstractNumId]


class Numbs:
    """Класс содержит список используемых в документе тэгов ``w:num``

    he class contains a list of the ``w:num`` tags described in the document.
    """
    def __init__(self):
        self.nums = []

    def add_num(self, n:Num):
        """Метод добавляет объект Num в список nums объекта по индексу, равному его номеру - numId

        Add a Num object to the nums list at the index equal to its num_id.
        """
        if n.numId < len(self.nums):
            self.nums[n.numId] = n
        else:
            self.nums.insert(n.numId,n)


class NumPr:
    """Класс содержит описание параграфа, содержащего элемент списка

    Representation of a paragraph containing a list item in a Word document.
    """
    def __init__(self,src,numId:int,ilvl:int,paraId:str):
        """src - исходный объект, описывающий параграф,
        numId - номер используемого набора уровней,
        ilvl - номер уровня в используемом наборе,
        paraId - идентификатор параграфа - атрибут w14:paraId.

        Initialize a paragraph that participates in a numbered/bulleted list.

        Parameters
        ----------
        src : object
            The source XML element (<w:p>) describing the paragraph.
        num_id : int
            The ID of the numbering instance (<w:num>) applied to this paragraph.
            If the paragraph is not part of a list, this value is None.
        ilvl : int
            The indentation/list level index (``ilvl`` attribute). Represents the
            level within the numbering style. If not part of a list, this is None.
        para_id : str
            The paragraph identifier (``w14:paraId`` attribute).
        """
        self.numId = numId
        self.ilvl = ilvl
        self.paraId = paraId
        self.numberP = None
        self.src = src


class DocuNum:
    """Класс содержит результирующие значения всех элементов списков по параграфам документа.
    document - объект с документом word типа skelmis.docx.Document.

    This class aggregates the final computed value for every paragraph if document
    that participates in a numbered or
    bulleted list.

    Attributes
    ----------
    document : skelmis.docx.Document
        The Word document object.
    """

    def __init__(self, document: Document)-> None:
        """document - объект с документом word типа skelmis.docx.Document

        Parameters
        ----------
        document : skelmis.docx.Document
            The Word document object.
        """
        self.document = document
        # Получаем список описанных в документе наборов списков
        self.abstractNumbers = self.get_abstract_numbers()
        # Получаем список используемых в документе наборов списков
        self.numbers = self.get_numbering()
        # Получаем список и словарь с объектами NumPr, описывающими параграфы, в которых используются списки, по всему документу
        self.numPr_lst,self.numPr_dict = self.get_paragraphs()
        # Создаем номера для всех параграфов, в которых используются списки
        self.fill_numbers()
        self.set_num()

    def get_abstract_numbers(self)->Numberings:
        """Метод возвращает объект Numberings, в котором содержится список всех описанных в документе наборов уровней.

        Returns a Numberings object containing the list of all level sets
        described in the document.
        """
        doc_numbering = Numberings()
        # Получаем список наборов уровней, описанных в документе
        num_def = self.document.part.numbering_part.numbering_definitions._numbering.abstractNum_lst
        for nums in num_def:
            # Для каждого набора создаем объект
            abst_num_obj = AbstructNum(nums,
                                     nums.abstractNumId,
                                     int(nums.attrib[get_attribute_uri_mame(nums.nsmap, 'w15:restartNumberingAfterBreak')]),
                                     nums.multiLevelType_lst[0].val)
            # Перебираем все описанные в наборе уровни списка
            for levels in nums.lvl_lst:
                # Для каждого уровня создаем отдельный объект
                level_obj = Lvl(levels,
                               levels.ilvl,
                               levels.start_lst[0].val,
                               levels.numFmt_lst[0].val,
                               levels.lvlText_lst[0].val)
                # Добавляем объект с описанием уровня в набор
                abst_num_obj.add_level(level_obj)
            # Добавляем набор уровней в результирующий объект
            doc_numbering.add_num(abst_num_obj)
        return doc_numbering

    def get_numbering(self)->Numbs:
        """Метод возвращает объект Numbs, в котором содержится список всех используемых в документе наборов уровней.

        Returns a Numbs object containing the list of all level sets used in the document.
        """
        
        doc_numbs = Numbs()
        # Получаем список наборов уровней, используемых в документе
        num_def = self.document.part.numbering_part.numbering_definitions._numbering.num_lst
        for nums in num_def:
            # Для каждого набора создаем объект, который ссылается на ранее созданный объект с описанным в документе набором уровней
            num_obj = Num(nums,
                         nums.numId,
                         self.abstractNumbers.get_abstract_num(nums.abstractNumId.val))
            # Копируем список уровней используемого набора из описанного в документе набора
            num_obj.levels = num_obj.abstructNum.levels[:]
            # Если в текущем наборе уровней есть переопределенные стартовые номера, то заменяем на них
            if len(nums.lvlOverride_lst) > 0:
                for level in nums.lvlOverride_lst:
                    num_obj.levels[level.ilvl].start = level.startOverride.val
            # Добавляем набор уровней в результирующий объект
            doc_numbs.add_num(num_obj)
        return doc_numbs

    def get_paragraphs(self)-> tuple[list, dict]:
        """Метод возвращет все параграфы документа, в которых есть элементы списков, кортежем
        первым элементом которого является список, вторым словарь.
        В списке параграфы в виде объектов размещены в порядке их появления.
        В словаре ключом является атрибут ``w14:paraId`` параграфа, значением - объект параграф.

        Returns all paragraphs in the document that contain list elements as a tuple.

        The first element of the tuple is a list containing the paragraph objects
        in the order they appear in the document.

        The second element is a dictionary where:
            - the key is the paragraph's ``w14:paraId`` attribute,
            - the value is the corresponding paragraph object.
        """

        numPr_lst = []
        numPr_dict = {}
        # Создаем объект для парсинга документа в виде xml
        doc_xml = ET.fromstring(self.document.element.xml)
        # Получаем список из всех параграфов документа
        paragraphs = doc_xml.findall('.//w:p',namespaces=doc_xml.nsmap)
        for paragraph in paragraphs:
            # Получаем параграф в виде объекта skelmis.docx
            test_obj = parse_xml(ET.tostring(paragraph))
            # Если в параграфе есть атрибуты pPr и numPr, то он содержит элемент списка
            if test_obj.pPr is not None and test_obj.pPr.numPr is not None:
                # Получаем используемый номер набора уровней
                num_id = test_obj.pPr.numPr.numId.val
                # Получаем номер уровня списка в наборе
                ilvl = test_obj.pPr.numPr.ilvl.val
                # Получаем значение атрибута w14:paraId параграфа
                para_id = self.get_para_id(paragraph)
                # Создаем объект NumPr, который содержит описание параграфа, содержащего элемент списка
                para_num = NumPr(test_obj,num_id,ilvl,para_id)
                # Добавляем созданный объект в результирующий список
                numPr_lst.append(para_num)
                # Добавляем созданный объект в результирующий словарь
                numPr_dict[para_id] = para_num
        return numPr_lst, numPr_dict

    def get_current_number(self, numId:int, ilvl:int)->str|None:
        """Метод получает текущее значение элемента списка по номеру набору и уровню.
        Если значение не удалось определить, возвращается значение None.
         numId - номер набора уровней списков,
         ilvl - номер уровня в наборе.

         Returns the current value of a list element for the given numbering set and level.
        If the value cannot be determined, returns None.

        Parameters
        ----------
        num_id : int
            The ID of the numbering set (list style).
        ilvl : int
            The level index within the numbering set.
        """

        # Создаем регулярное выражение для поиска шаблонов, указывающих на уровни, в строке,
        # описывающей формат значения номера для текущего уровня.
        # Искомое значение это знак % за каторым следует одна или 2 цифры.
        # При этом могут использоваться значения из предыдущих уровней.
        # Например, строка %1.%2.%3 указывает на необходимость использования текущие значения элементов списков
        # из 1, 2 и 3 уровня в текущем наборе
        lvl_re = re.compile(r'(%\d{1,2})')
        # Проверяем не выходит ли номер запрашиваемого набора за пределы списка найденных наборов
        if len(self.numbers.nums) >= numId:
            # Получаем объект с набором уровней
            num_obj = self.numbers.nums[numId-1]
            # Получаем объект с описанием необходимого уровня
            lvl_obj = num_obj.levels[ilvl]
            # Получаем строку, описывающую формат значения элемента списка для текущего уровня
            text_str = lvl_obj.lvlText
            # Получаем список, содержащий номера всех, используемых для формирования элемента списка, уровней
            text_str_list = lvl_re.findall(text_str)
            if len(text_str_list) > 0:
                # Если полученный список не пуст, перебираем все используемые уровни
                for item in text_str_list:
                    # Получаем номер уровня
                    level = int(item[1:])-1
                    # Получаем объект, описывающий уровень списка, по его номеру
                    curLvl = num_obj.levels[level]
                    # Получаем строку со значением для этого уровня списка
                    levelText = curLvl.get_value(ilvl)
                    # Замещаем обозначение уровня, его значением
                    text_str = text_str.replace(item,levelText)
                return text_str
            else:
                return None
        else:
            return None

    def fill_numbers(self)->None:
        """Метод производит заполнение значений элементов списков по всем найденным в документе параграфам.

        Populates the list element values for all paragraphs found in the document that contain list items."""

        # Перебираем список объектов, в которых содержатся описания параграфов, содержащих элементы списков,
        # в порядке их появления в документе
        for numPr in self.numPr_lst:
            # Устанавливаем значение элемента списка
            numPr.numberP = self.get_current_number(numPr.numId, numPr.ilvl)

    def get_para_id(self,paragraph: BaseOxmlElement)->str:
        """Метод возвращает строку содержащую значение атрибута ``w14:paraId`` параграфа.

        Returns a string containing the value of the ``w14:paraId`` attribute for the paragraph."""
        if getattr(paragraph,'attrib',None) is not None:
            return paragraph.attrib[get_attribute_uri_mame(paragraph.nsmap, 'w14:paraId')]
        else:
            return paragraph._element.attrib[get_attribute_uri_mame(paragraph._element.nsmap, 'w14:paraId')]

    def get_num_by_paraId(self, paraId:str)->str|None:
        """Метод возвращает строку со значением элемента списка по атрибуту ``w14:paraId`` параграфа, в котором он используется.
        Если такого параграфа не нашлось, возвращается значение None.

        paraId - строка со значением атрибута ``w14:paraId``.

        Returns a string with the list element value associated with the paragraph
        identified by the ``w14:paraId`` attribute.
        If no such paragraph is found, returns None.

        Parameters
        ----------
        para_id : str
            The value of the ``w14:paraId`` attribute for the target paragraph.
        """

        if paraId in self.numPr_dict:
            return self.numPr_dict[paraId].numberP
        else:
            return None

    def get_num_by_paragraf(self, paragraph: BaseOxmlElement)->str|None:
        """Метод возвращает строку со значением элемента списка для параграфа, в котором он используется.
        Если такого параграфа не нашлось, возвращается значение None.
        paragraph - объект параграф, для которого выполняется поиск

        Returns a string with the list element value for the paragraph in which it is used.

        If no such paragraph is found, returns None.

        Parameters
        ----------
        paragraph : Paragraph
            The paragraph object for which the search is performed.
        """

        return self.get_num_by_paraId(self.get_para_id(paragraph))

    def set_num_in_paragraf(self, paragraph: BaseOxmlElement)->None:
        """Метод устанавливает для параграфа в свойстве _element.pPr.numPr.text значение элемента списка
         (если в нем используется список)\n
         paragraph - объект параграф, для которого выполняется установка значения"""
        # Проверяем использование в параграфе списка
        if paragraph._element.pPr is not None and paragraph._element.pPr.numPr is not None:
            # Если список используется, устанавливаем значение свойства
            result_num = self.get_num_by_paragraf(paragraph)
            paragraph._element.pPr.numPr.text = result_num

    def set_num(self)->None:
        """Метод устанавливает значения элементов списков по отдельным параграфам и ячейкам всех таблиц документа"""
        # Перебираем все отдельные параграфы документа и устанавливаем значения (если в них используются списки)
        for paragraph in self.document.paragraphs:
            self.set_num_in_paragraf(paragraph)
        # Перебираем все таблицы документа
        for table in self.document.tables:
            # Перебираем все строки таблицы
            for row in table.rows:
                # Перебираем все ячейки в строке
                for cell in row.cells:
                    # Перебираем все параграфы в ячейке и устанавливаем значения (если в них используются списки)
                    for paragraph in cell.paragraphs:
                        self.set_num_in_paragraf(paragraph)
