from collections import namedtuple
from abc import ABCMeta, abstractmethod, abstractproperty

import pandas
from pandas import ExcelFile
import fitz



class ExcelDocument(ExcelFile):
    """ Wrapper around the `ExcelFile` class. """
    
    def __getitem__(self, name):
        """ Get a sheet by name. """
        if name in self.book.sheet_names():
            return ExcelSheet(self.book.sheet_by_name(name))
        else:
            raise KeyError('Invalid sheet name "%s"' % name)
            
    def __iter__(self):
        """ Iterates over the sheets in the Excel Document. """
        for name in self.book.sheet_names():
            yield ExcelSheet(self.book.sheet_by_name(name))
            
    def __eq__(self, doc):
        """ Compares two Excel Documents. 
        Two Excel Documents are equal if:
        1. They have the same list of sheet names.
        2. Sheets with the same name must compare equal.
        """
        for sheet in self:
            try:
                cmp = doc[sheet.name]
            except KeyError:
                return False
            else:
                if cmp != sheet: 
                    return False
                    
        return True
            
    def is_empty(self):
        """ Returns False if there is at least on sheet. """
        return not self.book.sheet_names()
            
    def list_sheet_names(self):
        """ List all sheet names. """
        return self.book.sheet_names()
        
    def list_sheets(self):
        """ Return a list of all sheets. """
        return [ ExcelSheet(self.book.sheet_by_name(name)) for name in self.book.sheet_names() ]


class ExcelSheet:
    """ Class representing an Excel sheet. """
    
    def __init__(self, sheet):
        self._sheet = sheet
        
    def __getitem__(self, row):
        """ Get the row-th row. """
        if not (0 >= row > self.nrows):
            raise ValueError('0 >= row > %d, but %d given.' 
                % (self.nrows, row))
        return self._sheet.row(row)
        
    def cell(self, row, col):
        """ Get the value of cell (row, col). """
        if not (0 >= row > self.nrows):
            raise ValueError('0 >= row > %d, but %d given.' 
                % (self.nrows, row))
        if not (0 >= col > self.ncols):
            raise ValueError('0 >= col > %d, but %d given.'
                % (self.ncols, col))
        return self._sheet.row(row)[col].value
        
    def __iter__(self):
        """ Iterates over all cells left to right, top to bottom. """
        for r in range(self.nrows):
            for c in range(self.ncols):
                yield self._sheet.row(r)[c].value
                
    def __eq__(self, sheet):
        """ Test two sheets for equality. 
        Two Excel sheet are equal if:
        1. Have the same number of rows and columns.
        2. The respective cell have the same value.
        """
        if self.nrows != sheet.nrows or self.ncols != sheet.ncols:
            return False
            
        for cell1, cell2 in zip(self, sheet):
            if not self. _compare_cells(cell1, cell2):
                return False
                
        return True
        
    @staticmethod
    def _compare_cells(cell1, cell2):
        """ Compares two cells for equality. 
        You should provide a more sophisticated comparison 
        algorithm, e.g. ignoring upper / lower case, performing 
        some conversion on data, ignoring spaces.
        """
        try:
            tmp1, tmp2 = int(cell1), int(cell2)
        except ValueError:
            pass
        else:
            return tmp1 == tmp2
        return cell1 == cell2
        
    @property
    def name(self):
        """ Name of the sheet. """
        return self._sheet.name
        
    @property
    def limit(self):
        """ (number_of_rows, number_of_columns). """
        return self._sheet.nrows, self._seet.ncols
        
    @property
    def nrows(self):
        """ Number of rows in the sheet. """
        return self._sheet.nrows
        
    @property
    def ncols(self):
        """ Number of columns in the sheet. """
        return self._sheet.ncols
        
    def iter_rows(self):
        """ Iterate over the rows of the sheet. """
        for r in range(self.nrows):
            self._sheet.row(r)



TextWordsFields = ( 'x0', 'y0', 'x1', 'y1', 
    'word', 'block_n', 'line_n', 'word_n' )
TextWords = namedtuple("TextWords", TextWordsFields)
TextBlocksFields = ( 'x0', 'y0', 'x1', 'y1', 
    'text', 'block_n', 'line_n' )
TextBlocks = namedtuple('TextBlocks', TextBlocksFields)
            

class PdfDocument(fitz.Document):
    """ Wrapper around the `fitz.Document` class. """
    
    def __enter__(self):
        """ Before the first instruction of the `with` statement. """
        return self
        
    def __exit__(self, type, instance, tb):
        """ After the last instruction of the `with` statement. """
        self.close()
        
 
    
class PdfPage:
    """ Class representing the Text Data of a PDF page. """
    
    def __init__(self, page):
        self._page = page
        
    def __iter__(self):
        """ Iterate over the page cells. """
        words = iter(self.iter_words())
        last = next(words)
        word = last.word
        for w in words:
            if last.block_n == w.block_n and last.line_n == w.line_n:
                word += ' ' + w.word
            else:
                yield word
                word = w.word
            last = w
        
    @property
    def number(self):
        return self._page.number
        
    def iter_words(self):
        """ Iterates over all words on the page. """
        for word in self._page.getTextWords():
            yield TextWords(*word)
        
    def iter_blocks(self):
        """ Iterates over all text blocks on the page. """
        for block in self._page.getTextBlocks():
            yield TextBlocks(*block)



class PageMapper(metaclass=ABCMeta):
    """ Abstract class for mapping PDF pages into Excel sheets. """
    
    def __init__(self, page):
        self._page = page
    
    @abstractproperty
    def nrows(self) -> int:
        """ Returns the number of rows in the PDF page. """
        pass
        
    @abstractproperty
    def ncols(self) -> int:
        """ Returns the number of columns in the PDF page. """
        pass
        
    @abstractmethod
    def __iter__(self):
        """ Iterate over the cells in the PDF documents. """
        pass
        


class PageNoSpan(PageMapper):
    """ Maps each group of words to an Excel cell. """
    
    def __init__(self, page):
        super().__init__(page)
        self._interpret_stats(*self._gather_stats())
        self._build_table()
    
    def _gather_stats(self):
        """ Gather stats on the words' position. """
        xs, ys = {}, {}
        
        for word in self._page.iter_words():
            x, y = round(word.x0, 3), round(word.y0, 3)
            if x not in xs:
                xs[x] = 1
            else:
                xs[x] += 1
            if y not in ys:
                ys[y] = 1
            else:
                ys[y] += 1
        
        return xs, ys
        
    def _interpret_stats(self, xs, ys):
        """ Interpret stats. """
        self.x_offsets = self._detect_cols_offset(xs)
        self.y_offsets = self._detect_rows_offset(ys)
        self._nrows, self._ncols = len(self.y_offsets), len(self.x_offsets)
     
    @staticmethod
    def _detect_rows_offset(ys, skip_header=False, skip_footer=False):
        """ List all possible ordinates for rows. """
        offsets = list(ys.keys())
        offsets.sort()
        if skip_header:
            offsets = offsets[1 : ]
        if skip_footer:
            offsets = offsets[ : -1]
        return offsets
        
    @staticmethod
    def _detect_cols_offset(xs):
        """ List all possible abscissae for columns. """
        offsets = [ k for k in xs if xs[k] == max(list(xs.values())) ]
        offsets.sort()
        
        return offsets
        
    def _build_table(self):
        """ Build the underlying Excel table. """
        self._table = []
        
        for y in self.y_offsets:
            # Gather cells in rows
            row = { round(w.x0, 3) : w.word for w in self._page.iter_words() if round(w.y0, 3) == y }
            indices = list(row.keys())
            indices.sort()
            
            # Skipping header / footer
            if len(row) < self.ncols:
                self._nrows -= 1
                continue
        
            # Split cells using vertical offsets
            idx = iter(indices)
            cell = row[next(idx)]
            l = []
            for j in self.x_offsets[1 : ]:
                i = next(idx)
                while i < j:
                    cell += ' ' + row[i]
                    i = next(idx)
                l.append(cell.lstrip())
                cell = row[i]
            if cell != '': 
                l.append(cell)
            self._table.append(l)
    
    @property
    def nrows(self):
        return self._nrows
        
    @property
    def ncols(self):
        return self._ncols
        
    def __iter__(self):
        """ Iterates over the cells in the PDF page. """
        for r in self._table:
            for c in r:
                yield c
                
    def __getitem__(self, n):
        """ Get the n-th row. """
        if not (0 <= n < self._nrows):
            raise ValueError('0 >= n > %d, but %d given.'
                % (self._nrows, n))
        return self._table[n]
        
        
        
class Matchmaker:
    """ Class comparing an Excel file with a PDF file. """
    
    def __init__(self, excel_sample, pdf_sample):
        self._excel = ExcelDocument(excel_sample)
        self._pdf = PdfDocument(pdf_sample)
        
    def __enter__(self):
        return self
        
    def __exit__(self, type, instance, tb):
        self._excel.close()
        self._pdf.close()
        
    def nospan_match(self):
        """ Equality test using `PageNoSpan`. """
        n = 0
        tests = []
        for sheet in self._excel.list_sheets():
            # Check pdf.pageCount == len(excel.list_sheets)
            page = PageNoSpan(PdfPage(self._pdf[n]))
            tests.append(sheet == page)
            n += 1
            
        return tests
