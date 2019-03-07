from compare import *

def test_matchmaker():
    """ Test MatchMaker. """
    xlsx_path = 'amaranthe.xlsx'
    pdf_path  = 'amaranthe.pdf'
    
    match = Matchmaker(xlsx_path, pdf_path)
    print(match.nospan_match())
    
    
test_matchmaker()
