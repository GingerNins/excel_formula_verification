"""
Module for creating Excel formula verification tests
- Searches through Excel document and extracts all formulas
- Searches through Excel document and extracts all named ranges
- Creates a test for each formula incorporating the named ranges when applicable
- Exports the tests to a word document
"""
# TODO: Create UI for this

from datetime import datetime
from docx import Document
from template_file import process_template_file
from verification_document import create_document, create_filename


def main():
    template_fname = (r'C:/Users/erins/OneDrive - University of North Carolina at Chapel Hill/Protocols and SOPs/' 
                r'MARG-PTC-001 Primary Cell Spinoculation and Latency/MARG-PTC-001a-v1_1-Lewin Template.xlsx')

    out_dir = r'C:/Users/erins/OneDrive - University of North Carolina at Chapel Hill/Protocols and SOPs/' \
              r'MARG-PTC-001 Primary Cell Spinoculation and Latency/'

    out_pre = datetime.now().strftime('%Y%m%d') + '_ELS_'
    out_suf = '_' + datetime.now().strftime('%H%M%S')
    outfile = out_dir + out_pre + r'list_formulas_and_names' + out_suf + '.xlsx'

    template_data: dict = process_template_file(template_fname, outfile)

    template_name: str = 'test'
    template_desc: str = 'test'

    filename: str = create_filename(out_dir, template_desc)
    document: Document = create_document(template_data, template_name, template_desc)

    document.save(filename)


if __name__ == '__main__':
    main()
