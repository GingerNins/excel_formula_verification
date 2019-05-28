"""
Module for creating Excel formula verification tests
- Searches through Excel document and extracts all formulas
- Searches through Excel document and extracts all named ranges
- Creates a test for each formula incorporating the named ranges when applicable
- Exports the tests to a word document
"""
# TODO: Create UI for this
from docx import Document
from template_file import process_template_file
from verification_document import create_document
from utils import create_filename


def main():
    # NOTE: All the following information will come from the UI once it's created
    template_fname = (r'C:/Users/erins/OneDrive - University of North Carolina at Chapel Hill/Protocols and SOPs/' 
                r'MARG-PTC-001 Primary Cell Spinoculation and Latency/MARG-PTC-001a-v1_1-Lewin Template.xlsx')

    out_dir = r'C:/Users/erins/OneDrive - University of North Carolina at Chapel Hill/Protocols and SOPs/' \
              r'MARG-PTC-001 Primary Cell Spinoculation and Latency/'

    template_name: str = 'test'
    template_desc: str = 'test'

    ########
    outfile: str = create_filename(out_dir, 'formulas and names', extension='xlsx')
    template_data: dict = process_template_file(template_fname, outfile)

    filename: str = create_filename(out_dir, template_desc)
    document: Document = create_document(template_data, template_name, template_desc)

    document.save(filename)


if __name__ == '__main__':
    main()
