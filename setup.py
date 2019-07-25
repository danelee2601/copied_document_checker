from setuptools import setup, find_packages

setup(
    name             = 'copied_document_checker',
    version          = '1.3',
    description      = 'Find out matched documents that are likely to be copied.',
    author           = 'Daesoo Lee',
    author_email     = 'danelee2601@naver.com',
    url              = 'https://github.com/danelee2601/copied_document_checker',
    install_requires = [ ],
    packages         = find_packages(),
    keywords         = ['copied documents', 'plagiarism', 'plagiarize'],
    python_requires  = '>=3',
    package_data     =  {'': ['*.docx', '*.pdf']},
)
