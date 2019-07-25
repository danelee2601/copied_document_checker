# copied_document_checker

<h2>Description</h2>
It finds out copied documents among multiple documents.<br>
[NOTE] This code can only accept file extentions of '.doc', '.docx'(ms word files), '.pdf' <br>

<h2>Installation</h2>
pip install copied-document-checker<br>

<h2>Dependencies</h2>
numpy, pandas, matplotlib, scikit-learn, pdfminer.six, docx, comtypes

<h2>Quick Start</h2>
<pre>
import os
import copied_document_checker
from copied_document_checker import copied_doc_checker

\# path of the directory that contains the document files that you want to inspect.
example_path = os.path.dirname(copied_document_checker.__file__) + '/students_homeworks_example'  # you can put your directory
print('\n# example_path: ', example_path, end='\n\n')

\# run
checker = copied_doc_checker.CopiedDocumentChecker(example_path)
checker.run(n_top_likely=15)   # number of documents that are the most likely to be copied.
</pre>

<h2>Based Algorithms/Knowledge</h2>
Document parsing: n-gram parsing, Bag Of Words (BOW)<br>
Measuring similarity: euclidean distance (modified by giving additional penalties for the matched word counts)
