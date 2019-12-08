sources\_to\_docx
================

Little script to write a bunch of files with source code to .docx.

Requirements: python3, python-docx.


Just listings (probably, ugly looking):
```
# List source files
$ find some_project -name '*.cpp' -or -name '*.h' > some_project.txt
$ find some_library -name '*.cpp' -or -name '*.h' > some_library.txt

$ sources_to_docx.py --input *.txt
                     --output code_listings.docx
```


Use template .docx file as reference and apply style named 'My\_code\_style' to listings:  
```
$ sources_to_docx.py --input *.txt
                     --output code_listings.docx
                     --template template.docx 
                     --code_style 'My_code_style'
```

Use .docx file as reference, apply code style and print code instead of paragraph with text '\<CODE\>':
```
$ sources_to_docx.py --input *.txt
                     --output document_with_code_listings.docx
                     --template document.docx 
                     --code_style 'My_code_style'
                     --code_place_mark '<CODE>'
```

Print heading "Library \<file\_list\_name\>" for each file list,
apply some custom styles:
```
$ sources_to_docx.py --input *.txt
                     --output code_listings.docx
                     --template template.docx 
                     --code_style code
                     --with_headings
                     --heading_style 'Heading 3'
                     --text_style 'My_text_style'
                     --code_style 'My_code_style'
                     --file_list_description 'Library {}'
                     --file_description 'File {}:'
```

Usage of GOST 19.401-78 ("ЕСПД") template:
```
$ sources_to_docx.py --input *.txt
                     --output code_listings.docx
                     --template gost_espd/template.docx 

                     --with_headings # Optional
                     --heading_style 'td_toc_caption_level_1'
                     --code_style td_code
                     --text_style td_text

                     --file_list_description 'Подпрограмма {}'
                     --file_description 'Исходный модуль {}:'
```

