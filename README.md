# Spanish Natural Language Analyzer

This is a freelance project for a high school Spanish teacher. It is a word frequency analyzer for .txt, .docx, and .doc that will catch and organize unique words into a list organized by most frequent to least frequent as well as alphabetical order among frequency levels. There are options for analyzing single files, copy pasted text or analyzing folders recursively. There is also a menu to add filters in case you don't want the parser to catch a specific word. 

Logic wise, the program does not differentiate between Spanish and English. It will capture unicode and ascii words. Therefore, if you want to use this to analyze your own text documents that aren't Spanish, it will still work.

### Credit to other packages that helped make this open source project possible:
- Xceed's free package (for extracting .DocX xml)
- NPOI.HPWF (for extracting text from .Doc)
