# Spanish Natural Language Analyzer

This is a freelance project for a high school Spanish teacher. The end goal is to create a program that parses a document, which can be .txt, .docx, or .doc(originally .pdf was included but has since been removed), and produce a list of words ordered from most frequent to least frequent. 

Other caveats desired:

1.  Group together like verbs
    1. This will require a dictionary of verbs to accomplish(There's at least 501 with multiple conjugations). The program will also not be able to differentiate between verbs/nouns for similarly spelled words without significantly more work.
2. Extract words from tables/charts/graphs if possible rather than just paragraphs for the end frequency count.

C# and WPF is the focus of this project, but I will transition the logic to other platforms as needed.
