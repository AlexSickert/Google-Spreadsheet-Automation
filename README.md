# Google-Spreadsheet-Automation
JavaScript code used in Google Docs and Google Spreadsheets

##vocabulary-processor.js
This script is a tool that helps me studying new words in various languages. 

This script provides the following functionality: 

1. pull URLs from cells in a table
2. get the content of the URL by doing an HTTP request
3. parse the content and ret rid of all characters not part of the latin or Cyrillic alphabet
4. extract words 
5. calculate the frequency of each word
6. sort the words by frequency
7. check if the words with highest frequency are already in a list of known words
8. write the remaining non-known words to a table
9. automatically translate each word by using google translate
