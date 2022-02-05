# DOCXTemplate

Windows aplication for Word docx documents generation based on Word docx templates and variables

## How to use

1. In the ***templates*** child directory create a Word docx document with a text for template
2. In the text of the template document insert a variables in format ***{$name$}***
3. Open the app, on the start it should scan the templates and variables
4. If the app is already open, press the ***Scan template folder*** button
5. On the left side of the app window there will be a list of found templates
6. On the right side there will be a list of variables, found in the checked templates
7. Enter the values for the variables, if no value, the variable will pass into the generated document
8. Press the ***Generate*** button, it will generate documents with the checked templates and variable values into the ***output*** child directory
9. The ***Generate*** button will recreate all existing output documents with the checked template names

### Known issues

If accidently the variable ***{$name$}*** will be split into several styles, it will break the app functionality, the fix is on the way to combine the variable styles in the broken template automatically 