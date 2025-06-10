# DocxTemplateTool
DocxTemplateTool is a tool that allows you to quickly edit docx files by making customized templates. This is useful for repetitive texts, such as cover letters or email.

## Create a template
The creation of templates is very easy. Here are the necessary steps:
- Create a docx file using MS Word, LibreOffice or any other text editor that handles docx files.
- Write a text
- Edit the text that can change to put it as a placeholder name inside curly brackets

Example: `"Hello {name1} and {name2}, my name is {my_name} and today is {date}"`

To further understand, you may want to look at [the example I used](https://github.com/ZackPedretti/DocxTemplateTool/blob/main/base_files/base_letter.docx).

## Use the tool on a template
To use the tool on a template, you need to put it in the base_files directory of the project. Then, you need to edit the "base_file_name" variable to the name of your base file.
You then need to run the tool. It will ask you, for each file, what the directory of the file should be, and each of the arguments of the template.
After each file, the tool will ask you if you want to continue. Enter 'n' if you want to stop, or anything else if you want to keep going.

## Use default values
You can set default values so the edition is even quicker. When a default value is set, the tool will ask you to write anything to replace the placeholder, or to press enter to use the default value for this placeholder. It can be very handy for the date for example, as you can set the default value to today's date.

To set default values, you will need to write them in the `default_replacements` dictionnary:
- The placeholder name should be the key
- The default value should be the value

As an example, for the example: `"Hello {name1} and {name2}, my name is {my_name} and today is {date}"`

The `default_replacements` dictionnary could be:
```python
default_replacements = {
    "my_name": "Zack Pedretti"
    "date": datetime.datetime.now().strftime("%d/%m/%y"),
}
```

## Export to PDF
To export to pdf, you will need to install [LibreOffice](https://fr.libreoffice.org/download/telecharger-libreoffice/), as it will be used for the export. You will need to verify that `C:\Program Files\LibreOffice\program\soffice.exe` exists.
The file will then be automatically exported.
