# Blackbird.io Google Sheets

Blackbird is the new automation backbone for the language technology industry. Blackbird provides enterprise-scale automation and orchestration with a simple no-code/low-code platform. Blackbird enables ambitious organizations to identify, vet and automate as many processes as possible. Not just localization workflows, but any business and IT process. This repository represents an application that is deployable on Blackbird and usable inside the workflow editor.

## Introduction

<!-- begin docs -->

Google Sheets is a cloud-based spreadsheet tool that enables easy collaboration on creating, editing, and analyzing data online. With real-time collaboration features and a user-friendly interface, it's a versatile platform for organizing information, managing projects, and creating charts and graphs.

## Actions

### Spreadsheet Actions

- **Add new sheet row** Adds a new row to the first empty line of the sheet
- **Create sheet** Creates sheet
- **Download sheet CSV file** Downloads CSV file
- **Find sheet row** Providing a column address and a value, returns row number where said value is located
- **Get column** Gets column values
- **Get range** Gets specific range
- **Get sheet cell** Gets cell by address
- **Get sheet row** Gets sheet row by address
- **Get sheet used range** Gets used range
- **Update sheet cell** Updates cell by address
- **Update sheet row** Updates row by start address
- **Update sheet column** Updates column by start address

### Glossary Actions

- **Import glossary** Imports glossary as a sheet
- **Export glossary** Exports glossary from a sheet

To utilize the **Export glossary** action, ensure that the Google sheet mirrors the structure obtained from the **Import glossary** action result. Follow these guidelines:

- **Sheet structure**:
    - The first row serves as column names, representing properties of the glossary entity: _ID_, _Definition_, _Subject field_, _Notes_, _Term (language code)_, _Variations (language code)_, _Notes (language code)_.
    - Include columns for each language present in the glossary. For instance, if the glossary includes English and Spanish, the column names will be: _ID_, _Definition_, _Subject field_, _Notes_, _Term (en)_, _Variations (en)_, _Notes (en)_, _Term (es)_, _Variations (es)_, _Notes (es)_.
- **Optional fields**:
    - _Definition_, _Subject field_, _Notes_, _Variations (language code)_, _Notes (language code)_ are optional and can be left empty.
- **Main term and synonyms**:
    - _Term (language code)_ represents the primary term in the specified language for the glossary.
    - _Variations (language code)_ includes synonymous values for the term.
- **Notes handing**:
    - Notes in the _Notes_ column should be separated by ';' if there are multiple notes for a given entry.
- **Variations handling**:
    - Variations in the _Variations (language code)_ column should be separated by ';' if there are multiple variations for a given term.
- **Terms notes format**:
    - Each note in the _Notes (language code)_ column should follow this structure: **Term or variation: note**.
    - Notes for terms should be separated by ';;'. For example, 'money: may refer to physical or banked currency;; cash: refers to physical currency.'

## Feedback

Do you want to use this app or do you have feedback on our implementation? Reach out to us using the [established channels](https://www.blackbird.io/) or create an issue.

<!-- end docs -->
