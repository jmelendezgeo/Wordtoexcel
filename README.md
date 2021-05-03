# Wordtoexcel

## Description
This code was used to move a database in Word files into a more structured form. It has functions that look for the specific pattern and apply a cleanup flow.

## Requirements
* Pandas
* docx2txt
* re

## How does it work?

Initially there was a database of users in many .docx files. The records had the following pattern:

```   
    Claim Number:	  S0M3C0D3C3Details
    Claim Number Cross Reference:
    Name:	  PEDRO I PEREZ
    Birth Date:	  05/24/1949
    Date of Death:
    Sex:	  M
    Address:	  12345 112TH ST
      JAMAICA, NY  12345-6789
    Most recent State:	  NY (33)
    Most recent County:	  QUEENS (590)

 ```
In the main function you enter the **path** to a folder containing all .docx files.

### Reading files
The __main()__ function applies the workflow for reading all documents in **path**. Here, the function leer_documento() saves the information of the .docx file in a string and looks for a regex _pattern_ to save the information in dictionary groups and converts it into a _Pandas DataFrame_.

### Cleaning files
With all records stored in a Pandas DataFrame, a cleanup workflow is executed. This consists of:

* Removes blanks at the beginning and end of each record.
* Separate the ClaimNumber columns into two desired codes: Code1 and Code2.
* Separate the Address column into Address, county, state and zip code.
* Remove unwanted strings written by people in .docx files as: "Paso", "Roll", etc

### Saving files
With our database in a Pandas Dataframe, it is saved in a **.csv** file and in a **.xlsx** file with _guardar()_ function.

## What did we accomplish?

In minutes, we converted a customer database distributed in 206 .docx files to +60k clean and structured records.
After using this code, it was easier to enrich the database and store it in MySQL to optimize the company's workflows.
