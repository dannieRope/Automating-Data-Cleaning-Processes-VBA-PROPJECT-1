# AUTOMATING DATA CLEANING PROCESSES- VBA PROJECT

## INTRODUCTION

Before analysing data, it is important that the data is clean, meaning it is free from errors, null values, blank values, inconsistent data types, duplicates, and other issues. 
Data cleaning improves the quality of the data, reduces the time required for analysis, and helps prevent inaccurate conclusions.

  However, data cleaning can often be tedious, especially large datasets or dealing with data from the same source in the same format repeatedly. This is where automating the data cleaning process becomes valuable.
Imagine loading your data and simply clicking a button on your spreadsheet to clean it. 
This project outlines how to automate the data cleaning process using VBA programming, making it efficient and less monotonous.

## PROBLEM STATEMENT
Data cleaning process is a very important activity of a data analyst before analysing any type of data. If data is not cleaned, it may lead to inaccurate presumptions about data-driven insights and poor decisions will be made based on those insights.
Data cleaning can also be tedious and time consuming and boring particualarly whe dealing with large datasets the same source in a consistent format. This means the analyst would have to be cleaning same data anytime it comes in.
The repetitive nature of the task makes the data a prime candidate for automation. 

## PROJECT OBJECTIVE

1. Develop a VBA-based solution that automates the data cleaning process
2. Eliminate common data issues such as errors, null values, blank values, inconsistent data types, and duplicates.
3. Enhance the user experience by allowing data cleaning to be performed with a simple button click within the spreadsheet.

## THE DATASET

The data provided is made up of 4 columns and 5000 rows.
Find the description of the columns of the dataset below: 

`Customer Id` - Unique identification number for a customer

`Customer name` - Name of customer

`customer email` - Email address of customer

`Referrer email` - Email address of referrer

`Sales amount` - Amount spent by each customer

`Date` - Date the purchase was done

Check out the preview of the dataset below: 

![Screenshot 2024-08-03 145505](https://github.com/user-attachments/assets/2f66d218-2b68-40b4-8c46-5f50c41d9c6b)


## DATA CLEANING
Note: Date the cleaning process is done in the Visual Basic Editor availble in the developer tab in Microsoft excel

**Step 1.  Spliting the Customer Name column**

The customer Name column consist of both firt name and second name combine together with a comma delimeter. Both Names were extracted and place in seperate columns

![Screenshot 2024-08-03 150605](https://github.com/user-attachments/assets/9e7d4007-2dfe-4c7e-bf81-2d3ed6310de3)

Result

![Screenshot 2024-08-03 150722](https://github.com/user-attachments/assets/101c4fff-54ab-49e5-aeff-835c2c8bb175)


**Step 2.  Extract Email from the Referrer Email Column**

The formatting of the referrer email column is inconsistent.
some of the emails are enclosed in <> and others enclosed in []. A user-defined function "EmailExtract()" was created to help extract the correct email format as shown below. 

![Screenshot 2024-08-03 151656](https://github.com/user-attachments/assets/29596369-b0e1-44b8-afad-808dcdcdeb98)


![Screenshot 2024-08-03 151233](https://github.com/user-attachments/assets/6a5b962f-b1cc-4ab7-8916-2e0f6c4f053a)

Result

![Screenshot 2024-08-03 151747](https://github.com/user-attachments/assets/27cc0ed6-ce4c-4790-9b93-f18015fee845)

**Step 3.  Extracting the Id number and Formating the Date Column**

The the id before the number was remove leaving only the number on the customer ID column also
the date format in Date column is not consistent. It is made up of "dd-mm-yyyy" and "dd/mm/yyyy" format. Changed all dateFormat to dd/mm/yyyy

![Screenshot 2024-08-03 152731](https://github.com/user-attachments/assets/37c9fd30-f2b2-472a-918f-d6a3a85f877c)

Result

![Screenshot 2024-08-03 153235](https://github.com/user-attachments/assets/146798b1-f2f3-44aa-afc4-18d0f4d6b811)

**Step 4.  Removing outliers in Sales Amount column**

The sales amount column consists of 187 outliers which was removed. 

![Screenshot 2024-08-03 153434](https://github.com/user-attachments/assets/cd3bdbb6-8031-4a30-a9dd-c7b7119de70f)


Result
![Screenshot 2024-08-03 153835](https://github.com/user-attachments/assets/fe2ef971-6a8e-45f8-88d5-53a117d1b79e)

**Step 5.  Creating a button to run the code when clicked**

A button was created to enhance the user experience by allowing data cleaning to be performed with a simple button click within the spreadsheet

![Screenshot 2024-08-03 155039](https://github.com/user-attachments/assets/5b6a93dd-5047-4e13-8641-7140a4b774d6)

## FINAL RESULT

### Video Demo
Watch a demo video here: https://www.youtube.com/watch?v=fq9q-0oTQCM

Here is how the cleaned data looks like
![Screenshot 2024-08-05 100940](https://github.com/user-attachments/assets/754859cf-f1f0-4fa6-a5ec-0e36522f76e2)

Find the VBA script here
Find the Workbook here 

# THANKS FOR READING 
















