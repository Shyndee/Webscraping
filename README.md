**Web Scraping and Excel Automation Project**


**Overview**

This project is a web scraping and data processing tool that extracts information from GitHub's search results page, processes the data, and stores it in an Excel file. The project demonstrates the use of Python, regular expressions, and Excel automation with OpenPyXL. Additionally, a bar chart is generated to visualize the number of stars for different repositories.

**Features**

Scrapes GitHub search results for repository information.

Extracts relevant details such as:

Repository titles

Last updated date

Programming language

Number of stars

Repository links

Converts short links into complete GitHub URLs.

Populates an Excel sheet with the extracted data.

Converts star counts from shorthand (e.g., "1.2k") into numeric values.

Generates a bar chart to visualize repository popularity.

**Technologies Used** 

Python

urllib.request for web requests

re (Regular Expressions) for data extraction

OpenPyXL for Excel manipulation

BarChart (OpenPyXL) for data visualization


**How It Works**



**Web Scraping:**

Uses urllib.request to fetch the HTML content of GitHub's search results.

Extracts repository details using regular expressions.

**Data Processing:**

Short GitHub links are converted into full URLs.

Star counts with 'k' notation are converted to numeric values.

**Excel Automation:**

Data is stored in an Excel sheet with appropriate headers.

Star counts are stored as numbers for visualization.

**Chart Generation:**

A bar chart is created in Excel to represent repository popularity based on stars.

**License**

This project is for educational purposes and not intended for commercial use.

**Author**

Shyndee - Developed for learning web scraping and data automation with Python.

