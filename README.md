# Web Scraping and Text Analysis Project  

This project is a Python-based application designed to scrape web pages, extract textual content, and perform advanced text analysis. The analysis includes sentiment evaluation, readability scoring, and linguistic feature extraction. It outputs processed results in a structured Excel format and saves the extracted text as files for further use.  

---

## Table of Contents  
- [Features](#features)  
- [Technologies Used](#technologies-used)  
- [Setup and Usage](#setup-and-usage)  
- [How It Works](#how-it-works)  

---

## Features  

1. **Text Extraction**  
   - Scrapes content from URLs provided in an Excel file.  
   - Saves the text from each URL as a `.txt` file in a structured directory.  

2. **Text Analysis**  
   - Calculates sentiment metrics: positive score, negative score, polarity, and subjectivity.  
   - Computes readability metrics: Fog Index, average sentence length, percentage of complex words.  
   - Performs linguistic analysis: word count, syllable count, complex word count, and personal pronoun count.  

3. **Output Storage**  
   - Extracted text files are saved for archival and manual review.  
   - Analysis results are logged in an Excel file for structured reporting.  

---

## Technologies Used  

### 1. **Libraries for Web Scraping and Analysis**  
- **BeautifulSoup4**: For HTML parsing and content extraction.  
- **Requests**: For fetching web page HTML.  
- **NLTK**: For tokenizing and processing text during sentiment analysis.  
- **TextStat**: For calculating readability metrics.  

### 2. **Excel Handling Tools**  
- **OpenPyXL**: For reading input URLs from Excel and writing analysis results.  

### 3. **Python Environment**  
- **Python 3.8 or higher**: Required for running the project scripts.  

---

## Setup and Usage  

This project is designed to run on any Python-supported environment. Follow these steps to set up and execute the project:  

### Steps  

1. **Clone or Download the Project**  
   - Clone this repository or download the files to your local machine.  

2. **Install Dependencies**  
   - Use pip to install the required libraries:  
     ```bash
     pip install requests beautifulsoup4 nltk openpyxl textstat
     ```  

3. **Prepare the Input File**  
   - Ensure `Input.xlsx` is present in the project root directory.  
   - Populate it with two columns:  
     - `URL_ID`: Unique ID for each URL.  
     - `URL`: The web page link to scrape.  

4. **Run the Text Extraction Script**  
   - Execute the extraction script to scrape and save web page content:  
     ```bash
     python DataExtraction.py
     ```  
   - Extracted content will be saved in the `extracted_texts` directory.  

5. **Run the Text Analysis Script**  
   - Analyze the extracted text files and generate metrics:  
     ```bash
     python text_analysis.py
     ```  
   - Results will be saved in the `analysis_results.xlsx` file.  

---

## How It Works  

The project processes web pages and performs advanced text analysis through the following steps:  

1. **Text Extraction**  
   - Reads URLs from `Input.xlsx`.  
   - Fetches the HTML content using `Requests` and parses it with `BeautifulSoup4`.  
   - Extracted text is cleaned and saved as `.txt` files in the `extracted_texts` folder.  

2. **Text Analysis**  
   - Extracted text files are read and analyzed for:  
     - **Sentiment Metrics**: Positive score, negative score, polarity, and subjectivity.  
     - **Readability Metrics**: Fog Index, average sentence length, and percentage of complex words.  
     - **Linguistic Features**: Word count, complex word count, and personal pronoun count.  
   - Results are stored in `analysis_results.xlsx`.  

3. **Output and Reporting**  
   - The extracted text and analysis results are structured for easy review and integration with other workflows.  

This workflow ensures efficient web scraping and robust analysis, making it suitable for a variety of use cases such as content research, sentiment evaluation, and readability optimization.  
