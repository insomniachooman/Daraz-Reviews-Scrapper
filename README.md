Daraz Reviews Scraper
This Python script allows you to scrape reviews for a specific item from the Daraz website. It uses Selenium WebDriver to automate the browser interactions and saves the collected reviews in an Excel file.

Prerequisites
Before running the script, make sure you have the following dependencies installed:

Python (version 3.6 or higher)
Selenium WebDriver
openpyxl library
You can install the required libraries using pip:


pip install selenium openpyxl
Additionally, you need to have the Chrome WebDriver executable installed and added to your system's PATH.


Make sure to download the version that matches your installed Chrome browser version.

Usage
Clone or download the script to your local machine.

Open a terminal or command prompt and navigate to the directory where the script is located.

Run the script using the following command:


python daraz_reviews_scraper.py
When prompted, enter the item you want to scrape reviews for.

Optionally, you can specify the maximum number of pages to collect reviews from. If you leave it blank, the script will scrape all available pages.

The script will start scraping the reviews for the specified item. It will display the progress and the number of reviews collected.

If the script encounters any errors or missing reviews for a particular item, it will skip that item and move on to the next one.

Once the scraping is complete or if you interrupt the script using Ctrl+C, the collected reviews will be saved in an Excel file named daraz_reviews.xlsx in the same directory as the script.

The script will display the total number of collected reviews and the file path where the Excel file is saved.

Notes
The script uses the Chrome browser in incognito mode to scrape the reviews.

If the daraz_reviews.xlsx file already exists, the script will append the newly collected reviews to the existing file. If the file doesn't exist, it will create a new one.

The script saves the reviews in the format of "Item" and "Review" columns in the Excel file.

If you interrupt the script using Ctrl+C, it will gracefully save the collected reviews and exit.

The script has error handling to skip items that have missing reviews or encounter errors during the scraping process.

License
This script is released under the MIT License.
