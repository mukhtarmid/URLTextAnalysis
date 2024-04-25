# NLP Article Analysis

This project contains Python code for analyzing articles from a given list of URLs. The code calculates various scores and metrics related to the articles' content, such as polarity, subjectivity, complexity, and readability.

This project encompasses a Python script designed to analyze articles sourced from a specified list of URLs. The script is engineered to compute a range of scores and metrics that offer insights into the content of the articles. These metrics include:
- Polarity: Indicates the sentiment of the text, whether it leans towards positive, negative, or neutral.
- Subjectivity: Reflects how subjective or objective the text is in its tone and content.
- Complexity: Measures the complexity of the language used in the articles, often based on the presence of intricate or specialized vocabulary.
- Readability: Assesses the ease with which the text can be read and understood by the audience.
By calculating these scores and metrics, the script provides a comprehensive analysis of the articles' content, enabling a deeper understanding of their characteristics and potential impact on readers.

## Installation

To install the project, follow these steps:

1. Clone the repository: `https://github.com/mukhtarmid/URLTextAnalysis`
2. Install dependencies: 
- `pip install beautifulsoup4`
- `pip install selenium`
- `pip install requests`
- `pip install pandas`
- `pip install string`
- `pip install nltk`
- `pip install openpyxl`
3. Run the project: `python TextAnalysis.py`

## Setup

To setup, follow these steps:

1. Download the Input.xlsx file from the provided directory and place it in your working directory.
2. Update the file paths in the code to match the location of the Input.xlsx file and the stop words files.

## Usage

To use the project, follow these steps:

1. Open the project in your preferred code editor.
2. Make changes to the code as necessary.
3. Run the project: `python TextAnalysis.py`

## Code Overview

The code consists of several functions and steps to perform the following tasks:
1. Read the `Input.xlsx` file: The code reads the `Input.xlsx` file, which contains a list of URLs and their corresponding IDs.
2. Extract article content: The code uses the requests library to fetch the content of each URL and parses it using `BeautifulSoup`.
3. Preprocess the text: The code tokenizes the text, removes stop words, calculates syllable counts, and performs other preprocessing tasks.
4. Calculate scores and metrics: The code calculates various scores and metrics for each article, such as polarity, subjectivity, complexity, and readability.
5. Save the results: The code saves the calculated scores and metrics in a new Excel file named `Output Data Structure.xlsx`.

## Contributing

If you'd like to contribute to the project, please follow these guidelines:

1. Fork the repository.
2. Create a new branch for your changes: `git checkout -b my-new-feature`
3. Make your changes and commit them: `git commit -m "Add new feature"`
4. Push your changes to your forked repository: `git push origin my-new-feature`
5. Submit a pull request.

## Contact

If you have any questions or comments, please contact me at [mukhtarmid@gmail.com](mailto:mukhtarmid@gmail.com).