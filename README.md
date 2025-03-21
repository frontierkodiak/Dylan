# Dylan

A collection of helpful Python scripts for research and data processing tasks.

## Setup Instructions

### Clone the Repository

```bash
git clone https://github.com/yourusername/Dylan.git
cd Dylan
```

### Create a Virtual Environment

```bash
# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

### Install Required Packages

```bash
pip install pandas openpyxl biopython tqdm
```

## Available Scripts

### PubMed Metadata Fetcher

A tool to fetch metadata for a list of PubMed IDs and export the results to CSV and Excel files.

#### Usage

1. Create a text file with one PubMed ID per line (e.g., `pubmed_ids.txt`)

2. Run the script:

```bash
python pubmed_meta_fetcher.py path/to/pubmed_ids.txt
```

For example:

```bash
python pubmed_meta_fetcher.py data/pubmed_ids.txt
```

#### Output

The script will create two files in the same directory as your input file:
- `institution_publications_metadata.csv` - CSV format of the metadata
- `institution_publications_metadata.xlsx` - Excel format of the metadata

Each file contains the following data for each PubMed article:
- PubMed ID
- Title
- Authors
- Journal
- Publication Year