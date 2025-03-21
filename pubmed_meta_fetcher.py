#!/usr/bin/env python3
"""
fetch_pubmed_metadata.py

A standalone script to read PubMed IDs (one ID per line) from a provided text file,
query the NCBI Entrez (PubMed) API for each ID's metadata, and export the results
to both CSV and XLSX files.

Usage:
    python fetch_pubmed_metadata.py <path_to_pubmed_ids_txt>
    python fetch_pubmed_metadata.py --debug (to run in debug mode)

Example:
    python fetch_pubmed_metadata.py "/Users/caleb/Downloads/pubmed_ids.txt"

Requirements:
    - Python 3.x
    - pandas
    - openpyxl (for writing XLSX files)
    - biopython (for accessing Bio.Entrez)
    - tqdm (for progress bar and ETA)
    - pip install pandas openpyxl biopython tqdm

Output:
    - institution_publications_metadata.csv
    - institution_publications_metadata.xlsx
"""

import sys
import logging
import time
import pandas as pd
import os

from tqdm import tqdm
from Bio import Entrez
from Bio.Entrez import HTTPError

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def test_pubmed_api() -> bool:
    """
    Test the PubMed API with a known working PubMed ID to verify the connection
    and response format.
    
    :return: True if the connection and data fetch are successful, False otherwise.
    """
    test_id = "33176117"  # A known valid PubMed ID
    
    logging.info(f"Testing PubMed API with known valid ID: {test_id}")
    try:
        # Entrez requires a valid email address
        Entrez.email = "test@example.com"
        
        handle = Entrez.efetch(db="pubmed", id=test_id, rettype="xml", retmode="xml")
        records = Entrez.read(handle)
        handle.close()
        
        if "PubmedArticle" in records and len(records["PubmedArticle"]) > 0:
            logging.info("✅ PubMed API test successful - received valid data")
            article = records["PubmedArticle"][0]
            title = article.get("MedlineCitation", {}).get("Article", {}).get("ArticleTitle", "")
            logging.info(f"   Example Title: {title}")
            return True
        else:
            logging.error("❌ PubMed API test failed - received empty response for valid ID")
            logging.debug(f"Response keys: {list(records.keys())}")
            return False
    except Exception as e:
        logging.error(f"❌ PubMed API test failed with error: {e}")
        return False

def search_pubmed_id(search_term: str) -> str:
    """
    Search PubMed for a term and return the first matching PubMed ID.
    Useful for converting DOIs, PMC IDs, or other identifiers to PubMed IDs.
    
    :param search_term: The term to search for (DOI, PMC ID, title, etc.)
    :return: The first matching PubMed ID or None if no matches.
    """
    Entrez.email = "test@example.com"
    
    try:
        handle = Entrez.esearch(db="pubmed", term=search_term, retmax=1)
        results = Entrez.read(handle)
        handle.close()
        
        id_list = results.get("IdList", [])
        if id_list:
            pmid = id_list[0]
            logging.info(f"Found PubMed ID {pmid} by searching for '{search_term}'")
            return pmid
        else:
            logging.warning(f"No PubMed records found for search term: '{search_term}'")
            return None
    except Exception as e:
        logging.error(f"Error searching PubMed for '{search_term}': {e}")
        return None

def fetch_pubmed_record(pmid: str) -> dict:
    """
    Fetch metadata for a single PubMed ID using NCBI Entrez efetch (XML format).
    
    :param pmid: A string representing the PubMed ID to fetch.
    :return: A dictionary containing keys:
        - 'PubMed_ID'
        - 'Title'
        - 'Authors'
        - 'Journal'
        - 'Year'
      If retrieval fails or metadata is incomplete, returns partial data
      or an empty dict.
    """
    Entrez.email = "test@example.com"

    logging.debug(f"Making API request for PubMed ID: {pmid}")
    try:
        # Typical PubMed IDs are up to ~8 digits, so if the ID is longer, it may be invalid.
        if len(pmid) > 8 and pmid.isdigit():
            logging.warning(f"PubMed ID {pmid} is unusually long; it may not be valid.")
        
        handle = Entrez.efetch(db="pubmed", id=pmid, rettype="xml", retmode="xml")
        records = Entrez.read(handle)
        handle.close()
        
        if not records:
            logging.warning(f"No response data for PubMed ID {pmid}.")
            # Attempt fallback search
            fallback_id = search_pubmed_id(pmid)
            if fallback_id and fallback_id != pmid:
                return fetch_pubmed_record(fallback_id)
            return {}
        
        if "PubmedArticle" not in records or len(records["PubmedArticle"]) == 0:
            logging.warning(f"No valid PubmedArticle found for PubMed ID {pmid}.")
            # Attempt fallback search
            fallback_id = search_pubmed_id(pmid)
            if fallback_id and fallback_id != pmid:
                return fetch_pubmed_record(fallback_id)
            return {}
        
        article_data = records["PubmedArticle"][0]
        medline_citation = article_data.get("MedlineCitation", {})
        article = medline_citation.get("Article", {})

        # PubMed ID (usually the same as the input, but we can confirm from the record)
        record_pmid = medline_citation.get("PMID", pmid)

        # Title
        title = article.get("ArticleTitle", "")

        # Journal
        journal_info = article.get("Journal", {})
        journal_title = journal_info.get("Title", "")

        # Authors (join ForeName + LastName)
        author_list = []
        authors = article.get("AuthorList", [])
        for author in authors:
            last_name = author.get("LastName", "")
            fore_name = author.get("ForeName", "")
            if last_name or fore_name:
                author_list.append(f"{fore_name} {last_name}".strip())
            elif "CollectiveName" in author:
                author_list.append(author["CollectiveName"])
        authors_str = ", ".join(author_list)

        # Year (preferred from JournalIssue -> PubDate -> Year)
        journal_issue = journal_info.get("JournalIssue", {})
        pub_date = journal_issue.get("PubDate", {})
        year = pub_date.get("Year", "")
        if not year:
            date_created = medline_citation.get("DateCreated", {})
            year = date_created.get("Year", "")

        return {
            "PubMed_ID": str(record_pmid),
            "Title": title,
            "Authors": authors_str,
            "Journal": journal_title,
            "Year": year
        }

    except HTTPError as e:
        logging.error(f"HTTPError while fetching PubMed ID {pmid}: {e}")
        # Attempt fallback search
        fallback_id = search_pubmed_id(pmid)
        if fallback_id and fallback_id != pmid:
            return fetch_pubmed_record(fallback_id)
        return {}
    except Exception as e:
        logging.error(f"Unexpected error while fetching PubMed ID {pmid}: {e}")
        return {}

def validate_pubmed_ids(ids_list) -> list:
    """
    Validate a list of potential PubMed IDs to ensure they are in a correct numeric or PMC format.
    Also attempts conversion of PMC IDs to corresponding PubMed IDs if possible.

    :param ids_list: A list of ID strings to validate.
    :return: A list of valid (and cleaned) PubMed IDs.
    """
    valid_ids = []
    pmc_ids_to_convert = []
    
    for raw_id in ids_list:
        candidate = raw_id.strip()
        if not candidate:
            continue

        # Check if it is a PMC ID
        if candidate.upper().startswith("PMC"):
            pmc_ids_to_convert.append(candidate)
            continue

        # If purely numeric, or possibly longer numeric that might still be valid
        if candidate.isdigit():
            # If it's a typical length (<=8 digits), accept it
            if len(candidate) <= 8:
                valid_ids.append(candidate)
            else:
                # Log and add it anyway, or attempt to take last 8 digits
                logging.warning(f"ID {candidate} is longer than 8 digits. Attempting last 8.")
                last_8 = candidate[-8:]
                valid_ids.append(last_8)
        else:
            # Not numeric, attempt to do a fallback search by the string
            logging.warning(f"ID {candidate} is not a pure digit nor PMC, searching for match.")
            fallback_id = search_pubmed_id(candidate)
            if fallback_id:
                valid_ids.append(fallback_id)

    # Convert PMC IDs to PubMed IDs if possible
    for pmc_id in pmc_ids_to_convert:
        pmid = search_pubmed_id(pmc_id)
        if pmid:
            valid_ids.append(pmid)
        else:
            logging.warning(f"Unable to convert {pmc_id} to PubMed ID.")

    # Deduplicate while preserving order
    deduped = list(dict.fromkeys(valid_ids))
    return deduped

def main(input_txtfile: str):
    """
    Main execution function:
      1. Test PubMed API connectivity.
      2. Read PubMed IDs from the provided text file.
      3. Validate and deduplicate the IDs (including PMC conversions, if any).
      4. Fetch metadata for each ID, tracking progress via tqdm and periodic logging.
      5. Export results as both CSV and XLSX in the same directory as the input text file.
    """
    # 1. Test PubMed API
    if not test_pubmed_api():
        logging.error("Failed to connect to PubMed API. Please check your internet connection.")
        sys.exit(1)

    # 2. Read lines from the input text file
    if not os.path.exists(input_txtfile):
        logging.error(f"Input file not found: {input_txtfile}")
        sys.exit(1)

    try:
        with open(input_txtfile, 'r') as f:
            pubmed_ids_raw = [line.strip() for line in f if line.strip()]
    except Exception as e:
        logging.error(f"Error reading input text file '{input_txtfile}': {e}")
        sys.exit(1)

    if not pubmed_ids_raw:
        logging.warning("No PubMed IDs found in the provided text file.")
        sys.exit(0)

    # 3. Validate and deduplicate IDs
    pubmed_ids_cleaned = validate_pubmed_ids(pubmed_ids_raw)

    if not pubmed_ids_cleaned:
        logging.warning("No valid PubMed IDs after validation/conversion.")
        sys.exit(0)

    logging.info(f"Found {len(pubmed_ids_cleaned)} unique valid PubMed IDs to fetch.")

    # 4. Fetch metadata with progress tracking
    rows = []
    success_count = 0
    fail_count = 0
    last_print_time = time.time()

    for pmid in tqdm(pubmed_ids_cleaned, desc="Fetching PubMed metadata", unit="ID"):
        metadata = fetch_pubmed_record(pmid)
        if metadata and "PubMed_ID" in metadata:
            rows.append(metadata)
            success_count += 1
        else:
            fail_count += 1

        # Periodically print how many have succeeded/failed
        if (time.time() - last_print_time) >= 10:
            logging.info(f"{success_count} records found, {fail_count} not found so far.")
            last_print_time = time.time()

    if not rows:
        logging.warning("No valid metadata could be retrieved. Exiting.")
        sys.exit(0)

    df_result = pd.DataFrame(rows, columns=["PubMed_ID", "Title", "Authors", "Journal", "Year"])
    logging.info(f"Successfully retrieved metadata for {len(df_result)} articles.")

    # 5. Write output to CSV and XLSX in the same directory as the input text file
    input_dir = os.path.dirname(os.path.abspath(input_txtfile))
    output_csv = os.path.join(input_dir, "institution_publications_metadata.csv")
    output_xlsx = os.path.join(input_dir, "institution_publications_metadata.xlsx")

    df_result.to_csv(output_csv, index=False)
    df_result.to_excel(output_xlsx, index=False)

    logging.info(f"Exported metadata to: {output_csv}")
    logging.info(f"Exported metadata to: {output_xlsx}")
    logging.info("Done.")

if __name__ == "__main__":
    # Check for debug mode
    if len(sys.argv) > 1 and sys.argv[1] == "--debug":
        logging.info("Running in debug mode with test PubMed ID.")
        if not test_pubmed_api():
            logging.error("PubMed API test failed. Check your configuration/internet.")
            sys.exit(1)
        test_id = "33176117"  # Known valid ID
        logging.info(f"Testing fetch with ID: {test_id}")
        result = fetch_pubmed_record(test_id)
        if result:
            logging.info(f"Fetch successful: {result}")
        else:
            logging.error(f"Fetch failed for ID: {test_id}")
        sys.exit(0)

    if len(sys.argv) != 2:
        print("Usage: python fetch_pubmed_metadata.py <path_to_pubmed_ids_txt>")
        print("       python fetch_pubmed_metadata.py --debug (to run in debug mode)")
        sys.exit(1)

    input_file_path = sys.argv[1]
    main(input_file_path)