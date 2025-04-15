Okay, I understand the requirements for refactoring the SEO Content Optimizer script. This is a significant refactoring task. Here's the plan:

1.  **Integrate Embeddings:** Fetch competitor content, generate embeddings, and use these embeddings to find relevant competitor text snippets. These snippets will be used as context/inspiration for the LLM (Claude) during new content generation and updates.
2.  **Minimize Updates & Color-Coding:** Implement a diffing mechanism between original and updated content sections. Use special markers or HTML tags (`<ins>`, `<del>`) in the generated HTML. Modify the Word document generation function to parse these markers/tags and apply color/strikethrough formatting using `python-docx`.
3.  **Heading Prefixes:** Add "(H2)" and "(H3)" prefixes to headings in the Word document outputs.
4.  **Remove Skeleton Output:** Eliminate the separate "Recommended Content Structure" section from the final Word brief.
5.  **Improve Content Flow:** Modify LLM prompts to explicitly request better transitions and referencing of previous sections, potentially providing the last paragraph of the previous section as context.
6.  **Slim Down & Improve:** Refactor for clarity, remove redundancy, improve error handling, consolidate where possible, and enhance comments/docstrings. Standardize HTML parsing for document generation.

I will provide the refactored code broken down into chunks as requested.

**Note on Complexity:** Implementing the fine-grained, color-coded diffing in the Word document (`<ins>`/`<del>` tag parsing and applying formatting) is complex and can be fragile depending on the LLM's output consistency and the intricacies of HTML-to-DOCX conversion. I will implement this as requested, but be aware it might require further refinement based on real-world testing.

---

**Refactored Code - Chunk 1/4**

```python
# -*- coding: utf-8 -*-
"""
Refactored SEO Content Optimizer Script

This script analyzes SERPs, competitor content, and existing content 
to provide recommendations and generate/update SEO-optimized articles.

Key Refactoring Changes:
- Uses competitor content embeddings to inform content generation and updates.
- Implements minimal updates with color-coded diff tracking in Word output.
- Adds (H2)/(H3) prefixes to Word document headings.
- Removes the standalone structure skeleton from the Word brief output.
- Enhances LLM prompts for better content flow between sections.
- General code slimming, improved error handling, and comments.
"""

import streamlit as st
import pandas as pd
import numpy as np
import requests
import json
import time
import re
from bs4 import BeautifulSoup
import trafilatura
import openai
import anthropic
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import base64
import random
from typing import List, Dict, Any, Tuple, Optional
import logging
import traceback
import openpyxl
import matplotlib.pyplot as plt
import altair as alt
import difflib # Added for content diffing

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# --- Configuration Constants ---
DEFAULT_LOCATION_CODE = 2840  # USA
DEFAULT_LANGUAGE_CODE = "en"
DEFAULT_DEVICE = "desktop"
DEFAULT_OS = "windows"
SERP_DEPTH = 30
MAX_ORGANIC_RESULTS = 10
MAX_RELATED_KEYWORDS = 20
MAX_KEYWORD_SUGGESTIONS = 20
SCRAPE_TIMEOUT = 15
EMBEDDING_MODEL_OPENAI = "text-embedding-3-large" # Or "text-embedding-3-small"
EMBEDDING_DIM_LARGE = 3072
EMBEDDING_DIM_SMALL = 1536
ANTHROPIC_MODEL = "claude-3-5-sonnet-20240620" # Use the latest Sonnet model
TEXT_ANALYSIS_MAX_LENGTH = 15000 # Max length for text sent to LLM for analysis
ARTICLE_GENERATION_MAX_TOKENS = 6000
ARTICLE_TARGET_WORD_COUNT = 1500
INTERNAL_LINK_BATCH_SIZE = 20
INTERNAL_LINK_MIN_SIMILARITY = 0.70 # Increased threshold slightly
INTERNAL_LINK_MAX_COUNT_FACTOR = 8 # Links per 1000 words
INTERNAL_LINK_MIN_COUNT = 3

# Color definitions for diffing
COLOR_ADDED = RGBColor(0, 128, 0) # Green
COLOR_DELETED = RGBColor(255, 0, 0) # Red

# User Agents for Scraping
USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36'
]


# Set page configuration (should be the first Streamlit command)
st.set_page_config(
    page_title="SEO Content Optimizer v2",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

###############################################################################
# 1. Utility Functions
###############################################################################

def display_error(error_msg: str, exception: Optional[Exception] = None):
    """Display error message in Streamlit and log it."""
    st.error(f"Error: {error_msg}")
    logger.error(error_msg)
    if exception:
        logger.error(f"Exception details: {exception}")
        logger.error(traceback.format_exc())

def get_download_link_html(file_bytes: bytes, filename: str, link_label: str) -> str:
    """Returns an HTML link to download file (for use in Markdown)."""
    b64 = base64.b64encode(file_bytes).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{link_label}</a>'
    return href

def format_time(seconds: float) -> str:
    """Format time in seconds to readable string 'X min Y.Z sec' or 'Y.Z sec'."""
    if seconds < 60:
        return f"{seconds:.1f} seconds"
    else:
        minutes = int(seconds // 60)
        sec = seconds % 60
        return f"{minutes} min {sec:.1f} sec"

def clean_html(html_content: str) -> str:
    """Basic HTML cleaning for text extraction."""
    if not html_content:
        return ""
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        # Remove script, style, comments
        for element in soup(["script", "style", "comment"]):
            element.extract()
        text = soup.get_text(separator='\n', strip=True)
        # Normalize whitespace
        text = re.sub(r'\n{3,}', '\n\n', text)
        return text.strip()
    except Exception as e:
        logger.warning(f"HTML cleaning failed: {e}. Returning original content.")
        # Fallback: try basic regex removal if BS4 fails
        text = re.sub(r'<script.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<style.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', ' ', text) # Replace tags with spaces
        text = re.sub(r'\s{2,}', ' ', text).strip() # Normalize whitespace
        return text

def truncate_text(text: str, max_length: int) -> str:
    """Truncate text to a maximum length."""
    if len(text) > max_length:
        return text[:max_length] + "..."
    return text

def count_words(text: str) -> int:
    """Count words in a given text string."""
    if not text:
        return 0
    return len(re.findall(r'\b\w+\b', text))

def safe_json_loads(json_string: str) -> Optional[Dict[str, Any]]:
    """Attempts to parse JSON, handling potential errors and formats."""
    if not json_string:
        return None
        
    # Try direct parsing first
    try:
        return json.loads(json_string)
    except json.JSONDecodeError:
        pass # Continue to other methods

    # Try extracting JSON from markdown code blocks
    match = re.search(r'```json\s*(\{.*?\})\s*```', json_string, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass

    # Try finding the first valid JSON object structure
    match = re.search(r'(\{.*?\})', json_string, re.DOTALL)
    if match:
        try:
            # Attempt to repair common issues like trailing commas
            repaired_json = re.sub(r',(\s*[}\]])', r'\1', match.group(1))
            return json.loads(repaired_json)
        except json.JSONDecodeError as e:
            logger.warning(f"Failed to parse extracted JSON: {e}")
            return None # Give up if even the extracted part fails

    logger.warning("Could not find or parse valid JSON in the provided string.")
    return None

###############################################################################
# 2. API Integration - DataForSEO (SERP & Keywords)
###############################################################################

class DataForSEOClient:
    """Client for interacting with the DataForSEO API."""

    def __init__(self, api_login: str, api_password: str):
        if not api_login or not api_password:
            raise ValueError("DataForSEO API Login and Password are required.")
        self.api_login = api_login
        self.api_password = api_password
        self.base_url = "https://api.dataforseo.com/v3"

    def _post_request(self, endpoint: str, data: List[Dict]) -> Tuple[Optional[Dict], str]:
        """Helper function to make POST requests to DataForSEO API."""
        url = f"{self.base_url}/{endpoint}"
        headers = {'Content-Type': 'application/json'}
        try:
            response = requests.post(
                url,
                auth=(self.api_login, self.api_password),
                headers=headers,
                json=data,
                timeout=30 # Increased timeout
            )
            response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
            
            response_data = response.json()
            if response_data.get('status_code') == 20000:
                 # Check if tasks and results are present before accessing
                if response_data.get('tasks') and len(response_data['tasks']) > 0 and \
                   response_data['tasks'][0].get('result') and len(response_data['tasks'][0]['result']) > 0:
                    return response_data['tasks'][0]['result'][0], "Success"
                else:
                     # Handle cases where the structure is unexpected but status is 20000
                    logger.warning(f"API returned success code but unexpected result structure for {endpoint}: {response_data}")
                    return None, "API Error: Unexpected result structure."
            else:
                error_msg = f"API Error ({response_data.get('status_code')}): {response_data.get('status_message')}"
                logger.error(error_msg)
                return None, error_msg

        except requests.exceptions.RequestException as e:
            error_msg = f"HTTP Request Error for {endpoint}: {e}"
            logger.error(error_msg)
            return None, error_msg
        except Exception as e:
            error_msg = f"Unexpected Error during API call to {endpoint}: {e}"
            logger.error(error_msg, exc_info=True)
            return None, error_msg

    def classify_page_type(self, url: str, title: str, snippet: str) -> str:
        """Classify page type based on URL, title, and snippet patterns."""
        title_lower = title.lower() if title else ""
        snippet_lower = snippet.lower() if snippet else ""
        url_lower = url.lower() if url else ""

        # Prioritize E-commerce if specific keywords appear strongly
        commerce_keywords = ['product', '/p/', 'buy', 'shop', 'add to cart', 'checkout', 'price']
        if any(keyword in url_lower or keyword in title_lower or keyword in snippet_lower for keyword in commerce_keywords):
             # Stronger check: URL structure often indicates product/category pages
            if '/product/' in url_lower or '/p/' in url_lower or '/shop/' in url_lower or '/buy/' in url_lower or '/category/' in url_lower:
                return "E-commerce"
             # Weaker check: keywords in title/snippet
            if any(pattern in title_lower or pattern in snippet_lower for pattern in ['buy', 'shop', 'purchase', 'price', 'discount', 'sale', 'order', 'checkout', 'cart']):
                 return "E-commerce"

        # Check for forums
        forum_patterns = ['forum', 'community', 'discussion', 'thread', 'topic', '/t/', '/f/', 'ask', 'question', 'answer']
        if any(pattern in url_lower or pattern in title_lower or pattern in snippet_lower for pattern in forum_patterns):
             # Stronger check for URL patterns
            if '/forum/' in url_lower or '/community/' in url_lower or '/discussion/' in url_lower or '/thread/' in url_lower or '/topic/' in url_lower or '/t/' in url_lower or '/f/' in url_lower:
                return "Forum/Community"
            if any(pattern in title_lower or pattern in snippet_lower for pattern in ['forum', 'community', 'discussion', 'thread', 'question', 'answer', 'replies', 'comments']):
                return "Forum/Community"


        # Check for Reviews/Comparisons
        review_patterns = ['review', 'comparison', 'vs', 'versus', 'top 10', 'best', 'rating', 'rated', 'alternative']
        if any(pattern in title_lower or pattern in snippet_lower or pattern in url_lower for pattern in review_patterns):
            # Stronger check for URL patterns
            if '/review/' in url_lower or '/comparison/' in url_lower or '/vs/' in url_lower:
                 return "Review/Comparison"
            if any(pattern in title_lower for pattern in review_patterns):
                return "Review/Comparison"


        # Check for Articles/Blogs
        article_patterns = ['blog', 'article', 'news', 'post', 'how to', 'guide', 'tips', 'tutorial', 'learn', 'what is', 'why', 'when']
        if any(pattern in url_lower or pattern in title_lower or pattern in snippet_lower for pattern in article_patterns):
             # Stronger check for URL patterns
            if '/blog/' in url_lower or '/article/' in url_lower or '/news/' in url_lower or '/post/' in url_lower or '/guide/' in url_lower:
                return "Article/Blog"
            if any(pattern in title_lower for pattern in article_patterns):
                return "Article/Blog"

        # Default type
        return "Informational"


    def fetch_serp_results(self, keyword: str) -> Tuple[Optional[List[Dict]], Optional[List[Dict]], Optional[List[Dict]], str]:
        """Fetch SERP results, classify pages, and extract PAA/Features."""
        logger.info(f"Fetching SERP results for keyword: {keyword}")
        post_data = [{
            "keyword": keyword,
            "location_code": DEFAULT_LOCATION_CODE,
            "language_code": DEFAULT_LANGUAGE_CODE,
            "device": DEFAULT_DEVICE,
            "os": DEFAULT_OS,
            "depth": SERP_DEPTH
        }]
        
        result_data, status_msg = self._post_request("serp/google/organic/live/advanced", post_data)
        
        if result_data is None:
            return None, None, None, status_msg
            
        organic_results = []
        paa_questions = []
        serp_features_dict = {}

        items = result_data.get('items', [])
        if not items:
             logger.warning(f"No items found in SERP result for keyword: {keyword}")
             return [], [], [], "Success (No items found)"


        # --- Pass 1: Extract Organic and PAA ---
        organic_count = 0
        for item in items:
            item_type = item.get('type')
            
            # Organic Results
            if item_type == 'organic' and organic_count < MAX_ORGANIC_RESULTS:
                url = item.get('url')
                title = item.get('title')
                snippet = item.get('breadcrumb') or item.get('snippet') # Prefer breadcrumb for snippet if available
                
                if url and title: # Require URL and Title
                    page_type = self.classify_page_type(url, title, snippet if snippet else "")
                    organic_results.append({
                        'url': url,
                        'title': title,
                        'snippet': snippet if snippet else "",
                        'rank_group': item.get('rank_group'),
                        'page_type': page_type
                    })
                    organic_count += 1

            # People Also Ask (handle multiple structures)
            elif item_type == 'people_also_ask':
                for paa_item in item.get('items', []):
                    if paa_item.get('type') == 'people_also_ask_element':
                        question_data = {
                            'question': paa_item.get('title', ''),
                            'expanded': [] # Placeholder for potential future expansion
                        }
                        # Extract expanded element data if available
                        for expanded in paa_item.get('expanded_element', []):
                             if expanded.get('type') == 'people_also_ask_expanded_element':
                                 question_data['expanded'].append({
                                     'url': expanded.get('url', ''),
                                     'title': expanded.get('title', ''),
                                     'description': expanded.get('description', '')
                                 })

                        if question_data['question']:
                           paa_questions.append(question_data)

            elif item_type == 'people_also_ask_element': # Direct PAA element
                question_data = {
                    'question': item.get('title', ''),
                    'expanded': []
                }
                for expanded in item.get('expanded_element', []):
                     if expanded.get('type') == 'people_also_ask_expanded_element':
                         question_data['expanded'].append({
                             'url': expanded.get('url', ''),
                             'title': expanded.get('title', ''),
                             'description': expanded.get('description', '')
                         })
                if question_data['question']:
                   # Avoid duplicates if already found via container
                   if not any(q['question'] == question_data['question'] for q in paa_questions):
                       paa_questions.append(question_data)

        # --- Pass 2: Extract SERP Features (excluding organic) ---
        for item in items:
            item_type = item.get('type')
            if item_type != 'organic':
                serp_features_dict[item_type] = serp_features_dict.get(item_type, 0) + 1
                
        serp_features = [{'feature_type': ft, 'count': ct} for ft, ct in serp_features_dict.items()]
        serp_features = sorted(serp_features, key=lambda x: x['count'], reverse=True)

        logger.info(f"Extracted {len(organic_results)} organic results, {len(paa_questions)} PAA questions, {len(serp_features)} feature types.")
        return organic_results, serp_features, paa_questions, "Success"


    def _fetch_keywords_from_endpoint(self, endpoint: str, keyword: str, limit: int) -> Tuple[Optional[List[Dict]], str]:
        """Internal helper to fetch keywords from different DataForSEO Labs endpoints."""
        logger.info(f"Fetching keywords for '{keyword}' from endpoint: {endpoint}")
        post_data = [{
            "keyword": keyword,
            "location_code": DEFAULT_LOCATION_CODE,
            "language_code": DEFAULT_LANGUAGE_CODE,
            "limit": limit
        }]
        
        # Add endpoint specific parameters if needed
        if endpoint == "dataforseo_labs/google/keyword_suggestions/live":
            post_data[0]["include_serp_info"] = True # Get SV etc.
            post_data[0]["include_seed_keyword"] = True
        elif endpoint == "dataforseo_labs/google/related_keywords/live":
            post_data[0]["language_name"] = "English" # API uses name here
            # Optional: post_data[0]["depth"] = 3 # Example depth

        result_data, status_msg = self._post_request(endpoint, post_data)
        
        if result_data is None:
            return None, status_msg

        keywords_list = []
        extracted_keywords = set() # To avoid duplicates

        items = result_data.get('items', [])

        # Handle seed keyword data specifically for keyword_suggestions
        if endpoint == "dataforseo_labs/google/keyword_suggestions/live" and 'seed_keyword_data' in result_data:
             seed_data = result_data['seed_keyword_data']
             if seed_data and 'keyword_info' in seed_data:
                 kw_info = seed_data['keyword_info']
                 kw = result_data.get('seed_keyword', '')
                 if kw and kw not in extracted_keywords:
                     keywords_list.append({
                         'keyword': kw,
                         'search_volume': kw_info.get('search_volume'),
                         'cpc': kw_info.get('cpc'),
                         'competition': kw_info.get('competition')
                     })
                     extracted_keywords.add(kw)


        # Process items array (common structure)
        if items and isinstance(items, list):
            for item in items:
                # Structure varies slightly between endpoints
                kw_data = item.get('keyword_data') or item # Use item directly if keyword_data not present
                kw = kw_data.get('keyword')
                kw_info = kw_data.get('keyword_info')

                if kw and kw_info and kw not in extracted_keywords:
                     # Validate data types
                     sv = kw_info.get('search_volume')
                     cpc = kw_info.get('cpc')
                     comp = kw_info.get('competition')
                     
                     keywords_list.append({
                         'keyword': kw,
                         'search_volume': int(sv) if sv is not None else None,
                         'cpc': float(cpc) if cpc is not None else None,
                         'competition': float(comp) if comp is not None else None
                     })
                     extracted_keywords.add(kw)

        if not keywords_list:
            logger.warning(f"No keywords found in the response from {endpoint} for '{keyword}'.")
            # Return empty list but success status if API call itself was okay
            if status_msg == "Success":
                return [], "Success (No keywords found)"
            else:
                # Propagate the API error message
                return None, status_msg


        # Sort by search volume (descending), handling None values
        keywords_list.sort(key=lambda x: x.get('search_volume') or 0, reverse=True)
        
        logger.info(f"Successfully extracted {len(keywords_list)} keywords from {endpoint} for '{keyword}'.")
        return keywords_list, "Success"

    def fetch_related_keywords(self, keyword: str) -> Tuple[Optional[List[Dict]], str]:
        """Fetch related keywords, falling back to suggestions if needed."""
        keywords, status = self._fetch_keywords_from_endpoint(
            "dataforseo_labs/google/related_keywords/live", 
            keyword, 
            MAX_RELATED_KEYWORDS
        )
        
        # Fallback to Keyword Suggestions if Related Keywords fails or returns empty
        if keywords is None or (status == "Success (No keywords found)" and not keywords):
            logger.warning(f"Related Keywords failed or returned empty for '{keyword}'. Falling back to Keyword Suggestions.")
            keywords, status = self._fetch_keywords_from_endpoint(
                "dataforseo_labs/google/keyword_suggestions/live", 
                keyword, 
                MAX_KEYWORD_SUGGESTIONS
            )
            
        return keywords, status

###############################################################################
# 3. Web Scraping and Content Analysis
###############################################################################

def scrape_webpage(url: str) -> Tuple[Optional[str], str]:
    """
    Scrape main content from a webpage using Trafilatura with fallback.
    Returns: content (str) or None, status_message (str)
    """
    logger.info(f"Attempting to scrape content from: {url}")
    content = None
    status_message = "Scraping failed"

    # Attempt 1: Trafilatura (often best for article content)
    try:
        downloaded = trafilatura.fetch_url(url, timeout=SCRAPE_TIMEOUT)
        if downloaded:
            # Extract main content, favoring text content, include tables
            content = trafilatura.extract(
                downloaded,
                include_comments=False,
                include_tables=True,
                favor_precision=True, # Prioritize cleaner extraction
                output_format='text' # Get plain text
            )
            if content and len(content) > 100: # Basic check for meaningful content
                logger.info(f"Successfully scraped content using Trafilatura for {url} ({len(content)} chars)")
                return content, "Success (Trafilatura)"
            else:
                 logger.warning(f"Trafilatura extracted minimal content from {url}. Trying fallback.")
        else:
            logger.warning(f"Trafilatura failed to download {url}. Trying fallback.")
    except Exception as e:
        logger.warning(f"Trafilatura failed for {url}: {e}. Trying fallback.")

    # Attempt 2: Requests + BeautifulSoup (Fallback)
    try:
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': 'https://www.google.com/'
        }
        response = requests.get(url, headers=headers, timeout=SCRAPE_TIMEOUT)
        response.raise_for_status() # Check for HTTP errors

        soup = BeautifulSoup(response.text, 'html.parser')

        # Remove non-content elements more aggressively
        for element in soup(["script", "style", "header", "footer", "nav", "aside", "form", "noscript", "iframe", "button", "input", "select", "textarea", ".sidebar", ".ad", ".advertisement", ".popup", ".modal", ".cookie-banner", ".social-share"]):
             try:
                 if isinstance(element, str): # Check if it's a selector
                      for el in soup.select(element):
                          el.decompose()
                 else: # Assume it's a tag name
                     for el in soup.find_all(element):
                          el.decompose()
             except Exception as decompose_error:
                  logger.debug(f"Could not decompose element {element}: {decompose_error}")


        # Try finding common main content containers
        main_content_selectors = ['main', 'article', '[role="main"]', '.main-content', '.post-content', '.entry-content', '#content', '#main', '.content']
        main_div = None
        for selector in main_content_selectors:
             try:
                 main_div = soup.select_one(selector)
                 if main_div:
                      logger.debug(f"Found main content using selector: {selector}")
                      break
             except Exception as selector_error:
                 logger.debug(f"Selector error for {selector}: {selector_error}")

        if main_div:
             text = main_div.get_text(separator='\n', strip=True)
        else:
             # Fallback to body if no main container found
             logger.warning(f"No main content container found for {url}. Extracting from body.")
             body = soup.find('body')
             text = body.get_text(separator='\n', strip=True) if body else soup.get_text(separator='\n', strip=True)


        # Clean up extracted text
        lines = (line.strip() for line in text.splitlines())
        # Keep only lines with substance (e.g., more than 2 words)
        meaningful_lines = [line for line in lines if count_words(line) > 2]
        cleaned_text = '\n\n'.join(meaningful_lines) # Use double newline for paragraphs


        if cleaned_text and len(cleaned_text) > 100:
            logger.info(f"Successfully scraped content using BeautifulSoup fallback for {url} ({len(cleaned_text)} chars)")
            return cleaned_text, "Success (BeautifulSoup Fallback)"
        else:
            logger.warning(f"BeautifulSoup fallback extracted minimal content from {url}")
            return None, "Scraping Failed (Fallback - No Content)"


    except requests.exceptions.HTTPError as http_err:
         if http_err.response.status_code == 403:
             status_message = "Scraping Failed (403 Forbidden)"
             logger.warning(f"{status_message} for URL: {url}")
         elif http_err.response.status_code == 404:
             status_message = "Scraping Failed (404 Not Found)"
             logger.warning(f"{status_message} for URL: {url}")
         else:
             status_message = f"Scraping Failed (HTTP {http_err.response.status_code})"
             logger.warning(f"{status_message} for URL: {url} - {http_err}")
         return None, status_message
    except requests.exceptions.RequestException as req_err:
        status_message = f"Scraping Failed (Request Exception: {req_err})"
        logger.error(status_message, exc_info=True)
        return None, status_message
    except Exception as e:
        status_message = f"Scraping Failed (Unexpected Error: {e})"
        logger.error(status_message, exc_info=True)
        return None, status_message

    # If we reach here, both methods failed significantly
    logger.error(f"All scraping attempts failed for {url}. Final status: {status_message}")
    return None, status_message


def extract_headings(url: str) -> Tuple[Optional[Dict[str, List[str]]], str]:
    """Extract H1, H2, H3 headings from a webpage."""
    logger.info(f"Extracting headings from: {url}")
    try:
        headers = {
            'User-Agent': random.choice(USER_AGENTS),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        }
        response = requests.get(url, headers=headers, timeout=SCRAPE_TIMEOUT)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        
        headings = {
            'h1': [h.get_text(strip=True) for h in soup.find_all('h1') if h.get_text(strip=True)],
            'h2': [h.get_text(strip=True) for h in soup.find_all('h2') if h.get_text(strip=True)],
            'h3': [h.get_text(strip=True) for h in soup.find_all('h3') if h.get_text(strip=True)]
        }
        
        # Also try to find H4-H6 for more context if needed later
        headings['h4'] = [h.get_text(strip=True) for h in soup.find_all('h4') if h.get_text(strip=True)]
        headings['h5'] = [h.get_text(strip=True) for h in soup.find_all('h5') if h.get_text(strip=True)]
        headings['h6'] = [h.get_text(strip=True) for h in soup.find_all('h6') if h.get_text(strip=True)]

        logger.info(f"Extracted headings for {url}: H1({len(headings['h1'])}), H2({len(headings['h2'])}), H3({len(headings['h3'])}).")
        return headings, "Success"

    except requests.exceptions.RequestException as e:
        error_msg = f"Heading extraction failed for {url}: Request Exception - {e}"
        logger.error(error_msg)
        return None, error_msg
    except Exception as e:
        error_msg = f"Heading extraction failed for {url}: Unexpected Error - {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg

#==============================================================================
# End of Chunk 1/4
#==============================================================================
```

---

**Refactored Code - Chunk 2/4**

```python
#==============================================================================
# Start of Chunk 2/4
#==============================================================================

###############################################################################
# 4. Embeddings and Semantic Analysis (Using OpenAI & Anthropic)
###############################################################################

def generate_embedding(text: str, openai_api_key: str, model: str = EMBEDDING_MODEL_OPENAI) -> Tuple[Optional[List[float]], str]:
    """
    Generate text embedding using OpenAI API.
    Returns: embedding list or None, status message.
    """
    if not text:
        return None, "Input text is empty"
    if not openai_api_key:
         return None, "OpenAI API key is missing"

    # Truncate text to avoid exceeding model limits (check model's max tokens)
    # text-embedding-3-large has 8191 context length. Truncate safely.
    max_input_chars = 30000 # Approx limit based on token estimate
    truncated_text = text[:max_input_chars]
    if len(text) > max_input_chars:
         logger.warning(f"Input text truncated to {max_input_chars} chars for embedding.")


    try:
        # Initialize OpenAI client (consider initializing once if possible)
        client = openai.OpenAI(api_key=openai_api_key)

        response = client.embeddings.create(
            model=model,
            input=[truncated_text] # API expects a list of strings
        )
        
        if response.data and len(response.data) > 0:
             embedding = response.data[0].embedding
             # Verify expected dimension based on model
             expected_dim = EMBEDDING_DIM_LARGE if "large" in model else EMBEDDING_DIM_SMALL
             if len(embedding) == expected_dim:
                 # logger.debug(f"Generated embedding with dimension {len(embedding)} using {model}")
                 return embedding, "Success"
             else:
                  error_msg = f"Generated embedding dimension mismatch: Got {len(embedding)}, expected {expected_dim} for model {model}."
                  logger.error(error_msg)
                  return None, error_msg

        else:
             logger.error(f"OpenAI embedding API returned no data: {response}")
             return None, "API Error: No embedding data returned"

    except openai.APIError as api_err:
        error_msg = f"OpenAI API Error generating embedding: {api_err}"
        logger.error(error_msg)
        return None, f"API Error: {api_err}"
    except Exception as e:
        error_msg = f"Unexpected error generating embedding: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg

def get_anthropic_client(api_key: str) -> Optional[anthropic.Anthropic]:
     """Initializes and returns an Anthropic client."""
     if not api_key:
         logger.error("Anthropic API key is missing.")
         return None
     try:
         return anthropic.Anthropic(api_key=api_key)
     except Exception as e:
         logger.error(f"Failed to initialize Anthropic client: {e}")
         return None

def analyze_semantic_structure(competitor_contents: List[Dict], anthropic_api_key: str) -> Tuple[Optional[Dict], str]:
    """
    Analyze competitor content using Anthropic to determine optimal semantic hierarchy.
    Returns: semantic analysis dict or None, status message.
    """
    client = get_anthropic_client(anthropic_api_key)
    if not client:
        return None, "Anthropic client initialization failed"
        
    # Combine content intelligently, prioritizing titles and headings
    combined_input = ""
    for i, comp in enumerate(competitor_contents):
        title = comp.get('title', f'Competitor {i+1}')
        h1 = comp.get('headings', {}).get('h1', [""])[0] if comp.get('headings') else ""
        h2s = comp.get('headings', {}).get('h2', []) if comp.get('headings') else []
        content_snippet = clean_html(comp.get('content', ''))[:1000] # Snippet of main text

        combined_input += f"--- Competitor {i+1}: {title} ---\n"
        if h1: combined_input += f"H1: {h1}\n"
        if h2s: combined_input += f"H2s: {', '.join(h2s[:5])}\n" # First 5 H2s
        combined_input += f"Content Snippet:\n{content_snippet}\n\n"

    combined_input = truncate_text(combined_input, TEXT_ANALYSIS_MAX_LENGTH)

    prompt = f"""
    Analyze the following combined content summaries from top-ranking pages for a specific keyword. 
    Recommend an optimal semantic hierarchy (headings structure) for a new, comprehensive article on this topic.

    Analysis Task:
    1. Identify the core themes and sub-themes consistently covered.
    2. Determine a logical flow for presenting the information.
    3. Suggest a primary H1 title capturing the main topic.
    4. Propose 5-8 main H2 section headings covering key aspects.
    5. Suggest 2-4 relevant H3 subheadings under EACH H2, detailing specific points within that section.

    Input Content Summaries:
    {combined_input}

    Output Format:
    Return ONLY a valid JSON object adhering strictly to this structure:
    {{
        "h1": "Recommended H1 Title",
        "sections": [
            {{
                "h2": "First H2 Section Title",
                "subsections": [
                    {{"h3": "First H3 Subsection under H2-1"}},
                    {{"h3": "Second H3 Subsection under H2-1"}},
                    ...
                ]
            }},
            {{
                "h2": "Second H2 Section Title",
                "subsections": [
                    {{"h3": "First H3 Subsection under H2-2"}},
                    ...
                ]
            }},
            ... more H2 sections ...
        ]
    }}
    Ensure the JSON is well-formed. Do not include any text before or after the JSON object.
    """

    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=2000, # Increased token limit for structure
            system="You are an expert SEO Content Strategist specializing in analyzing competitor content to devise optimal article structures. You provide only valid JSON output.",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2 # Lower temperature for predictable structure
        )
        
        response_text = response.content[0].text
        semantic_analysis = safe_json_loads(response_text)

        if semantic_analysis and "h1" in semantic_analysis and "sections" in semantic_analysis:
             logger.info("Successfully analyzed semantic structure.")
             # Basic validation of structure
             if not isinstance(semantic_analysis["sections"], list):
                  logger.warning("Semantic analysis 'sections' is not a list. Attempting recovery.")
                  semantic_analysis["sections"] = [] # Reset if invalid
             return semantic_analysis, "Success"
        else:
             error_msg = "Failed to parse valid semantic structure JSON from LLM response."
             logger.error(f"{error_msg} Response was: {response_text[:500]}...")
             return None, error_msg

    except anthropic.APIError as api_err:
        error_msg = f"Anthropic API Error analyzing structure: {api_err}"
        logger.error(error_msg)
        return None, f"API Error: {api_err}"
    except Exception as e:
        error_msg = f"Unexpected error analyzing semantic structure: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg


def extract_important_terms(competitor_contents: List[Dict], anthropic_api_key: str, keyword: str) -> Tuple[Optional[Dict], str]:
    """
    Extract important terms, topics, and questions from competitor content using Anthropic.
    Returns: term data dict or None, status message.
    """
    client = get_anthropic_client(anthropic_api_key)
    if not client:
        return None, "Anthropic client initialization failed"

    # Combine content, prioritizing unique text
    combined_content = ""
    word_limit = 10000 # Limit words to manage context window
    current_words = 0
    seen_paragraphs = set()

    for comp in competitor_contents:
         content_text = clean_html(comp.get('content', ''))
         if not content_text:
             continue

         paragraphs = content_text.split('\n\n')
         added_text = ""
         for para in paragraphs:
             para_hash = hash(para) # Simple hash to detect duplicate paragraphs
             if para and para_hash not in seen_paragraphs:
                  para_words = count_words(para)
                  if current_words + para_words <= word_limit:
                      added_text += para + "\n\n"
                      seen_paragraphs.add(para_hash)
                      current_words += para_words
                  else:
                      break # Stop adding paragraphs if limit reached
         
         if added_text:
              combined_content += f"--- Content from {comp.get('url', 'competitor')} ---\n{added_text}\n"

         if current_words >= word_limit:
             break # Stop processing competitors if limit reached


    if not combined_content:
         return None, "No usable competitor content found for term extraction."

    # Truncate just in case, though word limit should handle it
    combined_content = truncate_text(combined_content, TEXT_ANALYSIS_MAX_LENGTH + 5000) # Allow slightly larger input

    prompt = f"""
    Analyze the following combined competitor content related to the primary keyword "{keyword}". 
    Extract key semantic entities relevant to this topic.

    Analysis Tasks:
    1.  **Primary Terms:** Identify the top 10-15 most critical single or multi-word terms (nouns, noun phrases) directly related to "{keyword}". These are essential concepts.
    2.  **Secondary Terms:** Identify 15-25 supporting terms or concepts that add context, detail, or related information to the primary terms.
    3.  **Key Topics:** Summarize the 8-12 main thematic areas or subjects discussed across the content. These should represent distinct sections or major ideas.
    4.  **Questions Answered:** List 5-10 specific questions that the content implicitly or explicitly answers related to "{keyword}".

    Input Content:
    {combined_content}

    Output Format:
    Return ONLY a valid JSON object adhering strictly to this structure:
    {{
        "primary_terms": [
            {{"term": "term1", "importance": 0.95, "recommended_usage": 5}}, 
            {{"term": "term2", "importance": 0.90, "recommended_usage": 4}},
            ...
        ],
        "secondary_terms": [
            {{"term": "termA", "importance": 0.75, "recommended_usage": 2}},
            {{"term": "termB", "importance": 0.70, "recommended_usage": 1}},
            ...
        ],
        "topics": [
            {{"topic": "Topic 1 Name", "description": "Brief description of what this topic covers..."}},
            {{"topic": "Topic 2 Name", "description": "Brief description..."}},
            ...
        ],
        "questions": [
            "Question 1 answered by the content?",
            "Question 2 answered by the content?",
            ...
        ]
    }}
    - Assign an 'importance' score (0.0 to 1.0) based on frequency and centrality.
    - Suggest 'recommended_usage' count based on importance (e.g., higher for more important terms).
    - Ensure the JSON is well-formed. No text outside the JSON structure.
    """

    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=2500, # Allow ample tokens for terms/topics
            system="You are an expert SEO Analyst specializing in semantic content analysis and entity extraction. You provide only valid JSON output.",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3 # Slightly creative but focused on extraction
        )
        
        response_text = response.content[0].text
        term_data = safe_json_loads(response_text)
        
        # Validate the structure
        if term_data and \
           isinstance(term_data.get('primary_terms'), list) and \
           isinstance(term_data.get('secondary_terms'), list) and \
           isinstance(term_data.get('topics'), list) and \
           isinstance(term_data.get('questions'), list):
            logger.info("Successfully extracted important terms, topics, and questions.")
            return term_data, "Success"
        else:
            error_msg = "Failed to parse valid term data JSON from LLM response."
            logger.error(f"{error_msg} Response was: {response_text[:500]}...")
            # Return a default structure on failure
            default_terms = {
                "primary_terms": [], "secondary_terms": [],
                "topics": [], "questions": []
                }
            return default_terms, error_msg

    except anthropic.APIError as api_err:
        error_msg = f"Anthropic API Error extracting terms: {api_err}"
        logger.error(error_msg)
        return None, f"API Error: {api_err}"
    except Exception as e:
        error_msg = f"Unexpected error extracting terms: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg


###############################################################################
# 5. Content Scoring Functions
###############################################################################

def get_score_grade(score: float) -> str:
    """Convert numeric score (0-100) to letter grade."""
    if score >= 97: return "A+"
    elif score >= 90: return "A"
    elif score >= 87: return "A-"
    elif score >= 83: return "B+"
    elif score >= 80: return "B"
    elif score >= 77: return "B-"
    elif score >= 73: return "C+"
    elif score >= 70: return "C"
    elif score >= 67: return "C-"
    elif score >= 60: return "D"
    else: return "F"

def score_content(content_html: str, term_data: Dict, keyword: str) -> Tuple[Optional[Dict], str]:
    """
    Score content based on keyword usage, term coverage, topic coverage, and question answering.
    Accepts HTML content for analysis.
    Returns: score data dict or None, status message.
    """
    if not term_data or not keyword:
        return None, "Missing term data or keyword for scoring."

    # Extract clean text from HTML for analysis
    content_text = clean_html(content_html)
    if not content_text:
         return None, "Content is empty after HTML cleaning."

    content_lower = content_text.lower()
    word_count = count_words(content_text)

    # --- Scoring Components ---
    keyword_score = 0.0
    primary_terms_score = 0.0
    secondary_terms_score = 0.0
    topic_coverage_score = 0.0
    question_coverage_score = 0.0

    # 1. Keyword Score (Primary Keyword Usage)
    keyword_count = len(re.findall(r'\b' + re.escape(keyword.lower()) + r'\b', content_lower))
    # More sophisticated density calculation (aim for 0.5% to 1.5%)
    ideal_density_low = 0.005
    ideal_density_high = 0.015
    actual_density = keyword_count / word_count if word_count > 0 else 0

    if actual_density == 0:
         keyword_score = 0.0
    elif ideal_density_low <= actual_density <= ideal_density_high:
         keyword_score = 100.0 # Perfect score within ideal range
    elif actual_density < ideal_density_low:
         # Scale score based on how close it is to the lower bound
         keyword_score = (actual_density / ideal_density_low) * 80 # Max 80 if below range
    else: # actual_density > ideal_density_high
         # Penalize for over-optimization, but less harshly
         overage_factor = actual_density / ideal_density_high
         penalty = min(40, (overage_factor - 1) * 50) # Max 40 point penalty
         keyword_score = max(60.0, 100.0 - penalty) # Floor score at 60


    # 2. Primary Terms Score
    primary_term_details = {}
    primary_terms_list = term_data.get('primary_terms', [])
    primary_terms_total = len(primary_terms_list)
    primary_score_total_weighted = 0.0
    primary_max_score_weighted = 0.0

    if primary_terms_total > 0:
        for term_info in primary_terms_list:
            term = term_info.get('term', '').lower()
            importance = term_info.get('importance', 0.5) # Default importance
            recommended = term_info.get('recommended_usage', 1)
            if not term: continue

            term_count = len(re.findall(r'\b' + re.escape(term) + r'\b', content_lower))
            primary_term_details[term] = {'count': term_count, 'recommended': recommended, 'importance': importance}

            # Calculate score for this term based on usage vs recommendation
            term_score = 0.0
            if term_count > 0:
                 if term_count >= recommended:
                     # Bonus for hitting target, penalize slightly for massive overuse
                     overuse_ratio = term_count / recommended if recommended > 0 else 1
                     term_score = 100.0 - min(20, (overuse_ratio - 1.5) * 20) # Max 20 penalty starting at 1.5x
                     term_score = max(80.0, term_score) # Floor at 80 if overused
                 else:
                     # Score based on proportion of recommended usage achieved
                     term_score = (term_count / recommended) * 100.0 if recommended > 0 else 0.0

            # Weight the term score by its importance
            primary_score_total_weighted += term_score * importance
            primary_max_score_weighted += 100.0 * importance

        if primary_max_score_weighted > 0:
             primary_terms_score = (primary_score_total_weighted / primary_max_score_weighted) * 100.0


    # 3. Secondary Terms Score (Simpler: Presence weighted by importance)
    secondary_term_details = {}
    secondary_terms_list = term_data.get('secondary_terms', [])
    secondary_terms_total = len(secondary_terms_list)
    secondary_score_total_weighted = 0.0
    secondary_max_score_weighted = 0.0

    if secondary_terms_total > 0:
        for term_info in secondary_terms_list:
            term = term_info.get('term', '').lower()
            importance = term_info.get('importance', 0.3) # Lower default importance
            recommended = term_info.get('recommended_usage', 1)
            if not term: continue

            term_count = len(re.findall(r'\b' + re.escape(term) + r'\b', content_lower))
            secondary_term_details[term] = {'count': term_count, 'recommended': recommended, 'importance': importance}

            # Score based on presence (100 if present, 0 otherwise)
            term_score = 100.0 if term_count > 0 else 0.0

            # Weight the term score by its importance
            secondary_score_total_weighted += term_score * importance
            secondary_max_score_weighted += 100.0 * importance

        if secondary_max_score_weighted > 0:
             secondary_terms_score = (secondary_score_total_weighted / secondary_max_score_weighted) * 100.0

    # 4. Topic Coverage Score
    topic_coverage_details = {}
    topics_list = term_data.get('topics', [])
    topics_total = len(topics_list)
    topics_covered_count = 0

    if topics_total > 0:
        common_words = {'the', 'a', 'an', 'is', 'are', 'to', 'in', 'on', 'and', 'or', 'of', 'for', 'with'}
        for topic_info in topics_list:
            topic_name = topic_info.get('topic', '')
            description = topic_info.get('description', '')
            if not topic_name: continue

            # Extract keywords from topic name and description
            topic_keywords = set(re.findall(r'\b\w{3,}\b', topic_name.lower())) - common_words
            desc_keywords = set(re.findall(r'\b\w{3,}\b', description.lower())) - common_words
            all_keywords = topic_keywords.union(desc_keywords)

            # Simple check: does the topic name appear? Or a significant number of keywords?
            found_keywords = {kw for kw in all_keywords if kw in content_lower}
            coverage_ratio = len(found_keywords) / len(all_keywords) if all_keywords else 0
            
            is_covered = (topic_name.lower() in content_lower) or (coverage_ratio >= 0.5 and len(all_keywords) > 1) or (coverage_ratio >= 0.7)

            topic_coverage_details[topic_name] = {
                'covered': is_covered,
                'match_ratio': round(coverage_ratio, 2),
                'description': description
            }
            if is_covered:
                topics_covered_count += 1

        topic_coverage_score = (topics_covered_count / topics_total) * 100.0

    # 5. Question Coverage Score
    question_coverage_details = {}
    questions_list = term_data.get('questions', [])
    questions_total = len(questions_list)
    questions_answered_count = 0

    if questions_total > 0:
        for question in questions_list:
            if not question: continue
            
            # Extract keywords from question
            question_keywords = set(re.findall(r'\b\w{4,}\b', question.lower())) - common_words - {'what', 'how', 'why', 'when', 'where', 'who', 'which'}
            
            # Check if a good portion of keywords are present in the content
            found_keywords = {kw for kw in question_keywords if kw in content_lower}
            match_ratio = len(found_keywords) / len(question_keywords) if question_keywords else 0
            
            is_answered = (question.lower() in content_lower) or (match_ratio >= 0.6) # Threshold for answering

            question_coverage_details[question] = {
                'answered': is_answered,
                'match_ratio': round(match_ratio, 2)
            }
            if is_answered:
                questions_answered_count += 1
        
        question_coverage_score = (questions_answered_count / questions_total) * 100.0

    # Calculate Overall Score (Adjusted Weights)
    overall_score = (
        keyword_score * 0.15 +             # Reduced weight
        primary_terms_score * 0.35 +       # Increased weight
        secondary_terms_score * 0.15 +     # Maintained weight
        topic_coverage_score * 0.25 +      # Increased weight
        question_coverage_score * 0.10     # Maintained weight
    )
    overall_score = max(0, min(100, round(overall_score))) # Ensure score is between 0 and 100

    # Compile results
    score_data = {
        'overall_score': overall_score,
        'grade': get_score_grade(overall_score),
        'components': {
            'keyword_score': round(keyword_score),
            'primary_terms_score': round(primary_terms_score),
            'secondary_terms_score': round(secondary_terms_score),
            'topic_coverage_score': round(topic_coverage_score),
            'question_coverage_score': round(question_coverage_score)
        },
        'details': {
            'word_count': word_count,
            'keyword_info': {'keyword': keyword, 'count': keyword_count, 'density': round(actual_density * 100, 2)},
            'primary_term_analysis': primary_term_details,
            'secondary_term_analysis': secondary_term_details,
            'topic_coverage_analysis': topic_coverage_details,
            'question_coverage_analysis': question_coverage_details,
        }
    }
    
    logger.info(f"Content scoring complete for keyword '{keyword}'. Overall Score: {overall_score}")
    return score_data, "Success"

#==============================================================================
# End of Chunk 2/4
#==============================================================================
```

---

**Refactored Code - Chunk 3/4**

```python
#==============================================================================
# Start of Chunk 3/4
#==============================================================================

def highlight_keywords_in_content(content_html: str, term_data: Dict, keyword: str) -> Tuple[str, str]:
    """
    Highlight primary keyword, primary terms, and secondary terms in HTML content.
    Uses specific colors for each category.
    Returns: highlighted_html, status_message
    """
    if not term_data or not keyword:
        return content_html, "Missing term data or keyword for highlighting."

    try:
        # Use BeautifulSoup to parse and modify the HTML safely
        soup = BeautifulSoup(content_html, 'html.parser')

        # Define colors
        color_primary_keyword = "#FFEB9C" # Yellow
        color_primary_term = "#CDFFD8" # Green
        color_secondary_term = "#E6F3FF" # Blue

        # 1. Highlight Primary Keyword
        keyword_pattern = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)
        for text_node in soup.find_all(string=True):
             # Avoid highlighting within script/style tags etc.
            if text_node.parent.name in ['script', 'style', 'a', 'button']:
                 continue

            original_text = str(text_node)
            new_html = keyword_pattern.sub(
                lambda m: f'<span style="background-color: {color_primary_keyword};">{m.group(0)}</span>',
                original_text
            )
            if new_html != original_text:
                 # Replace the text node with the new parsed HTML
                 new_soup = BeautifulSoup(new_html, 'html.parser')
                 text_node.replace_with(new_soup)


        # Need to re-parse the soup after modifications before next highlighting step
        soup = BeautifulSoup(str(soup), 'html.parser')

        # 2. Highlight Primary Terms (excluding the main keyword)
        primary_terms = {info['term'].lower() for info in term_data.get('primary_terms', []) if info.get('term')} - {keyword.lower()}
        if primary_terms:
             primary_pattern = re.compile(r'\b(' + '|'.join(re.escape(term) for term in primary_terms) + r')\b', re.IGNORECASE)
             for text_node in soup.find_all(string=True):
                  if text_node.parent.name in ['script', 'style', 'a', 'button'] or text_node.parent.get('style'): # Avoid double highlighting
                      continue
                  original_text = str(text_node)
                  new_html = primary_pattern.sub(
                      lambda m: f'<span style="background-color: {color_primary_term};">{m.group(0)}</span>',
                      original_text
                  )
                  if new_html != original_text:
                      new_soup = BeautifulSoup(new_html, 'html.parser')
                      text_node.replace_with(new_soup)


        # Re-parse again
        soup = BeautifulSoup(str(soup), 'html.parser')

        # 3. Highlight Secondary Terms
        secondary_terms = {info['term'].lower() for info in term_data.get('secondary_terms', []) if info.get('term')}
        # Avoid highlighting terms already highlighted as primary
        secondary_terms -= primary_terms
        secondary_terms -= {keyword.lower()}

        if secondary_terms:
            secondary_pattern = re.compile(r'\b(' + '|'.join(re.escape(term) for term in secondary_terms) + r')\b', re.IGNORECASE)
            for text_node in soup.find_all(string=True):
                 if text_node.parent.name in ['script', 'style', 'a', 'button'] or text_node.parent.get('style'): # Avoid double highlighting
                      continue
                 original_text = str(text_node)
                 new_html = secondary_pattern.sub(
                     lambda m: f'<span style="background-color: {color_secondary_term};">{m.group(0)}</span>',
                     original_text
                 )
                 if new_html != original_text:
                      new_soup = BeautifulSoup(new_html, 'html.parser')
                      text_node.replace_with(new_soup)


        return str(soup), "Success"

    except Exception as e:
        error_msg = f"Exception during keyword highlighting: {e}"
        logger.error(error_msg, exc_info=True)
        # Return original content on error
        return content_html, f"Highlighting Failed: {e}"


def get_content_improvement_suggestions(content_html: str, term_data: Dict, score_data: Dict, keyword: str) -> Tuple[Optional[Dict], str]:
    """
    Generate detailed suggestions for improving content based on scoring results.
    Returns: suggestions dict or None, status message.
    """
    if not term_data or not score_data or not keyword:
         return None, "Missing term data, score data, or keyword for suggestions."

    suggestions = {
        'missing_primary_terms': [],
        'underused_primary_terms': [],
        'missing_secondary_terms': [], # Suggest only important ones
        'missing_topics': [],
        'partial_topics': [], # Topics needing expansion
        'unanswered_questions': [],
        'readability_suggestions': [],
        'structure_suggestions': []
    }

    content_text = clean_html(content_html)
    content_lower = content_text.lower()
    details = score_data.get('details', {})

    # 1. Term Usage Suggestions
    primary_analysis = details.get('primary_term_analysis', {})
    for term, info in primary_analysis.items():
        if info['count'] == 0:
            suggestions['missing_primary_terms'].append({
                'term': term,
                'importance': info.get('importance', 0),
                'recommended_usage': info.get('recommended', 1)
            })
        elif info['count'] < info.get('recommended', 1):
             suggestions['underused_primary_terms'].append({
                 'term': term,
                 'importance': info.get('importance', 0),
                 'current_usage': info['count'],
                 'recommended_usage': info.get('recommended', 1)
             })

    secondary_analysis = details.get('secondary_term_analysis', {})
    for term, info in secondary_analysis.items():
         # Only suggest missing secondary terms if they are reasonably important
         if info['count'] == 0 and info.get('importance', 0) >= 0.5:
             suggestions['missing_secondary_terms'].append({
                 'term': term,
                 'importance': info.get('importance', 0),
                 'recommended_usage': info.get('recommended', 1)
             })
             
    # Sort missing terms by importance
    suggestions['missing_primary_terms'].sort(key=lambda x: x['importance'], reverse=True)
    suggestions['missing_secondary_terms'].sort(key=lambda x: x['importance'], reverse=True)
    suggestions['underused_primary_terms'].sort(key=lambda x: x['importance'], reverse=True)


    # 2. Topic Coverage Suggestions
    topic_analysis = details.get('topic_coverage_analysis', {})
    for topic, info in topic_analysis.items():
        if not info.get('covered', False):
            suggestions['missing_topics'].append({
                'topic': topic,
                'description': info.get('description', '')
            })
        # Check match ratio even if considered 'covered' by basic check
        elif info.get('match_ratio', 0) < 0.6: # Threshold for needing expansion
             suggestions['partial_topics'].append({
                 'topic': topic,
                 'description': info.get('description', ''),
                 'match_ratio': info.get('match_ratio', 0),
                 'suggestion': f"Expand coverage for '{topic}'. It seems partially addressed but could be more comprehensive."
             })


    # 3. Question Answering Suggestions
    question_analysis = details.get('question_coverage_analysis', {})
    for question, info in question_analysis.items():
        if not info.get('answered', False):
            suggestions['unanswered_questions'].append(question)

    # 4. Readability & Structure Suggestions
    word_count = details.get('word_count', 0)
    if word_count < 500:
         suggestions['readability_suggestions'].append("Content is very short. Consider expanding significantly (aim for 1000+ words) for better depth and ranking potential.")
    elif word_count < 1000:
         suggestions['readability_suggestions'].append("Content length is moderate. Expanding towards 1200-1800 words could improve comprehensiveness.")

    # Basic structure checks (can be enhanced)
    try:
        soup = BeautifulSoup(content_html, 'html.parser')
        h2_count = len(soup.find_all('h2'))
        h3_count = len(soup.find_all('h3'))
        list_count = len(soup.find_all(['ul', 'ol']))
        p_tags = soup.find_all('p')
        long_paragraphs = sum(1 for p in p_tags if count_words(p.get_text()) > 150) # Shorter threshold

        if h2_count < 3 and word_count > 800:
             suggestions['structure_suggestions'].append("Add more H2 headings to break up content into logical sections.")
        if h3_count < h2_count and h2_count > 0: # If H2s exist but few H3s
             suggestions['structure_suggestions'].append("Consider adding H3 subheadings under relevant H2 sections for better granularity.")
        if list_count < 2 and word_count > 800:
             suggestions['structure_suggestions'].append("Incorporate bulleted or numbered lists to improve scannability, especially for steps or key points.")
        if long_paragraphs > 1:
             suggestions['structure_suggestions'].append(f"Break up {long_paragraphs} long paragraph(s) (over 150 words) into shorter ones (ideally under 100 words).")

    except Exception as e:
        logger.warning(f"Could not perform structure analysis due to parsing error: {e}")
        suggestions['structure_suggestions'].append("Could not automatically analyze structure. Manually review heading usage, paragraph length, and list usage.")


    return suggestions, "Success"


def create_content_scoring_brief(keyword: str, term_data: Dict, score_data: Dict, suggestions: Dict) -> Tuple[Optional[BytesIO], str]:
    """
    Create a Word document summarizing the content score and improvement suggestions.
    Returns: BytesIO stream or None, status message.
    """
    if not term_data or not score_data or not suggestions or not keyword:
        return None, "Missing data required for scoring brief."

    try:
        doc = Document()
        # Basic styles
        # doc.styles['Normal'].font.name = 'Calibri'
        # doc.styles['Normal'].font.size = Pt(11)

        # --- Header ---
        doc.add_heading(f'Content Optimization Brief: {keyword}', 0)
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        doc.add_paragraph() # Spacing

        # --- Overall Score ---
        doc.add_heading('Content Score Summary', level=1)
        score_para = doc.add_paragraph()
        score_para.add_run("Overall Score: ").bold = True
        overall_score = score_data.get('overall_score', 0)
        grade = score_data.get('grade', 'F')
        score_run = score_para.add_run(f"{overall_score} ({grade})")
        score_run.font.bold = True
        score_run.font.size = Pt(14)
        # Apply color based on score
        if overall_score >= 70: score_run.font.color.rgb = COLOR_ADDED
        elif overall_score < 50: score_run.font.color.rgb = COLOR_DELETED
        else: score_run.font.color.rgb = RGBColor(255, 165, 0) # Orange

        # Component Scores Table
        components = score_data.get('components', {})
        if components:
             doc.add_paragraph("Score Breakdown:")
             table = doc.add_table(rows=1, cols=2)
             table.style = 'Table Grid'
             table.autofit = False
             table.columns[0].width = Inches(2.5)
             table.columns[1].width = Inches(1.0)

             hdr_cells = table.rows[0].cells
             hdr_cells[0].text = 'Component'
             hdr_cells[1].text = 'Score'
             hdr_cells[0].paragraphs[0].runs[0].bold = True
             hdr_cells[1].paragraphs[0].runs[0].bold = True

             for component, score in components.items():
                 row_cells = table.add_row().cells
                 row_cells[0].text = component.replace('_score', '').replace('_', ' ').title()
                 row_cells[1].text = str(score)
             doc.add_paragraph() # Spacing


        # --- Key Improvement Areas ---
        doc.add_heading('Key Improvement Areas', level=1)

        # Missing Primary Terms
        if suggestions.get('missing_primary_terms'):
            doc.add_heading('Add These Primary Terms', level=2)
            for term_info in suggestions['missing_primary_terms']:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(f"{term_info['term']} ").bold = True
                p.add_run(f"(Importance: {term_info['importance']:.2f}, Recommended: {term_info['recommended_usage']})")

        # Underused Primary Terms
        if suggestions.get('underused_primary_terms'):
             doc.add_heading('Increase Usage of These Primary Terms', level=2)
             for term_info in suggestions['underused_primary_terms']:
                  p = doc.add_paragraph(style='List Bullet')
                  p.add_run(f"{term_info['term']} ").bold = True
                  p.add_run(f"(Current: {term_info['current_usage']}, Recommended: {term_info['recommended_usage']})")

        # Missing Secondary Terms (Important ones)
        if suggestions.get('missing_secondary_terms'):
            doc.add_heading('Consider Adding These Secondary Terms', level=2)
            for term_info in suggestions['missing_secondary_terms'][:10]: # Limit suggestions
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(f"{term_info['term']} ").bold = True
                p.add_run(f"(Importance: {term_info['importance']:.2f})")


        # Missing Topics
        if suggestions.get('missing_topics'):
            doc.add_heading('Address These Content Gaps (Missing Topics)', level=2)
            for topic in suggestions['missing_topics']:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(f"{topic.get('topic', '')}").bold = True
                if topic.get('description'):
                     p.add_run(f": {topic.get('description', '')}")

        # Partial Topics
        if suggestions.get('partial_topics'):
             doc.add_heading('Expand on These Partially Covered Topics', level=2)
             for topic in suggestions['partial_topics']:
                  p = doc.add_paragraph(style='List Bullet')
                  p.add_run(f"{topic.get('topic', '')} ").bold = True
                  p.add_run(f"(Current coverage: {topic.get('match_ratio', 0)*100:.0f}%)")
                  if topic.get('suggestion'):
                       p.add_run(f" - Suggestion: {topic.get('suggestion', '')}")


        # Unanswered Questions
        if suggestions.get('unanswered_questions'):
            doc.add_heading('Answer These Questions', level=2)
            for question in suggestions['unanswered_questions']:
                doc.add_paragraph(question, style='List Bullet')

        # Structure & Readability
        structure_suggestions = suggestions.get('structure_suggestions', [])
        readability_suggestions = suggestions.get('readability_suggestions', [])
        if structure_suggestions or readability_suggestions:
            doc.add_heading('Improve Structure & Readability', level=2)
            for suggestion in structure_suggestions:
                doc.add_paragraph(suggestion, style='List Bullet')
            for suggestion in readability_suggestions:
                doc.add_paragraph(suggestion, style='List Bullet')
        
        # --- Detailed Term Analysis (Optional Appendix?) ---
        # Consider adding the full term usage tables here if needed,
        # similar to the original `create_word_document` function's term tables.
        # For brevity in the main brief, the key areas above might suffice.

        # Save document to memory stream
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        return doc_stream, "Success"

    except Exception as e:
        error_msg = f"Exception creating content scoring brief: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg

###############################################################################
# 6. Meta Title and Description Generation
###############################################################################

def generate_meta_tags(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], term_data: Dict, 
                      anthropic_api_key: str) -> Tuple[Optional[str], Optional[str], str]:
    """
    Generate optimized meta title and description using Anthropic.
    Returns: meta_title, meta_description, status_message
    """
    client = get_anthropic_client(anthropic_api_key)
    if not client:
        return None, None, "Anthropic client initialization failed"
        
    # Context for generation
    h1 = semantic_structure.get('h1', f"Guide to {keyword}")
    top_related_kws = [kw.get('keyword', '') for kw in related_keywords[:5] if kw.get('keyword')]
    primary_terms = [t.get('term') for t in term_data.get('primary_terms', [])[:5] if t.get('term')]
    # Combine related and primary for more context
    key_terms = list(set(primary_terms + top_related_kws))[:8] # Limit combined terms

    prompt = f"""
    Generate an SEO-optimized meta title and meta description for an article about "{keyword}".

    Context:
    - Main Keyword: {keyword}
    - Article H1 Title: {h1}
    - Key Terms/Concepts: {', '.join(key_terms)}
    
    Guidelines:
    - Meta Title: 50-60 characters. Include "{keyword}" near the beginning. Be compelling and clear.
    - Meta Description: 150-160 characters. Include "{keyword}" and 1-2 other key terms naturally. Summarize the article's value and include a subtle call-to-action (e.g., "Learn more", "Discover how", "Explore...", "Find out...").
    - Tone: Informative, engaging, and trustworthy. Avoid hype or clickbait.

    Output Format:
    Return ONLY a valid JSON object like this:
    {{
        "meta_title": "Your optimized meta title here",
        "meta_description": "Your optimized meta description here"
    }}
    No extra text before or after the JSON.
    """

    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=200,
            system="You are an expert SEO copywriter specializing in crafting compelling meta tags that improve click-through rates. You provide only valid JSON output.",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.6 # Slightly more creative for meta tags
        )

        response_text = response.content[0].text
        meta_data = safe_json_loads(response_text)

        if meta_data and "meta_title" in meta_data and "meta_description" in meta_data:
             meta_title = meta_data['meta_title']
             meta_description = meta_data['meta_description']

             # Basic length enforcement (conservative)
             if len(meta_title) > 60: meta_title = meta_title[:57] + "..."
             if len(meta_description) > 160: meta_description = meta_description[:157] + "..."

             logger.info(f"Successfully generated meta tags for '{keyword}'.")
             return meta_title, meta_description, "Success"
        else:
             error_msg = "Failed to parse valid meta tags JSON from LLM response."
             logger.error(f"{error_msg} Response was: {response_text[:500]}...")
             # Provide defaults on failure
             default_title = f"{keyword.title()} | Comprehensive Guide & Tips"
             default_desc = f"Explore the complete guide to {keyword}. Learn key concepts, expert tips, and best practices. Find out more today!"
             return default_title[:60], default_desc[:160], error_msg

    except anthropic.APIError as api_err:
        error_msg = f"Anthropic API Error generating meta tags: {api_err}"
        logger.error(error_msg)
        return None, None, f"API Error: {api_err}"
    except Exception as e:
        error_msg = f"Unexpected error generating meta tags: {e}"
        logger.error(error_msg, exc_info=True)
        return None, None, error_msg

#==============================================================================
# End of Chunk 3/4
#==============================================================================
```

---

**Refactored Code - Chunk 4/4**

```python
#==============================================================================
# Start of Chunk 4/4
#==============================================================================

###############################################################################
# 7. Content Generation (New Article)
###############################################################################

def generate_article_section(client: anthropic.Anthropic, keyword: str, heading_level: str, heading_text: str, 
                             previous_section_summary: str, terms_to_include: Dict, 
                             competitor_context: str, target_word_count: int) -> Tuple[str, str]:
    """Generates content for a single article section using Anthropic."""
    
    primary_terms_str = "\n".join([f"- {t['term']} (Usage: {t['recommended_usage']})" for t in terms_to_include.get('primary', [])])
    secondary_terms_str = "\n".join([f"- {t['term']}" for t in terms_to_include.get('secondary', [])])
    topics_str = "\n".join([f"- {t['topic']}: {t.get('description','')}" for t in terms_to_include.get('topics', [])])
    questions_str = "\n".join([f"- {q}" for q in terms_to_include.get('questions', [])])
    
    prompt = f"""
    Write a concise, informative, and engaging content section for an article about "{keyword}".

    Section Details:
    - Heading Level: {heading_level.upper()}
    - Heading Text: {heading_text}
    
    Context & Requirements:
    1.  **Previous Section Summary:** "{previous_section_summary}" (Ensure smooth transition from this).
    2.  **Target Word Count:** Approximately {target_word_count} words for this section. Be concise but thorough.
    3.  **Content Source:** Draw information and insights primarily from the relevant competitor context provided below. Synthesize, do not copy.
    4.  **Flow:** Ensure this section logically follows the previous one and sets up the next (if applicable). Use transition phrases.
    5.  **Keyword Focus:** Maintain relevance to the main keyword "{keyword}".
    6.  **Term Inclusion:** Naturally integrate the following terms where relevant:
        *Primary Terms to prioritize:*
        {primary_terms_str if primary_terms_str else "N/A"}
        *Secondary Terms to include if possible:*
        {secondary_terms_str if secondary_terms_str else "N/A"}
    7.  **Topic Coverage:** Address aspects related to these key topics if they fit this section:
        {topics_str if topics_str else "N/A"}
    8.  **Answer Questions:** Address these questions if relevant to this section:
        {questions_str if questions_str else "N/A"}
    9.  **Style:** Write clearly with short paragraphs (2-4 sentences). Use bullet points for lists. Avoid overly complex language or jargon unless necessary and explained.
    10. **Output Format:** Return ONLY the HTML content for this section, starting directly with paragraph (<p>) or list (<ul>/<ol>) tags. Do NOT include the heading tag itself.

    Relevant Competitor Context for "{heading_text}":
    ```
    {competitor_context if competitor_context else "No specific competitor context provided for this section. Write based on general knowledge and the requirements."}
    ```
    """
    
    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=int(target_word_count * 2.5), # Allow ample tokens for generation + overhead
            system="You are an expert SEO Content Writer creating well-structured, informative article sections based on competitor analysis and specific instructions. You follow formatting requirements precisely.",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4 # Balanced temperature
        )
        
        section_content = response.content[0].text.strip()
        
        # Basic cleanup: remove potential leading/trailing ```
        section_content = re.sub(r'^```html\s*', '', section_content)
        section_content = re.sub(r'\s*```$', '', section_content)

        # Ensure it starts with a valid block tag
        if not section_content.startswith(('<p>', '<ul>', '<ol>')):
             section_content = f"<p>{section_content}</p>" # Wrap if needed

        return section_content, "Success"

    except anthropic.APIError as api_err:
        error_msg = f"Anthropic API Error generating section '{heading_text}': {api_err}"
        logger.error(error_msg)
        return f"<p><i>Error generating content for this section: {api_err}</i></p>", f"API Error: {api_err}"
    except Exception as e:
        error_msg = f"Unexpected error generating section '{heading_text}': {e}"
        logger.error(error_msg, exc_info=True)
        return f"<p><i>Error generating content for this section. Please try again.</i></p>", error_msg


def find_relevant_competitor_context(heading_text: str, competitor_embeddings: List[Dict], openai_api_key: str, top_n: int = 2) -> str:
    """
    Finds relevant text snippets from competitor content for a given heading using embeddings.
    
    Args:
        heading_text: The heading of the section to generate content for.
        competitor_embeddings: List of dicts, each containing 'url', 'embedding', 'content'.
        openai_api_key: OpenAI API key.
        top_n: Number of top competitor snippets to return.

    Returns:
        A string containing concatenated relevant text snippets.
    """
    if not competitor_embeddings or not heading_text:
        return ""

    # Generate embedding for the heading
    heading_embedding, status = generate_embedding(heading_text, openai_api_key)
    if not heading_embedding:
        logger.warning(f"Could not generate embedding for heading '{heading_text}': {status}")
        return ""

    # Calculate similarity between heading and competitor content
    similarities = []
    for comp in competitor_embeddings:
        comp_embedding = comp.get('embedding')
        if comp_embedding and len(heading_embedding) == len(comp_embedding):
            try:
                # Cosine similarity
                similarity = np.dot(heading_embedding, comp_embedding) / (np.linalg.norm(heading_embedding) * np.linalg.norm(comp_embedding))
                similarities.append({'url': comp['url'], 'content': comp['content'], 'score': similarity})
            except Exception as e:
                 logger.warning(f"Error calculating similarity for {comp.get('url', 'unknown')}: {e}")
        elif comp_embedding:
             logger.warning(f"Dimension mismatch: Heading ({len(heading_embedding)}) vs Competitor ({len(comp_embedding)}) for {comp.get('url', 'unknown')}")


    # Sort by similarity and get top N
    similarities.sort(key=lambda x: x['score'], reverse=True)
    
    relevant_context = ""
    for item in similarities[:top_n]:
        # Extract a relevant snippet (e.g., first few paragraphs)
        content_text = clean_html(item['content'])
        snippet = "\n\n".join(content_text.split('\n\n')[:3]) # First 3 paragraphs approx
        relevant_context += f"--- Context from {item['url']} (Similarity: {item['score']:.2f}) ---\n{snippet}\n\n"

    return relevant_context.strip()


def generate_full_article(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], 
                         paa_questions: List[Dict], term_data: Dict, 
                         competitor_embeddings: List[Dict], # Added competitor embeddings
                         anthropic_api_key: str, openai_api_key: str # Added openai key for context finding
                         ) -> Tuple[Optional[str], str]:
    """
    Generates a full article sequentially, section by section, using competitor context.
    Returns: Full HTML article content or None, status message.
    """
    client = get_anthropic_client(anthropic_api_key)
    if not client:
        return None, "Anthropic client initialization failed"
        
    if not semantic_structure or not term_data:
        return None, "Missing semantic structure or term data for article generation."

    full_article_html = []
    h1 = semantic_structure.get('h1', f"Guide to {keyword}")
    full_article_html.append(f"<h1>{h1}</h1>")
    
    last_section_summary = f"Introduction section focusing on '{h1}'." # Initial context

    # Prepare term/topic/question context once
    terms_to_include = {
        'primary': term_data.get('primary_terms', []),
        'secondary': term_data.get('secondary_terms', []),
        'topics': term_data.get('topics', []),
        'questions': term_data.get('questions', []) + [q.get('question') for q in paa_questions if q.get('question')]
    }
    
    # Estimate word counts per section
    num_h2 = len(semantic_structure.get('sections', []))
    num_h3_total = sum(len(sec.get('subsections', [])) for sec in semantic_structure.get('sections', []))
    total_sections = 1 + num_h2 + num_h3_total # H1 + H2s + H3s (+ Conclusion)
    
    # Allocate words, reserving some for intro/conclusion
    allocatable_words = ARTICLE_TARGET_WORD_COUNT * 0.9
    words_per_section_avg = allocatable_words / max(1, total_sections - 1) # Exclude H1/Conclusion approx


    # Generate Intro (using H1 context) - Special call? Or assume H1 implies intro
    intro_target_words = int(words_per_section_avg * 0.75) # Shorter intro
    intro_context = find_relevant_competitor_context(h1, competitor_embeddings, openai_api_key, top_n=1)
    intro_content, status = generate_article_section(client, keyword, "Introduction", h1, "Start of the article.", terms_to_include, intro_context, intro_target_words)
    if status == "Success":
         full_article_html.append(intro_content)
         last_section_summary = clean_html(intro_content)[:150] + "..." # Update summary
    else:
         logger.error(f"Failed to generate introduction: {status}")
         full_article_html.append("<p><i>Error generating introduction.</i></p>")


    # Generate H2 sections and their H3 subsections sequentially
    for section in semantic_structure.get('sections', []):
        h2_text = section.get('h2')
        if not h2_text: continue
        
        full_article_html.append(f"<h2>{h2_text}</h2>")
        
        # Find relevant competitor context for this H2
        competitor_context_h2 = find_relevant_competitor_context(h2_text, competitor_embeddings, openai_api_key)
        
        # Determine word count for H2 (consider if it has H3s)
        num_h3_in_section = len(section.get('subsections', []))
        h2_target_words = int(words_per_section_avg * (1.0 if num_h3_in_section == 0 else 0.5)) # More words if no H3s

        h2_content, status = generate_article_section(client, keyword, "H2", h2_text, last_section_summary, terms_to_include, competitor_context_h2, h2_target_words)
        if status == "Success":
             full_article_html.append(h2_content)
             last_section_summary = f"Section '{h2_text}'. " + clean_html(h2_content)[:150] + "..."
        else:
             logger.error(f"Failed to generate H2 section '{h2_text}': {status}")
             full_article_html.append(f"<p><i>Error generating content for {h2_text}.</i></p>")
             last_section_summary = f"Section '{h2_text}' (Error generating content)."


        # Generate H3 subsections
        h3_target_words = int(words_per_section_avg * 0.7 / max(1, num_h3_in_section)) if num_h3_in_section > 0 else 0
        
        for subsection in section.get('subsections', []):
            h3_text = subsection.get('h3')
            if not h3_text: continue
            
            full_article_html.append(f"<h3>{h3_text}</h3>")
            
            competitor_context_h3 = find_relevant_competitor_context(h3_text, competitor_embeddings, openai_api_key, top_n=1) # Maybe only top 1 for H3
            
            h3_content, status = generate_article_section(client, keyword, "H3", h3_text, last_section_summary, terms_to_include, competitor_context_h3, h3_target_words)
            if status == "Success":
                 full_article_html.append(h3_content)
                 # Update summary based on H3, but keep H2 context
                 last_section_summary = f"Subsection '{h3_text}' under '{h2_text}'. " + clean_html(h3_content)[:100] + "..."
            else:
                 logger.error(f"Failed to generate H3 subsection '{h3_text}': {status}")
                 full_article_html.append(f"<p><i>Error generating content for {h3_text}.</i></p>")
                 last_section_summary = f"Subsection '{h3_text}' under '{h2_text}' (Error generating content)."


    # Generate Conclusion
    conclusion_target_words = int(words_per_section_avg * 0.75) # Shorter conclusion
    conclusion_context = find_relevant_competitor_context(f"Conclusion for {keyword}", competitor_embeddings, openai_api_key, top_n=1)
    full_article_html.append("<h2>Conclusion</h2>")
    conclusion_content, status = generate_article_section(client, keyword, "Conclusion", "Conclusion", last_section_summary, terms_to_include, conclusion_context, conclusion_target_words)
    if status == "Success":
         full_article_html.append(conclusion_content)
    else:
         logger.error(f"Failed to generate conclusion: {status}")
         full_article_html.append("<p><i>Error generating conclusion.</i></p>")

    final_html = "\n".join(full_article_html)
    final_word_count = count_words(clean_html(final_html))
    logger.info(f"Generated article for '{keyword}' with {final_word_count} words.")
    
    return final_html, "Success"


###############################################################################
# 8. Internal Linking (Using Embeddings)
###############################################################################

# (Keep existing internal linking functions: parse_site_pages_spreadsheet, 
# embed_site_pages, generate_internal_links_with_embeddings - they seem okay,
# ensure generate_internal_links uses the correct embedding model based on detected dim)

def parse_site_pages_spreadsheet(uploaded_file) -> Tuple[Optional[List[Dict]], str]:
    """Parse uploaded CSV/Excel with site pages."""
    try:
        filename = uploaded_file.name
        if filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
        else:
            return None, f"Unsupported file type: {filename}. Please use CSV or Excel."

        required_columns = ['URL', 'Title', 'Meta Description']
        if not all(col in df.columns for col in required_columns):
             missing = [col for col in required_columns if col not in df.columns]
             return None, f"Missing required columns in spreadsheet: {', '.join(missing)}"
        
        # Basic cleaning and validation
        df = df[required_columns].dropna(subset=['URL', 'Title']) # Require URL and Title
        df['URL'] = df['URL'].astype(str).str.strip()
        df['Title'] = df['Title'].astype(str).str.strip()
        df['Meta Description'] = df['Meta Description'].astype(str).str.strip().fillna('') # Fill NaN descriptions

        # Filter out invalid URLs (basic check)
        df = df[df['URL'].str.startswith(('http://', 'https://'))]

        if df.empty:
             return [], "No valid pages found in the spreadsheet after cleaning."

        pages = df.to_dict('records')
        logger.info(f"Successfully parsed {len(pages)} pages from spreadsheet.")
        return pages, "Success"

    except Exception as e:
        error_msg = f"Failed to parse site pages spreadsheet: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg

def embed_site_pages(pages: List[Dict], openai_api_key: str, batch_size: int = INTERNAL_LINK_BATCH_SIZE) -> Tuple[Optional[List[Dict]], str]:
    """Generate embeddings for site pages using OpenAI, processing in batches."""
    if not pages:
        return [], "No pages provided for embedding."
    if not openai_api_key:
        return None, "OpenAI API key is missing."

    client = openai.OpenAI(api_key=openai_api_key)
    
    texts_to_embed = []
    page_indices = [] # Track original index for mapping back
    for i, page in enumerate(pages):
         # Combine key fields for semantic meaning
         text = f"Title: {page.get('title', '')}\nURL: {page.get('url', '')}\nDescription: {page.get('description', '')}"
         texts_to_embed.append(text)
         page_indices.append(i)

    all_embeddings = [None] * len(pages) # Initialize list to store embeddings
    model_used = EMBEDDING_MODEL_OPENAI
    
    logger.info(f"Generating embeddings for {len(texts_to_embed)} pages using {model_used} in batches of {batch_size}...")
    
    try:
        for i in range(0, len(texts_to_embed), batch_size):
            batch_texts = texts_to_embed[i : i + batch_size]
            batch_indices = page_indices[i : i + batch_size]
            
            logger.debug(f"Processing batch {i//batch_size + 1}...")
            response = client.embeddings.create(
                model=model_used,
                input=batch_texts
            )

            if response.data:
                 for embedding_data, original_index in zip(response.data, batch_indices):
                     all_embeddings[original_index] = embedding_data.embedding
            else:
                 logger.warning(f"No embeddings returned for batch starting at index {i}.")

            time.sleep(0.1) # Small delay between batches

        # Add embeddings back to the page data
        pages_with_embeddings = []
        successful_embeddings = 0
        for i, page in enumerate(pages):
            page_copy = page.copy()
            embedding = all_embeddings[i]
            if embedding:
                 page_copy['embedding'] = embedding
                 successful_embeddings += 1
            else:
                 page_copy['embedding'] = None # Ensure key exists but is None if failed
                 logger.warning(f"Failed to generate embedding for page: {page.get('url')}")
            pages_with_embeddings.append(page_copy)

        if successful_embeddings == 0:
             return None, "Embedding generation failed for all pages."
             
        logger.info(f"Successfully generated embeddings for {successful_embeddings}/{len(pages)} pages.")
        return pages_with_embeddings, "Success"

    except openai.APIError as api_err:
        error_msg = f"OpenAI API Error embedding site pages: {api_err}"
        logger.error(error_msg)
        return None, f"API Error: {api_err}"
    except Exception as e:
        error_msg = f"Unexpected error embedding site pages: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg


def generate_internal_links_with_embeddings(article_html: str, pages_with_embeddings: List[Dict], 
                                           openai_api_key: str, # Keep openai key for consistency
                                           word_count: int) -> Tuple[str, List[Dict], str]:
    """
    Identifies internal linking opportunities using semantic similarity between article paragraphs and site pages.
    Uses simple keyword matching for anchor text selection to avoid extra LLM calls.
    Returns: article HTML with links, list of added links, status message.
    """
    if not article_html or not pages_with_embeddings:
        return article_html, [], "Missing article content or site pages for linking."
    if not openai_api_key:
        return article_html, [], "OpenAI API key needed for paragraph embeddings."
        
    client = openai.OpenAI(api_key=openai_api_key)

    # Calculate target number of links
    max_links = min(15, max(INTERNAL_LINK_MIN_COUNT, int(word_count / 1000) * INTERNAL_LINK_MAX_COUNT_FACTOR))
    logger.info(f"Aiming for up to {max_links} internal links for content with {word_count} words.")

    # 1. Extract Paragraphs from Article HTML
    soup = BeautifulSoup(article_html, 'html.parser')
    paragraphs = []
    for p_tag in soup.find_all('p'):
        para_text = p_tag.get_text(strip=True)
        if count_words(para_text) > 15: # Only consider paragraphs with substance
            paragraphs.append({
                'text': para_text,
                'html_tag': p_tag # Keep reference to the tag object
            })
            
    if not paragraphs:
        logger.warning("No suitable paragraphs found in the article for internal linking.")
        return article_html, [], "Success (No paragraphs found)"

    # 2. Generate Embeddings for Paragraphs
    paragraph_texts = [p['text'] for p in paragraphs]
    paragraph_embeddings = []
    
    # Determine embedding model based on page embeddings
    first_page_embedding = next((p['embedding'] for p in pages_with_embeddings if p.get('embedding')), None)
    if not first_page_embedding:
         return article_html, [], "Error: No valid embeddings found in site pages data."
         
    embedding_dim = len(first_page_embedding)
    if embedding_dim == EMBEDDING_DIM_LARGE:
        para_embedding_model = "text-embedding-3-large"
    elif embedding_dim == EMBEDDING_DIM_SMALL:
        para_embedding_model = "text-embedding-3-small"
    else:
        # Fallback or error if dimension is unexpected
        logger.warning(f"Unexpected page embedding dimension: {embedding_dim}. Defaulting to large model for paragraphs.")
        para_embedding_model = "text-embedding-3-large"

    logger.info(f"Generating paragraph embeddings using {para_embedding_model}...")
    try:
         # Batch paragraph embedding generation
         batch_size = INTERNAL_LINK_BATCH_SIZE * 2 # Larger batch for paragraphs
         for i in range(0, len(paragraph_texts), batch_size):
             batch_texts = paragraph_texts[i : i + batch_size]
             response = client.embeddings.create(model=para_embedding_model, input=batch_texts)
             if response.data:
                 paragraph_embeddings.extend([item.embedding for item in response.data])
             else:
                 # Add placeholders for failed batch
                 paragraph_embeddings.extend([None] * len(batch_texts))
             time.sleep(0.05) # Small delay


         if len(paragraph_embeddings) != len(paragraphs):
              raise ValueError("Mismatch between number of paragraphs and generated embeddings.")

         # Add embeddings to paragraph data
         for i, p in enumerate(paragraphs):
             p['embedding'] = paragraph_embeddings[i]

    except Exception as e:
        error_msg = f"Failed to generate paragraph embeddings: {e}"
        logger.error(error_msg, exc_info=True)
        return article_html, [], f"Error: {error_msg}"
        
    # Filter out paragraphs where embedding failed
    paragraphs_with_embeddings = [p for p in paragraphs if p.get('embedding')]
    if not paragraphs_with_embeddings:
         return article_html, [], "Error: Embedding generation failed for all paragraphs."

    # 3. Find Best Page Match for Each Paragraph
    potential_links = []
    valid_pages = [p for p in pages_with_embeddings if p.get('embedding')]

    for para_idx, paragraph in enumerate(paragraphs_with_embeddings):
        para_embedding = paragraph['embedding']
        best_match = {'score': -1, 'page': None, 'para_index': para_idx}

        for page in valid_pages:
            page_embedding = page['embedding']
            # Ensure dimensions match before calculating similarity
            if len(para_embedding) == len(page_embedding):
                 try:
                     similarity = np.dot(para_embedding, page_embedding) / (np.linalg.norm(para_embedding) * np.linalg.norm(page_embedding))
                     if similarity > best_match['score']:
                          best_match['score'] = similarity
                          best_match['page'] = page
                 except Exception as e:
                      logger.debug(f"Similarity calculation error: {e}")
            # else: logger.debug("Skipping page due to embedding dimension mismatch.")


        if best_match['page'] and best_match['score'] >= INTERNAL_LINK_MIN_SIMILARITY:
            potential_links.append({
                'paragraph_index': best_match['para_index'], # Index within paragraphs_with_embeddings
                'paragraph_object': paragraph, # Reference to original paragraph dict
                'page_url': best_match['page']['url'],
                'page_title': best_match['page']['title'],
                'similarity_score': best_match['score']
            })

    # Sort potential links by score (highest similarity first)
    potential_links.sort(key=lambda x: x['similarity_score'], reverse=True)

    # 4. Select Links and Identify Anchor Text
    links_added = []
    used_paragraphs = set() # Track index in original paragraphs list
    used_pages = set()
    
    # Helper to find paragraph index in the original list
    original_para_indices = {id(p['html_tag']): i for i, p in enumerate(paragraphs)}

    for link_info in potential_links:
        if len(links_added) >= max_links:
            break

        paragraph_obj = link_info['paragraph_object']
        para_text = paragraph_obj['text']
        page_url = link_info['page_url']
        page_title = link_info['page_title']
        
        # Find original index using the tag object's ID
        original_para_idx = original_para_indices.get(id(paragraph_obj['html_tag']))
        
        if original_para_idx is None or original_para_idx in used_paragraphs or page_url in used_pages:
            continue # Skip if paragraph already used, page linked, or index not found

        # Simple Anchor Text Strategy: Match page title keywords
        title_keywords = set(re.findall(r'\b\w{4,}\b', page_title.lower())) # Keywords >= 4 chars
        
        best_anchor = ""
        max_matched_keywords = 0

        # Try finding multi-word phrases containing title keywords
        sentences = re.split(r'[.?!]\s+', para_text) # Split into sentences
        for sentence in sentences:
             words = sentence.split()
             for i in range(len(words)):
                  for j in range(i + 2, min(i + 7, len(words) + 1)): # Phrases of 2-6 words
                      phrase = " ".join(words[i:j])
                      # Clean phrase for matching (remove punctuation at ends)
                      clean_phrase = re.sub(r"^[^\w\s]+|[^\w\s]+$", "", phrase)
                      phrase_keywords = set(re.findall(r'\b\w{4,}\b', clean_phrase.lower()))
                      matched_keywords = len(phrase_keywords.intersection(title_keywords))

                      if matched_keywords > max_matched_keywords:
                           # Basic check: avoid anchors ending in common prepositions/articles
                           if not re.search(r'\b(in|on|at|to|for|a|an|the)$', clean_phrase.lower()):
                               best_anchor = clean_phrase
                               max_matched_keywords = matched_keywords
                      # Prefer anchors with more matching keywords, then shorter anchors
                      elif matched_keywords == max_matched_keywords and matched_keywords > 0:
                           if len(clean_phrase) < len(best_anchor):
                                best_anchor = clean_phrase


        # Fallback: use the longest matching title keyword if no phrase found
        if not best_anchor and title_keywords:
             found_kws_in_para = title_keywords.intersection(set(re.findall(r'\b\w{4,}\b', para_text.lower())))
             if found_kws_in_para:
                  best_anchor = max(found_kws_in_para, key=len)


        # Final check: Anchor must exist in the paragraph text
        if best_anchor and re.search(r'\b' + re.escape(best_anchor) + r'\b', para_text, re.IGNORECASE):
            # Add the link
            link_info['anchor_text'] = best_anchor
            link_info['original_paragraph_index'] = original_para_idx # Store original index
            links_added.append(link_info)
            used_paragraphs.add(original_para_idx)
            used_pages.add(page_url)
        else:
             logger.debug(f"Could not find suitable anchor text for page '{page_title}' in paragraph {original_para_idx}.")


    # 5. Apply Links to Article HTML (using the stored tag objects)
    if not links_added:
        logger.info("No internal links added based on similarity and anchor text criteria.")
        return article_html, [], "Success (No links added)"

    logger.info(f"Attempting to add {len(links_added)} internal links.")
    
    modified_count = 0
    # We modify the soup *in place* by accessing the original paragraph tags
    for link in links_added:
        p_tag = link['paragraph_object']['html_tag']
        anchor = link['anchor_text']
        url = link['page_url']
        
        # Use regex to replace only the first occurrence of the anchor within the tag's text content
        try:
             # Get current text content of the paragraph tag
             current_text = p_tag.decode_contents()
             
             # Case-insensitive replacement of the first match
             pattern = re.compile(r'(\b' + re.escape(anchor) + r'\b)', re.IGNORECASE)
             new_html_content, num_subs = pattern.subn(f'<a href="{url}" title="{link["page_title"]}">{r"\1"}</a>', current_text, count=1)
             
             if num_subs > 0:
                 # Clear the original tag content and append the modified HTML
                 p_tag.clear()
                 # Parse the new content and append it (handles nested tags)
                 p_tag.append(BeautifulSoup(new_html_content, 'html.parser'))
                 modified_count += 1
                 # Add context for reporting
                 context_match = re.search(r'(\b' + re.escape(anchor) + r'\b)', link['paragraph_object']['text'], re.IGNORECASE)
                 if context_match:
                     start = max(0, context_match.start() - 30)
                     end = min(len(link['paragraph_object']['text']), context_match.end() + 30)
                     context = f"...{link['paragraph_object']['text'][start:context_match.start()]}<mark>[{context_match.group(0)}]</mark>{link['paragraph_object']['text'][context_match.end():end]}..."
                     link['context'] = context
                 else: link['context'] = "Context unavailable"

             else:
                 logger.warning(f"Could not find anchor text '{anchor}' in paragraph tag for URL {url}. Skipping link.")

        except Exception as e:
             logger.error(f"Error applying link for anchor '{anchor}' to URL {url}: {e}", exc_info=True)


    logger.info(f"Successfully added {modified_count} internal links.")

    # Prepare output list
    links_output = [{
        "url": link['page_url'],
        "anchor_text": link['anchor_text'],
        "context": link.get('context', 'N/A'),
        "page_title": link['page_title'],
        "similarity_score": round(link['similarity_score'], 2)
    } for link in links_added if 'context' in link] # Only return successfully applied links

    return str(soup), links_output, "Success"


###############################################################################
# 9. Document Generation (Word DOCX)
###############################################################################

def add_html_to_docx(doc: Document, html_content: str):
    """
    Parses basic HTML (p, h1-h6, ul, ol, li, strong, em, ins, del, a) 
    and adds it to the python-docx Document object.
    Handles color-coding for <ins> and <del> tags.
    """
    if not html_content:
        return

    # Preprocess: Ensure block tags are separated by newlines for easier parsing
    html_content = re.sub(r'</(h[1-6]|p|li|ul|ol)>\s*<', r'</\1>\n<', html_content)
    
    soup = BeautifulSoup(html_content, 'html.parser')

    # Function to add run with potential formatting
    def add_run_with_format(paragraph, text, bold=False, italic=False, color=None, strike=False, link_url=None):
         run = paragraph.add_run(text)
         run.bold = bold
         run.italic = italic
         if color:
             run.font.color.rgb = color
         if strike:
             run.font.strike = True
         # Add hyperlink if URL is provided
         if link_url:
             # Create hyperlink element
             hyperlink = OxmlElement('w:hyperlink')
             hyperlink.set(qn('r:id'), doc.part.relate_to(link_url, openai.RELATIONSHIP_TYPE.HYPERLINK, is_external=True))
             
             # Create the run properties element for styling the link
             run_props = OxmlElement('w:rPr')
             style = OxmlElement('w:rStyle')
             style.set(qn('w:val'), 'Hyperlink') # Apply hyperlink style
             run_props.append(style)
             
             # Create the run element and add properties and text
             run_element = OxmlElement('w:r')
             run_element.append(run_props)
             run_element.append(OxmlElement('w:t'))
             run_element.xpath('.//w:t')[0].text = text

             hyperlink.append(run_element)
             paragraph._p.append(hyperlink)
             # Need to return None because the run was added via oxml
             return None
         return run


    # Function to process element's children recursively
    def process_node(node, current_paragraph, list_style=None, is_bold=False, is_italic=False, current_color=None, is_strike=False, link_url=None):
        if node.name is None: # Text node
            # Only add text if it's not just whitespace inside a block element
            if node.strip() or node.parent.name not in ['p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                add_run_with_format(current_paragraph, str(node), bold=is_bold, italic=is_italic, color=current_color, strike=is_strike, link_url=link_url)
            return

        # Handle specific tags
        new_paragraph = current_paragraph
        new_list_style = list_style
        new_bold = is_bold or node.name in ['strong', 'b']
        new_italic = is_italic or node.name in ['em', 'i']
        new_color = current_color
        new_strike = is_strike
        new_link_url = link_url

        if node.name == 'ins':
             new_color = COLOR_ADDED
        elif node.name == 'del':
             new_color = COLOR_DELETED
             new_strike = True
        elif node.name == 'a' and node.get('href'):
             new_link_url = node['href']


        if node.name in ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li']:
            text_content = node.get_text(strip=True)
            if not text_content: return # Skip empty block elements
            
            if node.name.startswith('h'):
                level = int(node.name[1])
                 # Add prefix for H2/H3 as requested
                prefix = f"(H{level}) " if level in [2, 3] else ""
                new_paragraph = doc.add_heading(prefix + text_content, level=level)
                # Set bold/italic based on level for visual distinction
                if level <= 2: new_paragraph.runs[0].bold = True
                else: new_paragraph.runs[0].italic = True
                # We processed the text in add_heading, so don't recurse children directly
                return # Stop processing children for headings handled this way
            elif node.name == 'li':
                style = 'List Number' if list_style == 'ol' else 'List Bullet'
                new_paragraph = doc.add_paragraph(style=style)
            else: # Paragraph
                new_paragraph = doc.add_paragraph()

        elif node.name == 'ul':
             new_list_style = 'ul'
        elif node.name == 'ol':
             new_list_style = 'ol'
        elif node.name == 'br':
             if current_paragraph: # Add line break within the current paragraph
                  current_paragraph.add_run().add_break(WD_BREAK.LINE)
             return # Don't process children


        # Recursively process children
        if node.name not in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']: # Avoid reprocessing heading text
             for child in node.children:
                 process_node(child, new_paragraph, new_list_style, new_bold, new_italic, new_color, new_strike, new_link_url)


    # Process top-level elements in the parsed HTML
    for element in soup.contents:
        process_node(element, None) # Start with no paragraph context


def create_word_document(keyword: str, serp_results: List[Dict], related_keywords: List[Dict],
                        semantic_structure: Dict, article_content_html: str, meta_title: str, 
                        meta_description: str, paa_questions: List[Dict], term_data: Dict = None,
                        score_data: Dict = None, internal_links: List[Dict] = None, 
                        guidance_only: bool = False) -> Tuple[Optional[BytesIO], str]:
    """
    Create the main SEO Brief Word document.
    Includes SERP, keywords, terms, score, content (HTML parsed), and internal links.
    Removes the standalone structure skeleton. Adds (H2)/(H3) prefixes.
    Returns: BytesIO stream or None, status message.
    """
    if not keyword or not serp_results or not semantic_structure:
        return None, "Missing essential data for brief generation (keyword, SERP, structure)."

    try:
        doc = Document()
        # --- Header ---
        doc.add_heading(f'SEO Brief: {keyword}', 0)
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        # --- Meta Tags ---
        doc.add_heading('Meta Tags', level=1)
        p = doc.add_paragraph()
        p.add_run("Meta Title: ").bold = True
        p.add_run(meta_title if meta_title else "N/A")
        p = doc.add_paragraph()
        p.add_run("Meta Description: ").bold = True
        p.add_run(meta_description if meta_description else "N/A")
        doc.add_paragraph()

        # --- SERP Analysis ---
        doc.add_heading('SERP Analysis', level=1)
        doc.add_paragraph('Top Organic Results:')
        if serp_results:
             table = doc.add_table(rows=1, cols=4)
             table.style = 'Table Grid'
             hdr_cells = table.rows[0].cells
             hdr_cells[0].text = 'Rank'; hdr_cells[1].text = 'Title'; hdr_cells[2].text = 'URL'; hdr_cells[3].text = 'Page Type'
             for cell in hdr_cells: cell.paragraphs[0].runs[0].bold = True
             for i, result in enumerate(serp_results[:MAX_ORGANIC_RESULTS]):
                 row_cells = table.add_row().cells
                 row_cells[0].text = str(i + 1)
                 row_cells[1].text = result.get('title', '')
                 row_cells[2].text = result.get('url', '')
                 row_cells[3].text = result.get('page_type', 'Unknown')
        else:
             doc.add_paragraph("No organic results data available.")
        doc.add_paragraph()

        # PAA Questions
        if paa_questions:
            doc.add_heading('People Also Asked', level=2)
            for i, q_data in enumerate(paa_questions, 1):
                q = q_data.get('question')
                if q:
                     doc.add_paragraph(f"{i}. {q}", style='List Number')
                     # Optionally add expanded answer if needed
                     # if q_data.get('expanded'):
                     #     for exp in q_data['expanded']:
                     #         doc.add_paragraph(f"   - {exp.get('description', '')}", style='List Bullet 2')
            doc.add_paragraph()

        # --- Related Keywords ---
        doc.add_heading('Related Keywords', level=1)
        if related_keywords:
            table = doc.add_table(rows=1, cols=4) # Added Competition
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Keyword'; hdr_cells[1].text = 'Volume'; hdr_cells[2].text = 'CPC ($)'; hdr_cells[3].text = 'Competition'
            for cell in hdr_cells: cell.paragraphs[0].runs[0].bold = True
            for kw in related_keywords:
                row_cells = table.add_row().cells
                row_cells[0].text = kw.get('keyword', '')
                row_cells[1].text = str(kw['search_volume']) if kw.get('search_volume') is not None else 'N/A'
                row_cells[2].text = f"{kw['cpc']:.2f}" if kw.get('cpc') is not None else 'N/A'
                row_cells[3].text = f"{kw['competition']:.2f}" if kw.get('competition') is not None else 'N/A' # Format competition
        else:
            doc.add_paragraph("No related keywords data available.")
        doc.add_paragraph()

        # --- Important Terms (if available) ---
        if term_data:
            doc.add_heading('Important Terms', level=1)
            # Primary Terms
            if term_data.get('primary_terms'):
                doc.add_heading('Primary Terms', level=2)
                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Term'; hdr_cells[1].text = 'Importance'; hdr_cells[2].text = 'Recommended Usage'
                for cell in hdr_cells: cell.paragraphs[0].runs[0].bold = True
                for term in term_data['primary_terms']:
                    row_cells = table.add_row().cells
                    row_cells[0].text = term.get('term', '')
                    row_cells[1].text = f"{term.get('importance', 0):.2f}"
                    row_cells[2].text = str(term.get('recommended_usage', 1))
            # Secondary Terms
            if term_data.get('secondary_terms'):
                 doc.add_heading('Secondary Terms', level=2)
                 table = doc.add_table(rows=1, cols=2)
                 table.style = 'Table Grid'
                 hdr_cells = table.rows[0].cells
                 hdr_cells[0].text = 'Term'; hdr_cells[1].text = 'Importance'
                 for cell in hdr_cells: cell.paragraphs[0].runs[0].bold = True
                 for term in term_data['secondary_terms'][:15]: # Limit display
                     row_cells = table.add_row().cells
                     row_cells[0].text = term.get('term', '')
                     row_cells[1].text = f"{term.get('importance', 0):.2f}"
            doc.add_paragraph()

        # --- Content Score (if available and not guidance) ---
        if score_data and not guidance_only:
            doc.add_heading('Content Score', level=1)
            score_para = doc.add_paragraph()
            score_para.add_run("Overall Score: ").bold = True
            overall_score = score_data.get('overall_score', 0)
            grade = score_data.get('grade', 'F')
            score_run = score_para.add_run(f"{overall_score} ({grade})")
            if overall_score >= 70: score_run.font.color.rgb = COLOR_ADDED
            elif overall_score < 50: score_run.font.color.rgb = COLOR_DELETED
            else: score_run.font.color.rgb = RGBColor(255, 165, 0) # Orange
            # Optionally add component scores table here if desired
            doc.add_paragraph()


        # --- Semantic Structure (REMOVED as per requirement) ---
        # The section displaying the H1/H2/H3 skeleton is intentionally omitted.

        # --- Generated Content ---
        content_title = "Generated Content Guidance" if guidance_only else "Generated Article Content"
        doc.add_heading(content_title, level=1)
        if article_content_html:
             add_html_to_docx(doc, article_content_html) # Use the HTML parsing function
        else:
             doc.add_paragraph("No content was generated.")
        doc.add_paragraph()

        # --- Internal Linking Summary (if available and not guidance) ---
        if internal_links and not guidance_only:
            doc.add_heading('Internal Linking Summary', level=1)
            doc.add_paragraph(f"Summary of {len(internal_links)} internal links added:")
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Anchor Text'; hdr_cells[1].text = 'Target URL'; hdr_cells[2].text = 'Context Snippet'
            for cell in hdr_cells: cell.paragraphs[0].runs[0].bold = True
            for link in internal_links:
                 row_cells = table.add_row().cells
                 row_cells[0].text = link.get('anchor_text', '')
                 row_cells[1].text = link.get('url', '')
                 # Clean context snippet for display
                 context_html = link.get('context', '')
                 context_text = clean_html(context_html).replace('<mark>', '[').replace('</mark>', ']')
                 row_cells[2].text = context_text
            doc.add_paragraph()


        # --- Save Document ---
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        return doc_stream, "Success"

    except Exception as e:
        error_msg = f"Exception creating Word brief: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg

###############################################################################
# 10. Content Update Functions
###############################################################################

def parse_word_document(uploaded_file) -> Tuple[Optional[Dict], str]:
    """Parse uploaded Word document to extract text and basic structure."""
    try:
        doc = Document(BytesIO(uploaded_file.getvalue()))
        content = {'title': '', 'headings': [], 'paragraphs': [], 'full_text': ''}
        full_text_parts = []
        current_heading_obj = None

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue

            full_text_parts.append(text)
            style_name = para.style.name

            if style_name.startswith('Heading'):
                try:
                    level = int(style_name.split(' ')[-1])
                except:
                    level = 1 # Default for "Heading" or malformed styles

                heading_obj = {'text': text, 'level': level, 'paragraphs': []}
                content['headings'].append(heading_obj)
                current_heading_obj = heading_obj
                if level == 1 and not content['title']:
                    content['title'] = text
            else:
                content['paragraphs'].append({'text': text, 'style': style_name})
                if current_heading_obj:
                    current_heading_obj['paragraphs'].append(text)

        content['full_text'] = '\n\n'.join(full_text_parts)
        logger.info(f"Parsed Word document. Title: '{content['title']}'. Found {len(content['headings'])} headings, {len(content['paragraphs'])} paragraphs.")
        return content, "Success"

    except Exception as e:
        error_msg = f"Failed to parse Word document: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg


def analyze_content_gaps(existing_content: Dict, competitor_contents: List[Dict], 
                        semantic_structure: Dict, term_data: Dict, score_data: Dict, 
                        anthropic_api_key: str, keyword: str, 
                        paa_questions: List[Dict] = None) -> Tuple[Optional[Dict], str]:
    """
    Analyze gaps between existing content and competitor/recommended structure using Anthropic.
    Includes score data in the analysis context.
    Returns: Content gaps dictionary or None, status message.
    """
    client = get_anthropic_client(anthropic_api_key)
    if not client:
        return None, "Anthropic client initialization failed"
        
    if not existing_content or not semantic_structure or not term_data:
        return None, "Missing existing content, structure, or term data for gap analysis."

    # --- Prepare Context for LLM ---
    existing_title = existing_content.get('title', 'N/A')
    existing_headings_text = [h.get('text', '') for h in existing_content.get('headings', [])]
    existing_full_text_snippet = existing_content.get('full_text', '')[:6000] # Limit context size

    recommended_h1 = semantic_structure.get('h1', 'N/A')
    recommended_sections = []
    for sec in semantic_structure.get('sections', []):
         h2 = sec.get('h2')
         if h2:
             recommended_sections.append(f"H2: {h2}")
             for sub in sec.get('subsections', []):
                 h3 = sub.get('h3')
                 if h3: recommended_sections.append(f"  H3: {h3}")

    # Competitor context (Summaries)
    competitor_summary = ""
    for i, comp in enumerate(competitor_contents[:3]): # Use top 3 competitors for context
         title = comp.get('title', f'Competitor {i+1}')
         content_snippet = clean_html(comp.get('content', ''))[:1500]
         competitor_summary += f"--- Competitor {i+1}: {title} ---\n{content_snippet}\n\n"

    # Score context
    score_summary = "N/A"
    if score_data:
         overall = score_data.get('overall_score', 'N/A')
         grade = score_data.get('grade', 'N/A')
         components = score_data.get('components', {})
         score_summary = f"Overall Score: {overall} ({grade})\nComponents:\n"
         for comp, val in components.items():
             score_summary += f"- {comp.replace('_score','').title()}: {val}\n"
         # Add key issues from score details (optional, can make prompt long)
         # Example: Add top 3 missing primary terms
         if score_data.get('details', {}).get('primary_term_analysis'):
              missing_primary = [t for t, info in score_data['details']['primary_term_analysis'].items() if info['count'] == 0]
              if missing_primary:
                   score_summary += f"Missing Primary Terms Sample: {', '.join(missing_primary[:3])}\n"

    # Term data context
    term_summary = "Primary Terms:\n" + "\n".join([f"- {t['term']}" for t in term_data.get('primary_terms', [])[:5]]) + "\n"
    term_summary += "Key Topics:\n" + "\n".join([f"- {t['topic']}" for t in term_data.get('topics', [])[:5]])

    # PAA context
    paa_summary = "People Also Asked:\n" + "\n".join([f"- {q['question']}" for q in paa_questions[:5] if q.get('question')]) if paa_questions else ""


    # --- Construct the Prompt ---
    prompt = f"""
    Perform a detailed content gap analysis for an existing article targeting the keyword "{keyword}". 
    Compare the existing content against the recommended structure (derived from top competitors) and competitor summaries. 
    Identify specific areas for improvement.

    **Existing Content Summary:**
    - Title: {existing_title}
    - Headings: {json.dumps(existing_headings_text)}
    - Current Score Analysis: {score_summary}
    - Text Snippet (first 6000 chars): 
    {existing_full_text_snippet}

    **Target/Recommended Structure:**
    - H1: {recommended_h1}
    - Sections:
    {chr(10).join(recommended_sections)} 

    **Key Terms & Topics from Analysis:**
    {term_summary}

    **Relevant Competitor Content Snippets:**
    {competitor_summary}
    
    **Relevant Questions (PAA):**
    {paa_summary if paa_summary else "N/A"}

    **Analysis Tasks:**
    1.  **Missing Sections:** Identify recommended H2/H3 sections completely missing from the existing content. Suggest where to insert them.
    2.  **Heading Revisions:** Suggest improvements/renaming for existing headings to better align with the recommended structure or keyword focus.
    3.  **Content Gaps:** List key sub-topics or points covered by competitors (or implied by terms/topics) that are absent or underdeveloped in the existing content.
    4.  **Expansion Areas:** Pinpoint existing sections that are too brief or lack depth compared to competitors or the recommended scope.
    5.  **Semantic Relevancy:** Identify parts of the existing content that deviate significantly from the core topic "{keyword}" or lack sufficient keyword integration.
    6.  **Term Usage:** Check if important primary/secondary terms (from Term Analysis) are adequately used. List key missing/underused terms.
    7.  **Question Answering:** Identify which PAA questions are not adequately addressed.

    **Output Format:**
    Return ONLY a valid JSON object using EXACTLY this structure. Provide specific, actionable recommendations.
    ```json
    {{
        "missing_sections": [
            {{ 
                "heading": "Recommended Heading Text", 
                "level": 2, // or 3
                "suggested_content_summary": "Brief description of content needed.",
                "insert_after_heading": "Name of existing heading to insert after (or 'START'/'END')" 
            }}
        ],
        "revised_headings": [
            {{ "original_heading": "Current Heading Text", "suggested_heading": "Improved Heading Text", "reason": "Brief rationale (e.g., better keyword focus, clarity)." }}
        ],
        "content_gaps": [
            {{ "topic_gap": "Specific missing sub-topic or information point", "relevant_section": "Existing section where it fits (or suggest new)", "recommendation": "Briefly describe content to add." }}
        ],
        "expansion_areas": [
            {{ "section_to_expand": "Heading of the section needing more depth", "reason": "Why it needs expansion (e.g., too brief, lacks examples).", "suggested_additions": "Examples of content to add." }}
        ],
        "semantic_relevancy_issues": [
            {{ 
                "section_heading": "Heading of the off-topic section", 
                "issue_description": "How it deviates or lacks focus on '{keyword}'.", 
                "recommendation": "How to refocus or integrate '{keyword}' better." 
            }}
        ],
        "term_usage_issues": [
            {{ "term": "Missing or underused term", "suggestion": "Suggest section(s) to add it and how (e.g., 'Integrate into Introduction and Key Benefits section naturally')." }}
        ],
        "unanswered_paa": [
             {{ "question": "PAA question not answered", "recommendation": "Suggest section to add the answer or create a dedicated FAQ." }}
        ]
    }}
    ```
    Ensure the JSON is valid. Do not include explanations outside the JSON structure.
    """
    
    try:
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=3500, # Needs ample tokens for detailed analysis
            system="You are an expert SEO Content Auditor. You meticulously analyze content gaps and provide structured, actionable recommendations in valid JSON format.",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1 # Low temperature for factual gap analysis
        )

        response_text = response.content[0].text
        content_gaps = safe_json_loads(response_text)

        # Validate the structure
        required_keys = ["missing_sections", "revised_headings", "content_gaps", "expansion_areas", 
                         "semantic_relevancy_issues", "term_usage_issues", "unanswered_paa"]
        if content_gaps and all(key in content_gaps for key in required_keys):
             # Further validation: ensure values are lists
             if all(isinstance(content_gaps[key], list) for key in required_keys):
                 logger.info(f"Successfully performed content gap analysis for '{keyword}'.")
                 return content_gaps, "Success"
             else:
                  logger.warning("Gap analysis response JSON has incorrect types for keys (expected lists).")
                  # Attempt to fix: ensure all keys have lists, even if empty
                  for key in required_keys:
                       if not isinstance(content_gaps.get(key), list):
                           content_gaps[key] = []
                  return content_gaps, "Success (with potential type correction)"
        else:
            error_msg = "Failed to parse valid content gap JSON or missing required keys."
            logger.error(f"{error_msg} Response: {response_text[:500]}...")
            # Return default structure on failure
            default_gaps = {key: [] for key in required_keys}
            return default_gaps, error_msg


    except anthropic.APIError as api_err:
        error_msg = f"Anthropic API Error analyzing content gaps: {api_err}"
        logger.error(error_msg)
        return None, f"API Error: {api_err}"
    except Exception as e:
        error_msg = f"Unexpected error analyzing content gaps: {e}"
        logger.error(error_msg, exc_info=True)
        return None, error_msg


def generate_optimized_article_with_tracking(existing_content: Dict, competitor_contents: List[Dict], 
                                             semantic_structure: Dict, term_data: Dict, 
                                             content_gaps: Dict, # Use gap analysis results
                                             anthropic_api_key: str, keyword: str,
                                             competitor_embeddings: List[Dict], openai_api_key: str
                                             ) -> Tuple[Optional[str], Optional[str], str]:
    """
    Generates an optimized article by intelligently merging existing content 
    with recommendations from gap analysis and competitor context. Tracks changes.
    Returns: Optimized HTML content, Change summary HTML, Status message.
    """
    client = get_anthropic_client(anthropic_api_key)
    if not client:
        return None, None, "Anthropic client initialization failed"
        
    if not existing_content or not semantic_structure or not term_data or not content_gaps:
         return None, None, "Missing required data for optimized article generation."

    logger.info(f"Generating optimized article for '{keyword}', incorporating existing content and gap analysis.")
    
    optimized_html_sections = []
    change_log = {'added': [], 'revised': [], 'deleted': [], 'expanded': [], 'refocused': []}
    
    # --- Map existing content by heading for easier lookup ---
    existing_sections = {h['text'].lower(): {'level': h['level'], 'content': "\n\n".join(h['paragraphs'])} 
                         for h in existing_content.get('headings', []) if h.get('text')}

    # --- Process Recommended Structure ---
    
    # 1. Handle H1
    recommended_h1 = semantic_structure.get('h1', f"Guide to {keyword}")
    original_h1 = existing_content.get('title', '')
    if original_h1.lower() != recommended_h1.lower():
         change_log['revised'].append(f"H1 Title changed from '{original_h1}' to '{recommended_h1}'")
    optimized_html_sections.append(f"<h1>{recommended_h1}</h1>")
    
    # Find content potentially belonging to the intro (before first H2)
    intro_content = ""
    first_h2_text = semantic_structure.get('sections', [{}])[0].get('h2','').lower() if semantic_structure.get('sections') else None
    intro_paras = []
    for para in existing_content.get('paragraphs', []):
         belongs_to_heading = False
         for h in existing_content.get('headings', []):
             if h['text'] in para.get('heading', ''): # Simple check if para associated with any heading
                  belongs_to_heading = True
                  break
         # Rough check if paragraph appears before the first H2 content might start
         is_potentially_intro = True
         if first_h2_text and first_h2_text in existing_content.get('full_text','').lower():
              if existing_content['full_text'].lower().find(para['text'].lower()) > existing_content['full_text'].lower().find(first_h2_text):
                   is_potentially_intro = False
                   
         if not belongs_to_heading and is_potentially_intro:
             intro_paras.append(para['text'])
    original_intro = "\n\n".join(intro_paras)


    # Generate/Refine Intro using H1 context and original intro
    intro_target_words = 150
    competitor_context_intro = find_relevant_competitor_context(recommended_h1, competitor_embeddings, openai_api_key, top_n=1)
    
    prompt_intro = f"""
    Refine or generate an introduction for an article titled "{recommended_h1}" about "{keyword}".
    
    Original Intro Content (if any):
    {original_intro if original_intro else "N/A"}
    
    Context from Competitors:
    {competitor_context_intro}

    Requirements:
    - Briefly introduce the topic "{keyword}".
    - State the article's purpose or what the reader will learn.
    - Engage the reader.
    - Target word count: ~{intro_target_words} words.
    - Integrate 1-2 primary terms naturally: {', '.join([t['term'] for t in term_data.get('primary_terms',[])[:2]])}
    - Format as HTML paragraphs (<p> tags).

    Output ONLY the HTML paragraph(s) for the introduction.
    """
    try:
        intro_response = client.messages.create(
            model=ANTHROPIC_MODEL, max_tokens=intro_target_words * 3,
            system="You are an expert copywriter creating engaging article introductions.",
            messages=[{"role": "user", "content": prompt_intro}], temperature=0.5
        )
        intro_html = intro_response.content[0].text.strip()
        optimized_html_sections.append(intro_html)
        if original_intro and difflib.SequenceMatcher(None, original_intro, clean_html(intro_html)).ratio() < 0.7:
             change_log['revised'].append("Introduction section significantly revised for clarity and engagement.")
        elif not original_intro:
              change_log['added'].append("Generated new introduction section.")

    except Exception as e:
        logger.error(f"Failed to generate introduction: {e}")
        optimized_html_sections.append("<p><i>Error generating introduction.</i></p>")


    # --- Process H2 and H3 Sections ---
    processed_existing_headings = {original_h1.lower()} if original_h1 else set()

    for section in semantic_structure.get('sections', []):
        h2_text = section.get('h2')
        if not h2_text: continue

        optimized_html_sections.append(f"<h2>{h2_text}</h2>")
        processed_existing_headings.add(h2_text.lower()) # Mark recommended heading as 'processed' conceptually
        
        # Find if this H2 corresponds to a revised heading or a missing section
        original_heading_match = None
        revision_reason = ""
        for rev in content_gaps.get('revised_headings', []):
             if rev.get('suggested_heading', '').lower() == h2_text.lower():
                 original_heading_match = rev.get('original_heading')
                 revision_reason = rev.get('reason', '')
                 if original_heading_match:
                      processed_existing_headings.add(original_heading_match.lower())
                 break
                 
        is_new_section = h2_text.lower() not in existing_sections and not original_heading_match
        for miss in content_gaps.get('missing_sections',[]):
             if miss.get('heading','').lower() == h2_text.lower():
                 is_new_section = True
                 break

        original_h2_content = existing_sections.get(h2_text.lower(), {}).get('content', '')
        if not original_h2_content and original_heading_match: # Content might be under original name
             original_h2_content = existing_sections.get(original_heading_match.lower(), {}).get('content', '')


        # --- Generate/Update H2 Content ---
        # Gather context specific to this section
        section_context = f"Section: {h2_text}\n"
        section_context += f"Original Content Snippet:\n{original_h2_content[:1000]}\n\n" if original_h2_content else "This is potentially a new section.\n"
        if revision_reason: section_context += f"Revision Reason: {revision_reason}\n"
        
        # Find relevant gaps/expansions/terms/questions
        gaps_for_section = [g['recommendation'] for g in content_gaps.get('content_gaps', []) if h2_text.lower() in g.get('relevant_section', '').lower()]
        expansions_for_section = [e['suggested_additions'] for e in content_gaps.get('expansion_areas', []) if h2_text.lower() in e.get('section_to_expand', '').lower()]
        relevancy_issues = [r['recommendation'] for r in content_gaps.get('semantic_relevancy_issues', []) if h2_text.lower() in r.get('section_heading', '').lower()]
        terms_for_section = [t['term'] for t in content_gaps.get('term_usage_issues', []) if h2_text.lower() in t.get('suggestion', '').lower()]
        questions_for_section = [q['question'] for q in content_gaps.get('unanswered_paa', []) if h2_text.lower() in q.get('recommendation', '').lower()]
        
        section_context += "Specific Improvement Areas:\n"
        if gaps_for_section: section_context += f"- Address Gaps: {'; '.join(gaps_for_section)}\n"
        if expansions_for_section: section_context += f"- Expand Content: {'; '.join(expansions_for_section)}\n"
        if relevancy_issues: section_context += f"- Improve Relevancy: {'; '.join(relevancy_issues)}\n"
        if terms_for_section: section_context += f"- Include Terms: {', '.join(terms_for_section)}\n"
        if questions_for_section: section_context += f"- Answer Questions: {', '.join(questions_for_section)}\n"
        if not any([gaps_for_section, expansions_for_section, relevancy_issues, terms_for_section, questions_for_section]):
             section_context += "- Focus on clarity, depth, and competitor insights.\n"

        competitor_context_h2 = find_relevant_competitor_context(h2_text, competitor_embeddings, openai_api_key)
        section_context += f"\nCompetitor Context Snippets:\n{competitor_context_h2}"

        # Prepare prompt for H2 update/generation
        h2_prompt = f"""
        Update or generate content for the section "{h2_text}" in an article about "{keyword}".
        
        Context & Instructions:
        {section_context}

        Requirements:
        - If original content exists, **preserve valuable information** while incorporating improvements. Synthesize, don't just append.
        - If this is a new section, generate content based on the improvement areas and competitor context.
        - Address the specific improvement points listed above.
        - Integrate primary/secondary terms naturally.
        - Ensure logical flow from the previous section.
        - Target word count: ~200-300 words (adjust based on subsections).
        - Format as HTML paragraphs (<p>), using lists (<ul>, <ol>) where appropriate.
        - **Crucially, use `<ins>` tags for added text and `<del>` tags for removed text compared to the original snippet provided.** If it's entirely new content, use `<ins>` for the whole section. Wrap the final output in a single block, e.g., <p><ins>new text...</ins></p> or <p>existing text <del>old</del><ins>new</ins> existing text</p>.

        Output ONLY the HTML content for this section (paragraphs, lists), including the `<ins>`/`<del>` tags for changes. Do not include the `<h2>` tag itself.
        """

        try:
            h2_response = client.messages.create(
                model=ANTHROPIC_MODEL, max_tokens=1000,
                system="You are an expert Content Editor. You rewrite and generate article sections, preserving value, incorporating specific feedback, and meticulously tracking changes using <ins> and <del> HTML tags.",
                messages=[{"role": "user", "content": h2_prompt}], temperature=0.3
            )
            h2_html_diff = h2_response.content[0].text.strip()
            optimized_html_sections.append(h2_html_diff)
            
            # Log changes based on presence of <ins>/<del> or original content
            if is_new_section:
                 change_log['added'].append(f"Added new section: '{h2_text}'.")
            elif original_heading_match and original_heading_match.lower() != h2_text.lower():
                 change_log['revised'].append(f"Revised heading '{original_heading_match}' to '{h2_text}' and updated content.")
            elif '<ins>' in h2_html_diff or '<del>' in h2_html_diff:
                 change_log['revised'].append(f"Revised content for section: '{h2_text}'.")
            if expansions_for_section: change_log['expanded'].append(f"Expanded section: '{h2_text}'.")
            if relevancy_issues: change_log['refocused'].append(f"Refocused section for keyword relevancy: '{h2_text}'.")


        except Exception as e:
            logger.error(f"Failed to generate/update H2 section '{h2_text}': {e}")
            optimized_html_sections.append(f"<p><i>Error processing content for {h2_text}.</i></p>")

        # --- Process H3 Subsections ---
        for subsection in section.get('subsections', []):
             h3_text = subsection.get('h3')
             if not h3_text: continue
             
             optimized_html_sections.append(f"<h3>{h3_text}</h3>")
             processed_existing_headings.add(h3_text.lower())

             # Check if H3 existed, was revised, or is new (similar logic as H2)
             original_h3_match = None
             for rev in content_gaps.get('revised_headings', []):
                  if rev.get('suggested_heading', '').lower() == h3_text.lower():
                      original_h3_match = rev.get('original_heading')
                      if original_h3_match: processed_existing_headings.add(original_h3_match.lower())
                      break
                      
             is_new_h3 = h3_text.lower() not in existing_sections and not original_h3_match
             for miss in content_gaps.get('missing_sections',[]):
                  if miss.get('heading','').lower() == h3_text.lower():
                      is_new_h3 = True
                      break
                      
             original_h3_content = existing_sections.get(h3_text.lower(), {}).get('content', '')
             if not original_h3_content and original_h3_match:
                  original_h3_content = existing_sections.get(original_h3_match.lower(), {}).get('content', '')

             # Generate/Update H3 content (simplified context gathering for brevity)
             h3_improvement_context = f"Focus on adding specific details for '{h3_text}' under the main topic '{h2_text}'.\n"
             h3_improvement_context += f"Original Content Snippet:\n{original_h3_content[:500]}\n\n" if original_h3_content else "This is potentially a new subsection.\n"
             h3_competitor_context = find_relevant_competitor_context(h3_text, competitor_embeddings, openai_api_key, top_n=1)
             h3_improvement_context += f"Competitor Context Snippet:\n{h3_competitor_context}"
             
             h3_prompt = f"""
             Update or generate content for the subsection "{h3_text}" (under "{h2_text}") in an article about "{keyword}".
             
             Context & Instructions:
             {h3_improvement_context}
             
             Requirements:
             - Be concise and focused on the specific H3 topic. Target ~100-200 words.
             - If original content exists, preserve key info and integrate improvements.
             - Use competitor context for insights if generating new content or expanding.
             - Format as HTML paragraphs (<p>), using lists (<ul>, <ol>) if needed.
             - **Use `<ins>` for added text and `<del>` for removed text compared to the original snippet.** Use `<ins>` for the whole section if new.
             
             Output ONLY the HTML content (paragraphs, lists) with `<ins>`/`<del>` tags. Do not include the `<h3>` tag.
             """
             
             try:
                  h3_response = client.messages.create(
                      model=ANTHROPIC_MODEL, max_tokens=600,
                      system="You are an expert Content Editor specializing in subsections. You preserve value, incorporate feedback, and track changes with <ins>/<del> tags.",
                      messages=[{"role": "user", "content": h3_prompt}], temperature=0.3
                  )
                  h3_html_diff = h3_response.content[0].text.strip()
                  optimized_html_sections.append(h3_html_diff)
                  # Log H3 changes (simplified)
                  if is_new_h3: change_log['added'].append(f"Added new subsection: '{h3_text}'.")
                  elif '<ins>' in h3_html_diff or '<del>' in h3_html_diff: change_log['revised'].append(f"Revised subsection: '{h3_text}'.")

             except Exception as e:
                  logger.error(f"Failed to generate/update H3 subsection '{h3_text}': {e}")
                  optimized_html_sections.append(f"<p><i>Error processing content for {h3_text}.</i></p>")


    # --- Identify and Mark Deleted Sections ---
    deleted_sections_html = []
    for heading, data in existing_sections.items():
         if heading not in processed_existing_headings:
             # Add deleted heading and content, wrapped in <del>
             level = data.get('level', 2) # Assume H2 if level unknown
             deleted_heading_tag = f"h{level}"
             deleted_sections_html.append(f"<{deleted_heading_tag}><del>{heading.title()}</del></{deleted_heading_tag}>")
             # Wrap existing paragraphs in <p><del>...</del></p>
             deleted_content = "\n".join([f"<p><del>{p}</del></p>" for p in data.get('content', '').split('\n\n') if p])
             deleted_sections_html.append(deleted_content)
             change_log['deleted'].append(f"Removed section: '{heading.title()}'.")
             
    # Append deleted sections at the end (or integrate them based on structure analysis - harder)
    if deleted_sections_html:
         optimized_html_sections.append("<hr/><h2>Deleted Sections (For Reference)</h2>")
         optimized_html_sections.extend(deleted_sections_html)


    # --- Compile Final Article and Change Summary ---
    final_optimized_html = "\n".join(optimized_html_sections)

    # Create HTML change summary
    summary_html = f"""
    <div style="border: 1px solid #ccc; padding: 15px; margin-bottom: 20px; background-color: #f9f9f9; font-family: sans-serif;">
        <h2 style="margin-top: 0;">Optimization Change Summary for "{keyword}"</h2>
        <p>This summary highlights the key modifications made to the original content.</p>
    """
    if change_log['added']:
        summary_html += '<h3>Sections Added:</h3><ul>'
        for item in change_log['added']: summary_html += f'<li>{item}</li>'
        summary_html += '</ul>'
    if change_log['revised']:
        summary_html += '<h3>Sections Revised / Renamed:</h3><ul>'
        for item in change_log['revised']: summary_html += f'<li>{item}</li>'
        summary_html += '</ul>'
    if change_log['expanded']:
        summary_html += '<h3>Sections Expanded:</h3><ul>'
        for item in change_log['expanded']: summary_html += f'<li>{item}</li>'
        summary_html += '</ul>'
    if change_log['refocused']:
        summary_html += '<h3>Sections Refocused:</h3><ul>'
        for item in change_log['refocused']: summary_html += f'<li>{item}</li>'
        summary_html += '</ul>'
    if change_log['deleted']:
        summary_html += '<h3>Sections Deleted:</h3><ul>'
        for item in change_log['deleted']: summary_html += f'<li>{item}</li>'
        summary_html += '</ul>'
        
    summary_html += '<p><i>Review the document below for detailed changes marked with <ins style="background-color:#d4edda; text-decoration:none; color:#155724;">green</ins> for additions and <del style="background-color:#f8d7da; color:#721c24;">red strikethrough</del> for deletions.</i></p>'
    summary_html += "</div>"

    return final_optimized_html, summary_html, "Success"


def create_word_document_with_changes(html_content_with_diff: str, keyword: str, change_summary_html: str = "") -> Optional[BytesIO]:
    """
    Creates a Word document from HTML content containing <ins> and <del> tags,
    applying specific formatting for additions (green) and deletions (red strikethrough).
    Includes the change summary at the beginning.
    """
    try:
        doc = Document()
        # --- Header ---
        doc.add_heading(f'Optimized Content (with Changes): {keyword}', 0)
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        doc.add_paragraph()

        # --- Change Summary ---
        if change_summary_html:
             doc.add_heading('Change Summary', level=1)
             # Basic parsing of the summary HTML (assumes simple ul/li structure)
             summary_soup = BeautifulSoup(change_summary_html, 'html.parser')
             for element in summary_soup.find_all(['h2', 'h3', 'p', 'ul']):
                 if element.name == 'h2':
                      doc.add_heading(element.get_text(strip=True), level=2)
                 elif element.name == 'h3':
                      doc.add_heading(element.get_text(strip=True), level=3)
                 elif element.name == 'p':
                      # Handle the instruction paragraph with formatting
                      p_tag = doc.add_paragraph()
                      if "Review the document below" in element.get_text():
                           p_tag.add_run("Review the document below for detailed changes marked with ")
                           run_ins = p_tag.add_run("green")
                           run_ins.font.color.rgb = COLOR_ADDED
                           p_tag.add_run(" for additions and ")
                           run_del = p_tag.add_run("red strikethrough")
                           run_del.font.color.rgb = COLOR_DELETED
                           run_del.font.strike = True
                           p_tag.add_run(" for deletions.")
                           p_tag.italic = True
                      else:
                           p_tag.add_run(element.get_text(strip=True))

                 elif element.name == 'ul':
                      for li in element.find_all('li'):
                          doc.add_paragraph(li.get_text(strip=True), style='List Bullet')
             doc.add_page_break() # Separate summary from content


        # --- Add Optimized Content with Diff Formatting ---
        doc.add_heading('Optimized Content', level=1)
        add_html_to_docx(doc, html_content_with_diff) # Use the enhanced parser

        # --- Save Document ---
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        return doc_stream

    except Exception as e:
        error_msg = f"Exception creating Word document with changes: {e}"
        logger.error(error_msg, exc_info=True)
        return None


###############################################################################
# 11. Main Streamlit App
###############################################################################

def main():
    st.title("🚀 SEO Content Optimizer v2.0")
    
    # --- Sidebar ---
    st.sidebar.header("🔑 API Credentials")
    dataforseo_login = st.sidebar.text_input("DataForSEO API Login", type="password", key="d4s_login")
    dataforseo_password = st.sidebar.text_input("DataForSEO API Password", type="password", key="d4s_pass")
    openai_api_key = st.sidebar.text_input("OpenAI API Key", type="password", key="openai_key")
    anthropic_api_key = st.sidebar.text_input("Anthropic API Key", type="password", key="anthropic_key")
    
    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ Settings")
    # Add any future settings here, e.g., target word count, model selection

    # --- Initialize Session State ---
    required_keys = ['results', 'keyword_input', 'current_content', 'analysis_complete', 'generation_complete']
    for key in required_keys:
        if key not in st.session_state:
            if key == 'results': st.session_state.results = {}
            else: st.session_state[key] = None
            
    # --- Main App Logic ---
    
    # Input Area
    st.header("1. Target Keyword & SERP Analysis")
    keyword = st.text_input("Enter Target Keyword:", value=st.session_state.get('keyword_input', ''), key="keyword_input_field")
    
    # Use columns for better layout
    col1, col2 = st.columns(2)
    
    with col1:
         if st.button("Analyze SERP & Competitors", key="analyze_serp_button"):
             st.session_state.keyword_input = keyword # Store keyword input
             if not keyword:
                 st.error("Please enter a target keyword.")
             elif not all([dataforseo_login, dataforseo_password, openai_api_key, anthropic_api_key]):
                 st.error("Please provide all API credentials in the sidebar.")
             else:
                 st.session_state.results = {'keyword': keyword} # Reset results for new keyword
                 st.session_state.analysis_complete = False
                 st.session_state.generation_complete = False
                 status_placeholder = st.empty()
                 progress_bar = st.progress(0)
                 start_time_analysis = time.time()

                 try:
                     # --- Step 1: Fetch SERP Data ---
                     status_placeholder.info("🔄 Fetching SERP data from DataForSEO...")
                     d4s_client = DataForSEOClient(dataforseo_login, dataforseo_password)
                     organic_results, serp_features, paa_questions, serp_status = d4s_client.fetch_serp_results(keyword)
                     
                     if organic_results is None: # Check specifically for None, empty list is okay
                         raise Exception(f"Failed to fetch SERP data: {serp_status}")
                         
                     st.session_state.results['organic_results'] = organic_results
                     st.session_state.results['serp_features'] = serp_features or []
                     st.session_state.results['paa_questions'] = paa_questions or []
                     progress_bar.progress(15)

                     # --- Step 2: Fetch Related Keywords ---
                     status_placeholder.info("🔄 Fetching related keywords...")
                     related_keywords, kw_status = d4s_client.fetch_related_keywords(keyword)
                     if related_keywords is None:
                         logger.warning(f"Could not fetch related keywords: {kw_status}. Proceeding without them.")
                         st.session_state.results['related_keywords'] = []
                     else:
                          st.session_state.results['related_keywords'] = related_keywords
                     progress_bar.progress(30)

                     # --- Step 3: Scrape Competitor Content ---
                     status_placeholder.info("🔄 Scraping content from top competitors...")
                     scraped_contents = []
                     if not organic_results:
                          st.warning("No organic results found to scrape.")
                     else:
                         for i, result in enumerate(organic_results):
                             status_placeholder.info(f"🔄 Scraping {result.get('url', '')}...")
                             content_text, scrape_status = scrape_webpage(result['url'])
                             headings_data, heading_status = extract_headings(result['url'])
                             
                             if content_text and scrape_status.startswith("Success"):
                                  # Generate embedding immediately after scraping
                                  embedding, embed_status = generate_embedding(content_text, openai_api_key)
                                  if embedding:
                                      scraped_contents.append({
                                          'url': result['url'],
                                          'title': result.get('title',''),
                                          'content': content_text, # Store raw-ish text
                                          'headings': headings_data if headings_data else {},
                                          'embedding': embedding
                                      })
                                  else:
                                      logger.warning(f"Failed to generate embedding for {result['url']}: {embed_status}")
                                      # Store without embedding if failed
                                      scraped_contents.append({
                                          'url': result['url'], 'title': result.get('title',''),
                                          'content': content_text, 'headings': headings_data if headings_data else {},
                                          'embedding': None
                                      })
                             else:
                                  logger.warning(f"Failed to scrape {result['url']}: {scrape_status}")
                             
                             progress_bar.progress(30 + int((i + 1) / len(organic_results) * 30)) # 30% for scraping

                     st.session_state.results['scraped_contents'] = scraped_contents
                     st.session_state.results['competitor_embeddings'] = [c for c in scraped_contents if c.get('embedding')] # Separate list with embeddings

                     if not st.session_state.results['competitor_embeddings']:
                          st.warning("Could not generate embeddings for any competitor content. Context features will be limited.")


                     # --- Step 4: Analyze Structure & Terms ---
                     status_placeholder.info("🔄 Analyzing semantic structure using Anthropic...")
                     semantic_structure, struct_status = analyze_semantic_structure(scraped_contents, anthropic_api_key)
                     if not semantic_structure:
                          st.warning(f"Failed to analyze semantic structure: {struct_status}. Using default structure.")
                          st.session_state.results['semantic_structure'] = {"h1": f"Guide to {keyword}", "sections": []}
                     else:
                         st.session_state.results['semantic_structure'] = semantic_structure
                     progress_bar.progress(75)
                     
                     status_placeholder.info("🔄 Extracting important terms using Anthropic...")
                     term_data, term_status = extract_important_terms(scraped_contents, anthropic_api_key, keyword)
                     if not term_data:
                         st.warning(f"Failed to extract terms: {term_status}. Content scoring/generation might be less effective.")
                         st.session_state.results['term_data'] = {"primary_terms": [], "secondary_terms": [], "topics": [], "questions": []}
                     else:
                         st.session_state.results['term_data'] = term_data
                     progress_bar.progress(90)
                     
                     # --- Step 5: Generate Meta Tags ---
                     status_placeholder.info("🔄 Generating meta tags using Anthropic...")
                     meta_title, meta_description, meta_status = generate_meta_tags(
                         keyword,
                         st.session_state.results['semantic_structure'],
                         st.session_state.results.get('related_keywords', []),
                         st.session_state.results['term_data'],
                         anthropic_api_key
                     )
                     st.session_state.results['meta_title'] = meta_title
                     st.session_state.results['meta_description'] = meta_description
                     if not meta_title:
                         st.warning(f"Failed to generate meta tags: {meta_status}")
                     progress_bar.progress(100)
                     
                     status_placeholder.success(f"✅ Analysis complete for '{keyword}' in {format_time(time.time() - start_time_analysis)}!")
                     st.session_state.analysis_complete = True
                     # Force rerun to update display sections
                     st.experimental_rerun()

                 except Exception as e:
                     display_error(f"An error occurred during analysis: {e}", e)
                     st.session_state.analysis_complete = False
                     status_placeholder.error(f"Analysis failed. See logs for details.")

    with col2:
        # Display Analysis Summary / Status
        st.write("**Analysis Status:**")
        if st.session_state.get('analysis_complete'):
            st.success(f"Analysis complete for: **{st.session_state.results.get('keyword','N/A')}**")
            st.markdown(f"**Meta Title:** {st.session_state.results.get('meta_title','N/A')}")
            st.markdown(f"**Meta Desc:** {st.session_state.results.get('meta_description','N/A')}")
        elif st.session_state.get('results') and 'keyword' in st.session_state.results:
            st.info("Analysis started but may not be complete. Check status messages.")
        else:
            st.warning("Enter a keyword and click 'Analyze SERP & Competitors'.")
            
    st.markdown("---")

    # --- Main Tabs for Results & Actions ---
    if st.session_state.get('analysis_complete'):
        
        tab_titles = ["📊 SERP & Keywords", "📝 Terms & Structure", "✍️ Generate Content", "🔄 Update Content", "🔗 Internal Linking", "📄 Download Brief"]
        tabs = st.tabs(tab_titles)

        # Tab 1: SERP & Keywords
        with tabs[0]:
            st.subheader("Top Organic Results")
            st.dataframe(pd.DataFrame(st.session_state.results.get('organic_results', [])))
            
            col1, col2 = st.columns(2)
            with col1:
                 st.subheader("Related Keywords")
                 st.dataframe(pd.DataFrame(st.session_state.results.get('related_keywords', [])))
            with col2:
                 st.subheader("People Also Asked")
                 paa = st.session_state.results.get('paa_questions', [])
                 if paa:
                     for q in paa: st.write(f"- {q.get('question', '')}")
                 else: st.write("None found.")


        # Tab 2: Terms & Structure
        with tabs[1]:
             st.subheader("Recommended Content Structure")
             struct = st.session_state.results.get('semantic_structure', {})
             st.markdown(f"**H1:** {struct.get('h1', 'N/A')}")
             for i, sec in enumerate(struct.get('sections', []), 1):
                 st.markdown(f"**H2 ({i}):** {sec.get('h2', 'N/A')}")
                 for j, sub in enumerate(sec.get('subsections', []), 1):
                     st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**H3 ({i}.{j}):** {sub.get('h3', 'N/A')}")
             
             st.subheader("Extracted Terms & Topics")
             term_data = st.session_state.results.get('term_data', {})
             with st.expander("Primary Terms"):
                 st.dataframe(pd.DataFrame(term_data.get('primary_terms', [])))
             with st.expander("Secondary Terms"):
                 st.dataframe(pd.DataFrame(term_data.get('secondary_terms', [])))
             with st.expander("Key Topics"):
                 st.dataframe(pd.DataFrame(term_data.get('topics', [])))
             with st.expander("Questions Answered by Competitors"):
                 questions = term_data.get('questions', [])
                 if questions:
                      for q in questions: st.write(f"- {q}")
                 else: st.write("None extracted.")


        # Tab 3: Generate Content
        with tabs[2]:
            st.subheader("Generate New Article")
            st.warning("This will generate a completely new article based on the analysis.")
            
            if st.button("✨ Generate New Article", key="generate_new_article"):
                 with st.spinner("✍️ Generating article... This may take a few minutes."):
                     start_time_gen = time.time()
                     article_html, gen_status = generate_full_article(
                         st.session_state.results['keyword'],
                         st.session_state.results['semantic_structure'],
                         st.session_state.results.get('related_keywords', []),
                         st.session_state.results.get('paa_questions', []),
                         st.session_state.results['term_data'],
                         st.session_state.results.get('competitor_embeddings', []), # Pass embeddings
                         anthropic_api_key,
                         openai_api_key # Pass openai key
                     )
                     
                     if article_html:
                         st.session_state.results['generated_article_html'] = article_html
                         st.session_state.results['generated_article_change_summary'] = None # No diff for new article
                         st.session_state.results['last_action'] = 'generate'
                         
                         # Score the newly generated content
                         score_data, score_status = score_content(article_html, st.session_state.results['term_data'], st.session_state.results['keyword'])
                         st.session_state.results['generated_article_score'] = score_data
                         if not score_data: st.warning(f"Could not score generated content: {score_status}")
                         
                         st.success(f"Generated new article in {format_time(time.time() - start_time_gen)}.")
                         st.session_state.generation_complete = True
                         # Rerun to display
                         st.experimental_rerun()
                     else:
                         st.error(f"Article generation failed: {gen_status}")


            if st.session_state.get('generation_complete') and st.session_state.results.get('last_action') == 'generate':
                 st.markdown("---")
                 st.subheader("Generated Article Preview")
                 score_data = st.session_state.results.get('generated_article_score')
                 if score_data:
                      score = score_data.get('overall_score', 0); grade = score_data.get('grade','F')
                      st.metric("Content Score", f"{score} ({grade})")
                 
                 # Display article (consider limiting height)
                 article_html_to_display = st.session_state.results.get('generated_article_html','')
                 st.markdown(f'<div style="max-height: 500px; overflow-y: auto; border: 1px solid #eee; padding: 10px;">{article_html_to_display}</div>', unsafe_allow_html=True)
                 
                 # Download button for generated article
                 doc_stream = create_word_document_with_changes(article_html_to_display, st.session_state.results['keyword'])
                 if doc_stream:
                      st.download_button(
                          label="Download Generated Article (.docx)",
                          data=doc_stream,
                          file_name=f"generated_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                          mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                          key="download_generated_article_button"
                      )
                 else: st.error("Failed to create download file for generated article.")


        # Tab 4: Update Content
        with tabs[4]:
            st.subheader("Update Existing Content")
            st.write("Upload your current article (.docx) to get optimization recommendations or generate an updated version.")
            
            uploaded_doc = st.file_uploader("Upload Word Document (.docx)", type="docx", key="update_doc_uploader")
            
            update_option = st.radio("Choose Action:", ("Get Recommendations Only", "Generate Updated Article with Changes"), key="update_option_radio")

            if st.button("Process Existing Content", key="process_update_button"):
                 if not uploaded_doc:
                     st.error("Please upload a document.")
                 else:
                     with st.spinner("🔄 Analyzing your document and preparing updates..."):
                         start_time_update = time.time()
                         
                         # 1. Parse uploaded document
                         existing_content, parse_status = parse_word_document(uploaded_doc)
                         if not existing_content:
                             st.error(f"Failed to parse document: {parse_status}")
                             st.stop()
                         st.session_state.results['existing_content'] = existing_content
                         
                         # 2. Score existing content
                         existing_html = f"<p>{existing_content['full_text'].replace('</p><p>', '</p>\\n<p>').replace('\\n\\n', '</p>\\n<p>')}</p>" # Basic HTML conversion
                         existing_score, score_status = score_content(existing_html, st.session_state.results['term_data'], st.session_state.results['keyword'])
                         st.session_state.results['existing_content_score'] = existing_score
                         if not existing_score: st.warning(f"Could not score existing content: {score_status}")
                         
                         # 3. Perform Gap Analysis
                         content_gaps, gap_status = analyze_content_gaps(
                             existing_content,
                             st.session_state.results['scraped_contents'],
                             st.session_state.results['semantic_structure'],
                             st.session_state.results['term_data'],
                             existing_score if existing_score else {},
                             anthropic_api_key,
                             st.session_state.results['keyword'],
                             st.session_state.results.get('paa_questions', [])
                         )
                         if not content_gaps:
                             st.error(f"Failed to perform gap analysis: {gap_status}")
                             st.stop()
                         st.session_state.results['content_gaps'] = content_gaps
                         
                         # --- Action based on user choice ---
                         if update_option == "Get Recommendations Only":
                             st.session_state.results['last_action'] = 'recommend'
                             # (Display logic handled below)
                             st.success(f"Generated update recommendations in {format_time(time.time() - start_time_update)}.")
                         
                         elif update_option == "Generate Updated Article with Changes":
                             st.session_state.results['last_action'] = 'update'
                             with st.spinner("✍️ Generating optimized article with changes tracked... This may take several minutes."):
                                 optimized_html, change_summary, update_status = generate_optimized_article_with_tracking(
                                     existing_content,
                                     st.session_state.results['scraped_contents'],
                                     st.session_state.results['semantic_structure'],
                                     st.session_state.results['term_data'],
                                     content_gaps,
                                     anthropic_api_key,
                                     st.session_state.results['keyword'],
                                     st.session_state.results.get('competitor_embeddings', []),
                                     openai_api_key
                                 )
                                 if optimized_html:
                                     st.session_state.results['optimized_article_html'] = optimized_html
                                     st.session_state.results['optimized_article_change_summary'] = change_summary
                                     # Score the updated content
                                     updated_score, score_status_upd = score_content(optimized_html, st.session_state.results['term_data'], st.session_state.results['keyword'])
                                     st.session_state.results['optimized_article_score'] = updated_score
                                     if not updated_score: st.warning(f"Could not score updated content: {score_status_upd}")
                                     st.success(f"Generated updated article in {format_time(time.time() - start_time_update)}.")
                                     st.session_state.generation_complete = True # Mark that *an* article exists
                                 else:
                                     st.error(f"Failed to generate updated article: {update_status}")

                         # Rerun to display results
                         st.experimental_rerun()


            # --- Display Update Results ---
            last_action = st.session_state.results.get('last_action')
            
            if last_action == 'recommend':
                st.markdown("---")
                st.subheader("Content Update Recommendations")
                gaps = st.session_state.results.get('content_gaps', {})
                score = st.session_state.results.get('existing_content_score')
                if score: st.metric("Current Content Score", f"{score.get('overall_score',0)} ({score.get('grade','F')})")
                
                # Display gap analysis results (simplified view)
                if gaps.get('missing_sections'): st.warning(f"**Missing Sections:** {len(gaps['missing_sections'])} recommended sections are missing.")
                if gaps.get('revised_headings'): st.info(f"**Heading Revisions:** {len(gaps['revised_headings'])} headings could be improved.")
                if gaps.get('content_gaps'): st.warning(f"**Content Gaps:** {len(gaps['content_gaps'])} specific topic points seem underdeveloped or missing.")
                if gaps.get('expansion_areas'): st.info(f"**Expansion Needed:** {len(gaps['expansion_areas'])} sections could benefit from more detail.")
                if gaps.get('semantic_relevancy_issues'): st.error(f"**Relevancy Issues:** {len(gaps['semantic_relevancy_issues'])} sections may deviate from the core keyword.")
                if gaps.get('term_usage_issues'): st.warning(f"**Term Usage:** {len(gaps['term_usage_issues'])} important terms are missing or underused.")
                if gaps.get('unanswered_paa'): st.info(f"**Unanswered PAA:** {len(gaps['unanswered_paa'])} common questions are not clearly addressed.")
                
                # Offer download for recommendations document (implementation needs a separate function)
                # Placeholder: Need a create_recommendations_document function
                # recommendations_doc = create_recommendations_document(...) 
                # if recommendations_doc: st.download_button(...)


            elif last_action == 'update':
                 st.markdown("---")
                 st.subheader("Updated Article with Changes Tracked")
                 
                 # Display score comparison
                 old_score_data = st.session_state.results.get('existing_content_score')
                 new_score_data = st.session_state.results.get('optimized_article_score')
                 if old_score_data and new_score_data:
                     col1, col2 = st.columns(2)
                     with col1: st.metric("Original Score", f"{old_score_data.get('overall_score',0)} ({old_score_data.get('grade','F')})")
                     with col2: st.metric("Updated Score", f"{new_score_data.get('overall_score',0)} ({new_score_data.get('grade','F')})", delta=f"{new_score_data.get('overall_score',0) - old_score_data.get('overall_score',0)} pts")

                 # Display change summary and content
                 summary_html = st.session_state.results.get('optimized_article_change_summary', '<p>Change summary not available.</p>')
                 article_html_diff = st.session_state.results.get('optimized_article_html', '<p>Updated article not available.</p>')
                 
                 st.markdown(summary_html, unsafe_allow_html=True)
                 st.markdown(f'<div style="max-height: 500px; overflow-y: auto; border: 1px solid #eee; padding: 10px; margin-top:10px;">{article_html_diff}</div>', unsafe_allow_html=True)

                 # Download button for updated article
                 doc_stream = create_word_document_with_changes(article_html_diff, st.session_state.results['keyword'], summary_html)
                 if doc_stream:
                      st.download_button(
                          label="Download Updated Article (.docx)",
                          data=doc_stream,
                          file_name=f"updated_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                          mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                          key="download_updated_article_button"
                      )
                 else: st.error("Failed to create download file for updated article.")


        # Tab 5: Internal Linking
        with tabs[5]:
            st.subheader("Suggest Internal Links")
            
            article_available = st.session_state.results.get('generated_article_html') or st.session_state.results.get('optimized_article_html')
            if not article_available:
                 st.warning("Generate or update an article first before suggesting internal links.")
            else:
                st.write("Upload a spreadsheet (CSV/Excel) of your site's pages containing columns: `URL`, `Title`, `Meta Description`.")
                
                # Sample template download
                if st.button("Download Site Pages Template (.csv)", key="download_template_button"):
                    sample_df = pd.DataFrame({
                        'URL': ['https://example.com/page1', 'https://example.com/page2'],
                        'Title': ['Example Page 1 Title', 'Example Page 2 Title'],
                        'Meta Description': ['Meta description for example page 1.', 'Meta description for example page 2.']
                    })
                    csv_data = sample_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Click to Download Template", # Button appears after click
                        data=csv_data,
                        file_name="site_pages_template.csv",
                        mime="text/csv",
                        key="template_download_actual"
                    )

                pages_file = st.file_uploader("Upload Site Pages File", type=['csv', 'xlsx', 'xls'], key="site_pages_uploader")
                
                if st.button("🔗 Find Linking Opportunities", key="find_links_button"):
                    if not pages_file:
                        st.error("Please upload the site pages file.")
                    else:
                         with st.spinner("🔄 Processing pages and finding link opportunities..."):
                             start_time_link = time.time()
                             
                             # Determine which article content to use
                             article_to_link = st.session_state.results.get('optimized_article_html') or st.session_state.results.get('generated_article_html')
                             
                             # 1. Parse pages file
                             site_pages, parse_status = parse_site_pages_spreadsheet(pages_file)
                             if not site_pages:
                                 st.error(f"Failed to process site pages file: {parse_status}")
                                 st.stop()
                                 
                             # 2. Embed site pages
                             pages_embed, embed_status = embed_site_pages(site_pages, openai_api_key)
                             if not pages_embed:
                                 st.error(f"Failed to embed site pages: {embed_status}")
                                 st.stop()
                                 
                             # 3. Generate links
                             article_linked_html, links_added, link_status = generate_internal_links_with_embeddings(
                                 article_to_link,
                                 pages_embed,
                                 openai_api_key,
                                 count_words(clean_html(article_to_link))
                             )
                             
                             if link_status.startswith("Success"):
                                 st.session_state.results['article_with_links'] = article_linked_html
                                 st.session_state.results['internal_links_added'] = links_added
                                 st.success(f"Found {len(links_added)} potential internal links in {format_time(time.time() - start_time_link)}.")
                                 # Rerun to display results
                                 st.experimental_rerun()
                             else:
                                 st.error(f"Internal linking failed: {link_status}")

            # Display linking results
            if 'internal_links_added' in st.session_state.results:
                 st.markdown("---")
                 st.subheader("Suggested Internal Links")
                 links = st.session_state.results['internal_links_added']
                 if not links:
                      st.info("No suitable internal linking opportunities were found based on the criteria.")
                 else:
                     st.dataframe(pd.DataFrame(links))
                     st.markdown("---")
                     st.subheader("Article Preview with Links")
                     article_linked_display = st.session_state.results.get('article_with_links','')
                     st.markdown(f'<div style="max-height: 500px; overflow-y: auto; border: 1px solid #eee; padding: 10px;">{article_linked_display}</div>', unsafe_allow_html=True)


        # Tab 6: Download Brief
        with tabs[6]:
            st.subheader("Download Full SEO Brief")
            st.write("Generates a comprehensive Word document containing all analysis and generated content.")
            
            if st.button("📄 Generate and Download Brief (.docx)", key="generate_brief_button"):
                 with st.spinner("Creating Word document..."):
                     # Determine which content to include
                     article_content = st.session_state.results.get('article_with_links') or \
                                       st.session_state.results.get('optimized_article_html') or \
                                       st.session_state.results.get('generated_article_html') or \
                                       "" # Fallback to empty if nothing generated/updated

                     # Determine score to include
                     score_data = st.session_state.results.get('optimized_article_score') or \
                                  st.session_state.results.get('generated_article_score') # Prioritize optimized/generated score
                     
                     # Links to include
                     internal_links = st.session_state.results.get('internal_links_added')

                     doc_stream, doc_status = create_word_document(
                         st.session_state.results['keyword'],
                         st.session_state.results['organic_results'],
                         st.session_state.results.get('related_keywords', []),
                         st.session_state.results['semantic_structure'],
                         article_content,
                         st.session_state.results.get('meta_title',''),
                         st.session_state.results.get('meta_description',''),
                         st.session_state.results.get('paa_questions', []),
                         st.session_state.results.get('term_data'),
                         score_data,
                         internal_links,
                         guidance_only=False # Brief always assumes full content for now
                     )
                     
                     if doc_stream:
                          st.session_state.results['final_brief_stream'] = doc_stream
                          st.success("Brief generated successfully!")
                          # Trigger download immediately after generation
                          st.download_button(
                              label="Download Brief Now", # Keep label consistent
                              data=doc_stream,
                              file_name=f"SEO_Brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                              mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                              key="final_brief_download_button"
                          )
                     else:
                          st.error(f"Failed to generate brief: {doc_status}")

            # Offer download if already generated
            if 'final_brief_stream' in st.session_state.results:
                st.download_button(
                     label="Download Previously Generated Brief",
                     data=st.session_state.results['final_brief_stream'],
                     file_name=f"SEO_Brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     key="download_previous_brief_button"
                 )

    else:
        st.info("Enter a keyword and click 'Analyze SERP & Competitors' to begin.")

if __name__ == "__main__":
    main()
#==============================================================================
# End of Chunk 4/4
#==============================================================================
```
