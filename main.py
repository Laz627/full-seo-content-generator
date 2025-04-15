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
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT # Corrected import
from io import BytesIO
import base64
import random
from typing import List, Dict, Any, Tuple, Optional
import logging
import traceback
import openpyxl
import altair as alt # Kept as it was in the original provided code for the scoring chart

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Set page configuration
st.set_page_config(
    page_title="SEO Content Optimizer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

###############################################################################
# 1. Utility Functions
###############################################################################

def display_error(error_msg: str):
    """Display error message in Streamlit"""
    st.error(f"Error: {error_msg}")
    logger.error(error_msg)

def get_download_link(file_bytes: bytes, filename: str, link_label: str) -> str:
    """Returns an HTML link to download file"""
    b64 = base64.b64encode(file_bytes).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{link_label}</a>'
    return href

def format_time(seconds: float) -> str:
    """Format time in seconds to readable format"""
    if seconds < 60:
        return f"{seconds:.1f} seconds"
    else:
        minutes = int(seconds // 60)
        sec = seconds % 60
        return f"{minutes} min {sec:.1f} sec"

###############################################################################
# 2. API Integration - DataForSEO
###############################################################################

def classify_page_type(url: str, title: str, snippet: str) -> str:
    """
    Simplified page type classification based on title and snippet patterns
    Returns: page_type
    """
    title_lower = title.lower() if title else ""
    snippet_lower = snippet.lower() if snippet else ""
    url_lower = url.lower() if url else ""

    # E-commerce indicators
    commerce_patterns = ['buy', 'shop', 'purchase', 'cart', 'checkout', 'price', 'discount', 'sale', 'product', 'order']
    if any(pattern in title_lower or pattern in snippet_lower or pattern in url_lower for pattern in commerce_patterns):
        return "E-commerce"

    # Article/Blog indicators
    article_patterns = ['blog', 'article', 'news', 'post', 'how to', 'guide', 'tips', 'tutorial', 'learn']
    if any(pattern in title_lower or pattern in snippet_lower or pattern in url_lower for pattern in article_patterns):
        return "Article/Blog"

    # Forum/Community indicators
    forum_patterns = ['forum', 'community', 'discussion', 'thread', 'reply', 'comment', 'question', 'answer']
    if any(pattern in title_lower or pattern in snippet_lower or pattern in url_lower for pattern in forum_patterns):
        return "Forum/Community"

    # Review indicators
    review_patterns = ['review', 'comparison', 'vs', 'versus', 'top 10', 'best', 'rating', 'rated']
    if any(pattern in title_lower or pattern in snippet_lower or pattern in url_lower for pattern in review_patterns):
        return "Review/Comparison"

    # Default to informational
    return "Informational"

def fetch_serp_results(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], List[Dict], List[Dict], bool]:
    """
    Fetch SERP results from DataForSEO API and classify pages
    Returns: organic_results, serp_features, paa_questions, success_status
    """
    try:
        url = "https://api.dataforseo.com/v3/serp/google/organic/live/advanced"
        headers = {
            'Content-Type': 'application/json',
        }

        # Prepare request data
        post_data = [{
            "keyword": keyword,
            "location_code": 2840,  # USA
            "language_code": "en",
            "device": "desktop",
            "os": "windows",
            "depth": 30  # Get more results to ensure we have at least 10 organic
        }]

        # Make API request
        response = requests.post(
            url,
            auth=(api_login, api_password),
            headers=headers,
            json=post_data
        )

        # Process response
        if response.status_code == 200:
            data = response.json()
            logger.info(f"SERP API Response status: {data.get('status_code')}")

            if data.get('status_code') == 20000:
                results = data['tasks'][0]['result'][0]

                # Extract organic results
                organic_results = []
                for item in results.get('items', []):
                    if item.get('type') == 'organic':
                        if len(organic_results) < 10:  # Limit to top 10
                            # Get title and snippet for classification
                            title = item.get('title', '')
                            snippet = item.get('snippet', '')
                            url = item.get('url', '')

                            # Classify page type using pattern matching
                            page_type = classify_page_type(url, title, snippet)

                            organic_results.append({
                                'url': url,
                                'title': title,
                                'snippet': snippet,
                                'rank_group': item.get('rank_group'),
                                'page_type': page_type
                            })

                # Extract People Also Asked questions - handle both possible formats
                paa_questions = []

                # Method 1: Look for a PAA container
                for item in results.get('items', []):
                    if item.get('type') == 'people_also_ask':
                        logger.info(f"Found PAA container with {len(item.get('items', []))} questions")

                        # Process PAA items within the container
                        for paa_item in item.get('items', []):
                            if paa_item.get('type') == 'people_also_ask_element':
                                question_data = {
                                    'question': paa_item.get('title', ''),
                                    'expanded': []
                                }

                                # Extract expanded element data if available
                                expanded_elements = paa_item.get('expanded_element', [])
                                if expanded_elements: # Check if list exists and is not empty
                                    for expanded in expanded_elements:
                                        if expanded.get('type') == 'people_also_ask_expanded_element':
                                            question_data['expanded'].append({
                                                'url': expanded.get('url', ''),
                                                'title': expanded.get('title', ''),
                                                'description': expanded.get('description', '')
                                            })

                                paa_questions.append(question_data)

                # Method 2 (fallback): Look for individual PAA elements directly in items
                if not paa_questions:
                    logger.info("No PAA container found, looking for individual PAA elements")
                    for item in results.get('items', []):
                        if item.get('type') == 'people_also_ask_element':
                            question_data = {
                                'question': item.get('title', ''),
                                'expanded': []
                            }

                            # Extract expanded element data if available
                            expanded_elements = item.get('expanded_element', [])
                            if expanded_elements: # Check if list exists and is not empty
                                for expanded in expanded_elements:
                                    if expanded.get('type') == 'people_also_ask_expanded_element':
                                        question_data['expanded'].append({
                                            'url': expanded.get('url', ''),
                                            'title': expanded.get('title', ''),
                                            'description': expanded.get('description', '')
                                        })

                            paa_questions.append(question_data)

                # Log how many PAA questions we found
                logger.info(f"Extracted {len(paa_questions)} PAA questions")

                # Extract SERP features
                serp_features = []
                feature_counts = {}
                for item in results.get('items', []):
                    item_type = item.get('type')
                    if item_type != 'organic':
                        if item_type in feature_counts:
                            feature_counts[item_type] += 1
                        else:
                            feature_counts[item_type] = 1

                for feature, count in feature_counts.items():
                    serp_features.append({
                        'feature_type': feature,
                        'count': count
                    })

                # Sort by count and limit to top 20
                serp_features = sorted(serp_features, key=lambda x: x['count'], reverse=True)[:20]

                return organic_results, serp_features, paa_questions, True
            else:
                error_msg = f"API Error: {data.get('status_message')}"
                logger.error(error_msg)
                return [], [], [], False
        else:
            error_msg = f"HTTP Error: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return [], [], [], False

    except Exception as e:
        error_msg = f"Exception in fetch_serp_results: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return [], [], [], False

###############################################################################
# 3. API Integration - DataForSEO for Keywords
###############################################################################

def fetch_keyword_suggestions(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], bool]:
    """
    Fetch keyword suggestions from DataForSEO to get accurate search volume data
    Returns: keyword_suggestions, success_status
    """
    try:
        url = "https://api.dataforseo.com/v3/dataforseo_labs/google/keyword_suggestions/live"
        headers = {
            'Content-Type': 'application/json',
        }

        # Prepare request data based on the sample JSON structure
        post_data = [{
            "keyword": keyword,
            "location_code": 2840,  # USA
            "language_code": "en",
            "include_serp_info": True, # Keep this as it might influence results
            "include_seed_keyword": True,
            "limit": 20  # Fetch top 20 suggestions
        }]

        # Log the request for debugging
        logger.info(f"Fetching keyword suggestions for: {keyword}")

        # Make API request
        response = requests.post(
            url,
            auth=(api_login, api_password),
            headers=headers,
            json=post_data
        )

        # Process response
        if response.status_code == 200:
            data = response.json()
            logger.info(f"Keyword Suggestions API Response status: {data.get('status_code')}")

            # Validate response
            if data.get('status_code') != 20000 or not data.get('tasks') or len(data['tasks']) == 0:
                logger.warning(f"Invalid API response for suggestions: {data.get('status_message')}")
                return [], False

            keyword_suggestions = []

            # Process results based on the specific JSON structure from sample
            for task in data['tasks']:
                if not task.get('result'):
                    continue

                # Results are within a list, iterate through it
                for result_item in task['result']:
                    # Check for seed keyword data first if include_seed_keyword is True
                    if 'seed_keyword_data' in result_item and result_item['seed_keyword_data']:
                        seed_data = result_item['seed_keyword_data']
                        if 'keyword_info' in seed_data:
                            keyword_info = seed_data['keyword_info']
                            keyword_suggestions.append({
                                'keyword': result_item.get('seed_keyword', keyword), # Use original if 'seed_keyword' missing
                                'search_volume': keyword_info.get('search_volume'), # Keep None if missing
                                'cpc': keyword_info.get('cpc'),
                                'competition': keyword_info.get('competition')
                            })

                    # Then look for items array which contains related keywords
                    if 'items' in result_item and isinstance(result_item['items'], list):
                        for item in result_item['items']:
                            if 'keyword_info' in item:
                                keyword_info = item['keyword_info']
                                keyword_suggestions.append({
                                    'keyword': item.get('keyword', ''),
                                    'search_volume': keyword_info.get('search_volume'),
                                    'cpc': keyword_info.get('cpc'),
                                    'competition': keyword_info.get('competition')
                                })

            # Check if we successfully found keywords
            if keyword_suggestions:
                # Filter out potential duplicates (e.g., seed keyword appearing again)
                seen_keywords = set()
                unique_suggestions = []
                for sugg in keyword_suggestions:
                    kw = sugg.get('keyword', '').lower()
                    if kw and kw not in seen_keywords:
                        unique_suggestions.append(sugg)
                        seen_keywords.add(kw)

                # Sort by search volume (descending), handle None values
                unique_suggestions.sort(key=lambda x: x.get('search_volume', 0) or 0, reverse=True)
                logger.info(f"Successfully extracted {len(unique_suggestions)} unique keyword suggestions")
                return unique_suggestions, True
            else:
                logger.warning(f"No keyword suggestions found in the response items for '{keyword}'")
                return [], True # Still success, just no suggestions found
        else:
            error_msg = f"HTTP Error fetching keyword suggestions: {response.status_code} - {response.text}"
            logger.error(error_msg)
            return [], False

    except Exception as e:
        error_msg = f"Exception in fetch_keyword_suggestions for {keyword}: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return [], False

def fetch_related_keywords_dataforseo(keyword: str, api_login: str, api_password: str) -> Tuple[List[Dict], bool]:
    """
    Fetch related keywords from DataForSEO Related Keywords API.
    Falls back to Keyword Suggestions if Related Keywords fails or returns no data.
    Returns: related_keywords, success_status
    """
    logger.info(f"Attempting to fetch related keywords for: {keyword}")
    try:
        url = "https://api.dataforseo.com/v3/dataforseo_labs/google/related_keywords/live"
        headers = {'Content-Type': 'application/json'}
        post_data = [{"keyword": keyword, "language_name": "English", "location_code": 2840, "limit": 20}]

        response = requests.post(url, auth=(api_login, api_password), headers=headers, json=post_data)

        if response.status_code == 200:
            data = response.json()
            logger.info(f"Related Keywords API Response status: {data.get('status_code')}")

            if data.get('status_code') == 20000 and data.get('tasks'):
                related_keywords = []
                for task in data['tasks']:
                     if task.get('result'):
                         # Results are within a list, iterate through it
                         for result_item in task['result']:
                            if 'items' in result_item and isinstance(result_item['items'], list):
                                for item in result_item['items']:
                                     # Structure seems to be: items -> keyword_data -> keyword_info
                                     if 'keyword_data' in item and 'keyword_info' in item['keyword_data']:
                                         kw_data = item['keyword_data']
                                         keyword_info = kw_data['keyword_info']
                                         related_keywords.append({
                                             'keyword': kw_data.get('keyword', ''),
                                             'search_volume': keyword_info.get('search_volume'),
                                             'cpc': keyword_info.get('cpc'),
                                             'competition': keyword_info.get('competition')
                                         })

                if related_keywords:
                    # Filter out potential duplicates
                    seen_keywords = set()
                    unique_keywords = []
                    for kw_data in related_keywords:
                        kw = kw_data.get('keyword', '').lower()
                        if kw and kw not in seen_keywords:
                            unique_keywords.append(kw_data)
                            seen_keywords.add(kw)

                    # Sort by search volume, handle None values
                    unique_keywords.sort(key=lambda x: x.get('search_volume', 0) or 0, reverse=True)
                    logger.info(f"Successfully extracted {len(unique_keywords)} unique related keywords")
                    return unique_keywords, True
                else:
                    logger.warning(f"No related keywords found for '{keyword}', falling back to suggestions.")
                    return fetch_keyword_suggestions(keyword, api_login, api_password)
            else:
                logger.warning(f"Related Keywords API Error: {data.get('status_message')}. Falling back to suggestions.")
                return fetch_keyword_suggestions(keyword, api_login, api_password)
        else:
            error_msg = f"HTTP Error fetching related keywords: {response.status_code} - {response.text}. Falling back to suggestions."
            logger.error(error_msg)
            return fetch_keyword_suggestions(keyword, api_login, api_password)

    except Exception as e:
        error_msg = f"Exception in fetch_related_keywords_dataforseo for {keyword}: {str(e)}. Falling back to suggestions."
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return fetch_keyword_suggestions(keyword, api_login, api_password)

###############################################################################
# 4. Web Scraping and Content Analysis
###############################################################################

def scrape_webpage(url: str) -> Tuple[str, bool]:
    """
    Enhanced webpage scraping with better error handling and User-Agent rotation.
    Returns: content (string), success_status (boolean)
    """
    logger.info(f"Attempting to scrape: {url}")
    # Try trafilatura first
    try:
        downloaded = trafilatura.fetch_url(url)
        if downloaded:
            # Extract main content using trafilatura settings
            content = trafilatura.extract(downloaded, include_comments=False, include_tables=False, # Usually don't need tables/comments
                                          output_format='txt', # Get plain text
                                          favor_precision=True) # Try to be more precise
            if content and len(content) > 100:  # Basic check for meaningful content
                logger.info(f"Successfully scraped (Trafilatura): {url}")
                # Simple text cleaning
                content = re.sub(r'\s+\n', '\n', content) # Consolidate whitespace before newlines
                content = re.sub(r'\n{3,}', '\n\n', content) # Limit consecutive newlines
                return content.strip(), True
            else:
                 logger.warning(f"Trafilatura extracted little/no content from: {url}")
        else:
            logger.warning(f"Trafilatura failed to download: {url}")

    except Exception as e:
        logger.warning(f"Trafilatura exception for {url}: {e}")

    # Fallback to requests + BeautifulSoup
    logger.info(f"Falling back to requests+BeautifulSoup for: {url}")
    try:
        user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:108.0) Gecko/20100101 Firefox/108.0'
        ]
        headers = {
            'User-Agent': random.choice(user_agents),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': 'https://www.google.com/',
            'DNT': '1' # Do Not Track
        }
        response = requests.get(url, headers=headers, timeout=20, allow_redirects=True)
        response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)

        soup = BeautifulSoup(response.text, 'html.parser')

        # Remove script, style, nav, footer, header, forms, etc.
        for element in soup(["script", "style", "nav", "footer", "header", "aside", "form", "button", "iframe", "meta", "link"]):
            element.decompose()

        # Try finding common main content containers
        main_content = soup.find('main') or \
                       soup.find('article') or \
                       soup.find('div', role='main') or \
                       soup.find('div', id='main-content') or \
                       soup.find('div', class_='content') or \
                       soup.find('div', class_='post-content') or \
                       soup.find('div', class_='entry-content')

        if main_content:
            text = main_content.get_text(separator='\n', strip=True)
        else:
             # Fallback to body if specific containers not found
             body = soup.find('body')
             text = body.get_text(separator='\n', strip=True) if body else ""

        # Further clean up extracted text
        lines = (line.strip() for line in text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        text = '\n'.join(chunk for chunk in chunks if chunk) # Rejoin non-empty chunks

        if text and len(text) > 100:
             logger.info(f"Successfully scraped (BeautifulSoup): {url}")
             return text, True
        else:
            logger.warning(f"BeautifulSoup extracted little/no content from: {url}. Body length: {len(response.text)}")
            return "[Content not found or extracted]", False # More specific message

    except requests.exceptions.RequestException as e:
        logger.error(f"Request failed for {url}: {e}")
        return f"[Error retrieving content: {type(e).__name__}]", False
    except Exception as e:
        error_msg = f"Exception in BeautifulSoup fallback for {url}: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return f"[Error parsing content: {str(e)}]", False

def extract_headings(url: str) -> Dict[str, List[str]]:
    """
    Extract headings (H1, H2, H3) from a webpage. Uses similar request logic as scrape_webpage.
    """
    logger.info(f"Extracting headings from: {url}")
    headings = {'h1': [], 'h2': [], 'h3': []} # Default empty
    try:
        user_agents = [ # Keep list consistent with scrape_webpage
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:108.0) Gecko/20100101 Firefox/108.0'
        ]
        headers = {
            'User-Agent': random.choice(user_agents),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': 'https://www.google.com/',
             'DNT': '1'
        }
        response = requests.get(url, headers=headers, timeout=15, allow_redirects=True)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')

        for level in ['h1', 'h2', 'h3']:
            found_headings = soup.find_all(level)
            # Clean heading text: remove extra spaces and potential inline tags
            headings[level] = [re.sub(r'\s+', ' ', h.get_text()).strip() for h in found_headings if h.get_text().strip()]

        logger.info(f"Found headings for {url}: H1={len(headings['h1'])}, H2={len(headings['h2'])}, H3={len(headings['h3'])}")
        return headings

    except requests.exceptions.RequestException as e:
        logger.error(f"Request failed for headings extraction {url}: {e}")
        return headings # Return default empty dict on request failure
    except Exception as e:
        error_msg = f"Exception in extract_headings for {url}: {str(e)}"
        logger.error(error_msg)
        return headings # Return default empty dict on parsing failure


###############################################################################
# 5. Content Scoring Functions (Assumed Mostly Correct from Original)
###############################################################################

def extract_important_terms(competitor_contents: List[Dict], anthropic_api_key: str) -> Tuple[Dict, bool]:
    """
    Extract important terms and topics from competitor content using Claude.
    Returns: term_data (Dict), success_status (bool)
    """
    logger.info("Extracting important terms using Claude.")
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)

        # Combine content, prioritizing earlier parts if too long
        max_context_length = 18000 # Max chars for Claude context (adjust if needed)
        combined_content = ""
        for c in competitor_contents:
            content_piece = c.get('content', '')
            if content_piece:
                 if len(combined_content) + len(content_piece) < max_context_length:
                     combined_content += content_piece + "\n\n---\n\n"
                 else:
                     remaining_space = max_context_length - len(combined_content)
                     if remaining_space > 200: # Add partial if space allows
                         combined_content += content_piece[:remaining_space] + "...\n\n---\n\n"
                     break # Stop adding content

        if not combined_content:
             logger.error("No competitor content available to extract terms from.")
             return {}, False

        system_prompt = "You are an SEO expert analyzing competitor content. Extract key terms, topics, and questions based on the provided text. Output ONLY a valid JSON object matching the specified format exactly."
        user_prompt = f"""
        Analyze the following combined content from top-ranking pages and extract:

        1. Primary terms (most important concepts, max 15)
        2. Secondary terms (supporting keywords/phrases, max 25)
        3. Questions likely answered or implied (max 15)
        4. Key Topics covered comprehensively (max 15)

        For terms, estimate importance (0.0-1.0) and suggest a reasonable usage count per ~1000 words.

        Format your response strictly as JSON:
        ```json
        {{
            "primary_terms": [
                {{"term": "term1", "importance": 0.95, "recommended_usage": 5}},
                {{"term": "term2", "importance": 0.85, "recommended_usage": 3}}
            ],
            "secondary_terms": [
                {{"term": "termA", "importance": 0.75, "recommended_usage": 2}},
                {{"term": "termB", "importance": 0.65, "recommended_usage": 1}}
            ],
            "questions": [
                "Question 1?",
                "Question 2?"
            ],
            "topics": [
                {{"topic": "Topic A", "description": "Briefly describe what this topic covers in the content."}},
                {{"topic": "Topic B", "description": "Briefly describe what this topic covers."}}
            ]
        }}
        ```

        Content to analyze:
        {combined_content}
        """

        response = client.messages.create(
            model="claude-3-7-sonnet-20250219", # Or 3.5 model
            max_tokens=2000, # Ample space for JSON output
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
            temperature=0.1 # Low temp for structured output
        )

        raw_response_text = response.content[0].text
        logger.debug(f"Raw response from Claude for term extraction:\n{raw_response_text[:500]}...")

        # Robust JSON Extraction
        term_data = None
        json_match = re.search(r'```json\s*(\{.*?\})\s*```', raw_response_text, re.DOTALL)
        if json_match:
            json_string = json_match.group(1)
            try:
                term_data = json.loads(json_string)
                logger.info("Successfully parsed terms JSON from Claude response.")
            except json.JSONDecodeError as e:
                logger.error(f"Failed to decode terms JSON: {e}")
                 # Add repair attempt if needed
        else:
             logger.warning("Could not find ```json ... ``` for terms extraction.")
              # Fallback attempt if needed

        if term_data and isinstance(term_data, dict):
            # Basic validation
            if not term_data.get("primary_terms") and not term_data.get("secondary_terms"):
                logger.warning("Extracted term data seems empty.")
                return term_data, False # Or True if empty is acceptable
            return term_data, True
        else:
            logger.error("Failed to extract valid term data dictionary from Claude.")
            return {}, False

    except anthropic.APIError as e:
        error_msg = f"Anthropic API error during term extraction: {e}"
        logger.error(error_msg)
        return {}, False
    except Exception as e:
        error_msg = f"Exception in extract_important_terms: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {}, False

def score_content(content: str, term_data: Dict, keyword: str) -> Tuple[Dict, bool]:
    """
    Score content based on keyword usage, semantic relevance, and comprehensiveness, using term_data.
    Accepts HTML content and extracts text for analysis.
    Returns: score_data (Dict), success_status (boolean)
    """
    logger.info("Scoring content...")
    if not isinstance(term_data, dict) or not term_data:
        logger.error("Invalid or empty term_data provided for scoring.")
        return {'overall_score': 0, 'grade': 'F', 'error': 'Missing term data'}, False

    try:
        # Extract plain text from HTML for analysis
        soup = BeautifulSoup(content, 'html.parser')
        plain_content = soup.get_text(separator=' ', strip=True)
        content_lower = plain_content.lower()
        word_count = len(re.findall(r'\b\w+\b', content_lower))

        if word_count == 0:
            logger.warning("Content for scoring has zero words.")
            return {'overall_score': 0, 'grade': 'F', 'components': {}, 'details': {'word_count': 0}}, False # Return basic structure

        # --- Scoring Logic (based on original, ensure robustness) ---
        keyword_score = 0
        primary_terms_score = 0
        secondary_terms_score = 0
        topic_coverage_score = 0
        question_coverage_score = 0

        # 1. Keyword Score
        keyword_count = len(re.findall(r'\b' + re.escape(keyword.lower()) + r'\b', content_lower))
        # Adjusted optimal range: min 1 occurrence, density target 0.5% to 1.5%
        optimal_min = 1
        optimal_max = max(optimal_min + 1, int(word_count * 0.015)) # Allow slightly higher density
        target_density = 0.01 # Aim for ~1%

        if keyword_count == 0:
             keyword_score = 0
        elif keyword_count < optimal_min: # Should not happen if optimal_min is 1
             keyword_score = 30 # Penalize heavily if below minimum
        elif keyword_count <= optimal_max:
             # Score higher the closer to the target density, max 100 at optimal_max
             density = keyword_count / word_count
             proximity_bonus = 1.0 - abs(density - target_density) / target_density
             keyword_score = max(50, min(100, 50 + 50 * proximity_bonus)) # Scale from 50 to 100 based on density proximity
        else: # Over optimal max - penalize gradually
             overuse_ratio = keyword_count / optimal_max
             penalty = min(40, (overuse_ratio - 1) * 30) # Penalize up to 40 points
             keyword_score = max(0, 100 - penalty - 20) # Apply base penalty + overuse penalty

        # 2. Primary Terms Score
        primary_term_counts = {}
        primary_terms_data = term_data.get('primary_terms', [])
        primary_terms_total = len(primary_terms_data)
        primary_terms_found_weighted = 0.0

        if primary_terms_total > 0:
            for term_info in primary_terms_data:
                term = term_info.get('term')
                importance = term_info.get('importance', 0.5) # Default importance
                recommended = term_info.get('recommended_usage', 1)
                if term:
                    count = len(re.findall(r'\b' + re.escape(term.lower()) + r'\b', content_lower))
                    primary_term_counts[term] = {'count': count, 'importance': importance, 'recommended': recommended}
                    if count > 0:
                         # Weighted score: Found * Importance
                         # Bonus for meeting/exceeding recommendation (up to a limit)
                         usage_bonus = min(1.2, (count / recommended) if recommended > 0 else 1.0) # Cap bonus at 20%
                         primary_terms_found_weighted += importance * usage_bonus

            # Normalize weighted score to 0-100 range
            max_possible_weighted_score = sum(t.get('importance', 0.5) * 1.2 for t in primary_terms_data) # Max score with bonus
            if max_possible_weighted_score > 0:
                primary_terms_score = min(100, (primary_terms_found_weighted / max_possible_weighted_score) * 100)
            else:
                primary_terms_score = 100 # If no primary terms defined, score is 100

        # 3. Secondary Terms Score (Similar logic, maybe lower weight on bonus)
        secondary_term_counts = {}
        secondary_terms_data = term_data.get('secondary_terms', [])
        secondary_terms_total = len(secondary_terms_data)
        secondary_terms_found_weighted = 0.0

        if secondary_terms_total > 0:
            for term_info in secondary_terms_data:
                term = term_info.get('term')
                importance = term_info.get('importance', 0.3) # Lower default importance
                recommended = term_info.get('recommended_usage', 1)
                if term:
                    count = len(re.findall(r'\b' + re.escape(term.lower()) + r'\b', content_lower))
                    secondary_term_counts[term] = {'count': count, 'importance': importance, 'recommended': recommended}
                    if count > 0:
                         usage_bonus = min(1.1, (count / recommended) if recommended > 0 else 1.0) # Lower bonus cap
                         secondary_terms_found_weighted += importance * usage_bonus

            max_possible_weighted_score = sum(t.get('importance', 0.3) * 1.1 for t in secondary_terms_data)
            if max_possible_weighted_score > 0:
                secondary_terms_score = min(100, (secondary_terms_found_weighted / max_possible_weighted_score) * 100)
            else:
                 secondary_terms_score = 100

        # 4. Topic Coverage Score
        topic_coverage = {}
        topics_data = term_data.get('topics', [])
        topics_total = len(topics_data)
        topics_covered_count = 0

        if topics_total > 0:
            for topic_info in topics_data:
                topic = topic_info.get('topic')
                description = topic_info.get('description', '')
                if topic:
                    # Simple check: presence of topic name or keywords from description
                    topic_lower = topic.lower()
                    desc_words = set(re.findall(r'\b\w{4,}\b', description.lower())) # Keywords from description
                    
                    # Check if topic name is present or if >50% of description keywords are present
                    covered = topic_lower in content_lower or \
                              (desc_words and sum(1 for w in desc_words if w in content_lower) / len(desc_words) > 0.5)
                    
                    topic_coverage[topic] = {'covered': covered, 'description': description}
                    if covered:
                        topics_covered_count += 1

            topic_coverage_score = (topics_covered_count / topics_total) * 100 if topics_total > 0 else 100

        # 5. Question Coverage Score
        question_coverage = {}
        questions_data = term_data.get('questions', [])
        questions_total = len(questions_data)
        questions_answered_count = 0

        if questions_total > 0:
            for question in questions_data:
                question_lower = question.lower().replace('?','')
                # Simple check: Look for core non-stop words from the question in the content
                q_words = set(re.findall(r'\b\w{4,}\b', question_lower)) # Significant words
                common_words = {'what', 'when', 'where', 'which', 'who', 'why', 'how', 'your', 'does', 'this', 'that'}
                q_words -= common_words

                answered = False
                if q_words:
                     matches = sum(1 for word in q_words if word in content_lower)
                     match_ratio = matches / len(q_words)
                     answered = match_ratio > 0.6 # Needs >60% of keywords present
                else: # If question has no significant words, check if full phrase (minus ?) is present
                    answered = question_lower in content_lower

                question_coverage[question] = {'answered': answered}
                if answered:
                    questions_answered_count += 1

            question_coverage_score = (questions_answered_count / questions_total) * 100 if questions_total > 0 else 100

        # Calculate overall score (Adjusted weights)
        overall_score = (
            keyword_score * 0.15 +          # Lower weight for exact keyword
            primary_terms_score * 0.35 +    # High weight for primary terms
            secondary_terms_score * 0.20 +  # Moderate weight for secondary
            topic_coverage_score * 0.20 +   # Moderate weight for topics
            question_coverage_score * 0.10  # Lower weight for questions
        )
        overall_score = round(max(0, min(100, overall_score))) # Ensure score is 0-100


        # Compile detailed results
        score_data_out = {
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
                'keyword_count': keyword_count,
                'optimal_keyword_range': f"{optimal_min}-{optimal_max}",
                'primary_terms_found': sum(1 for info in primary_term_counts.values() if info['count'] > 0),
                'primary_terms_total': primary_terms_total,
                'primary_term_counts': primary_term_counts,
                'secondary_terms_found': sum(1 for info in secondary_term_counts.values() if info['count'] > 0),
                'secondary_terms_total': secondary_terms_total,
                'secondary_term_counts': secondary_term_counts,
                'topics_covered': topics_covered_count,
                'topics_total': topics_total,
                'topic_coverage': topic_coverage,
                'questions_answered': questions_answered_count,
                'questions_total': questions_total,
                'question_coverage': question_coverage
            }
        }
        logger.info(f"Content scoring complete. Overall score: {overall_score}")
        return score_data_out, True

    except Exception as e:
        error_msg = f"Exception in score_content: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {'overall_score': 0, 'grade': 'F', 'error': str(e)}, False

def get_score_grade(score: float) -> str:
    """Convert numeric score to letter grade"""
    if score >= 90: return "A+"
    elif score >= 85: return "A"
    elif score >= 80: return "A-"
    elif score >= 75: return "B+"
    elif score >= 70: return "B"
    elif score >= 65: return "B-"
    elif score >= 60: return "C+"
    elif score >= 55: return "C"
    elif score >= 50: return "C-"
    elif score >= 40: return "D"
    else: return "F"

def highlight_keywords_in_content(content: str, term_data: Dict, keyword: str) -> Tuple[str, bool]:
    """
    Highlight primary and secondary keywords in HTML content with different background colors.
    Uses span tags. Handles potential HTML tags within the content carefully.
    Returns: highlighted_html (string), success_status (boolean)
    """
    logger.info("Highlighting keywords in content.")
    if not isinstance(term_data, dict):
        logger.error("Invalid term_data for highlighting.")
        return content, False
    try:
        # Use BeautifulSoup to parse the HTML structure first
        soup = BeautifulSoup(content, 'html.parser')

        # Define highlight colors
        colors = {
            'primary_keyword': "#FFEB9C", # Yellowish
            'primary_term': "#CDFFD8",    # Greenish
            'secondary_term': "#E6F3FF"     # Bluish
        }

        # Combine all terms for efficient searching, prioritizing longer terms first to avoid partial matches
        terms_to_highlight = []
        # Add primary keyword
        terms_to_highlight.append({'term': keyword, 'type': 'primary_keyword'})
        # Add primary terms (excluding keyword itself)
        for term_info in term_data.get('primary_terms', []):
            term = term_info.get('term')
            if term and term.lower() != keyword.lower():
                terms_to_highlight.append({'term': term, 'type': 'primary_term'})
        # Add secondary terms
        for term_info in term_data.get('secondary_terms', []):
            term = term_info.get('term')
            if term:
                terms_to_highlight.append({'term': term, 'type': 'secondary_term'})

        # Sort by length descending
        terms_to_highlight.sort(key=lambda x: len(x['term']), reverse=True)

        # Process text nodes within the HTML
        text_nodes = soup.find_all(string=True) # Find all text content

        for node in text_nodes:
             # Skip text within tags we don't want to modify (like <script>, <style>)
             if node.parent.name in ['script', 'style', 'head', 'title', 'meta']:
                 continue

             original_text = str(node)
             new_html_parts = []
             last_index = 0

             # Create a regex pattern for all terms, case-insensitive, whole words only
             # Escape special characters in terms
             term_patterns = []
             term_map = {}
             for item in terms_to_highlight:
                 term_escaped = re.escape(item['term'])
                 term_patterns.append(r'\b' + term_escaped + r'\b')
                 # Store mapping from lower case pattern to type
                 # Using the full pattern as the key simplifies lookup
                 pattern_key = r'\b' + term_escaped.lower() + r'\b'
                 term_map[pattern_key] = item['type']

            # Combine patterns efficiently if possible, or iterate if too many
             if len(term_patterns) > 0:
                 combined_pattern = re.compile('|'.join(term_patterns), re.IGNORECASE)

                 for match in combined_pattern.finditer(original_text):
                     start, end = match.span()
                     term_text = match.group(0)

                     # Find which term type matched
                     matched_type = None
                     # Match the found text (lowercase) against the pattern map
                     pattern_key_match = r'\b' + re.escape(term_text).lower() + r'\b'

                     if pattern_key_match in term_map:
                           matched_type = term_map[pattern_key_match]
                     else: # Fallback if direct key match fails (less likely with full pattern key)
                       current_term_list = [t for t in terms_to_highlight if t['term'].lower() == term_text.lower()]
                       if current_term_list: matched_type = current_term_list[0]['type']


                     if matched_type:
                         # Add text before the match
                         if start > last_index:
                             new_html_parts.append(original_text[last_index:start])
                         # Add the highlighted term
                         color = colors.get(matched_type, '#FFFFFF') # Default white if type unknown
                         new_html_parts.append(f'<span style="background-color: {color};">{term_text}</span>')
                         last_index = end

                 # Add any remaining text after the last match
                 if last_index < len(original_text):
                     new_html_parts.append(original_text[last_index:])

                 # Replace the original text node with the new HTML parts
                 if new_html_parts:
                      new_soup = BeautifulSoup(''.join(new_html_parts), 'html.parser')
                      # Replace node content carefully to avoid issues with parent structure
                      node.replace_with(new_soup)


        # Return the modified HTML as a string
        highlighted_html = str(soup)
        logger.info("Keyword highlighting complete.")
        return highlighted_html, True

    except Exception as e:
        error_msg = f"Exception in highlight_keywords_in_content: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return content, False # Return original content on error


def get_content_improvement_suggestions(content: str, term_data: Dict, score_data: Dict, keyword: str) -> Tuple[Dict, bool]:
    """
    Generate suggestions for improving content based on scoring results.
    More robust checks for score_data structure.
    Returns: suggestions (Dict), success_status (boolean)
    """
    logger.info("Generating content improvement suggestions.")
    suggestions = {
        'missing_terms': [], 'underused_terms': [], 'missing_topics': [],
        'partial_topics': [], 'unanswered_questions': [],
        'readability_suggestions': [], 'structure_suggestions': []
    }
    if not isinstance(term_data, dict) or not isinstance(score_data, dict):
        logger.error("Invalid term_data or score_data for suggestions.")
        return suggestions, False

    try:
        content_details = score_data.get('details', {})
        if not content_details: # Check if details exist
             logger.warning("Score data details are missing, cannot generate full suggestions.")
             # Attempt to generate based on term_data only if possible
             # return suggestions, False # Or return partially filled suggestions

        # Extract text content if HTML is passed
        plain_content = content
        if '<' in content and '>' in content: # Basic check for HTML
             soup = BeautifulSoup(content, 'html.parser')
             plain_content = soup.get_text(separator=' ', strip=True)
        content_lower = plain_content.lower()


        # --- Missing/Underused Terms ---
        primary_term_counts = content_details.get('primary_term_counts', {})
        for term_info in term_data.get('primary_terms', []):
            term = term_info.get('term', '')
            if term:
                recommended = term_info.get('recommended_usage', 1)
                current_count = primary_term_counts.get(term, {}).get('count', 0)
                importance = term_info.get('importance', 0.5)

                if current_count == 0:
                    suggestions['missing_terms'].append({
                        'term': term, 'importance': importance, 'recommended_usage': recommended,
                        'type': 'primary', 'current_usage': 0
                    })
                elif current_count < recommended:
                    suggestions['underused_terms'].append({
                        'term': term, 'importance': importance, 'recommended_usage': recommended,
                        'current_usage': current_count, 'type': 'primary'
                    })

        secondary_term_counts = content_details.get('secondary_term_counts', {})
        for term_info in term_data.get('secondary_terms', [])[:20]:  # Limit suggestions for secondary
            term = term_info.get('term', '')
            if term:
                importance = term_info.get('importance', 0.3)
                if importance > 0.4: # Only suggest relatively important secondary terms
                    current_count = secondary_term_counts.get(term, {}).get('count', 0)
                    recommended = term_info.get('recommended_usage', 1) # Could be optional for secondary
                    if current_count == 0:
                         suggestions['missing_terms'].append({
                             'term': term, 'importance': importance, 'recommended_usage': recommended,
                             'type': 'secondary', 'current_usage': 0
                         })
                    # Optionally add underused check for secondary terms too

        # --- Missing/Partial Topics ---
        topic_coverage_data = content_details.get('topic_coverage', {})
        for topic_info in term_data.get('topics', []):
            topic = topic_info.get('topic', '')
            description = topic_info.get('description', '')
            if topic:
                 coverage_info = topic_coverage_data.get(topic, {})
                 is_covered = coverage_info.get('covered', False)
                 # Using a simple covered flag here. Could use match_ratio if available.
                 if not is_covered:
                      suggestions['missing_topics'].append({'topic': topic, 'description': description})
                 # else: Could add logic for partial coverage if scoring provides more detail

        # --- Unanswered Questions ---
        question_coverage_data = content_details.get('question_coverage', {})
        for question in term_data.get('questions', []):
             coverage_info = question_coverage_data.get(question, {})
             is_answered = coverage_info.get('answered', False)
             if not is_answered:
                 suggestions['unanswered_questions'].append(question)

        # --- Readability & Structure ---
        word_count = content_details.get('word_count', 0)
        if word_count < 500:
            suggestions['readability_suggestions'].append("Content is quite short (< 500 words). Consider significant expansion for comprehensive coverage.")
        elif word_count < 1000:
            suggestions['readability_suggestions'].append("Content is under 1000 words. Adding depth or covering related subtopics could improve performance.")

        # Basic structure checks (can be enhanced)
        # Use BeautifulSoup if content is HTML
        if '<' in content and '>' in content:
             soup = BeautifulSoup(content, 'html.parser')
             headings = soup.find_all(['h2', 'h3', 'h4']) # Check for subheadings
             paragraphs = soup.find_all('p')
             lists = soup.find_all(['ul', 'ol'])

             if len(headings) < 3:
                 suggestions['structure_suggestions'].append("Consider adding more H2/H3 headings to break up content and improve scan-ability.")
             if paragraphs:
                long_paras = sum(1 for p in paragraphs if len(p.get_text().split()) > 150) # Shorter threshold
                if long_paras > 1:
                     suggestions['structure_suggestions'].append(f"At least {long_paras} paragraphs are over 150 words. Try breaking them into smaller chunks.")
             if not lists and word_count > 500:
                 suggestions['structure_suggestions'].append("Using bulleted or numbered lists can improve readability for complex information or steps.")

        logger.info("Content suggestions generated.")
        return suggestions, True

    except Exception as e:
        error_msg = f"Exception in get_content_improvement_suggestions: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        # Return empty suggestions on error, but indicate failure
        return suggestions, False

def create_content_scoring_brief(keyword: str, term_data: Dict, score_data: Dict, suggestions: Dict) -> BytesIO:
    """
    Create a downloadable content scoring brief Word document with recommendations.
    """
    logger.info(f"Creating content scoring brief for '{keyword}'.")
    if not isinstance(term_data, dict) or not isinstance(score_data, dict) or not isinstance(suggestions, dict):
         logger.error("Invalid data provided for scoring brief generation.")
         return BytesIO() # Return empty stream

    try:
        doc = Document()
        # --- Document Header ---
        doc.add_heading(f'Content Optimization Brief: {keyword}', 0)
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        # --- Overall Score ---
        doc.add_heading('Content Score', level=1)
        score_para = doc.add_paragraph()
        score_para.add_run(f"Overall Score: ").bold = True
        overall_score = score_data.get('overall_score', 0)
        score_run = score_para.add_run(f"{overall_score} ({score_data.get('grade', 'F')})")
        # Apply color based on score
        if overall_score >= 70: score_run.font.color.rgb = RGBColor(0, 128, 0) # Green
        elif overall_score < 50: score_run.font.color.rgb = RGBColor(255, 0, 0) # Red
        else: score_run.font.color.rgb = RGBColor(255, 165, 0) # Orange

        # --- Component Scores Table ---
        components = score_data.get('components', {})
        if components:
            doc.add_heading('Score Breakdown', level=2)
            table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
            hcells = table.rows[0].cells; hcells[0].text = 'Component'; hcells[1].text = 'Score'
            for component, score in components.items():
                formatted_comp = component.replace('_score', '').replace('_', ' ').title()
                rcells = table.add_row().cells; rcells[0].text = formatted_comp; rcells[1].text = str(round(score))

        # --- Term Usage Summary ---
        doc.add_heading('Term Usage Recommendations', level=1)
        primary_terms_data = term_data.get('primary_terms', [])
        score_details = score_data.get('details', {})
        primary_term_counts = score_details.get('primary_term_counts', {})

        if primary_terms_data:
            doc.add_heading('Primary Term Usage vs. Recommendations', level=2)
            term_table = doc.add_table(rows=1, cols=4); term_table.style = 'Table Grid'
            hcells = term_table.rows[0].cells
            hcells[0].text = 'Term'; hcells[1].text = 'Recommended'; hcells[2].text = 'Current'; hcells[3].text = 'Status'

            for term_info in primary_terms_data:
                term = term_info.get('term', '')
                recommended = term_info.get('recommended_usage', 1)
                current_count = primary_term_counts.get(term, {}).get('count', 0)

                rcells = term_table.add_row().cells
                rcells[0].text = term
                rcells[1].text = str(recommended)
                rcells[2].text = str(current_count)

                status_run = rcells[3].paragraphs[0].add_run()
                if current_count == 0:
                    status_run.text = "Missing"
                    status_run.font.color.rgb = RGBColor(255, 0, 0) # Red
                elif current_count < recommended:
                    status_run.text = "Underused"
                    status_run.font.color.rgb = RGBColor(255, 165, 0) # Orange
                else:
                    status_run.text = "OK"
                    status_run.font.color.rgb = RGBColor(0, 128, 0) # Green

        # Add secondary terms if needed, maybe just a list of missing ones from suggestions

        # --- Content Gap Recommendations ---
        doc.add_heading('Content Gap Recommendations', level=1)

        if suggestions.get('missing_topics') or suggestions.get('partial_topics'):
             doc.add_heading('Topic Coverage', level=2)
             for topic in suggestions.get('missing_topics', []):
                 p = doc.add_paragraph(style='List Bullet')
                 run_topic = p.add_run(f"Add Topic: {topic.get('topic', '')}")
                 run_topic.bold = True; run_topic.font.color.rgb = RGBColor(255, 165, 0) # Orange
                 p.add_run(f" - {topic.get('description', '')}")
             for topic in suggestions.get('partial_topics', []):
                 p = doc.add_paragraph(style='List Bullet')
                 run_topic = p.add_run(f"Expand Topic: {topic.get('topic', '')}")
                 run_topic.bold = True; run_topic.font.color.rgb = RGBColor(255, 165, 0) # Orange
                 p.add_run(f" - {topic.get('suggestion', '')}")

        if suggestions.get('unanswered_questions'):
            doc.add_heading('Questions to Answer', level=2)
            for question in suggestions.get('unanswered_questions', []):
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(question)

        # --- Structure & Readability ---
        if suggestions.get('structure_suggestions') or suggestions.get('readability_suggestions'):
            doc.add_heading('Structure & Readability', level=1)
            all_structure_suggestions = suggestions.get('structure_suggestions', []) + suggestions.get('readability_suggestions', [])
            for suggestion in all_structure_suggestions:
                doc.add_paragraph(suggestion, style='List Bullet')

        # --- Save Document ---
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        logger.info("Content scoring brief created successfully.")
        return doc_stream

    except Exception as e:
        error_msg = f"Exception in create_content_scoring_brief: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return BytesIO()


###############################################################################
# 6. Meta Title and Description Generation (Assumed Correct from Original)
###############################################################################

def generate_meta_tags(keyword: str, semantic_structure: Dict, related_keywords: List[Dict], term_data: Dict,
                      anthropic_api_key: str) -> Tuple[str, str, bool]:
    """
    Generate optimized meta title and description for the content using Claude.
    Returns: meta_title, meta_description, success_status
    """
    logger.info(f"Generating meta tags for {keyword}.")
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)

        # Extract key info for prompt
        h1 = semantic_structure.get('h1', f"Guide to {keyword}") if isinstance(semantic_structure, dict) else f"Guide to {keyword}"

        # Top related keywords (handle potential None)
        top_keywords_list = [kw.get('keyword') for kw in related_keywords[:5] if kw and kw.get('keyword')]
        top_keywords = ", ".join(top_keywords_list) if top_keywords_list else "N/A"

        # Primary terms (handle potential None/empty dict)
        primary_terms_list = []
        if isinstance(term_data, dict):
            primary_terms_list = [term.get('term') for term in term_data.get('primary_terms', [])[:5] if term and term.get('term')]
        primary_terms_str = ", ".join(primary_terms_list) if primary_terms_list else top_keywords # Fallback

        system_prompt = "You are an SEO expert creating compelling, optimized meta tags. Output ONLY a valid JSON object."
        user_prompt = f"""
        Create an SEO-optimized meta title and meta description for an article about "{keyword}".

        Article Title (H1): "{h1}"
        Primary Terms to Consider: {primary_terms_str}
        Related Keywords: {top_keywords}

        Guidelines:
        - Meta Title: 50-60 characters. Include "{keyword}" near the beginning. Be clear and concise.
        - Meta Description: 150-160 characters. Include "{keyword}" and 1-2 relevant primary/related terms. Summarize the article's value and include a subtle call to action.
        - Language: Engaging, natural, and accurate. Avoid excessive keywords or hype.

        Format your response strictly as JSON:
        ```json
        {{
            "meta_title": "Your optimized title (50-60 chars)",
            "meta_description": "Your optimized description (150-160 chars), incorporating keywords and a CTA."
        }}
        ```
        """

        response = client.messages.create(
            model="claude-3-7-sonnet-20250219", # or 3.5 sonnet
            max_tokens=200, # Ample for title + desc
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
            temperature=0.7 # Balance creativity and adherence
        )

        raw_response_text = response.content[0].text
        logger.debug(f"Raw response from Claude for meta tags:\n{raw_response_text}")

        # Robust JSON Extraction
        meta_data = None
        json_match = re.search(r'```json\s*(\{.*?\})\s*```', raw_response_text, re.DOTALL)
        if json_match:
            json_string = json_match.group(1)
            try:
                meta_data = json.loads(json_string)
                logger.info("Successfully parsed meta tags JSON.")
            except json.JSONDecodeError as e:
                logger.error(f"Failed to decode meta tags JSON: {e}")
        else:
             logger.warning("Could not find ```json ... ``` for meta tags.")
             # Fallback attempt?

        if meta_data and isinstance(meta_data, dict):
            meta_title = meta_data.get('meta_title', f"{h1} | Expert Guide")
            meta_description = meta_data.get('meta_description', f"Explore {keyword} in detail. Get expert insights, tips, and answers in our comprehensive guide. Read now!")

            # Truncate if necessary (conservative limits)
            max_title_len = 60
            max_desc_len = 160
            if len(meta_title) > max_title_len:
                meta_title = meta_title[:max_title_len-3] + "..."
            if len(meta_description) > max_desc_len:
                meta_description = meta_description[:max_desc_len-3] + "..."

            return meta_title, meta_description, True
        else:
            logger.error("Failed to extract valid meta data dictionary from Claude.")
            # Return defaults on failure
            return f"{keyword} - Complete Guide", f"Learn everything about {keyword} in our comprehensive guide. Discover expert tips and best practices.", False

    except anthropic.APIError as e:
        error_msg = f"Anthropic API error during meta tag generation: {e}"
        logger.error(error_msg)
        return f"{keyword} - Guide", f"Learn more about {keyword}.", False
    except Exception as e:
        error_msg = f"Exception in generate_meta_tags: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return f"{keyword} - Guide", f"Learn more about {keyword}.", False

###############################################################################
# 7. Embeddings and Semantic Analysis (Assumed Correct from Original)
###############################################################################

# Placeholder for generate_embedding - ensure you have a working implementation
def generate_embedding(text: str, openai_api_key: str, model: str = "text-embedding-3-small") -> Tuple[List[float], bool]:
    """
    Generate embedding for text using OpenAI API. Corrected default model.
    Returns: embedding (List[float]), success_status (bool)
    """
    logger.info(f"Generating embedding for text (length {len(text)})...")
    if not openai_api_key:
        logger.error("OpenAI API key missing for embedding generation.")
        return [], False
    try:
        # Ensure the openai library is configured with the key
        # This might need to be done globally or passed differently depending on how you manage the client
        # Assuming a global setup for simplicity here:
        openai.api_key = openai_api_key

        # Handle potential text length issues for the chosen model
        # Max tokens depend on the model (e.g., 8191 for text-embedding-3-small/large)
        # A simple character truncation is a basic approach. Better: use tiktoken to count tokens.
        max_chars_approx = 25000 # Approximate character limit based on token estimates
        text_to_embed = text[:max_chars_approx]
        if len(text) > max_chars_approx:
            logger.warning(f"Input text truncated from {len(text)} to {max_chars_approx} characters for embedding.")

        # Use the modern client interface if available, otherwise fallback
        try:
            client = openai.OpenAI(api_key=openai_api_key) # Use the newer client if possible
            response = client.embeddings.create(
                model=model,
                input=[text_to_embed] # API expects a list of strings
            )
            embedding = response.data[0].embedding
            logger.info(f"Embedding generated successfully with model {model}.")
            return embedding, True
        except AttributeError: # Fallback to older syntax if OpenAI client is older
            logger.warning("Using legacy OpenAI embedding syntax.")
            response = openai.Embedding.create(
                model=model,
                input=[text_to_embed] # API expects a list of strings
            )
            embedding = response['data'][0]['embedding']
            logger.info(f"Embedding generated successfully with legacy syntax (model {model}).")
            return embedding, True

    except openai.APIError as e:
        # Handle API error here, e.g. retry or log
        logger.error(f"OpenAI API returned an API Error: {e}")
        return [], False
    except openai.AuthenticationError as e:
        logger.error(f"OpenAI Authentication Error: {e}")
        return [], False
    except openai.RateLimitError as e:
        logger.error(f"OpenAI Rate limit exceeded: {e}")
        # Consider adding a sleep/retry mechanism here
        return [], False
    except Exception as e:
        error_msg = f"Unexpected exception in generate_embedding: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return [], False


def analyze_semantic_structure(contents: List[Dict], anthropic_api_key: str) -> Tuple[Dict, bool]:
    """
    Analyze semantic structure of competitor content to determine optimal hierarchy using Claude.
    Returns: semantic_analysis (Dict), success_status (bool)
    """
    logger.info("Analyzing semantic structure using Claude.")
    if not contents:
        logger.error("No competitor content provided for semantic analysis.")
        return {}, False
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)

        # Combine content, focusing on headings and early paragraphs
        max_context_length = 18000 # Max chars
        combined_context = ""
        for c in contents:
            # Maybe get headings explicitly if available (add to scrape_webpage result?)
            # For now, just use initial text
            content_piece = c.get('content', '')
            if content_piece:
                 snippet = content_piece[:500].strip() # Take first 500 chars as representative
                 if len(combined_context) + len(snippet) < max_context_length:
                     combined_context += f"--- Content from {c.get('url', 'source')} ---\n{snippet}\n\n"
                 else:
                     break

        if not combined_context:
             logger.error("Failed to build combined context for semantic analysis.")
             return {}, False

        system_prompt = "You are an expert SEO content strategist. Analyze the provided content snippets to recommend an optimal heading structure (H1, H2s, H3s) for a new, comprehensive article covering the same topic. Output ONLY a valid JSON object."
        user_prompt = f"""
        Analyze the following content snippets from top-ranking pages for the underlying topic.
        Recommend an optimal semantic heading structure for a new article covering this topic comprehensively.

        Include:
        1. A clear, concise H1 title summarizing the core topic.
        2. 5-8 logical H2 section headings covering the main sub-topics.
        3. For each H2, suggest 2-4 relevant H3 subheadings detailing specific aspects.

        Content Snippets to Analyze:
        {combined_context}

        Format your response strictly as JSON:
        ```json
        {{
            "h1": "Recommended H1 Title",
            "sections": [
                {{
                    "h2": "First Logical H2 Section Title",
                    "subsections": [
                        {{"h3": "First H3 Subheading for Section 1"}},
                        {{"h3": "Second H3 Subheading for Section 1"}}
                    ]
                }},
                {{
                    "h2": "Second Logical H2 Section Title",
                    "subsections": [
                        {{"h3": "First H3 Subheading for Section 2"}},
                        {{"h3": "Second H3 Subheading for Section 2"}},
                        {{"h3": "Third H3 Subheading for Section 2"}}
                    ]
                }}
            ]
        }}
        ```
        Ensure the headings flow logically and cover the topic well based on the analyzed content.
        """

        response = client.messages.create(
            model="claude-3-7-sonnet-20250219", # or 3.5 sonnet
            max_tokens=1500, # Should be enough for structure
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
            temperature=0.2 # Low temp for structure
        )

        raw_response_text = response.content[0].text
        logger.debug(f"Raw response from Claude for semantic structure:\n{raw_response_text[:500]}...")

        # Robust JSON Extraction
        semantic_analysis = None
        json_match = re.search(r'```json\s*(\{.*?\})\s*```', raw_response_text, re.DOTALL)
        if json_match:
            json_string = json_match.group(1)
            try:
                semantic_analysis = json.loads(json_string)
                logger.info("Successfully parsed semantic structure JSON.")
            except json.JSONDecodeError as e:
                logger.error(f"Failed to decode semantic structure JSON: {e}")
                # Add repair attempt if needed
        else:
             logger.warning("Could not find ```json ... ``` for semantic structure.")
             # Fallback attempt if needed

        if semantic_analysis and isinstance(semantic_analysis, dict) and "h1" in semantic_analysis and "sections" in semantic_analysis:
            return semantic_analysis, True
        else:
            logger.error("Failed to extract valid semantic structure dictionary from Claude.")
            return {}, False

    except anthropic.APIError as e:
        error_msg = f"Anthropic API error during semantic analysis: {e}"
        logger.error(error_msg)
        return {}, False
    except Exception as e:
        error_msg = f"Exception in analyze_semantic_structure: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {}, False


###############################################################################
# 8. Content Generation (REFACTORED)
###############################################################################

# --- PASTE REFACTORED generate_article FUNCTION HERE ---
def generate_article(keyword: str, semantic_structure: Dict, related_keywords: List[Dict],
                     serp_features: List[Dict], paa_questions: List[Dict], term_data: Dict,
                     anthropic_api_key: str, # Removed openai_api_key as it wasn't used here
                     competitor_contents: List[Dict], # Added competitor content for context
                     guidance_only: bool = False) -> Tuple[str, bool]:
    """
    Generates a cohesive article based on semantic structure, terms, and competitor context,
    or provide writing guidance. Uses simpler, more direct prompts.
    Returns: article_content (HTML or Markdown Guidance), success_status
    """
    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)
        logger.info(f"Starting article generation for '{keyword}'. Guidance only: {guidance_only}")

        h1 = semantic_structure.get('h1', f"Comprehensive Guide to {keyword}") if isinstance(semantic_structure, dict) else f"Comprehensive Guide to {keyword}"

        # Prepare context: Primary terms, PAA questions
        primary_terms_list = []
        if isinstance(term_data, dict):
            primary_terms_list = [t.get('term') for i, t in enumerate(term_data.get('primary_terms', [])) if t.get('term') and i < 10] # Top 10
        primary_terms_str = ", ".join(primary_terms_list) if primary_terms_list else "N/A"

        paa_questions_list = [q.get('question') for i, q in enumerate(paa_questions or []) if q and q.get('question') and i < 5] # Top 5, handle None paa_questions
        paa_questions_str = "\n - ".join(paa_questions_list) if paa_questions_list else "N/A"

        # Combine competitor snippets for context (limit length)
        competitor_context = ""
        char_limit = 3000
        for comp_content in competitor_contents:
             # Ensure comp_content is a dictionary
             if isinstance(comp_content, dict):
                 content_text = comp_content.get('content', '')
                 if content_text:
                     snippet = content_text[:250].strip() + "...\n\n"
                     if len(competitor_context) + len(snippet) < char_limit:
                         competitor_context += snippet
                     else:
                         break
             else:
                 logger.warning(f"Skipping competitor content item as it's not a dict: {type(comp_content)}")


        if guidance_only:
            logger.info("Generating writing guidance...")
            guidance = f"# Writing Guidance: {h1}\n\n"
            guidance += f"## Target Keyword: {keyword}\n\n"
            guidance += f"## Key Information:\n"
            guidance += f"- **Primary Terms to Include:** {primary_terms_str}\n"
            guidance += f"- **Questions to Address (from PAA):**\n - {paa_questions_str}\n\n"
            guidance += f"## Recommended Structure & Section Guidance:\n\n"

            # Introduction Guidance
            guidance += f"### Introduction (Write ~100-150 words)\n"
            guidance += f"- **Goal:** Briefly introduce '{keyword}', state the article's purpose, and outline the main topics (H2s) covered.\n"
            guidance += f"- **Keywords:** Naturally weave in '{keyword}' and 1-2 primary terms.\n\n"

            # Section Guidance
            sections = semantic_structure.get('sections', []) if isinstance(semantic_structure, dict) else []
            for i, section in enumerate(sections, 1):
                 # Ensure section is a dictionary
                 if isinstance(section, dict):
                     h2 = section.get('h2', f'Section {i}')
                     guidance += f"### H2: {h2} (Write ~150-250 words total for section)\n"
                     guidance += f"- **Focus:** Cover the main aspects of '{h2}'.\n"
                     guidance += f"- **Keywords:** Include relevant primary/secondary terms naturally.\n"
                     guidance += f"- **Consider:** Does this section help answer any PAA questions?\n"

                     subsections = section.get('subsections', [])
                     if isinstance(subsections, list):
                         for j, subsection in enumerate(subsections, 1):
                              if isinstance(subsection, dict):
                                  h3 = subsection.get('h3', f'Subsection {j}')
                                  guidance += f"  - **H3: {h3} (Write ~75-100 words)**\n"
                                  guidance += f"    - **Focus:** Detail a specific aspect of '{h2}'. Keep it concise and directly related to '{h3}'.\n"
                     guidance += "\n"
                 else:
                      logger.warning(f"Skipping section in guidance as it's not a dict: {type(section)}")


            # Conclusion Guidance
            guidance += f"### Conclusion (Write ~100 words)\n"
            guidance += f"- **Goal:** Briefly summarize the key takeaways regarding '{keyword}'. Provide a final thought or call to action.\n"

            logger.info("Writing guidance generated successfully.")
            return guidance, True

        else: # Generate Full Article
            logger.info("Generating full article content...")
            # Note: H1 added after generation below if missing
            max_tokens_per_call = 4000 # Increased tokens for full article generation in one go with Claude 3.5/3.7

            system_prompt = """You are an expert SEO content writer. Write clear, engaging, and informative content based on the provided outline and context.
            Use standard HTML formatting for headings (<h2>, <h3>) and paragraphs (<p>). Ensure smooth transitions between sections.
            Keep paragraphs relatively short (2-4 sentences). Focus on quality and relevance.
            Naturally incorporate the specified primary terms and address the key questions where appropriate within the relevant sections.
            Avoid jargon and write for a general audience unless the topic is inherently technical.
            Output ONLY the HTML content, starting directly with the first introduction paragraph(s) using <p> tags. Do not include the <h1> tag in your output. Do not add any preamble or explanation before the HTML."""

            # Create the structure string for the prompt, ensuring semantic_structure is a dict
            structure_prompt = f"ARTICLE H1 (Do NOT include this in your output): {h1}\n\nSECTIONS TO WRITE:\n"
            if isinstance(semantic_structure, dict):
                sections = semantic_structure.get('sections', [])
                if isinstance(sections, list):
                     for i, section in enumerate(sections, 1):
                          if isinstance(section, dict):
                             structure_prompt += f"\nSection {i}:\n"
                             structure_prompt += f"  H2: {section.get('h2', '')}\n"
                             subsections = section.get('subsections', [])
                             if isinstance(subsections, list):
                                  for j, sub in enumerate(subsections, 1):
                                      if isinstance(sub, dict):
                                         structure_prompt += f"    H3: {sub.get('h3', '')}\n"
                          else:
                              logger.warning(f"Skipping section in structure prompt generation as it's not a dict: {type(section)}")

                else:
                     logger.warning(f"Sections in semantic_structure are not a list: {type(sections)}")
            else:
                 logger.error(f"semantic_structure is not a dictionary: {type(semantic_structure)}. Cannot generate structure prompt.")
                 return f"<h1>{h1}</h1><p>[Error: Invalid internal structure data for generation.]</p>", False


            user_prompt = f"""
            Write a full article about "{keyword}" based on the following structure and context.

            {structure_prompt}

            KEY INFORMATION FOR WRITING:
            - Main Keyword: {keyword}
            - Primary Terms to Include Naturally: {primary_terms_str}
            - Key Questions to Address (from PAA): {paa_questions_str}

            COMPETITOR CONTENT SNIPPETS (for context, do not copy):
            {competitor_context}

            INSTRUCTIONS:
            1. Write an introduction paragraph FIRST (approx 100-150 words) related to the H1 '{h1}'. Use <p> tags.
            2. THEN, write content for EACH H2 section and its corresponding H3 subsections as outlined in the STRUCTURE above. Aim for appropriate length per section.
            3. Ensure H3 content is directly related to its parent H2.
            4. Naturally weave in the Primary Terms throughout the article where relevant.
            5. Address the Key Questions within the most logical sections.
            6. Write a concluding section LAST (approx 100 words) summarizing key points. Wrap it in <h2>Conclusion</h2> and <p> tags.
            7. Use ONLY <h2>, <h3>, and <p> HTML tags for the body content.
            8. Output the complete article body as a single block of HTML, starting with the introduction paragraph(s). DO NOT include the <h1> tag.
            """

            try:
                response = client.messages.create(
                    model="claude-3-5-sonnet-20240620", # Updated to latest Sonnet model
                    max_tokens=max_tokens_per_call,
                    system=system_prompt,
                    messages=[{"role": "user", "content": user_prompt}],
                    temperature=0.6 # Mitigated creativity for consistency
                )

                generated_body_content = response.content[0].text.strip()

                # Basic validation: Check if it generated something substantial and looks like HTML
                if len(generated_body_content) > 200 and ("<h2" in generated_body_content or "<h3" in generated_body_content) and "<p>" in generated_body_content:
                    logger.info(f"Full article body generated successfully. Length: {len(generated_body_content)} chars.")
                    # Prepend the H1 tag and return
                    full_article_html = f"<h1>{h1}</h1>\n{generated_body_content}"
                    return full_article_html, True
                else:
                    logger.warning(f"Generated content seems too short or lacks expected HTML structure. Length: {len(generated_body_content)} chars.")
                    # Return error state with H1 and partial content for debugging
                    error_html = f"<h1>{h1}</h1>\n<p>[Error: Generated content structure invalid or too short.]</p>\n<!--\n{generated_body_content[:500]}\n-->"
                    return error_html, False

            except anthropic.APIError as e:
                error_msg = f"Anthropic API error during article generation: {e}"
                logger.error(error_msg)
                return f"<h1>{h1}</h1><p>[Error: API call failed during generation. Details: {e}]</p>", False
            except Exception as e:
                 error_msg = f"Unexpected exception during article generation: {e}"
                 logger.error(error_msg)
                 logger.error(traceback.format_exc())
                 return f"<h1>{h1}</h1><p>[Error: An unexpected error occurred during content generation.]</p>", False

    except Exception as e:
        error_msg = f"Exception in generate_article outer scope: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        # Return minimal content with error message
        error_content = "# Guidance Generation Error" if guidance_only else f"<h1>Error Generating Article</h1>"
        error_content += f"\n\n<p>An error occurred: {str(e)}</p>" if not guidance_only else f"\n\nAn error occurred: {str(e)}"
        return error_content, False

###############################################################################
# 9. Internal Linking (Assumed Correct from Original)
###############################################################################

def parse_site_pages_spreadsheet(uploaded_file) -> Tuple[List[Dict], bool]:
    """
    Parse uploaded CSV/Excel with site pages. Added validation checks.
    Returns: pages (List[Dict]), success_status (bool)
    """
    logger.info(f"Parsing site pages spreadsheet: {uploaded_file.name}")
    try:
        file_name = uploaded_file.name.lower()
        required_columns = ['url', 'title', 'meta description'] # Check lower case

        if file_name.endswith('.csv'):
            # Sniff dialect and encoding if possible - robust CSV reading
            try:
                # Read a small sample to guess dialect
                sample = uploaded_file.read(2048)
                uploaded_file.seek(0) # Reset pointer
                dialect = pd.io.common.sniff_csv_dialect(sample.decode())
                df = pd.read_csv(uploaded_file, dialect=dialect)
            except Exception: # Fallback to basic read
                 uploaded_file.seek(0)
                 df = pd.read_csv(uploaded_file)

        elif file_name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
        else:
            logger.error(f"Unsupported file type: {uploaded_file.name}")
            return [], False

        # Standardize column names (lowercase, strip spaces)
        df.columns = [str(col).lower().strip() for col in df.columns]

        # Check required columns (after standardization)
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns in spreadsheet: {', '.join(missing_columns)}")
            st.error(f"Spreadsheet missing required columns: {', '.join(missing_columns)}. Needed: URL, Title, Meta Description.")
            return [], False

        # Convert dataframe to list of dicts, handling potential NaN/None values
        pages = []
        for _, row in df.iterrows():
             # Ensure URL is present and looks somewhat like a URL
             url_val = row.get('url')
             if not url_val or not isinstance(url_val, str) or not url_val.startswith(('http://', 'https://')):
                 logger.warning(f"Skipping row due to invalid URL: {url_val}")
                 continue

             pages.append({
                 'url': url_val,
                 'title': str(row.get('title', '')) if pd.notna(row.get('title')) else '',
                 'description': str(row.get('meta description', '')) if pd.notna(row.get('meta description')) else ''
             })

        if not pages:
             logger.error("No valid pages found after parsing spreadsheet.")
             return [], False

        logger.info(f"Successfully parsed {len(pages)} pages from spreadsheet.")
        return pages, True

    except Exception as e:
        error_msg = f"Exception parsing spreadsheet {uploaded_file.name}: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return [], False

def embed_site_pages(pages: List[Dict], openai_api_key: str, batch_size: int = 10) -> Tuple[List[Dict], bool]:
    """
    Generate embeddings for site pages in batches for faster processing using OpenAI.
    Returns: pages_with_embeddings, success_status
    """
    logger.info(f"Starting embedding for {len(pages)} site pages with batch size {batch_size}.")
    if not pages:
        logger.warning("No pages provided for embedding.")
        return [], True # Return empty list successfully
    if not openai_api_key:
         logger.error("OpenAI API key missing for embedding.")
         return pages, False

    try:
        # Configure OpenAI client (prefer new client)
        try:
            client = openai.OpenAI(api_key=openai_api_key)
            is_new_client = True
        except AttributeError:
            openai.api_key = openai_api_key # Fallback for older library version
            is_new_client = False
            logger.warning("Using legacy OpenAI library structure for embeddings.")

        # Prepare texts to embed (combine key fields)
        texts = []
        for page in pages:
            # Combine relevant fields for semantic meaning
            combined_text = f"Title: {page.get('title', '')}\nURL: {page.get('url', '')}\nDescription: {page.get('description', '')}"
            # Truncate reasonably if needed, although batching helps with overall length
            # Max tokens depends on model, but let's keep individual texts manageable
            texts.append(combined_text[:8000]) # Approx limit per text

        # Process in batches
        all_embeddings = []
        total_batches = (len(texts) + batch_size - 1) // batch_size
        logger.info(f"Processing {total_batches} batches...")
        progress_bar = st.progress(0) # Optional Streamlit progress bar

        batch_start_time = time.time()
        for i in range(total_batches):
            start_idx = i * batch_size
            end_idx = min(start_idx + batch_size, len(texts))
            batch_texts = texts[start_idx:end_idx]

            if not batch_texts: continue # Skip empty batches

            try:
                 # Make API call using appropriate client structure
                 model = "text-embedding-3-small" # Smaller model generally preferred for cost/speed
                 if is_new_client:
                     response = client.embeddings.create(model=model, input=batch_texts)
                     batch_embeddings = [item.embedding for item in response.data]
                 else: # Legacy
                     response = openai.Embedding.create(model=model, input=batch_texts)
                     batch_embeddings = [item['embedding'] for item in response['data']]

                 all_embeddings.extend(batch_embeddings)
                 logger.info(f"Processed batch {i+1}/{total_batches}...")
                 # Optional: Update progress bar
                 if 'progress_bar' in locals():
                     progress_bar.progress( (i+1) / total_batches )
                 # Optional: Add slight delay if hitting rate limits frequently
                 # time.sleep(0.5)

            except openai.RateLimitError as rle:
                 logger.error(f"Rate limit hit during page embedding batch {i+1}. Error: {rle}. Retrying may be needed.")
                 # Consider more robust retry logic here if needed
                 st.error(f"OpenAI rate limit hit. Please wait and try again later or reduce batch size.")
                 return pages, False # Fail the process on rate limit
            except openai.APIError as apie:
                 logger.error(f"OpenAI API error during page embedding batch {i+1}. Error: {apie}")
                 st.error(f"OpenAI API error: {apie}. Please check keys and try again.")
                 return pages, False
            except Exception as batch_exc:
                  logger.error(f"Error processing embedding batch {i+1}: {batch_exc}")
                  logger.error(traceback.format_exc())
                  # Decide how to handle partial failure - e.g., skip batch or fail all
                  # Failing all is safer to avoid data mismatch
                  st.error(f"Failed to process embedding batch {i+1}. See logs.")
                  return pages, False


        # Check if we got the expected number of embeddings
        if len(all_embeddings) != len(pages):
            logger.error(f"Mismatch in embedding count ({len(all_embeddings)}) vs page count ({len(pages)}). Aborting.")
            return pages, False

        # Add embeddings back to pages data
        pages_with_embeddings = []
        for i, page in enumerate(pages):
            page_with_embedding = page.copy()
            page_with_embedding['embedding'] = all_embeddings[i]
            pages_with_embeddings.append(page_with_embedding)

        total_time = time.time() - batch_start_time
        logger.info(f"Successfully generated embeddings for {len(pages_with_embeddings)} pages in {format_time(total_time)}.")
        return pages_with_embeddings, True

    except Exception as e:
        error_msg = f"Exception in embed_site_pages: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        st.error(f"Embedding process failed: {e}")
        return pages, False # Return original pages, indicating failure


def verify_semantic_match(anchor_text: str, page_title: str) -> float:
    """
    Verify and score the semantic match between anchor text and page title using simple word overlap.
    Returns a similarity score (0-1)
    """
    # Define common stop words (can be expanded)
    stop_words = {'a', 'an', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'with',
                  'by', 'about', 'as', 'is', 'are', 'was', 'were', 'of', 'from', 'into', 'during',
                  'after', 'before', 'above', 'below', 'between', 'under', 'over', 'through',
                  'how', 'what', 'when', 'where', 'why', 'which', 'who', 'com'} # Added common query words

    # Convert to lowercase, tokenize, remove stop words and short words
    anchor_words = {word for word in re.findall(r'\b\w+\b', anchor_text.lower()) if len(word) > 2 and word not in stop_words}
    title_words = {word for word in re.findall(r'\b\w+\b', page_title.lower()) if len(word) > 2 and word not in stop_words}

    if not anchor_words or not title_words:
        return 0.0 # Cannot compare if one set is empty

    # Find overlapping meaningful words
    overlaps = anchor_words.intersection(title_words)

    if not overlaps:
        return 0.0

    # Calculate scores based on Jaccard similarity (intersection / union)
    # This penalizes anchor text being too broad or too narrow relative to title
    union_size = len(anchor_words.union(title_words))
    similarity = len(overlaps) / union_size if union_size > 0 else 0.0

    return similarity

# Ensure the 'openai_api_key' is passed correctly. The original code was missing it in the function definition and call.
# Added 'anthropic_api_key' as it was used in the fallback logic (though now simplified).
def generate_internal_links_with_embeddings(article_content: str, pages_with_embeddings: List[Dict],
                                           openai_api_key: str, anthropic_api_key: str, word_count: int) -> Tuple[str, List[Dict], bool]:
    """
    Generate internal links using paragraph-level semantic matching with embeddings.
    Simplified anchor text selection to basic keyword matching.
    Returns: article_with_links (HTML string), links_added (List[Dict]), success_status (boolean)
    """
    logger.info("Generating internal links...")
    if not article_content:
        logger.warning("Cannot generate links: Article content is empty.")
        return "", [], False
    if not pages_with_embeddings:
        logger.warning("Cannot generate links: No site pages with embeddings provided.")
        return article_content, [], True # Return original content, success=True as no links *could* be added

    if not openai_api_key:
        logger.error("OpenAI API key missing for internal linking (paragraph embeddings).")
        return article_content, [], False

    try:
        # Configure OpenAI client
        try:
            client = openai.OpenAI(api_key=openai_api_key)
            is_new_client = True
        except AttributeError:
            openai.api_key = openai_api_key
            is_new_client = False
            logger.warning("Using legacy OpenAI library structure for paragraph embeddings.")


        # Calculate max links based on word count (e.g., Aim for ~1 link per 150-200 words)
        max_links = min(15, max(3, int(word_count / 175))) # Adjust density target as needed
        logger.info(f"Targeting up to {max_links} internal links for content with {word_count} words.")

        # 1. Extract Content Paragraphs (Robustly handle HTML/Plain Text)
        soup = BeautifulSoup(article_content, 'html.parser')
        paragraphs = []
        min_para_length = 20 # Minimum words for a paragraph to be linkable

        # Prioritize <p> tags
        p_tags = soup.find_all('p')
        if p_tags:
            for p_tag in p_tags:
                para_text = p_tag.get_text(strip=True)
                if len(para_text.split()) >= min_para_length:
                    paragraphs.append({'text': para_text, 'element': p_tag})
        else:
            # If no <p> tags, try splitting by double newline for plain text-like content
            text_chunks = article_content.split('\n\n')
            temp_soup = BeautifulSoup("", 'html.parser') # Dummy soup to create tag objects
            for i, chunk in enumerate(text_chunks):
                plain_chunk = chunk.strip()
                if len(plain_chunk.split()) >= min_para_length:
                     # Store as text, create a dummy 'element' for later replacement reference if needed
                     # This part is tricky if replacing back into plain text without proper HTML structure
                     # For simplicity, focus on HTML input first.
                     # If plain text input is expected, the replacement logic needs adjustment.
                     # Assuming HTML input for now based on other functions.
                     logger.warning("No <p> tags found, internal linking might be less accurate.")
                     # Need a strategy here if input isn't guaranteed HTML <p> tags. Skip for now.
                     pass

        if not paragraphs:
            logger.warning("No suitable paragraphs found in the article content for linking.")
            return article_content, [], True # No suitable places to link

        # 2. Generate Embeddings for Paragraphs
        logger.info(f"Generating embeddings for {len(paragraphs)} paragraphs...")
        paragraph_texts = [p['text'][:8000] for p in paragraphs] # Limit text length per paragraph

        # Determine embedding model/dimension from site pages (check first valid page)
        first_valid_page = next((p for p in pages_with_embeddings if p and isinstance(p.get('embedding'), list) and len(p['embedding']) > 0), None)
        if not first_valid_page:
             logger.error("No valid page embeddings found to determine model/dimension.")
             return article_content, [], False

        embed_dim = len(first_valid_page['embedding'])
        # Choose the appropriate model based on dimension (common OpenAI models)
        if embed_dim == 1536:
            embedding_model = "text-embedding-3-small"
        elif embed_dim == 3072:
            embedding_model = "text-embedding-3-large"
        elif embed_dim == 768: # E.g., text-embedding-ada-002 (older)
            embedding_model = "text-embedding-ada-002"
        else:
             logger.warning(f"Unknown embedding dimension {embed_dim}. Defaulting to text-embedding-3-small.")
             embedding_model = "text-embedding-3-small"
        logger.info(f"Using embedding model: {embedding_model} (dim: {embed_dim}) for paragraphs.")


        # Get embeddings for all paragraphs (can also batch this if needed for many paragraphs)
        try:
             if is_new_client:
                 response = client.embeddings.create(model=embedding_model, input=paragraph_texts)
                 paragraph_embeddings = [item.embedding for item in response.data]
             else:
                 response = openai.Embedding.create(model=embedding_model, input=paragraph_texts)
                 paragraph_embeddings = [item['embedding'] for item in response['data']]

             # Add embeddings to paragraph data
             if len(paragraph_embeddings) == len(paragraphs):
                 for i, embedding in enumerate(paragraph_embeddings):
                     paragraphs[i]['embedding'] = np.array(embedding) # Convert to numpy array for calculations
             else:
                  logger.error("Mismatch between paragraph count and generated embeddings. Aborting link generation.")
                  return article_content, [], False

        except Exception as embed_err:
            logger.error(f"Error generating paragraph embeddings: {embed_err}")
            return article_content, [], False

        # 3. Find Best Matches and Select Anchor Texts
        links_to_add = []
        used_page_urls = set()       # Track pages already linked to
        used_paragraph_indices = set() # Track paragraphs that already have a link

        # Filter pages to only those with valid embeddings matching dimension
        valid_pages_with_embeddings = [p for p in pages_with_embeddings if p and isinstance(p.get('embedding'), list) and len(p['embedding']) == embed_dim]
        if not valid_pages_with_embeddings:
             logger.error("No project pages found with valid embeddings matching article paragraphs. Cannot link.")
             return article_content, [], False


        # Pre-calculate norms for faster cosine similarity
        page_embeddings_np = np.array([p['embedding'] for p in valid_pages_with_embeddings])
        page_norms = np.linalg.norm(page_embeddings_np, axis=1)

        # Iterate through paragraphs to find link opportunities
        for para_idx, paragraph in enumerate(paragraphs):
            if len(links_to_add) >= max_links or para_idx in used_paragraph_indices:
                continue

            para_embedding = paragraph.get('embedding')
            if para_embedding is None: continue # Skip if embedding failed

            para_norm = np.linalg.norm(para_embedding)
            if para_norm == 0 or np.any(page_norms == 0): # Avoid division by zero
                logger.warning(f"Skipping paragraph {para_idx} due to zero norm.")
                continue


            # Calculate cosine similarities between this paragraph and all valid pages
            similarities = np.dot(page_embeddings_np, para_embedding) / (page_norms * para_norm)

            # Find the best *unused* page match above a threshold
            best_score = 0.70  # Similarity threshold (adjust as needed)
            best_page_idx = -1

            # Iterate through similarity scores, checking if page is used
            sorted_indices = np.argsort(similarities)[::-1] # Indices from highest to lowest similarity
            for current_page_idx in sorted_indices:
                  if similarities[current_page_idx] < best_score:
                       break # No more matches above threshold

                  page_url = valid_pages_with_embeddings[current_page_idx].get('url')
                  if page_url and page_url not in used_page_urls:
                       # Found the best unused match above threshold
                       best_score = similarities[current_page_idx]
                       best_page_idx = current_page_idx
                       break # Stop searching for this paragraph

            # If a good, unused match was found
            if best_page_idx != -1:
                best_page = valid_pages_with_embeddings[best_page_idx]
                page_title = best_page.get('title', '')
                page_url = best_page['url']
                para_text = paragraph['text']

                # --- Simplified Anchor Text Selection ---
                anchor_text = ""
                # Option 1: Look for exact page title match (if short enough)
                if len(page_title.split()) <= 5 and page_title.lower() in para_text.lower():
                     # Find the actual casing in the paragraph text
                     match = re.search(re.escape(page_title), para_text, re.IGNORECASE)
                     if match: anchor_text = match.group(0)
                # Option 2: Look for noun phrases containing keywords from title
                if not anchor_text:
                     # Simple keyword extraction from title (longer words)
                     title_keywords = {w for w in re.findall(r'\b\w{4,}\b', page_title.lower())}
                     # Look for sentence segments containing these keywords
                     sentences = re.split(r'[.!?]\s+', para_text)
                     for sentence in sentences:
                          s_lower = sentence.lower()
                          if any(kw in s_lower for kw in title_keywords):
                               # Try to find a 2-5 word phrase around a keyword
                               for kw in title_keywords:
                                    if kw in s_lower:
                                         # Find instance of keyword
                                         match = re.search(r'\b' + re.escape(kw) + r'\b', sentence, re.IGNORECASE)
                                         if match:
                                              start_idx = match.start()
                                              # Find word boundaries around match
                                              words_before = sentence[:start_idx].split()
                                              words_after = sentence[match.end():].split()
                                              # Construct a phrase (e.g., 1 before, keyword, 1 after)
                                              phrase_words = []
                                              if words_before: phrase_words.append(words_before[-1])
                                              phrase_words.append(match.group(0))
                                              if words_after: phrase_words.append(words_after[0])

                                              candidate = " ".join(phrase_words)
                                              # Basic check if it looks reasonable
                                              if 2 <= len(phrase_words) <= 5 and candidate in para_text:
                                                   anchor_text = candidate
                                                   break # Found a decent anchor
                               if anchor_text: break # Found anchor in this sentence

                # Option 3: Fallback to just the page title (if no better anchor found)
                if not anchor_text:
                     anchor_text = page_title if len(page_title.split()) <= 6 else keyword # Fallback further to main keyword

                # --- Final Check and Add Link ---
                # Ensure proposed anchor text actually exists in the paragraph
                if anchor_text and re.search(re.escape(anchor_text), para_text, re.IGNORECASE):
                     # Find the actual casing for replacement
                     match = re.search(re.escape(anchor_text), para_text, re.IGNORECASE)
                     final_anchor = match.group(0)

                     # Add to list
                     links_to_add.append({
                         'url': page_url,
                         'anchor_text': final_anchor, # Use the text as found in paragraph
                         'paragraph_index': para_idx,
                         'similarity_score': best_score,
                         'page_title': page_title
                     })
                     used_page_urls.add(page_url)
                     used_paragraph_indices.add(para_idx)
                     logger.info(f"Found link: '{final_anchor}' -> {page_url} (Score: {best_score:.3f})")
                else:
                     logger.warning(f"Could not confirm anchor text '{anchor_text}' in paragraph {para_idx}.")


        # 4. Apply the Links to the Article HTML
        if not links_to_add:
            logger.info("No internal links generated.")
            return article_content, [], True # Successful, but no links added

        # Use the SAME soup object we extracted paragraphs from
        modified_soup = soup
        final_links_added = []

        # Sort links by index to process correctly if multiple in same para (though unlikely now)
        links_to_add.sort(key=lambda x: x['paragraph_index'])

        linked_para_elements = {} # Store modified content for unique paragraphs

        for link in links_to_add:
             para_idx = link['paragraph_index']
             para_element = paragraphs[para_idx]['element'] # Get the original BS4 element

             # Get current HTML of the paragraph, process if not already done
             if para_idx not in linked_para_elements:
                 linked_para_elements[para_idx] = para_element.decode_contents() # Get inner HTML

             current_html = linked_para_elements[para_idx]
             anchor_text = link['anchor_text']
             url = link['url']

             # Replace the first occurrence of the anchor text within this paragraph's HTML
             # Use regex for case-insensitive replacement, ensuring it's not already linked
             # Pattern: Negative lookbehind/ahead for existing <a> tags around the anchor
             pattern = re.compile(
                 r'(?<!<a[^>]*?>\s*)' + # Not preceded by opening <a> tag
                 r'(' + re.escape(anchor_text) + r')' +
                 r'(?!\s*</a[^>]*?>)' # Not followed by closing </a> tag
                 , re.IGNORECASE
             )

             link_html = f'<a href="{url}" title="{link["page_title"]}">{anchor_text}</a>'
             
             new_html, num_replacements = pattern.subn(link_html, current_html, count=1)

             if num_replacements > 0:
                  linked_para_elements[para_idx] = new_html # Update stored HTML
                  # Add context for output summary
                  context_text = paragraphs[para_idx]['text']
                  context_match = re.search(re.escape(anchor_text), context_text, re.IGNORECASE)
                  if context_match:
                       start_pos = max(0, context_match.start() - 40)
                       end_pos = min(len(context_text), context_match.end() + 40)
                       context = ("..." if start_pos > 0 else "") + \
                                 context_text[start_pos:context_match.start()] + \
                                 f"**[{context_match.group(0)}]**" + \
                                 context_text[context_match.end():end_pos] + \
                                 ("..." if end_pos < len(context_text) else "")
                       link['context'] = context.replace('\n', ' ') # Clean context
                  else:
                       link['context'] = f"(Context extraction failed for: {anchor_text})"

                  final_links_added.append({
                      "url": url,
                      "anchor_text": anchor_text,
                      "context": link.get('context', ''),
                      "page_title": link['page_title'],
                      "similarity_score": round(link['similarity_score'], 3)
                   })
             else:
                  logger.warning(f"Could not replace anchor '{anchor_text}' in paragraph {para_idx} (already linked or not found in HTML).")


        # Apply changes back to the soup object
        for para_idx, new_inner_html in linked_para_elements.items():
             para_element = paragraphs[para_idx]['element']
             # Parse the new inner HTML and replace the element's contents
             new_content_soup = BeautifulSoup(new_inner_html, 'html.parser')
             para_element.clear() # Remove old contents
             # Append children from the parsed new HTML
             for child in new_content_soup.contents:
                 para_element.append(child.extract()) # Use extract to move node

        logger.info(f"Applied {len(final_links_added)} internal links to the article.")
        return str(modified_soup), final_links_added, True # Return modified HTML

    except Exception as e:
        error_msg = f"Exception in generate_internal_links_with_embeddings: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return article_content, [], False # Return original content on unexpected error


###############################################################################
# 10. Document Generation (REFACTORED)
###############################################################################

# --- Helper functions for HTML to Word ---
def parse_style_attribute(style_string: str) -> Dict:
    """ Parses a CSS style string into a dictionary of relevant properties for Word conversion. """
    styles = {}
    if not style_string or not isinstance(style_string, str): return styles
    try:
        attributes = style_string.split(';')
        for attr in attributes:
            if ':' in attr:
                key, value = attr.split(':', 1)
                key = key.strip().lower()
                value = value.strip().lower()

                if key == 'color':
                    # Handle common color names and hex codes used in the script
                    if value in ('red', '#ff0000'): styles['color'] = RGBColor(255, 0, 0)
                    elif value in ('orange', '#ff8c00', '#ffa500'): styles['color'] = RGBColor(255, 165, 0)
                    elif value in ('gray', 'grey', '#808080'): styles['color'] = RGBColor(128, 128, 128)
                    elif value in ('green', '#008000'): styles['color'] = RGBColor(0, 128, 0)
                    # Add more specific hex if needed here...
                    else:
                         # Basic hex color parsing (naive)
                         match = re.match(r'#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})', value)
                         if match:
                             try:
                                 r, g, b = [int(c, 16) for c in match.groups()]
                                 styles['color'] = RGBColor(r, g, b)
                             except ValueError: pass # Ignore invalid hex

                elif key == 'background-color':
                     # Keyword highlighting colors
                     if value == '#ffeb9c': styles['background'] = RGBColor(255, 235, 156) # Yellowish
                     elif value == '#cdffd8': styles['background'] = RGBColor(205, 255, 216) # Greenish
                     elif value == '#e6f3ff': styles['background'] = RGBColor(230, 243, 255) # Bluish
                     # Could add more general background mapping if needed

                elif key == 'text-decoration' and 'line-through' in value:
                    styles['strike'] = True
                elif key == 'font-weight' and ('bold' in value or value.isdigit() and int(value) >= 700):
                     styles['bold'] = True
                elif key == 'font-style' and 'italic' in value:
                     styles['italic'] = True
    except Exception as e:
        logger.warning(f"Could not parse style string '{style_string}': {e}")
    return styles

def add_formatted_text(doc_paragraph, beautifulsoup_element):
    """
    Adds text from a BeautifulSoup element to a docx paragraph, handling basic inline formatting
    (strong, em, b, i), spans with styles (color, strike, background via highlight), and links.
    """
    # Iterate through the element's contents (text nodes and tags)
    for content in beautifulsoup_element.contents:
        try:
            # Handle Text Nodes
            if isinstance(content, str):
                if content.strip(): # Avoid adding runs for just whitespace
                    doc_paragraph.add_run(content)

            # Handle Styled Spans
            elif content.name == 'span' and 'style' in content.attrs:
                run = doc_paragraph.add_run(content.get_text())
                styles = parse_style_attribute(content['style'])
                if styles.get('color'): run.font.color.rgb = styles['color']
                if styles.get('strike'): run.font.strike = True
                if styles.get('bold'): run.bold = True # Handle style-based bold
                if styles.get('italic'): run.italic = True # Handle style-based italic
                # Note: Background color mapping to Word highlight is approximate
                if styles.get('background'):
                    from docx.enum.text import WD_COLOR_INDEX
                    # Map specific background RGBs to Word's limited highlight palette
                    bg = styles['background']
                    if bg == RGBColor(255, 235, 156): run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    elif bg == RGBColor(205, 255, 216): run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
                    elif bg == RGBColor(230, 243, 255): run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE
                    else: run.font.highlight_color = WD_COLOR_INDEX.GRAY_25 # Default highlight

            # Handle Semantic Inline Tags
            elif content.name in ['strong', 'b']:
                run = doc_paragraph.add_run(content.get_text())
                run.bold = True
            elif content.name in ['em', 'i']:
                run = doc_paragraph.add_run(content.get_text())
                run.italic = True

             # Handle Links (Basic - no actual hyperlink in docx)
            elif content.name == 'a':
                 link_text = content.get_text()
                 href = content.get('href', '')
                 title = content.get('title', '') # Get title attribute if present
                 run = doc_paragraph.add_run(link_text if link_text else href) # Use href if no text
                 run.font.underline = True
                 run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
                 # Optionally add URL/Title in parentheses
                 if href:
                      tooltip = f" [Link: {href}" + (f' - {title}' if title else "") + "]"
                      run_tooltip = doc_paragraph.add_run(tooltip)
                      run_tooltip.font.size = Pt(8) # Smaller font for tooltip
                      run_tooltip.font.color.rgb = RGBColor(0x55, 0x55, 0x55) # Gray


            # Handle Line Breaks
            elif content.name == 'br':
                 doc_paragraph.add_run("\n") # Or add_break(WD_BREAK.LINE)? '\n' usually works.

            # Handle other inline tags by just adding their text
            elif hasattr(content, 'name'): # Check if it's a tag
                text = content.get_text().strip()
                if text:
                    doc_paragraph.add_run(text)

        except Exception as e:
                # Log error and add raw text as fallback
                raw_text = str(content) if isinstance(content, str) else content.get_text()
                logger.warning(f"Error processing inline element {content.name if hasattr(content, 'name') else type(content)}: {e}. Adding raw text: {raw_text[:50]}")
                if raw_text.strip():
                     doc_paragraph.add_run(f"[Parse Error: {raw_text[:50]}...]")

# --- PASTE REFACTORED create_word_document FUNCTION HERE ---
def create_word_document(keyword: str, serp_results: List[Dict], related_keywords: List[Dict],
                        semantic_structure: Dict, article_content: str, meta_title: str,
                        meta_description: str, paa_questions: List[Dict], term_data: Dict = None,
                        score_data: Dict = None, internal_links: List[Dict] = None,
                        guidance_only: bool = False) -> Tuple[BytesIO, bool]:
    """
    Create Word document with all components. Includes content score if available.
    Handles basic HTML in article_content (H1-H6, P, Links, Basic Inline) or Markdown for guidance.
    Color coding for scores is applied. Uses helpers for HTML/Markdown parsing.

    Returns: document_stream, success_status
    """
    logger.info(f"Creating main SEO brief Word document for '{keyword}'. Guidance: {guidance_only}")
    try:
        doc = Document()

        # --- Standard Sections (Meta, SERP, Keywords, Terms, Score) ---
        doc.add_heading(f'SEO Brief: {keyword}', 0)
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        # Meta Tags
        doc.add_heading('Meta Tags', level=1)
        meta_paragraph = doc.add_paragraph(); meta_paragraph.add_run("Meta Title: ").bold = True; meta_paragraph.add_run(meta_title or "N/A")
        desc_paragraph = doc.add_paragraph(); desc_paragraph.add_run("Meta Description: ").bold = True; desc_paragraph.add_run(meta_description or "N/A")

        # SERP Analysis
        doc.add_heading('SERP Analysis', level=1)
        if serp_results:
             doc.add_paragraph('Top 10 Organic Results:')
             table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
             hcells = table.rows[0].cells; hcells[0].text = 'Rank'; hcells[1].text = 'Title'; hcells[2].text = 'URL'; hcells[3].text = 'Page Type'
             for result in serp_results:
                 rcells = table.add_row().cells
                 rcells[0].text = str(result.get('rank_group', '')); rcells[1].text = result.get('title', ''); rcells[2].text = result.get('url', ''); rcells[3].text = result.get('page_type', '')
        else:
             doc.add_paragraph("No SERP results data available.")

        # People Also Asked
        paa_list = paa_questions or []
        if paa_list:
            doc.add_heading('People Also Asked', level=2)
            for i, question_data in enumerate(paa_list, 1):
                q_text = question_data.get('question', '')
                q_paragraph = doc.add_paragraph(style='List Number'); q_paragraph.add_run(q_text).bold = True
                # Optional: Add expanded answers if structure supports it
                # for expanded in question_data.get('expanded', []): ... doc.add_paragraph(expanded.get('description', ''), style='List Bullet 2')

        # Related Keywords
        doc.add_heading('Related Keywords', level=1)
        kw_list = related_keywords or []
        if kw_list:
             kw_table = doc.add_table(rows=1, cols=3); kw_table.style = 'Table Grid'
             hcells = kw_table.rows[0].cells; hcells[0].text = 'Keyword'; hcells[1].text = 'Search Volume'; hcells[2].text = 'CPC ($)'
             for kw in kw_list:
                 rcells = kw_table.add_row().cells
                 rcells[0].text = kw.get('keyword', '')
                 sv = kw.get('search_volume'); rcells[1].text = str(int(sv)) if sv is not None else 'N/A'
                 cpc = kw.get('cpc');
                 if cpc is not None:
                      try: rcells[2].text = f"${float(cpc):.2f}"
                      except (ValueError, TypeError): rcells[2].text = "N/A"
                 else: rcells[2].text = "N/A"
        else:
             doc.add_paragraph("No related keywords data available.")

        # Important Terms
        term_dict = term_data or {}
        if term_dict:
            doc.add_heading('Important Terms to Include', level=1)
            # Primary Terms
            primary = term_dict.get('primary_terms', [])
            if primary:
                 doc.add_heading('Primary Terms', level=2)
                 primary_table = doc.add_table(rows=1, cols=3); primary_table.style = 'Table Grid'
                 hcells = primary_table.rows[0].cells; hcells[0].text = 'Term'; hcells[1].text = 'Importance'; hcells[2].text = 'Recommended Usage'
                 for term in primary:
                     rcells = primary_table.add_row().cells
                     rcells[0].text = term.get('term', ''); rcells[1].text = f"{term.get('importance', 0):.2f}"; rcells[2].text = str(term.get('recommended_usage', 1))
            # Secondary Terms
            secondary = term_dict.get('secondary_terms', [])
            if secondary:
                 doc.add_heading('Secondary Terms', level=2)
                 secondary_table = doc.add_table(rows=1, cols=2); secondary_table.style = 'Table Grid'
                 hcells = secondary_table.rows[0].cells; hcells[0].text = 'Term'; hcells[1].text = 'Importance'
                 for term in secondary[:15]: # Limit display
                     rcells = secondary_table.add_row().cells
                     rcells[0].text = term.get('term', ''); rcells[1].text = f"{term.get('importance', 0):.2f}"

        # Content Score
        score_dict = score_data or {}
        if score_dict and not guidance_only:
            doc.add_heading('Content Score', level=1)
            score_para = doc.add_paragraph(); score_para.add_run("Overall Score: ").bold = True
            overall_score = score_dict.get('overall_score', 0)
            score_run = score_para.add_run(f"{overall_score} ({score_dict.get('grade', 'F')})")
            if overall_score >= 70: score_run.font.color.rgb = RGBColor(0, 128, 0)
            elif overall_score < 50: score_run.font.color.rgb = RGBColor(255, 0, 0)
            else: score_run.font.color.rgb = RGBColor(255, 165, 0)
            # Component scores
            components = score_dict.get('components', {})
            if components:
                doc.add_heading('Score Components', level=2)
                for component, value in components.items():
                     comp_para = doc.add_paragraph(style='List Bullet')
                     comp_name = component.replace('_score', '').replace('_', ' ').title()
                     comp_para.add_run(f"{comp_name}: ").bold = True; comp_para.add_run(f"{value}")

        # --- Generated Article Content or Guidance (REVISED PARSING) ---
        doc.add_heading('Generated Content' if not guidance_only else 'Writing Guidance', level=1)

        if article_content and isinstance(article_content, str):
            if guidance_only: # Handle Markdown Guidance
                logger.info("Adding Markdown guidance to Word doc.")
                lines = article_content.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line: continue # Skip blank lines

                    # Simple Markdown parsing
                    level = 0
                    text = line
                    if line.startswith('# '): level = 1; text = line[2:]
                    elif line.startswith('## '): level = 2; text = line[3:]
                    elif line.startswith('### '): level = 3; text = line[4:]
                    elif line.startswith('#### '): level = 4; text = line[5:]

                    if level > 0:
                        doc.add_heading(text.strip(), level=level)
                    elif line.startswith(('-', '*')):
                         doc.add_paragraph(line[1:].strip(), style='List Bullet')
                    elif re.match(r'^\d+\.\s', line):
                         doc.add_paragraph(re.sub(r'^\d+\.\s', '', line).strip(), style='List Number')
                    else:
                        # Regular paragraph with basic inline handling
                        para = doc.add_paragraph()
                        add_formatted_text(para, BeautifulSoup(f"<p>{text}</p>", 'html.parser').p) # Wrap in <p> for helper fn

            else: # Handle HTML Article Content
                logger.info("Adding HTML article content to Word doc using BeautifulSoup.")
                soup = BeautifulSoup(article_content, 'html.parser')
                # Process block-level elements sequentially
                for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'ul', 'ol'], recursive=False):
                     try:
                         text_content = element.get_text().strip()
                         if not text_content and element.name not in ['ul', 'ol']: continue

                         # Handle Headings (add Hx: prefix)
                         if element.name.startswith('h'):
                             level = int(element.name[1])
                             prefix = f"H{level}: " if level > 1 else ""
                             doc.add_heading(f"{prefix}{text_content}", level=level)

                         # Handle Paragraphs using helper
                         elif element.name == 'p':
                             para = doc.add_paragraph()
                             add_formatted_text(para, element)

                         # Handle Lists using helper
                         elif element.name in ['ul', 'ol']:
                             list_style = 'List Number' if element.name == 'ol' else 'List Bullet'
                             for li in element.find_all('li', recursive=False):
                                  if li.get_text().strip():
                                       list_para = doc.add_paragraph(style=list_style)
                                       add_formatted_text(list_para, li)

                     except Exception as parse_err:
                         logger.warning(f"Could not parse element: {element.name}. Error: {parse_err}. Adding raw text.")
                         doc.add_paragraph(f"[Parse Error] {element.get_text()[:100]}...")

        # Internal Linking Summary
        links_list = internal_links or []
        if links_list and not guidance_only:
            doc.add_heading('Internal Linking Summary', level=1)
            link_table = doc.add_table(rows=1, cols=3); link_table.style = 'Table Grid'
            hcells = link_table.rows[0].cells; hcells[0].text = 'URL'; hcells[1].text = 'Anchor Text'; hcells[2].text = 'Context'
            for link in links_list:
                 rcells = link_table.add_row().cells
                 rcells[0].text = link.get('url', ''); rcells[1].text = link.get('anchor_text', ''); rcells[2].text = link.get('context', '')

        # --- Final Save ---
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        logger.info("Word document generated successfully.")
        return doc_stream, True

    except Exception as e:
        error_msg = f"Exception in create_word_document: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return BytesIO(), False

###############################################################################
# 11. Content Update Functions (REFACTORED)
###############################################################################

# --- PASTE REFACTORED parse_word_document FUNCTION HERE ---
def parse_word_document(uploaded_file) -> Tuple[Dict, bool]:
    """
    Parse uploaded Word document to extract content structure. Added more error checking.
    Returns: document_content, success_status
    """
    logger.info(f"Parsing Word document: {uploaded_file.name}")
    try:
        doc = Document(BytesIO(uploaded_file.getvalue()))
        document_content = {
            'title': '',
            'headings': [], # List of {'text': '', 'level': int, 'paragraphs': []}
            'paragraphs': [], # List of {'text': '', 'heading': str or None}
            'full_text': ''
        }
        full_text_list = []
        current_heading_obj = None

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue # Skip empty paragraphs

            # Check style name robustness
            style_name = para.style.name.lower() if para.style and para.style.name else ''

            # Check if it's a heading (simple check, might need refinement based on specific styles)
            heading_level = 0
            if 'heading' in style_name:
                match = re.search(r'heading (\d+)', style_name)
                if match:
                    heading_level = int(match.group(1))
                elif style_name == 'heading': # Handle default heading style if unnamed
                    heading_level = 1
                # Add check for 'Title' style as H1 if needed
                elif 'title' in style_name and not document_content['title']:
                     heading_level = 1
                else:
                    # Attempt numeric check if style name *is* a number (less common)
                    try:
                        level_num = int(style_name)
                        if 1 <= level_num <= 6: heading_level = level_num
                    except ValueError: pass # Not a numeric heading style


            if heading_level > 0:
                 current_heading_obj = {
                     'text': text,
                     'level': heading_level,
                     'paragraphs': [] # Store paragraphs under this heading
                 }
                 document_content['headings'].append(current_heading_obj)

                 # Capture the first H1/Title as the document title
                 if heading_level == 1 and not document_content['title']:
                     document_content['title'] = text
                 full_text_list.append(f"[H{heading_level}] {text}") # Add heading marker for full text

            else: # It's a paragraph
                paragraph_obj = {
                    'text': text,
                    'heading': current_heading_obj['text'] if current_heading_obj else None
                }
                document_content['paragraphs'].append(paragraph_obj)

                 # Add paragraph text to the current heading's list
                if current_heading_obj:
                     current_heading_obj['paragraphs'].append(text)

                full_text_list.append(text) # Add paragraph text

        document_content['full_text'] = '\n\n'.join(full_text_list)
        logger.info(f"Parsed Word document. Found {len(document_content['headings'])} headings, {len(document_content['paragraphs'])} paragraphs.")
        return document_content, True

    except Exception as e:
        error_msg = f"Exception in parse_word_document for {uploaded_file.name}: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {}, False

# --- PASTE REFACTORED analyze_content_gaps FUNCTION HERE ---
def analyze_content_gaps(existing_content: Dict, competitor_contents: List[Dict], semantic_structure: Dict,
                        term_data: Dict, score_data: Dict, anthropic_api_key: str,
                        keyword: str, paa_questions: List[Dict] = None) -> Tuple[Dict, bool]:
    """
    Enhanced content gap analysis incorporating scoring data and competitor context.
    Focuses on actionable recommendations. Uses a more robust JSON extraction method.
    Returns: content_gaps (Dictionary), success_status
    """
    logger.info(f"Starting content gap analysis for '{keyword}'.")
    if not isinstance(existing_content, dict) or not existing_content.get('full_text'):
         logger.error("Invalid or empty existing_content provided for gap analysis.")
         return {"error": "Missing existing content data"}, False
    if not isinstance(semantic_structure, dict) or not semantic_structure.get('h1'):
         logger.error("Invalid or empty semantic_structure provided for gap analysis.")
         return {"error": "Missing recommended structure data"}, False
    if not isinstance(term_data, dict):
         logger.warning("Term data is missing or invalid for gap analysis.")
         # Proceed without term data if necessary, but recommendations will be limited
         term_data = {} # Ensure it's a dict to avoid errors later
    if not isinstance(score_data, dict):
         logger.warning("Score data is missing or invalid for gap analysis.")
         score_data = {} # Ensure it's a dict

    try:
        client = anthropic.Anthropic(api_key=anthropic_api_key)

        # --- Prepare inputs for the prompt ---
        # Existing Headings
        existing_headings_list = []
        if isinstance(existing_content.get('headings'), list):
             existing_headings_list = [f"[H{h.get('level', '?')}] {h.get('text', '')}" for h in existing_content['headings']]
        existing_headings_str = "\n".join(existing_headings_list) if existing_headings_list else "No headings found in existing content."

        # Recommended Structure
        recommended_h1 = semantic_structure.get('h1', 'N/A')
        recommended_sections = []
        sections_data = semantic_structure.get('sections', [])
        if isinstance(sections_data, list):
            for sec in sections_data:
                if isinstance(sec, dict):
                    h2 = sec.get('h2')
                    if h2:
                        recommended_sections.append(f"- H2: {h2}")
                        subsections_data = sec.get('subsections', [])
                        if isinstance(subsections_data, list):
                            for sub in subsections_data:
                                if isinstance(sub, dict):
                                    h3 = sub.get('h3')
                                    if h3:
                                        recommended_sections.append(f"  - H3: {h3}")
        recommended_structure_str = f"H1: {recommended_h1}\n" + "\n".join(recommended_sections)

        # Competitor Content Snippets
        competitor_context = ""
        char_limit = 4000
        comp_contents_list = competitor_contents or []
        for comp_content in comp_contents_list:
             if isinstance(comp_content, dict):
                 content_text = comp_content.get('content', '')
                 if content_text:
                     snippet = content_text[:300].strip() + "...\n\n"
                     if len(competitor_context) + len(snippet) < char_limit:
                         competitor_context += snippet
                     else:
                         break

        # PAA Questions String
        paa_questions_list = [q.get('question') for i, q in enumerate(paa_questions or []) if q and q.get('question') and i < 10]
        paa_questions_str = "\n - ".join(paa_questions_list) if paa_questions_list else "N/A"

        # Content Score Summary
        score_summary = "N/A"
        details = score_data.get('details', {}) if isinstance(score_data, dict) else {}
        if score_data and score_data.get('overall_score') is not None:
             missing_primary = []
             underused_primary = []
             # Check terms against counts from score_data['details']
             primary_term_counts = details.get('primary_term_counts', {}) if isinstance(details, dict) else {}
             primary_terms_list = term_data.get('primary_terms', []) if isinstance(term_data, dict) else []
             if isinstance(primary_terms_list, list):
                  for term_info in primary_terms_list:
                      if isinstance(term_info, dict):
                           term = term_info.get('term')
                           rec = term_info.get('recommended_usage', 1)
                           if term and isinstance(primary_term_counts, dict):
                               count = primary_term_counts.get(term, {}).get('count', 0) # Safely get count
                               if count == 0: missing_primary.append(term)
                               elif count < rec: underused_primary.append(f"{term} ({count}/{rec})")

             # Safely get unanswered questions
             question_coverage = details.get('question_coverage', {}) if isinstance(details, dict) else {}
             unanswered_q = [q for q, info in question_coverage.items() if isinstance(info, dict) and not info.get('answered')]

             score_summary = (
                 f"Overall Score: {score_data.get('overall_score', 0)} ({score_data.get('grade', 'F')})\n"
                 # f"Components: {score_data.get('components', {})}\n" # Maybe too verbose for prompt
                 f"Missing Primary Terms (from score): {', '.join(missing_primary) or 'None'}\n"
                 f"Underused Primary Terms (from score): {', '.join(underused_primary) or 'None'}\n"
                 f"Unanswered Questions (from score): {', '.join(unanswered_q) or 'None'}"
             )

        # Key Topics from Term Data
        topics_list = []
        topics_data = term_data.get('topics', []) if isinstance(term_data, dict) else []
        if isinstance(topics_data, list):
            topics_list = [f"- {t.get('topic')}: {t.get('description')}" for i, t in enumerate(topics_data) if isinstance(t, dict) and t.get('topic') and i < 10]
        topics_str = "\n".join(topics_list) if topics_list else "N/A"

        # --- Construct the Prompt ---
        system_prompt = "You are an expert SEO content analyst. Provide actionable recommendations to improve existing content based on competitor analysis, scoring data, and recommended structure. Output ONLY a valid JSON object matching the specified format exactly."

        user_prompt = f"""
        Analyze the existing content about "{keyword}" and provide specific recommendations for improvement.

        EXISTING CONTENT STRUCTURE:
        {existing_headings_str}

        RECOMMENDED CONTENT STRUCTURE (Based on competitor analysis):
        {recommended_structure_str}

        EXISTING CONTENT SCORE SUMMARY:
        {score_summary}

        KEY TOPICS TO COVER (From competitor analysis):
        {topics_str}

        PEOPLE ALSO ASKED QUESTIONS (Potential questions to address):
        {paa_questions_str}

        COMPETITOR CONTENT SNIPPETS (For topic context, do not copy directly):
        {competitor_context[:3500]}

        EXISTING CONTENT (First 1000 chars):
        {existing_content.get('full_text', '')[:1000]}

        TASK: Identify gaps and provide recommendations. Focus on:
        1.  MISSING SECTIONS: Suggest new H2/H3 sections based on the recommended structure or competitor gaps. Include where to insert them.
        2.  HEADING REVISIONS: Suggest improvements to existing headings for clarity or keyword relevance.
        3.  CONTENT GAPS: Identify key topics/subtopics discussed by competitors or required by term analysis but missing or underdeveloped in the existing content. Suggest specific content points to add.
        4.  TERM USAGE: List critical primary/secondary terms missing or significantly underused based on Score Summary or term analysis. Suggest where to incorporate them.
        5.  QUESTION COVERAGE: Identify which PAA questions or questions from Score Summary are not well-addressed and suggest where/how to answer them.
        6.  SEMANTIC RELEVANCE: Point out sections potentially off-topic or too broad for "{keyword}" and how to refocus them.

        OUTPUT FORMAT (Strictly follow this JSON structure):
        ```json
        {{
            "missing_sections": [
                {{
                    "suggested_heading": "New Section Title",
                    "level": 2 or 3,
                    "reason": "Why this section is needed (e.g., competitor coverage, topic gap)",
                    "suggested_content_points": ["Point 1 to cover", "Point 2 to cover"],
                    "insert_after_heading": "Existing Heading Name or 'Introduction' or 'End'"
                }}
            ],
            "revised_headings": [
                {{
                    "original_heading": "Old Heading Text",
                    "suggested_heading": "New Improved Heading Text",
                    "reason": "Reason for change (e.g., clarity, keyword focus)"
                }}
            ],
            "content_gaps_to_fill": [
                {{
                    "topic": "Missing or Underdeveloped Topic",
                    "in_section": "Target Existing Heading Name (or suggest creating new section)",
                    "details": "Specific points or information missing",
                    "competitor_example (optional)": "Brief mention of how competitors cover it"
                }}
            ],
            "term_usage_recommendations": [
                {{
                    "term": "Missing/Underused Term",
                    "type": "Primary or Secondary",
                    "recommendation": "Suggest section/context to add it naturally (e.g., 'In the Introduction', 'When discussing X')"
                }}
            ],
            "questions_to_answer": [
                {{
                    "question": "PAA Question or Question from Score Summary",
                    "recommendation": "Suggest section to add answer or create a FAQ section"
                }}
            ],
            "semantic_focus_recommendations": [
                {{
                    "section": "Existing Heading Name",
                    "issue": "Description of how it's off-topic or too broad",
                    "suggestion": "How to refocus it on '{keyword}'"
                }}
            ]
        }}
        ```
        Ensure the output is ONLY the JSON object, enclosed in ```json ... ```.
        """

        try:
            response = client.messages.create(
                model="claude-3-5-sonnet-20240620", # Use latest Sonnet
                max_tokens=3500, # Allow ample space for detailed JSON
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}],
                temperature=0.2 # Lower temperature for more consistent JSON structure
            )

            raw_response_text = response.content[0].text
            logger.debug(f"Raw response from Claude for gap analysis:\n{raw_response_text[:500]}...")

            # --- Robust JSON Extraction ---
            content_gaps = None
            json_match = re.search(r'```json\s*(\{.*?\})\s*```', raw_response_text, re.DOTALL)
            if json_match:
                json_string = json_match.group(1)
                try:
                    # Attempt to fix common JSON issues before parsing
                    # Remove trailing commas before braces/brackets using regex
                    json_string_fixed = re.sub(r',\s*([\}\]])', r'\1', json_string)
                    # Handle potential unterminated strings (more complex, basic check)
                    # json_string_fixed = json_string_fixed.replace('\n', '\\n') # Escape newlines within potential strings

                    content_gaps = json.loads(json_string_fixed)
                    logger.info("Successfully parsed JSON from Claude response.")
                except json.JSONDecodeError as e:
                    logger.error(f"Failed to decode JSON even after extraction/repair: {e}")
                    logger.error(f"Problematic JSON string snippet: {json_string[:500]}...")
                    content_gaps = {"error": f"JSON Decode Error: {e}", "raw_snippet": json_string[:200]} # Include error info
                    return content_gaps, False # Explicitly fail here

            else:
                 logger.warning("Could not find ```json ... ``` block in Claude response.")
                 content_gaps = {"error": "No JSON block found in LLM response."}
                 return content_gaps, False # Fail if can't find block


            # --- Validation and Return ---
            if content_gaps and isinstance(content_gaps, dict) and "error" not in content_gaps :
                 # Ensure all keys exist, even if empty, for consistency downstream
                 default_keys = {
                     "missing_sections": [], "revised_headings": [], "content_gaps_to_fill": [],
                     "term_usage_recommendations": [], "questions_to_answer": [], "semantic_focus_recommendations": []
                 }
                 for key, default_value in default_keys.items():
                     if key not in content_gaps or not isinstance(content_gaps[key], list): # Ensure keys exist and are lists
                         logger.warning(f"Gap analysis response missing or has invalid type for key: {key}. Setting to default empty list.")
                         content_gaps[key] = default_value
                 logger.info("Content gap analysis successful.")
                 return content_gaps, True
            else:
                 # Return the error structure if parsing failed
                 logger.error(f"Content gap analysis failed during JSON processing. Error: {content_gaps.get('error', 'Unknown parsing failure')}")
                 return content_gaps if isinstance(content_gaps, dict) else {"error": "Failed to generate or parse analysis."}, False


        except anthropic.APIError as e:
            error_msg = f"Anthropic API error during gap analysis: {e}"
            logger.error(error_msg)
            return {"error": error_msg}, False
        except Exception as e:
            error_msg = f"Unexpected exception during gap analysis API call: {e}"
            logger.error(error_msg)
            logger.error(traceback.format_exc())
            return {"error": str(e)}, False

    except Exception as e:
        error_msg = f"Exception in analyze_content_gaps outer scope: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return {"error": f"Outer scope error: {str(e)}"}, False

# --- PASTE REFACTORED create_updated_document FUNCTION HERE ---
def create_updated_document(existing_content: Dict, content_gaps: Dict, keyword: str, score_data: Dict = None) -> Tuple[BytesIO, bool]:
    """
    Creates a Word document outlining the *recommendations* for updating content.
    Uses color and formatting to highlight proposed changes (red for revisions/issues, orange for additions).
    Returns: document_stream, success_status
    """
    logger.info(f"Creating recommendations document for '{keyword}'.")
    if not isinstance(content_gaps, dict) or content_gaps.get("error"): # Check if analysis failed
         logger.error(f"Cannot create recommendations doc due to invalid content_gaps: {content_gaps.get('error', 'Data missing')}")
         return BytesIO(), False
    if not isinstance(existing_content, dict): # Need original for context, though not displayed
         logger.error("Existing content data missing for creating recommendations doc.")
         return BytesIO(), False


    try:
        doc = Document()

        # --- Header and Score Summary ---
        doc.add_heading(f'Content Update Recommendations: {keyword}', 0)
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        # Add content score if available and valid
        score_dict = score_data or {}
        if isinstance(score_dict, dict) and 'overall_score' in score_dict:
             doc.add_heading('Content Score Assessment', 1)
             overall_score = score_dict.get('overall_score', 0); grade = score_dict.get('grade', 'F')
             score_para = doc.add_paragraph(); score_para.add_run(f"Overall Score: ").bold = True
             score_run = score_para.add_run(f"{overall_score} ({grade})")
             if overall_score >= 70: score_run.font.color.rgb = RGBColor(0, 128, 0)
             elif overall_score < 50: score_run.font.color.rgb = RGBColor(255, 0, 0)
             else: score_run.font.color.rgb = RGBColor(255, 165, 0)
             # Optional: Add component score table if needed (similar to brief func)
             # Optional: Add projected score paragraph (similar to brief func)

        # --- Executive Summary ---
        doc.add_heading('Executive Summary of Recommendations', 1)
        summary_lines = [f"Based on analysis for '{keyword}', the following key actions are recommended:"]
        # Generate summary points based on non-empty lists in content_gaps
        if content_gaps.get('semantic_focus_recommendations'): summary_lines.append("- Refocus content sections for better semantic alignment.")
        if content_gaps.get('revised_headings'): summary_lines.append(f"- Revise {len(content_gaps['revised_headings'])} existing headings.")
        if content_gaps.get('missing_sections'): summary_lines.append(f"- Add {len(content_gaps['missing_sections'])} new content sections.")
        if content_gaps.get('content_gaps_to_fill'): summary_lines.append(f"- Address {len(content_gaps['content_gaps_to_fill'])} specific content gaps.")
        if content_gaps.get('term_usage_recommendations'): summary_lines.append(f"- Improve usage of {len(content_gaps['term_usage_recommendations'])} key terms.")
        if content_gaps.get('questions_to_answer'): summary_lines.append(f"- Answer {len(content_gaps['questions_to_answer'])} relevant questions.")

        for line in summary_lines:
             doc.add_paragraph(line, style='List Bullet' if line.startswith('-') else None)


        # --- Detailed Recommendations Sections (with checks for list existence) ---

        # 1. Semantic Focus Recommendations
        semantic_recs = content_gaps.get('semantic_focus_recommendations', [])
        if isinstance(semantic_recs, list) and semantic_recs:
            doc.add_heading('Semantic Focus Recommendations', 1)
            for item in semantic_recs:
                if isinstance(item, dict):
                     doc.add_heading(f"Refocus Section: {item.get('section', 'N/A')}", 2)
                     p_issue = doc.add_paragraph(); p_issue.add_run("Issue: ").bold = True
                     run_issue = p_issue.add_run(item.get('issue', 'N/A')); run_issue.font.color.rgb = RGBColor(255, 0, 0) # Red
                     p_sugg = doc.add_paragraph(); p_sugg.add_run("Suggestion: ").bold = True; p_sugg.add_run(item.get('suggestion', 'N/A'))
                     doc.add_paragraph() # Spacing

        # 2. Heading Revisions
        revised_h = content_gaps.get('revised_headings', [])
        if isinstance(revised_h, list) and revised_h:
            doc.add_heading('Heading Revisions', 1)
            table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
            hcells = table.rows[0].cells; hcells[0].text='Original Heading'; hcells[1].text='Suggested Heading'
            for item in revised_h:
                if isinstance(item, dict):
                     rcells = table.add_row().cells
                     run_orig = rcells[0].paragraphs[0].add_run(item.get('original_heading', '')); run_orig.font.strike = True; run_orig.font.color.rgb = RGBColor(128, 128, 128)
                     run_sugg = rcells[1].paragraphs[0].add_run(item.get('suggested_heading', '')); run_sugg.font.color.rgb = RGBColor(255, 0, 0) # Red
            doc.add_paragraph() # Spacing

        # 3. Missing Sections to Add
        missing_s = content_gaps.get('missing_sections', [])
        if isinstance(missing_s, list) and missing_s:
            doc.add_heading('Missing Sections to Add', 1)
            for item in missing_s:
                if isinstance(item, dict):
                     level = item.get('level', 2)
                     heading_text = item.get('suggested_heading', 'New Section')
                     h = doc.add_heading(heading_text, level=level)
                     for run in h.runs: run.font.color.rgb = RGBColor(255, 165, 0) # Orange

                     p_reason = doc.add_paragraph(); p_reason.add_run("Reason: ").bold = True; p_reason.add_run(item.get('reason', 'N/A'))
                     p_insert = doc.add_paragraph(); p_insert.add_run("Insert After: ").bold = True; p_insert.add_run(item.get('insert_after_heading', 'End'))

                     points = item.get('suggested_content_points', [])
                     if isinstance(points, list) and points:
                         p_points = doc.add_paragraph(); p_points.add_run("Content Points:").bold = True
                         for point in points:
                             p_bullet = doc.add_paragraph(style='List Bullet')
                             run_bullet = p_bullet.add_run(str(point)) # Ensure string
                             run_bullet.font.color.rgb = RGBColor(255, 165, 0) # Orange
                     doc.add_paragraph() # Spacing

        # 4. Content Gaps to Fill
        gaps_fill = content_gaps.get('content_gaps_to_fill', [])
        if isinstance(gaps_fill, list) and gaps_fill:
            doc.add_heading('Content Gaps to Fill', 1)
            for item in gaps_fill:
                if isinstance(item, dict):
                     doc.add_heading(f"Topic: {item.get('topic', 'N/A')}", 2)
                     p_section = doc.add_paragraph(); p_section.add_run("Target Section: ").bold = True; p_section.add_run(item.get('in_section', 'N/A'))
                     p_details = doc.add_paragraph(); p_details.add_run("Details to Add: ").bold = True
                     run_details = p_details.add_run(item.get('details', '')); run_details.font.color.rgb = RGBColor(255, 165, 0) # Orange
                     if item.get('competitor_example'):
                          p_comp = doc.add_paragraph(); p_comp.add_run("Competitor Note: ").italic = True; p_comp.add_run(item['competitor_example'])
                     doc.add_paragraph() # Spacing

        # 5. Term Usage Recommendations
        term_recs = content_gaps.get('term_usage_recommendations', [])
        if isinstance(term_recs, list) and term_recs:
            doc.add_heading('Term Usage Recommendations', 1)
            table = doc.add_table(rows=1, cols=3); table.style = 'Table Grid'
            hcells = table.rows[0].cells; hcells[0].text='Term'; hcells[1].text='Type'; hcells[2].text='Suggestion Area'
            for item in term_recs:
                if isinstance(item, dict):
                     rcells = table.add_row().cells
                     term_run = rcells[0].paragraphs[0].add_run(item.get('term', ''))
                     term_run.font.color.rgb = RGBColor(255, 0, 0) # Red
                     rcells[1].text = item.get('type', '')
                     rcells[2].text = item.get('recommendation', '')
            doc.add_paragraph() # Spacing

        # 6. Questions to Answer
        q_answer = content_gaps.get('questions_to_answer', [])
        if isinstance(q_answer, list) and q_answer:
             doc.add_heading('Questions to Answer', 1)
             for item in q_answer:
                 if isinstance(item, dict):
                     p_q = doc.add_paragraph(); p_q.add_run("Question: ").bold = True; p_q.add_run(item.get('question', '')).italic = True
                     p_rec = doc.add_paragraph(); p_rec.add_run("Recommendation: ").bold = True; p_rec.add_run(item.get('recommendation', ''))
                     doc.add_paragraph() # Spacing


        # --- Final Save ---
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        logger.info("Recommendations document created successfully.")
        return doc_stream, True

    except Exception as e:
        error_msg = f"Exception in create_updated_document: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return BytesIO(), False

# --- PASTE REFACTORED generate_optimized_article_with_tracking FUNCTION HERE ---
def generate_optimized_article_with_tracking(existing_content: Dict, competitor_contents: List[Dict],
                              semantic_structure: Dict, related_keywords: List[Dict],
                              keyword: str, paa_questions: List[Dict], term_data: Dict,
                              anthropic_api_key: str, target_word_count: int = 1800) -> Tuple[str, str, bool]:
    """
    Generates a *new* optimized article based on analysis, optionally using *sections* of existing content
    as context, but prioritizing the recommended structure and fresh writing.
    Change tracking is simplified to summary level. Uses color spans for *illustrative purposes only* in the summary.
    The generated article content itself will be clean HTML.

    Returns: optimized_html_content, change_summary (HTML), success_status
    """
    logger.info(f"Generating new optimized article for '{keyword}' based on analysis.")
    if not isinstance(semantic_structure, dict) or not semantic_structure.get('h1'):
        logger.error("Invalid semantic structure provided for generating optimized article.")
        return "", "<p>Error: Missing recommended structure.</p>", False

    try:
        # Use the standard article generation function, passing relevant context.
        optimized_html_content, success = generate_article(
            keyword=keyword,
            semantic_structure=semantic_structure, # Use the recommended structure directly
            related_keywords=related_keywords or [], # Ensure list
            serp_features=[], # Not directly used
            paa_questions=paa_questions or [], # Ensure list
            term_data=term_data or {}, # Ensure dict
            anthropic_api_key=anthropic_api_key,
            competitor_contents=competitor_contents or [], # Ensure list
            guidance_only=False # We want the full article
        )

        if not success:
            logger.error("Failed to generate the base optimized article.")
            # optimized_html_content might contain an error message from generate_article
            return optimized_html_content, f"<h3>Error generating optimized article.</h3><p>{optimized_html_content}</p>", False

        # --- Generate a Simplified Change Summary ---
        # Compare headings between original (if available) and new structure
        original_headings = {}
        if isinstance(existing_content, dict) and isinstance(existing_content.get('headings'), list):
             original_headings = {f"[H{h.get('level')}] {h.get('text')}": h for h in existing_content['headings'] if isinstance(h, dict)}

        new_h1 = semantic_structure.get('h1', '')
        new_headings_set = {new_h1} if new_h1 else set()
        new_headings_list = [f"[H1] {new_h1}"] if new_h1 else []

        sections = semantic_structure.get('sections', [])
        if isinstance(sections, list):
            for section in sections:
                if isinstance(section, dict):
                     h2 = section.get('h2')
                     if h2:
                         new_headings_set.add(h2)
                         new_headings_list.append(f"[H2] {h2}")
                         subsections = section.get('subsections', [])
                         if isinstance(subsections, list):
                             for sub in subsections:
                                 if isinstance(sub, dict):
                                     h3 = sub.get('h3')
                                     if h3:
                                         new_headings_set.add(h3)
                                         new_headings_list.append(f"[H3] {h3}")

        # Identify potentially kept/modified/removed sections at a high level
        kept_modified_headings = []
        removed_headings = []
        added_headings = new_headings_list[:] # Start with all new headings as added

        for orig_key, orig_heading_data in original_headings.items():
             orig_text = orig_heading_data.get('text')
             if orig_text and orig_text in new_headings_set:
                  # If original heading text found in the set of new headings
                  kept_modified_headings.append(orig_key)
                  # Remove corresponding item from added_headings list if found
                  possible_new_keys = [f"[H1] {orig_text}", f"[H2] {orig_text}", f"[H3] {orig_text}"] # Consider levels
                  for key_to_remove in possible_new_keys:
                      if key_to_remove in added_headings:
                          added_headings.remove(key_to_remove)
                          break # Assume only one match needed
             elif orig_text: # If original heading text not found in new set
                  removed_headings.append(orig_key)


        change_summary = f"""
        <div class="change-summary" style="padding: 15px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9; margin-bottom: 20px;">
            <h3 style="margin-top: 0;">Optimization Summary for "{keyword}"</h3>
            <p>This article was generated based on the recommended structure and analysis, aiming for optimal SEO performance. It incorporates key terms, addresses relevant questions, and aligns with competitor best practices.</p>

            <h4 style="margin-bottom: 5px;">Structural Overview (Approximate):</h4>
            <ul style="margin-top: 0; padding-left: 20px;">
                <li><strong>New Structure Adopted:</strong> The content follows the recommended semantic hierarchy.</li>
                {f"<li><strong>Sections Likely Kept/Modified:</strong> {len(kept_modified_headings)} original sections may form the basis of new sections.</li>" if kept_modified_headings else ""}
                {f"<li><strong>New Sections Likely Added:</strong> {len(added_headings)} new sections/subsections introduced.</li>" if added_headings else ""}
                {f"<li><strong>Original Sections Likely Replaced:</strong> {len(removed_headings)} original sections replaced or significantly altered.</li>" if removed_headings else ""}
            </ul>

            <h4 style="margin-bottom: 5px;">Content Improvements:</h4>
            <ul style="margin-top: 0; padding-left: 20px;">
                <li>Content automatically generated to target primary terms and topics from analysis.</li>
                <li>Relevant 'People Also Asked' questions considered during content generation.</li>
                <li>Focus maintained on the core keyword "{keyword}".</li>
                <li>Structure optimized for readability and search engine crawling based on analysis.</li>
            </ul>
            <p><i>Note: As this is newly generated content based on analysis, detailed change tracking is not applicable. The goal is an optimized final product.</i></p>
        </div>
        """

        logger.info("Optimized article and summary generated successfully.")
        return optimized_html_content, change_summary, True

    except Exception as e:
        error_msg = f"Exception in generate_optimized_article_with_tracking: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return "", f"<h3>Error generating optimized article: {str(e)}</h3>", False

# --- PASTE REFACTORED create_word_document_from_html FUNCTION HERE ---
def create_word_document_from_html(html_content: str, keyword: str, change_summary: str = "",
                                  score_data: Dict = None) -> BytesIO:
    """
    Creates a Word document from generated HTML content.
    Includes score data and change summary if provided.
    Uses helper functions to parse HTML and apply formatting, including colors/styles from spans.

    Returns: document_stream
    """
    logger.info(f"Creating Word document from HTML for '{keyword}'.")
    if not html_content or not isinstance(html_content, str):
         logger.error("Cannot create Word doc: HTML content is missing or invalid.")
         return BytesIO()
    try:
        doc = Document()

        # --- Header, Date, Score ---
        doc.add_heading(f'Optimized Content: {keyword}', 0)
        doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        if isinstance(score_data, dict) and 'overall_score' in score_data:
             doc.add_heading("Content Score", 1)
             overall_score = score_data.get('overall_score', 0); grade = score_data.get('grade', 'F')
             score_para = doc.add_paragraph(); score_para.add_run(f"Overall Score: ").bold = True
             score_run = score_para.add_run(f"{overall_score} ({grade})")
             if overall_score >= 70: score_run.font.color.rgb = RGBColor(0, 128, 0)
             elif overall_score < 50: score_run.font.color.rgb = RGBColor(255, 0, 0)
             else: score_run.font.color.rgb = RGBColor(255, 165, 0)

        # --- Change Summary ---
        if change_summary and isinstance(change_summary, str):
            doc.add_heading("Optimization Summary", 1)
            summary_soup = BeautifulSoup(change_summary, 'html.parser')
            # Parse summary HTML simply
            for element in summary_soup.find_all(['h3', 'h4', 'p', 'li']):
                 text = element.get_text().strip()
                 if not text: continue
                 if element.name == 'h3': doc.add_heading(text, level=2)
                 elif element.name == 'h4': doc.add_heading(text, level=3)
                 elif element.name == 'p': doc.add_paragraph(text)
                 elif element.name == 'li': doc.add_paragraph(text, style='List Bullet')

            doc.add_paragraph("_" * 50) # Separator

        # --- Main Content Parsing ---
        doc.add_heading("Optimized Article Content", 1)
        soup = BeautifulSoup(html_content, 'html.parser')

        # Process block-level elements sequentially
        for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'ul', 'ol'], recursive=False):
            try:
                text_content = element.get_text().strip()
                # Skip elements that only contain whitespace or are empty list containers
                if not text_content and element.name not in ['ul', 'ol']:
                     continue

                # Handle Headings (add Hx: prefix for clarity)
                if element.name.startswith('h'):
                    level = int(element.name[1])
                    prefix = f"H{level}: " if level > 1 else ""
                    heading = doc.add_heading(f"{prefix}" + text_content, level=level)
                    # Attempt to apply style from first span if present (basic check)
                    first_span = element.find('span', recursive=False)
                    if first_span and 'style' in first_span.attrs:
                        styles = parse_style_attribute(first_span['style'])
                        if styles.get('color'): heading.runs[0].font.color.rgb = styles['color']
                        if styles.get('strike'): heading.runs[0].font.strike = True

                # Handle Paragraphs using helper
                elif element.name == 'p':
                    para = doc.add_paragraph()
                    add_formatted_text(para, element)

                # Handle Lists using helper
                elif element.name in ['ul', 'ol']:
                    list_style = 'List Number' if element.name == 'ol' else 'List Bullet'
                    # Process only immediate children 'li' tags
                    for li in element.find_all('li', recursive=False):
                         if li.get_text(strip=True): # Only add if li has content
                             list_para = doc.add_paragraph(style=list_style)
                             add_formatted_text(list_para, li) # Apply formatting within li

            except Exception as parse_err:
                logger.warning(f"Could not parse element: {element.name}. Error: {parse_err}. Adding raw text.")
                doc.add_paragraph(f"[Parse Error] {element.get_text(strip=True)[:100]}...")

        # --- Final Save ---
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        logger.info("Word document from HTML created successfully.")
        return doc_stream

    except Exception as e:
        logger.error(f"Exception in create_word_document_from_html: {str(e)}")
        logger.error(traceback.format_exc())
        return BytesIO()


###############################################################################
# 12. Main Streamlit App
###############################################################################

def main():
    st.title("📊 SEO Content Optimizer")

    # Sidebar for API credentials
    st.sidebar.header("API Credentials")

    dataforseo_login = st.sidebar.text_input("DataForSEO API Login", type="password", key="dfs_login")
    dataforseo_password = st.sidebar.text_input("DataForSEO API Password", type="password", key="dfs_pass")

    openai_api_key = st.sidebar.text_input("OpenAI API Key", type="password", key="openai_key")
    anthropic_api_key = st.sidebar.text_input("Anthropic API Key", type="password", key="anthropic_key")

    # Initialize session state
    if 'results' not in st.session_state:
        st.session_state.results = {}
    # Ensure necessary keys exist even if empty after initialization or error
    default_keys = ['keyword', 'organic_results', 'serp_features', 'paa_questions',
                    'related_keywords', 'scraped_contents', 'semantic_structure',
                    'term_data', 'article_content', 'guidance_content', 'guidance_only',
                    'meta_title', 'meta_description', 'content_score', 'highlighted_content',
                    'content_suggestions', 'doc_stream', 'article_with_links', 'internal_links',
                    'existing_content', 'content_gaps', 'updated_doc', 'optimized_content',
                    'change_summary', 'existing_content_score', 'optimized_content_score']
    for key in default_keys:
        if key not in st.session_state.results:
             st.session_state.results[key] = None # Use None as default


    # Main content Tabs
    tabs = st.tabs([
        "Input & SERP Analysis",
        "Content Analysis",
        "Article Generation",
        "Internal Linking",
        "SEO Brief",
        "Content Updates",
        "Content Scoring"
    ])

    # Tab 1: Input & SERP Analysis
    with tabs[0]:
        st.header("Enter Target Keyword")
        keyword_input = st.text_input("Target Keyword", value=st.session_state.results.get('keyword', ''))

        if st.button("Fetch SERP Data", key="fetch_serp"):
            st.session_state.results = {} # Clear previous results on new fetch
            if not keyword_input:
                st.error("Please enter a target keyword.")
            elif not dataforseo_login or not dataforseo_password:
                st.error("Please enter DataForSEO API credentials.")
            elif not openai_api_key or not anthropic_api_key:
                st.error("Please enter API keys for OpenAI and Anthropic.")
            else:
                with st.spinner("Fetching SERP data and related keywords..."):
                    start_time = time.time()
                    # Clear previous results explicitly
                    st.session_state.results = {'keyword': keyword_input}

                    # Fetch SERP results
                    organic_results, serp_features, paa_questions, serp_success = fetch_serp_results(
                        keyword_input, dataforseo_login, dataforseo_password
                    )
                    if serp_success:
                        st.session_state.results['organic_results'] = organic_results
                        st.session_state.results['serp_features'] = serp_features
                        st.session_state.results['paa_questions'] = paa_questions
                        st.success("SERP data fetched.")

                        # Fetch related keywords using DataForSEO
                        st.text("Fetching related keywords...")
                        related_keywords, kw_success = fetch_related_keywords_dataforseo(
                            keyword_input, dataforseo_login, dataforseo_password
                        )
                        if kw_success:
                            # Validate related keywords data (ensure numeric types)
                            validated_keywords = []
                            for kw in related_keywords:
                                 try:
                                       sv = kw.get('search_volume')
                                       cpc = kw.get('cpc')
                                       comp = kw.get('competition')
                                       validated_kw = {
                                           'keyword': kw.get('keyword', ''),
                                           'search_volume': int(sv) if sv is not None else 0,
                                           'cpc': float(cpc) if cpc is not None else 0.0,
                                           'competition': float(comp) if comp is not None else 0.0
                                       }
                                       validated_keywords.append(validated_kw)
                                 except (ValueError, TypeError) as e:
                                      logger.warning(f"Skipping keyword due to type conversion error '{kw.get('keyword')}': {e}")
                            st.session_state.results['related_keywords'] = validated_keywords
                            st.success("Related keywords fetched.")
                        else:
                            st.warning("Failed to fetch related keywords, proceeding without them.")
                            st.session_state.results['related_keywords'] = []


                        st.success(f"SERP & Keyword analysis completed in {format_time(time.time() - start_time)}")
                        # Rerun to update display immediately after fetch
                        st.experimental_rerun() # Use st.rerun() in newer versions
                    else:
                        st.error("Failed to fetch SERP data. Please check API credentials or keyword.")
                        st.session_state.results = {} # Clear results on failure

        # Display fetched data if available
        if st.session_state.results.get('organic_results'):
            st.subheader("Top 10 Organic Results")
            try:
                 df_results = pd.DataFrame(st.session_state.results['organic_results'])
                 st.dataframe(df_results[['rank_group', 'title', 'url', 'page_type']])
            except Exception as e:
                 st.error(f"Error displaying SERP results: {e}")
                 st.write(st.session_state.results['organic_results']) # Show raw data on error

        if st.session_state.results.get('serp_features'):
             st.subheader("SERP Features")
             try:
                  df_features = pd.DataFrame(st.session_state.results['serp_features'])
                  st.dataframe(df_features)
             except Exception as e:
                  st.error(f"Error displaying SERP features: {e}")

        paa = st.session_state.results.get('paa_questions')
        if paa:
             st.subheader("People Also Asked")
             for q_data in paa:
                  if isinstance(q_data, dict): st.write(f"- {q_data.get('question', 'N/A')}")

        related_kws = st.session_state.results.get('related_keywords')
        if related_kws:
             st.subheader("Related Keywords")
             try:
                  df_keywords = pd.DataFrame(related_kws)
                  st.dataframe(df_keywords)
             except Exception as e:
                  st.error(f"Error displaying related keywords: {e}")


    # Tab 2: Content Analysis
    with tabs[1]:
        st.header("Content Analysis")

        if not st.session_state.results.get('organic_results'):
            st.warning("Please fetch SERP data first (in the 'Input & SERP Analysis' tab).")
        else:
            if st.button("Analyze Competitor Content", key="analyze_content"):
                if not anthropic_api_key:
                    st.error("Please enter Anthropic API key.")
                else:
                    with st.spinner("Scraping & analyzing top pages... This may take a minute."):
                        start_time = time.time()
                        st.session_state.results['scraped_contents'] = None # Clear previous
                        st.session_state.results['semantic_structure'] = None
                        st.session_state.results['term_data'] = None

                        # Scrape content
                        scraped_contents = []
                        urls_to_scrape = [res.get('url') for res in st.session_state.results['organic_results'] if res.get('url')]
                        progress_bar = st.progress(0)
                        status_text = st.empty()

                        for i, url in enumerate(urls_to_scrape):
                             status_text.text(f"Scraping ({i+1}/{len(urls_to_scrape)}): {url[:70]}...")
                             content, success = scrape_webpage(url)
                             if success and content and "[Error" not in content and "[Content not" not in content:
                                 # Get headings too
                                 # headings = extract_headings(url) # Can uncomment if headings needed separately later
                                 scraped_contents.append({
                                     'url': url,
                                     'title': next((r.get('title') for r in st.session_state.results['organic_results'] if r.get('url') == url),'N/A'),
                                     'content': content
                                     # 'headings': headings
                                 })
                             else:
                                 logger.warning(f"Failed to scrape or get meaningful content from: {url}")
                             progress_bar.progress((i + 1) / len(urls_to_scrape))
                        status_text.text("Scraping complete.")

                        if not scraped_contents:
                            st.error("Could not scrape sufficient content from competitors. Cannot perform analysis.")
                        else:
                            st.session_state.results['scraped_contents'] = scraped_contents

                            # Analyze semantic structure
                            status_text.text("Analyzing content structure...")
                            semantic_structure, structure_success = analyze_semantic_structure(
                                scraped_contents, anthropic_api_key
                            )
                            if structure_success:
                                st.session_state.results['semantic_structure'] = semantic_structure
                                logger.info("Semantic structure analysis successful.")
                            else:
                                 st.error("Failed to analyze semantic structure.")
                                 logger.error("Semantic structure analysis failed.")

                            # Extract important terms
                            status_text.text("Extracting key terms and topics...")
                            term_data, term_success = extract_important_terms(
                                scraped_contents, anthropic_api_key
                            )
                            if term_success:
                                st.session_state.results['term_data'] = term_data
                                logger.info("Term extraction successful.")
                            else:
                                st.error("Failed to extract important terms.")
                                logger.error("Term extraction failed.")


                            status_text.text(f"Content analysis completed in {format_time(time.time() - start_time)}")
                            st.success("Analysis complete!")
                            # Rerun to display results
                            st.experimental_rerun() # Use st.rerun() in newer versions


            # Display analysis results if available
            semantic_structure = st.session_state.results.get('semantic_structure')
            term_data = st.session_state.results.get('term_data')

            if semantic_structure:
                st.subheader("Recommended Semantic Structure")
                try:
                    st.write(f"**H1:** {semantic_structure.get('h1', 'N/A')}")
                    sections = semantic_structure.get('sections', [])
                    if isinstance(sections, list):
                         for i, section in enumerate(sections, 1):
                             if isinstance(section, dict):
                                 st.write(f"**H2 {i}:** {section.get('h2', 'N/A')}")
                                 subsections = section.get('subsections', [])
                                 if isinstance(subsections, list):
                                      for j, subsection in enumerate(subsections, 1):
                                          if isinstance(subsection, dict):
                                              st.write(f"  - **H3 {i}.{j}:** {subsection.get('h3', 'N/A')}")
                    else:
                         st.warning("Structure 'sections' format is invalid.")
                except Exception as e:
                    st.error(f"Error displaying semantic structure: {e}")
                    st.json(semantic_structure) # Show raw on error

            if term_data:
                with st.expander("View Extracted Terms & Topics", expanded=semantic_structure is None): # Expand if structure failed
                    st.subheader("Important Terms")
                    try:
                        primary = term_data.get('primary_terms')
                        if primary:
                            st.write("**Primary Terms:**")
                            st.dataframe(pd.DataFrame(primary))
                        secondary = term_data.get('secondary_terms')
                        if secondary:
                            st.write("**Secondary Terms:**")
                            st.dataframe(pd.DataFrame(secondary))
                        topics = term_data.get('topics')
                        if topics:
                             st.write("**Key Topics:**")
                             st.dataframe(pd.DataFrame(topics))
                        questions = term_data.get('questions')
                        if questions:
                             st.write("**Implied Questions:**")
                             for q in questions: st.write(f"- {q}")
                    except Exception as e:
                         st.error(f"Error displaying term data: {e}")
                         st.json(term_data) # Show raw on error

    # Tab 3: Article Generation
    with tabs[2]:
        st.header("Article Generation")

        # Check prerequisites
        analysis_complete = st.session_state.results.get('semantic_structure') and st.session_state.results.get('term_data')
        competitors_scraped = st.session_state.results.get('scraped_contents')

        if not analysis_complete or not competitors_scraped:
            st.warning("Please complete SERP Analysis and Content Analysis first.")
        else:
            content_type = st.radio(
                "Generation Type:",
                ["Full Article", "Writing Guidance Only"],
                key="gen_type",
                index=0 if not st.session_state.results.get('guidance_only') else 1 # Default based on previous run
            )
            guidance_only = (content_type == "Writing Guidance Only")

            # Prepare data needed for generation (handle potential None values)
            keyword = st.session_state.results.get('keyword', '')
            semantic_structure = st.session_state.results.get('semantic_structure', {})
            related_keywords = st.session_state.results.get('related_keywords', [])
            serp_features = st.session_state.results.get('serp_features', []) # Although not used directly in refactored fn
            paa_questions = st.session_state.results.get('paa_questions', [])
            term_data = st.session_state.results.get('term_data', {})
            competitor_contents = st.session_state.results.get('scraped_contents', [])

            # Button text depends on radio selection
            button_label = "Generate " + ("Guidance" if guidance_only else "Article") + " & Meta Tags"

            if st.button(button_label, key="generate_article"):
                if not anthropic_api_key:
                    st.error("Please enter Anthropic API key.")
                elif not keyword or not semantic_structure or not term_data:
                     st.error("Missing necessary analysis data (Keyword, Structure, Terms). Please re-run previous steps.")
                else:
                    with st.spinner(f"Generating {content_type.lower()}..."):
                        start_time = time.time()
                        st.session_state.results['article_content'] = None # Clear previous
                        st.session_state.results['guidance_content'] = None
                        st.session_state.results['meta_title'] = None
                        st.session_state.results['meta_description'] = None
                        st.session_state.results['content_score'] = None

                        # Generate Article or Guidance
                        article_content, article_success = generate_article(
                            keyword=keyword,
                            semantic_structure=semantic_structure,
                            related_keywords=related_keywords,
                            serp_features=serp_features,
                            paa_questions=paa_questions,
                            term_data=term_data,
                            anthropic_api_key=anthropic_api_key,
                            competitor_contents=competitor_contents,
                            guidance_only=guidance_only
                        )

                        st.session_state.results['guidance_only'] = guidance_only

                        if article_success:
                            if guidance_only:
                                st.session_state.results['guidance_content'] = article_content
                            else:
                                st.session_state.results['article_content'] = article_content

                            # Attempt Meta Tag generation regardless of article success for now
                            st.text("Generating meta tags...")
                            meta_title, meta_description, meta_success = generate_meta_tags(
                                keyword, semantic_structure, related_keywords, term_data, anthropic_api_key
                            )
                            if meta_success:
                                st.session_state.results['meta_title'] = meta_title
                                st.session_state.results['meta_description'] = meta_description
                            else:
                                st.warning("Failed to generate meta tags.")

                             # Attempt to score the content if full article
                            if not guidance_only and article_content:
                                st.text("Scoring generated content...")
                                score_data, score_success = score_content(
                                     article_content, term_data, keyword
                                )
                                if score_success:
                                     st.session_state.results['content_score'] = score_data
                                else:
                                     st.warning("Failed to score generated content.")

                            st.success(f"Generation completed in {format_time(time.time() - start_time)}")
                            st.experimental_rerun() # Update display
                        else:
                            st.error(f"Failed to generate {'guidance' if guidance_only else 'article'}.")
                            # Display the error content returned by the function
                            st.markdown("--- ERROR OUTPUT ---")
                            st.code(article_content, language='html' if not guidance_only else 'markdown')


        # Display generated content & meta tags if available
        meta_title = st.session_state.results.get('meta_title')
        meta_description = st.session_state.results.get('meta_description')
        article_content = st.session_state.results.get('article_content')
        guidance_content = st.session_state.results.get('guidance_content')
        is_guidance_mode = st.session_state.results.get('guidance_only', False)
        content_score_data = st.session_state.results.get('content_score')

        if meta_title or meta_description:
            st.subheader("Generated Meta Tags")
            st.write(f"**Title:** {meta_title or 'N/A Generation Failed'}")
            st.write(f"**Description:** {meta_description or 'N/A Generation Failed'}")

        if content_score_data and not is_guidance_mode:
             st.subheader("Content Score (Generated Article)")
             score = content_score_data.get('overall_score', 0); grade = content_score_data.get('grade', 'F')
             score_color = "green" if score >= 70 else "red" if score < 50 else "orange"
             st.markdown(f"""<div style="background-color: #f0f0f0; padding: 10px; border-radius: 5px;">
                         <h3 style="margin:0;">Score: <span style="color:{score_color};">{score} ({grade})</span></h3></div>""", unsafe_allow_html=True)

        if is_guidance_mode and guidance_content:
            st.subheader("Generated Writing Guidance")
            st.markdown(guidance_content) # Display Markdown
        elif not is_guidance_mode and article_content:
             st.subheader("Generated Article Content")
             st.markdown(article_content, unsafe_allow_html=True) # Display HTML

    # Tab 4: Internal Linking
    with tabs[3]:
        st.header("Internal Linking")

        # Check if we have a *generated full article*
        article_content = st.session_state.results.get('article_content')
        is_guidance_mode = st.session_state.results.get('guidance_only', False)

        if is_guidance_mode:
            st.warning("Internal linking is only available for fully generated articles, not writing guidance.")
        elif not article_content:
            st.warning("Please generate a full article first (in the 'Article Generation' tab).")
        else:
            st.write("Upload a spreadsheet (CSV/XLSX) with your site pages.")
            st.write("Required columns: `URL`, `Title`, `Meta Description` (case-insensitive)")

            # Sample template button
            if st.button("Download Sample CSV Template", key="download_tmpl"):
                sample_data = {'URL': ['https://example.com/page1'], 'Title': ['Example Page 1'], 'Meta Description': ['Description']}
                sample_df = pd.DataFrame(sample_data)
                csv_bytes = sample_df.to_csv(index=False).encode('utf-8')
                st.download_button(label="Click to Download", data=csv_bytes, file_name="site_pages_template.csv", mime="text/csv")

            pages_file = st.file_uploader("Upload Site Pages Spreadsheet", type=['csv', 'xlsx', 'xls'], key="pages_upload")

            batch_size = st.slider("Embedding Batch Size", 5, 50, defaultValue=20, key="batch_size_link",
                                   help="Larger is faster but uses more memory/API quota per call.")

            if st.button("Generate Internal Links", key="gen_links_btn"):
                if not openai_api_key: st.error("OpenAI API key required for embeddings.")
                # if not anthropic_api_key: st.error("Anthropic API key needed (if using fallback for anchor text).") # No longer strictly requires Anthropic
                elif not pages_file: st.error("Please upload the site pages spreadsheet.")
                else:
                    with st.spinner("Processing site pages and generating links..."):
                        start_time = time.time()
                        st.session_state.results['article_with_links'] = None # Clear previous
                        st.session_state.results['internal_links'] = None

                        pages, parse_success = parse_site_pages_spreadsheet(pages_file)
                        if not parse_success:
                            st.error("Failed to parse the uploaded spreadsheet. Check format and required columns.")
                        elif not pages:
                             st.error("No valid pages found in the spreadsheet.")
                        else:
                            # Embed Pages
                            status_text = st.empty(); status_text.text(f"Generating embeddings for {len(pages)} pages...")
                            pages_with_embeddings, embed_success = embed_site_pages(pages, openai_api_key, batch_size)
                            if not embed_success:
                                status_text.error("Failed to generate embeddings for site pages.")
                            else:
                                # Generate Links
                                status_text.text("Analyzing content and finding link opportunities...")
                                word_count = len(re.findall(r'\b\w+\b', article_content))
                                article_with_links, links_added, links_success = generate_internal_links_with_embeddings(
                                    article_content, pages_with_embeddings, openai_api_key, anthropic_api_key, word_count
                                ) # Pass keys needed by the function

                                if links_success:
                                    st.session_state.results['article_with_links'] = article_with_links
                                    st.session_state.results['internal_links'] = links_added
                                    status_text.success(f"Internal linking completed in {format_time(time.time() - start_time)}")
                                    st.experimental_rerun() # Update display
                                else:
                                    status_text.error("Failed to generate internal links.")
                                    # Maybe show the original article still?
                                    st.session_state.results['article_with_links'] = article_content # Show original on failure
                                    st.experimental_rerun()

            # Display article with links and summary table
            article_with_links = st.session_state.results.get('article_with_links')
            links_added = st.session_state.results.get('internal_links')

            if article_with_links:
                 st.subheader("Article with Internal Links")
                 st.markdown(article_with_links, unsafe_allow_html=True)

            if links_added:
                 st.subheader("Internal Links Added Summary")
                 try:
                     df_links = pd.DataFrame(links_added)
                     st.dataframe(df_links[['anchor_text', 'url', 'page_title', 'similarity_score', 'context']])
                 except Exception as e:
                     st.error(f"Error displaying links summary table: {e}")
                     st.write(links_added) # Show raw list on error


    # Tab 5: SEO Brief
    with tabs[4]:
        st.header("SEO Brief & Downloadable Report")

        # Check if essential data exists
        required_data_keys = ['keyword', 'organic_results', 'semantic_structure']
        # Plus either article_content or guidance_content
        has_content = st.session_state.results.get('article_content') or st.session_state.results.get('guidance_content')

        if not all(st.session_state.results.get(key) for key in required_data_keys) or not has_content:
            st.warning("Please complete Input, Analysis, and Generation steps first.")
        else:
            # Determine which content to use for the brief
            is_guidance = st.session_state.results.get('guidance_only', False)
            content_for_brief = st.session_state.results.get('guidance_content') if is_guidance else st.session_state.results.get('article_with_links', st.session_state.results.get('article_content'))
            internal_links_for_brief = st.session_state.results.get('internal_links') if not is_guidance else None
            score_for_brief = st.session_state.results.get('content_score') if not is_guidance else None

            if st.button("Generate SEO Brief Document", key="gen_brief_btn"):

                with st.spinner("Generating Word document..."):
                    start_time = time.time()
                    st.session_state.results['doc_stream'] = None # Clear previous

                    doc_stream, doc_success = create_word_document(
                        keyword=st.session_state.results['keyword'],
                        serp_results=st.session_state.results.get('organic_results', []),
                        related_keywords=st.session_state.results.get('related_keywords', []),
                        semantic_structure=st.session_state.results.get('semantic_structure', {}),
                        article_content=content_for_brief,
                        meta_title=st.session_state.results.get('meta_title', ''),
                        meta_description=st.session_state.results.get('meta_description', ''),
                        paa_questions=st.session_state.results.get('paa_questions', []),
                        term_data=st.session_state.results.get('term_data', {}),
                        score_data=score_for_brief,
                        internal_links=internal_links_for_brief,
                        guidance_only=is_guidance
                    )

                    if doc_success and doc_stream.getbuffer().nbytes > 0:
                        st.session_state.results['doc_stream'] = doc_stream
                        st.success(f"SEO Brief generated in {format_time(time.time() - start_time)}")
                        # Display download button immediately
                        st.download_button(
                            label="Download SEO Brief Now (.docx)",
                            data=doc_stream, # Use directly from variable
                            file_name=f"seo_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_brief_immediate"
                         )
                    else:
                        st.error("Failed to generate SEO brief document.")


            # Display Summary and persistent Download Button
            st.subheader("Generated Components Summary")
            components = [
                ("Keyword", 'keyword'), ("SERP Analysis", 'organic_results'),
                ("People Also Asked", 'paa_questions'), ("Related Keywords", 'related_keywords'),
                ("Competitor Content Scraped", 'scraped_contents'),
                ("Term Analysis", 'term_data'), ("Semantic Structure", 'semantic_structure'),
                ("Meta Tags", 'meta_title'),
                ("Generated Content", True if has_content else False), # Special check for content flag
                ("Content Score (if applicable)", 'content_score'),
                ("Internal Linking (if applicable)", 'internal_links')
            ]
            for name, key_or_flag in components:
                 status = False
                 if isinstance(key_or_flag, bool):
                     status = key_or_flag
                 else:
                     status = True if st.session_state.results.get(key_or_flag) else False
                 st.write(f"**{name}:** {'✅ Done' if status else '❌ Pending/Failed'}")


            doc_stream_saved = st.session_state.results.get('doc_stream')
            if doc_stream_saved and doc_stream_saved.getbuffer().nbytes > 0:
                 st.download_button(
                     label="Download Generated SEO Brief (.docx)",
                     data=doc_stream_saved,
                     file_name=f"seo_brief_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     key="download_brief_persistent"
                 )

    # Tab 6: Content Updates
    with tabs[5]:
        st.header("Content Update Recommendations")

        # Check prerequisites (Need SERP, Competitor Content, Structure, Terms)
        analysis_complete = all(st.session_state.results.get(k) for k in ['keyword', 'scraped_contents', 'semantic_structure', 'term_data'])

        if not analysis_complete:
             st.warning("Please complete SERP Analysis and Content Analysis first.")
        else:
             st.markdown("Upload your existing content document (`.docx`) to get recommendations or generate an optimized version based on the analysis.")
             content_file = st.file_uploader("Upload Existing Content Document", type=['docx'], key="update_upload")

             update_type = st.radio(
                 "Select Update Approach:",
                 ["Recommendations Only", "Generate Optimized Revision"],
                 key="update_type",
                 index=0 # Default to recommendations
             )

             if st.button("Analyze & Generate Updates", key="update_btn"):
                 if not anthropic_api_key: st.error("Anthropic API key required.")
                 elif not content_file: st.error("Please upload your existing content document.")
                 else:
                     with st.spinner(f"Processing document and generating {update_type}..."):
                          start_time = time.time()
                          # Clear previous update results
                          st.session_state.results['existing_content'] = None
                          st.session_state.results['content_gaps'] = None
                          st.session_state.results['updated_doc'] = None
                          st.session_state.results['optimized_content'] = None
                          st.session_state.results['change_summary'] = None
                          st.session_state.results['existing_content_score'] = None
                          st.session_state.results['optimized_content_score'] = None

                          # Parse existing Doc
                          existing_content, parse_success = parse_word_document(content_file)
                          if not parse_success:
                               st.error("Failed to parse the uploaded Word document.")
                          else:
                               st.session_state.results['existing_content'] = existing_content

                               # Score existing content (if possible)
                               term_data = st.session_state.results['term_data']
                               keyword = st.session_state.results['keyword']
                               st.text("Scoring existing content...")
                               score_data, score_success = score_content(
                                   existing_content.get('full_text',''), term_data, keyword
                               )
                               if score_success:
                                   st.session_state.results['existing_content_score'] = score_data
                                   logger.info(f"Existing content score: {score_data.get('overall_score')}")
                               else:
                                    st.warning("Could not score existing content.")

                               # Analyze Gaps
                               st.text("Analyzing content gaps...")
                               content_gaps, gap_success = analyze_content_gaps(
                                   existing_content,
                                   st.session_state.results['scraped_contents'],
                                   st.session_state.results['semantic_structure'],
                                   term_data,
                                   score_data or {}, # Pass score if available
                                   anthropic_api_key,
                                   keyword,
                                   st.session_state.results.get('paa_questions', [])
                               )
                               if not gap_success:
                                   st.error(f"Failed to analyze content gaps. Error: {content_gaps.get('error', 'Unknown')}")
                               else:
                                   st.session_state.results['content_gaps'] = content_gaps
                                   logger.info("Content gap analysis successful.")

                                   # --- Generate Output Based on Selection ---
                                   if update_type == "Recommendations Only":
                                       st.text("Generating recommendations document...")
                                       updated_doc, doc_success = create_updated_document(
                                           existing_content, content_gaps, keyword, score_data
                                       )
                                       if doc_success and updated_doc.getbuffer().nbytes > 0:
                                           st.session_state.results['updated_doc'] = updated_doc
                                           st.success(f"Recommendations generated in {format_time(time.time() - start_time)}")
                                       else:
                                           st.error("Failed to create recommendations document.")

                                   else: # Generate Optimized Revision
                                       st.text("Generating optimized revision...")
                                       optimized_content, change_summary, gen_success = generate_optimized_article_with_tracking(
                                           existing_content=existing_content, # Pass original for context/comparison
                                           competitor_contents=st.session_state.results['scraped_contents'],
                                           semantic_structure=st.session_state.results['semantic_structure'],
                                           related_keywords=st.session_state.results.get('related_keywords', []),
                                           keyword=keyword,
                                           paa_questions=st.session_state.results.get('paa_questions', []),
                                           term_data=term_data,
                                           anthropic_api_key=anthropic_api_key
                                       )
                                       if gen_success and optimized_content:
                                            st.session_state.results['optimized_content'] = optimized_content
                                            st.session_state.results['change_summary'] = change_summary

                                            # Score the new optimized content
                                            st.text("Scoring optimized content...")
                                            opt_score_data, opt_score_success = score_content(optimized_content, term_data, keyword)
                                            if opt_score_success:
                                                 st.session_state.results['optimized_content_score'] = opt_score_data
                                            else:
                                                 st.warning("Could not score the optimized content.")

                                            st.success(f"Optimized revision generated in {format_time(time.time() - start_time)}")
                                       else:
                                           st.error("Failed to generate optimized revision.")
                                           st.code(change_summary) # Show error summary if provided

                                   st.experimental_rerun() # Refresh display

             # Display results for Content Updates
             content_gaps = st.session_state.results.get('content_gaps')
             updated_doc = st.session_state.results.get('updated_doc')
             optimized_content = st.session_state.results.get('optimized_content')
             change_summary = st.session_state.results.get('change_summary')
             existing_score = st.session_state.results.get('existing_content_score')
             optimized_score = st.session_state.results.get('optimized_content_score')

             # Display Scores Side-by-Side if available
             if existing_score and optimized_score:
                  st.subheader("Score Comparison")
                  col1, col2 = st.columns(2)
                  with col1:
                       score = existing_score.get('overall_score', 0); grade = existing_score.get('grade', 'F')
                       s_color = "green" if score >= 70 else "red" if score < 50 else "orange"
                       st.markdown(f"**Original Score:** <span style='color:{s_color}; font-size: 1.2em;'>{score} ({grade})</span>", unsafe_allow_html=True)
                  with col2:
                       score = optimized_score.get('overall_score', 0); grade = optimized_score.get('grade', 'F')
                       s_color = "green" if score >= 70 else "red" if score < 50 else "orange"
                       st.markdown(f"**Optimized Score:** <span style='color:{s_color}; font-size: 1.2em;'>{score} ({grade})</span>", unsafe_allow_html=True)
                  improvement = optimized_score.get('overall_score', 0) - existing_score.get('overall_score', 0)
                  imp_color = 'green' if improvement > 0 else 'red' if improvement < 0 else 'grey'
                  st.markdown(f"**Improvement:** <span style='color:{imp_color}; font-weight: bold;'>{'+' if improvement > 0 else ''}{improvement} points</span>", unsafe_allow_html=True)
                  st.divider()


             # Display Recommendations or Optimized Content
             if updated_doc: # Display Recommendations Summary & Download
                  st.subheader("Content Update Recommendations")
                  if content_gaps: # Show a summary if analysis was successful
                       # Simplified summary display (Consider adding more detail back if needed)
                       if content_gaps.get("semantic_focus_recommendations"): st.write("- Semantic Focus issues found.")
                       if content_gaps.get("revised_headings"): st.write("- Heading revisions suggested.")
                       if content_gaps.get("missing_sections"): st.write("- New sections recommended.")
                       if content_gaps.get("content_gaps_to_fill"): st.write("- Content gaps identified.")
                       if content_gaps.get("term_usage_recommendations"): st.write("- Term usage improvements suggested.")
                       if content_gaps.get("questions_to_answer"): st.write("- Questions to answer identified.")
                  else:
                       st.warning("Could not display recommendation summary (analysis data missing).")

                  st.download_button(
                      label="Download Recommendations Document (.docx)",
                      data=updated_doc,
                      file_name=f"content_updates_{st.session_state.results.get('keyword', 'file').replace(' ', '_')}.docx",
                      mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                      key="download_update_recs"
                  )

             elif optimized_content: # Display Optimized Content & Download
                  st.subheader("Generated Optimized Revision")

                  opt_tabs = st.tabs(["Optimized Article", "Optimization Summary"])
                  with opt_tabs[0]:
                       st.markdown(optimized_content, unsafe_allow_html=True)
                  with opt_tabs[1]:
                       st.markdown(change_summary or "<p>No summary generated<p>", unsafe_allow_html=True)

                  # Generate Word doc for download
                  try:
                       word_doc_stream = create_word_document_from_html(
                           optimized_content,
                           st.session_state.results.get('keyword', 'optimized_doc'),
                           change_summary or "",
                           optimized_score # Pass score data
                       )
                       if word_doc_stream.getbuffer().nbytes > 0:
                            st.download_button(
                                label="Download Optimized Article (.docx)",
                                data=word_doc_stream,
                                file_name=f"optimized_{st.session_state.results.get('keyword', 'file').replace(' ', '_')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="download_opt_article"
                            )
                       else: st.error("Failed to generate Word document for download.")
                  except Exception as doc_gen_e:
                        st.error(f"Error generating download file: {doc_gen_e}")

    # Tab 7: Content Scoring
    with tabs[6]:
        st.header("Content Scoring & Optimization")
        
        if 'scraped_contents' not in st.session_state.results or 'keyword' not in st.session_state.results:
            st.warning("Please fetch SERP data and analyze content first (in the 'Input & SERP Analysis' and 'Content Analysis' tabs)")
        else:
            # Check if we've already extracted terms
            if 'term_data' not in st.session_state.results:
                if st.button("Extract Important Terms"):
                    if not anthropic_api_key:
                        st.error("Please enter Anthropic API key")
                    else:
                        with st.spinner("Extracting important terms and topics from top-ranking content..."):
                            start_time = time.time()
                            
                            term_data, success = extract_important_terms(
                                st.session_state.results['scraped_contents'], 
                                anthropic_api_key
                            )
                            
                            if success and term_data:
                                st.session_state.results['term_data'] = term_data
                                
                                # Show extracted terms
                                st.subheader("Top Primary Terms")
                                primary_df = pd.DataFrame(term_data.get('primary_terms', []))
                                if not primary_df.empty:
                                    st.dataframe(primary_df)
                                
                                st.subheader("Top Secondary Terms")
                                secondary_df = pd.DataFrame(term_data.get('secondary_terms', []))
                                if not secondary_df.empty:
                                    st.dataframe(secondary_df)
                                
                                st.success(f"Term extraction completed in {format_time(time.time() - start_time)}")
                            else:
                                st.error("Failed to extract terms")
            
            else:
                # Display previously extracted terms in a collapsible section
                with st.expander("View Extracted Terms & Topics", expanded=False):
                    st.subheader("Primary Terms")
                    primary_df = pd.DataFrame(st.session_state.results['term_data'].get('primary_terms', []))
                    if not primary_df.empty:
                        st.dataframe(primary_df)
                    
                    st.subheader("Secondary Terms")
                    secondary_df = pd.DataFrame(st.session_state.results['term_data'].get('secondary_terms', []))
                    if not secondary_df.empty:
                        st.dataframe(secondary_df)
                    
                    # Display topics if available
                    if 'topics' in st.session_state.results['term_data']:
                        st.subheader("Topics to Cover")
                        topics_df = pd.DataFrame(st.session_state.results['term_data'].get('topics', []))
                        if not topics_df.empty:
                            st.dataframe(topics_df)
                    
                    # Display questions if available
                    if 'questions' in st.session_state.results['term_data']:
                        st.subheader("Questions to Answer")
                        questions = st.session_state.results['term_data'].get('questions', [])
                        for q in questions:
                            st.write(f"- {q}")
                
                # Content input/editing area
                st.subheader("Enter or Paste Your Content")
                
                # Initialize content in session state if needed
                if 'current_content' not in st.session_state:
                    # Use previously generated content if available
                    if 'article_content' in st.session_state.results:
                        # Strip HTML tags for the textarea
                        soup = BeautifulSoup(st.session_state.results['article_content'], 'html.parser')
                        plain_text = soup.get_text()
                        st.session_state.current_content = plain_text
                    else:
                        st.session_state.current_content = ""
                
                # Function to update content in session state
                def update_content():
                    st.session_state.current_content = st.session_state.content_input
                    # Clear previous scoring results when content changes
                    if 'content_score' in st.session_state.results:
                        del st.session_state.results['content_score']
                    if 'highlighted_content' in st.session_state.results:
                        del st.session_state.results['highlighted_content']
                    if 'content_suggestions' in st.session_state.results:
                        del st.session_state.results['content_suggestions']
                
                content_input = st.text_area(
                    "Your content",
                    value=st.session_state.current_content,
                    height=400,
                    key="content_input",
                    on_change=update_content
                )
                
                if st.button("Score Content"):
                    if not st.session_state.current_content:
                        st.error("Please enter content to score")
                    else:
                        with st.spinner("Scoring content..."):
                            start_time = time.time()
                            
                            # Convert plain text to HTML for proper analysis
                            content_html = f"<p>{st.session_state.current_content.replace('</p><p>', '</p>\n<p>').replace('\n\n', '</p>\n<p>').replace('\n', '<br>')}</p>"
                            
                            # Score the content
                            score_data, score_success = score_content(
                                content_html, 
                                st.session_state.results['term_data'],
                                st.session_state.results['keyword']
                            )
                            
                            if score_success:
                                st.session_state.results['content_score'] = score_data
                                
                                # Highlight keywords in content
                                highlighted_content, highlight_success = highlight_keywords_in_content(
                                    content_html,
                                    st.session_state.results['term_data'],
                                    st.session_state.results['keyword']
                                )
                                
                                if highlight_success:
                                    st.session_state.results['highlighted_content'] = highlighted_content
                                
                                # Get content improvement suggestions
                                suggestions, suggestions_success = get_content_improvement_suggestions(
                                    content_html,
                                    st.session_state.results['term_data'],
                                    score_data,
                                    st.session_state.results['keyword']
                                )
                                
                                if suggestions_success:
                                    st.session_state.results['content_suggestions'] = suggestions
                                    st.success(f"Content scoring completed in {format_time(time.time() - start_time)}")
                                else:
                                    st.error("Failed to score content")
                
                # Display content score if available
                if 'content_score' in st.session_state.results:
                    score_data = st.session_state.results['content_score']
                    
                    # Create score display with CSS styling
                    col1, col2, col3 = st.columns([1, 1, 2])
                    
                    with col1:
                        overall_score = score_data.get('overall_score', 0)
                        st.markdown(f"""
                        <div style="text-align: center; padding: 20px; background-color: #f0f0f0; border-radius: 10px;">
                            <h2 style="margin:0; font-size: 18px;">Overall Score</h2>
                            <h1 style="margin:0; font-size: 48px; color: {'#28a745' if overall_score >= 70 else '#dc3545' if overall_score < 50 else '#ffc107'};">
                                {overall_score}
                            </h1>
                            <p style="margin:0; font-size: 24px; font-weight: bold;">{score_data.get('grade', 'F')}</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        component_scores = score_data.get('components', {})
                        
                        st.markdown(f"""
                        <div style="padding: 10px; background-color: #f0f0f0; border-radius: 10px;">
                            <h3 style="margin: 0 0 10px 0; font-size: 16px;">Component Scores</h3>
                            <p style="margin: 0; font-size: 14px;">Primary Keyword: <strong>{component_scores.get('keyword_score', 0)}</strong></p>
                            <p style="margin: 0; font-size: 14px;">Primary Terms: <strong>{component_scores.get('primary_terms_score', 0)}</strong></p>
                            <p style="margin: 0; font-size: 14px;">Secondary Terms: <strong>{component_scores.get('secondary_terms_score', 0)}</strong></p>
                            <p style="margin: 0; font-size: 14px;">Topic Coverage: <strong>{component_scores.get('topic_coverage_score', 0)}</strong></p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Create score visualization using Altair or Matplotlib
                    with col3:
                        # Create score components data
                        component_data = pd.DataFrame({
                            'Component': [
                                'Keyword Usage',
                                'Primary Terms',
                                'Secondary Terms',
                                'Topic Coverage',
                                'Questions'
                            ],
                            'Score': [
                                component_scores.get('keyword_score', 0),
                                component_scores.get('primary_terms_score', 0),
                                component_scores.get('secondary_terms_score', 0),
                                component_scores.get('topic_coverage_score', 0),
                                component_scores.get('question_coverage_score', 0)
                            ]
                        })
                        
                        # Add a color category column to the dataframe
                        component_data['Color Category'] = pd.cut(
                            component_data['Score'],
                            bins=[0, 49.99, 69.99, 100],
                            labels=['Poor', 'Medium', 'Good']
                        )
                        
                        # Create a bar chart with color based on the category
                        chart = alt.Chart(component_data).mark_bar().encode(
                            x='Score',
                            y=alt.Y('Component', sort=None),
                            color=alt.Color(
                                'Color Category:N',
                                scale=alt.Scale(
                                    domain=['Poor', 'Medium', 'Good'],
                                    range=['#dc3545', '#ffc107', '#28a745']
                                )
                            ),
                            tooltip=['Component', 'Score']
                        ).properties(
                            title='Score Components',
                            width=300,
                            height=200
                        )
                        
                        # Display the chart
                        st.altair_chart(chart, use_container_width=True)
                    
                    # Display content details
                    details = score_data.get('details', {})
                    st.markdown(f"""
                    <div style="margin: 20px 0; padding: 10px; background-color: #f8f9fa; border-radius: 5px;">
                        <h3 style="margin-top: 0;">Content Details</h3>
                        <p>Word Count: <strong>{details.get('word_count', 0)}</strong></p>
                        <p>Primary Keyword Count: <strong>{details.get('keyword_count', 0)}/{details.get('optimal_keyword_count', 0)} (optimal)</strong></p>
                        <p>Primary Terms Found: <strong>{details.get('primary_terms_found', 0)}/{details.get('primary_terms_total', 0)}</strong></p>
                        <p>Secondary Terms Found: <strong>{details.get('secondary_terms_found', 0)}/{details.get('secondary_terms_total', 0)}</strong></p>
                        <p>Topics Covered: <strong>{details.get('topics_covered', 0)}/{details.get('topics_total', 0)}</strong></p>
                        <p>Questions Answered: <strong>{details.get('questions_answered', 0)}/{details.get('questions_total', 0)}</strong></p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Display content with highlighted keywords
                    if 'highlighted_content' in st.session_state.results:
                        st.subheader("Content with Highlighted Keywords")
                        st.markdown("""
                        <div style="margin-bottom: 10px; font-size: 12px;">
                            <span style="background-color: #FFEB9C; padding: 2px 5px;">Primary Keyword</span>
                            <span style="background-color: #CDFFD8; padding: 2px 5px; margin-left: 10px;">Primary Terms</span>
                            <span style="background-color: #E6F3FF; padding: 2px 5px; margin-left: 10px;">Secondary Terms</span>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.markdown(f"""
                        <div style="padding: 15px; border: 1px solid #ddd; border-radius: 5px; max-height: 400px; overflow-y: auto;">
                            {st.session_state.results['highlighted_content']}
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Display content improvement suggestions
                    if 'content_suggestions' in st.session_state.results:
                        suggestions = st.session_state.results['content_suggestions']
                        
                        st.subheader("Content Improvement Suggestions")
                        
                        # Tabs for different suggestion types
                        suggestion_tabs = st.tabs([
                            "Missing Terms", 
                            "Content Gaps", 
                            "Questions to Answer", 
                            "Readability"
                        ])
                        
                        # Missing Terms tab
                        with suggestion_tabs[0]:
                            st.markdown("### Missing and Underused Terms")
                            
                            # Missing primary terms
                            missing_primary = [s for s in suggestions.get('missing_terms', []) if s.get('type') == 'primary']
                            if missing_primary:
                                st.markdown("#### Missing Primary Terms")
                                for term in missing_primary:
                                    st.markdown(f"""
                                    <div style="margin-bottom: 5px; padding: 5px 10px; background-color: #ffeeee; border-left: 3px solid #ff6666; border-radius: 3px;">
                                        <strong>{term.get('term')}</strong> - Importance: {term.get('importance', 0):.2f} - Recommended usage: {term.get('recommended_usage', 1)}
                                    </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.success("No important primary terms are missing!")
                            
                            # Underused terms
                            underused_terms = suggestions.get('underused_terms', [])
                            if underused_terms:
                                st.markdown("#### Underused Terms")
                                for term in underused_terms:
                                    st.markdown(f"""
                                    <div style="margin-bottom: 5px; padding: 5px 10px; background-color: #fff8e1; border-left: 3px solid #ffc107; border-radius: 3px;">
                                        <strong>{term.get('term')}</strong> - Current usage: {term.get('current_usage')}/{term.get('recommended_usage')} recommended
                                    </div>
                                    """, unsafe_allow_html=True)
                            
                            # Missing secondary terms
                            missing_secondary = [s for s in suggestions.get('missing_terms', []) if s.get('type') == 'secondary']
                            if missing_secondary:
                                st.markdown("#### Missing Secondary Terms")
                                for term in missing_secondary:
                                    st.markdown(f"""
                                    <div style="margin-bottom: 5px; padding: 5px 10px; background-color: #f0f0f0; border-left: 3px solid #808080; border-radius: 3px;">
                                        <strong>{term.get('term')}</strong> - Importance: {term.get('importance', 0):.2f}
                                    </div>
                                    """, unsafe_allow_html=True)
                        
                        # Content Gaps tab
                        with suggestion_tabs[1]:
                            st.markdown("### Content Topic Gaps")
                            
                            # Missing topics (completely missing)
                            missing_topics = suggestions.get('missing_topics', [])
                            if missing_topics:
                                st.markdown("#### Topics to Add")
                                for topic in missing_topics:
                                    st.markdown(f"""
                                    <div style="margin-bottom: 10px; padding: 10px; background-color: #e3f2fd; border-left: 3px solid #2196f3; border-radius: 3px;">
                                        <strong>Missing Topic: {topic.get('topic')}</strong>
                                        <p style="margin: 5px 0 0 0;">{topic.get('description', '')}</p>
                                    </div>
                                    """, unsafe_allow_html=True)
                            
                            # Partially covered topics (need enhancement)
                            partial_topics = suggestions.get('partial_topics', [])
                            if partial_topics:
                                st.markdown("#### Topics to Expand")
                                for topic in partial_topics:
                                    match_ratio = topic.get('match_ratio', 0)
                                    st.markdown(f"""
                                    <div style="margin-bottom: 10px; padding: 10px; background-color: #fff8e1; border-left: 3px solid #ffc107; border-radius: 3px;">
                                        <strong>Enhance Coverage: {topic.get('topic')}</strong> ({int(match_ratio * 100)}% covered)
                                        <p style="margin: 5px 0 0 0;">{topic.get('description', '')}</p>
                                        <p style="margin: 5px 0 0 0; font-style: italic;">{topic.get('suggestion', '')}</p>
                                    </div>
                                    """, unsafe_allow_html=True)
                            
                            if not missing_topics and not partial_topics:
                                st.success("Your content covers all the important topics comprehensively!")
                        
                        # Questions tab
                        with suggestion_tabs[2]:
                            st.markdown("### Questions to Answer")
                            
                            # Unanswered questions
                            unanswered = suggestions.get('unanswered_questions', [])
                            if unanswered:
                                for i, question in enumerate(unanswered, 1):
                                    st.markdown(f"""
                                    <div style="margin-bottom: 10px; padding: 10px; background-color: #e8f5e9; border-left: 3px solid #4caf50; border-radius: 3px;">
                                        <strong>{i}. {question}</strong>
                                    </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.success("Your content answers all the important questions!")
                        
                        # Readability tab
                        with suggestion_tabs[3]:
                            st.markdown("### Readability & Structure Suggestions")
                            
                            # Readability suggestions
                            readability = suggestions.get('readability_suggestions', [])
                            if readability:
                                for suggestion in readability:
                                    st.markdown(f"""
                                    <div style="margin-bottom: 5px; padding: 5px 10px; background-color: #f3e5f5; border-left: 3px solid #9c27b0; border-radius: 3px;">
                                        {suggestion}
                                    </div>
                                    """, unsafe_allow_html=True)
                            
                            # Structure suggestions
                            structure = suggestions.get('structure_suggestions', [])
                            if structure:
                                for suggestion in structure:
                                    st.markdown(f"""
                                    <div style="margin-bottom: 5px; padding: 5px 10px; background-color: #fce4ec; border-left: 3px solid #e91e63; border-radius: 3px;">
                                        {suggestion}
                                    </div>
                                    """, unsafe_allow_html=True)
                            
                            if not readability and not structure:
                                st.success("Your content has good readability and structure!")
                    
                    # Add download buttons for reports
                    st.subheader("Download Reports")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Create and offer content brief for download
                        brief_doc = create_content_scoring_brief(
                            st.session_state.results['keyword'],
                            st.session_state.results['term_data'],
                            score_data,
                            st.session_state.results.get('content_suggestions', {})
                        )
                        
                        st.download_button(
                            label="Download Content Optimization Brief",
                            data=brief_doc,
                            file_name=f"content_optimization_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    with col2:
                        # Create and offer highlighted content document for download
                        if 'highlighted_content' in st.session_state.results:
                            highlighted_doc = create_word_document_from_html(
                                st.session_state.results['highlighted_content'],
                                st.session_state.results['keyword'] + " - Highlighted Terms",
                                ""
                            )
                            
                            st.download_button(
                                label="Download Highlighted Content",
                                data=highlighted_doc,
                                file_name=f"highlighted_{st.session_state.results['keyword'].replace(' ', '_')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

if __name__ == "__main__":
    main()
